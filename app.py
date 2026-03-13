"""
IBP Forecast Compare – Streamlit Web App
Run locally:  streamlit run app.py
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import io
import traceback
from datetime import datetime

# ── Try pyxlsb ──────────────────────────────────────────────────────────────
try:
    import pyxlsb
    XLSB_OK = True
except ImportError:
    XLSB_OK = False

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IBP Forecast Compare Tool",
    page_icon="📊",
    layout="wide",
)

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1F4E79 0%, #2E75B6 100%);
        padding: 20px 30px; border-radius: 10px; margin-bottom: 24px;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 26px; }
    .main-header p  { color: #BDD7EE; margin: 4px 0 0 0; font-size: 14px; }
    .section-box {
        background: #F0F4F8; border-left: 4px solid #2E75B6;
        padding: 12px 16px; border-radius: 6px; margin-bottom: 16px;
    }
    .success-box {
        background: #E2EFDA; border-left: 4px solid #375623;
        padding: 12px 16px; border-radius: 6px;
    }
    .warning-box {
        background: #FFF2CC; border-left: 4px solid #C55A11;
        padding: 12px 16px; border-radius: 6px;
    }
    div[data-testid="stButton"] button {
        background: #217346; color: white; font-weight: bold;
        font-size: 16px; padding: 10px 32px; border-radius: 8px;
        border: none; width: 100%;
    }
    div[data-testid="stButton"] button:hover { background: #1a5c38; }
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
#  COLUMN CONFIGURATION  (update if Ahmed changes his file layout)
# ════════════════════════════════════════════════════════════════════════════
class AhmedFileCols:
    FCST_PRIMARY    = 5    # Col F → PRODUCT_MODEL_NR
    FCST_THEATER    = 6    # Col G → Reg / Theater
    FCST_HDR_ROW    = 0
    FCST_DATA_START = 1
    FCST_MO_START   = 9    # Col J → first forecast month
    FCST_MO_END     = 26   # Col AA → last forecast month (18 months)
    EX_PRIMARY      = 3    # Col D
    EX_THEATER      = 4    # Col E
    BRZ_BASE_PL     = 1
    BRZ_PLATFORM    = 2
    BRZ_PRIMARY     = 3
    BRZ_THEATER     = 4
    BRZ_KEYFIG      = 5

class MasterCols:
    KEY        = 0
    PLATFORM   = 2
    FACTOR     = 5
    PLC        = 6
    CANON      = 12
    PL         = 17
    DATA_START = 2

# ════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════════════════
def excel_col_to_index(col_letter: str) -> int:
    col_letter = col_letter.strip().upper()
    result = 0
    for ch in col_letter:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1

def read_xlsb(file_bytes: bytes, sheetname: str) -> pd.DataFrame:
    if not XLSB_OK:
        raise ImportError("pyxlsb not installed. See requirements.txt.")
    import pyxlsb, tempfile, os
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsb") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        rows = []
        with pyxlsb.open_workbook(tmp_path) as wb:
            with wb.get_sheet(sheetname) as ws:
                for row in ws.rows():
                    rows.append([c.v for c in row])
        return pd.DataFrame(rows) if rows else pd.DataFrame()
    finally:
        os.unlink(tmp_path)

def load_master(file_bytes: bytes, sheet: str = "Table-Nov25") -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None)
    df = df.iloc[MasterCols.DATA_START:].copy()
    df.columns = range(len(df.columns))
    df = df[df[MasterCols.KEY].notna()].copy()
    df[MasterCols.KEY] = df[MasterCols.KEY].astype(str).str.strip()
    df = df.drop_duplicates(subset=MasterCols.KEY)
    return df.set_index(MasterCols.KEY)

def vlook(key, master, col, default=""):
    try:
        v = master.loc[str(key).strip(), col]
        return v if pd.notna(v) else default
    except Exception:
        return default

# ════════════════════════════════════════════════════════════════════════════
#  CORE PROCESSING
# ════════════════════════════════════════════════════════════════════════════
def build_m0_fcst(ahmed_bytes, master):
    raw = read_xlsb(ahmed_bytes, "Final Supplies Fcst SKU to Base")
    header_row   = raw.iloc[AhmedFileCols.FCST_HDR_ROW].tolist()
    month_labels = header_row[AhmedFileCols.FCST_MO_START : AhmedFileCols.FCST_MO_END + 1]
    df = raw.iloc[AhmedFileCols.FCST_DATA_START:].copy().reset_index(drop=True)

    out = pd.DataFrame()
    out["PRODUCT_MODEL_NR"] = df.iloc[:, AhmedFileCols.FCST_PRIMARY].astype(str).str.strip()
    out["Reg"]              = df.iloc[:, AhmedFileCols.FCST_THEATER].astype(str).str.strip()

    valid = out["PRODUCT_MODEL_NR"].isin(master.index) & (out["PRODUCT_MODEL_NR"] != "nan")
    out = out[valid].copy().reset_index(drop=True)
    df  = df[valid].copy().reset_index(drop=True)

    out["PRODUCT_PLATFORM_NM"] = out["PRODUCT_MODEL_NR"].map(lambda x: vlook(x, master, MasterCols.PLATFORM))
    out["Canon-Shared Cap"]    = out["PRODUCT_MODEL_NR"].map(lambda x: vlook(x, master, MasterCols.CANON))
    out["PL"]                  = out["PRODUCT_MODEL_NR"].map(lambda x: vlook(x, master, MasterCols.PL))
    out["PLC"]                 = out["PRODUCT_MODEL_NR"].map(lambda x: vlook(x, master, MasterCols.PLC))
    out["ConC"]                = out["Reg"] + out["PRODUCT_MODEL_NR"]
    out["Values"]              = "Sum of M-0 FRCST_QT"
    out["Factor"]              = out["PRODUCT_MODEL_NR"].map(
        lambda x: pd.to_numeric(vlook(x, master, MasterCols.FACTOR, 1), errors='coerce') or 1)
    out["M1_Actuals"] = 0.0

    for i, col_idx in enumerate(range(AhmedFileCols.FCST_MO_START, AhmedFileCols.FCST_MO_END + 1)):
        lbl = month_labels[i]
        out[lbl] = pd.to_numeric(df.iloc[:, col_idx], errors='coerce').fillna(0)

    return out, month_labels

def build_trade_actuals(ahmed_bytes, master, actuals_col_idx):
    raw_ex = read_xlsb(ahmed_bytes, "L10 SHIP BASE ex-Brazil")
    raw_ex = raw_ex.iloc[1:].reset_index(drop=True)
    a = pd.DataFrame()
    a["WORLD_REGION_CD"] = raw_ex.iloc[:, AhmedFileCols.EX_THEATER].astype(str).str.strip()
    a["BASE_PROD_NR"]    = raw_ex.iloc[:, AhmedFileCols.EX_PRIMARY].astype(str).str.strip()
    a["Sum_SHIP_QT"]     = pd.to_numeric(raw_ex.iloc[:, actuals_col_idx], errors='coerce').fillna(0)
    a["Conc"]            = a["WORLD_REGION_CD"] + a["BASE_PROD_NR"]
    a = a[a["Sum_SHIP_QT"] != 0].copy().reset_index(drop=True)
    return a

def load_m1_fcst(prev_bytes):
    df = pd.read_excel(io.BytesIO(prev_bytes), sheet_name="PACKS-M0 Fcst",
                       header=None, engine="openpyxl")
    headers = df.iloc[2].tolist()
    data    = df.iloc[3:].copy()
    data.columns = headers
    data = data.reset_index(drop=True)
    if "Values" in data.columns:
        data = data.rename(columns={"Values": "KF"})
    if "KF" in data.columns:
        data["KF"] = "Sum of M-1 FRCST_QT"
    if "ConC" not in data.columns and "Reg" in data.columns:
        data["ConC"] = data["Reg"].astype(str) + data["PRODUCT_MODEL_NR"].astype(str)
    return data

def compute_changes(m0, m1, month_labels):
    id_cols = ["Reg","PRODUCT_MODEL_NR","PRODUCT_PLATFORM_NM","Canon-Shared Cap","PL","PLC","ConC"]
    changes = m0[id_cols].copy()
    changes["Values"] = "Sum of DELTA_M0-M1_QT"
    m1_lkp = m1.set_index("ConC") if (not m1.empty and "ConC" in m1.columns) else pd.DataFrame()
    for lbl in month_labels:
        m0_vals = m0.set_index("ConC")[lbl] if lbl in m0.columns else pd.Series(dtype=float)
        if not m1_lkp.empty and lbl in m1_lkp.columns:
            m1_vals = m1_lkp[lbl].apply(pd.to_numeric, errors='coerce').fillna(0)
            delta   = m0_vals.subtract(m1_vals, fill_value=0)
        else:
            delta = m0_vals
        changes[lbl] = changes["ConC"].map(delta).fillna(0)
    return changes

def build_compare_packs(m0, m1, changes, month_labels, lbl_m0, lbl_m1, m1_act_label):
    id_cols = ["Reg","PRODUCT_MODEL_NR","PRODUCT_PLATFORM_NM","Canon-Shared Cap","PL","PLC","ConC"]
    fcst_months = month_labels[1:]

    m0_rows = m0[id_cols].copy()
    m0_rows["Forecast Cycle"]      = lbl_m0
    m0_rows["M1_Actuals_Display"]  = m0["M1_Actuals"].values
    for lbl in fcst_months:
        m0_rows[lbl] = m0[lbl].values if lbl in m0.columns else 0

    m1_lkp = m1.set_index("ConC") if (not m1.empty and "ConC" in m1.columns) else pd.DataFrame()
    m1_rows = m0[id_cols].copy()
    m1_rows["Forecast Cycle"] = lbl_m1
    if not m1_lkp.empty and m1_act_label in m1_lkp.columns:
        m1_rows["M1_Actuals_Display"] = m1_rows["ConC"].map(
            m1_lkp[m1_act_label].apply(pd.to_numeric, errors='coerce').fillna(0)).fillna(0)
    else:
        m1_rows["M1_Actuals_Display"] = 0
    for lbl in fcst_months:
        if not m1_lkp.empty and lbl in m1_lkp.columns:
            m1_rows[lbl] = m1_rows["ConC"].map(
                m1_lkp[lbl].apply(pd.to_numeric, errors='coerce').fillna(0)).fillna(0)
        else:
            m1_rows[lbl] = 0

    delta_rows = changes[id_cols].copy()
    delta_rows["Forecast Cycle"]     = "DELTA"
    delta_rows["M1_Actuals_Display"] = (
        m0_rows.set_index("ConC")["M1_Actuals_Display"]
        .subtract(m1_rows.set_index("ConC")["M1_Actuals_Display"], fill_value=0)
    ).reindex(delta_rows["ConC"]).values
    for lbl in fcst_months:
        delta_rows[lbl] = changes[lbl].values if lbl in changes.columns else 0

    combined = []
    for conc in m0["ConC"].tolist():
        combined.append(m0_rows[m0_rows["ConC"] == conc])
        combined.append(m1_rows[m1_rows["ConC"] == conc])
        combined.append(delta_rows[delta_rows["ConC"] == conc])

    result = pd.concat(combined, ignore_index=True)
    result = result.rename(columns={"M1_Actuals_Display": m1_act_label})
    return result.drop(columns=["ConC"])

def build_platform_summary(changes, month_labels):
    fcst_months = month_labels[1:]
    avail = [l for l in fcst_months if l in changes.columns]
    pivot = changes.groupby("PRODUCT_PLATFORM_NM")[avail].sum().reset_index()
    return pivot.rename(columns={"PRODUCT_PLATFORM_NM": "Platform"})

# ════════════════════════════════════════════════════════════════════════════
#  EXCEL WRITER  (returns bytes for download)
# ════════════════════════════════════════════════════════════════════════════
CLR_DARK  = "1F4E79"
CLR_MID   = "2E75B6"
CLR_M0    = "D6E4F0"
CLR_M1    = "E2EFDA"
CLR_DELTA = "FFF2CC"
CLR_WHITE = "FFFFFF"

def _hdr(cell, bg=CLR_DARK, fg=CLR_WHITE, bold=True, sz=10):
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.font = Font(color=fg, bold=bold, size=sz, name="Calibri")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def _num(cell):
    cell.number_format = '#,##0'
    cell.alignment = Alignment(horizontal="right")

def write_excel(compare_df, platform_df, changes_df, month_labels) -> bytes:
    wb = openpyxl.Workbook()
    fcst_months = month_labels[1:]

    # ── Compare Packs ──────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Compare Packs"
    ws1.freeze_panes = "I2"
    row_fills = {
        "DELTA": PatternFill("solid", fgColor=CLR_DELTA),
        "M1":    PatternFill("solid", fgColor=CLR_M1),
        "M0":    PatternFill("solid", fgColor=CLR_M0),
    }
    if not compare_df.empty:
        hdrs = list(compare_df.columns)
        for c, h in enumerate(hdrs, 1):
            cell = ws1.cell(row=1, column=c, value=str(h) if h else "")
            _hdr(cell)
            ws1.column_dimensions[get_column_letter(c)].width = 18 if c <= 7 else 12
        for r, (_, row) in enumerate(compare_df.iterrows(), 2):
            cy = str(row.get("Forecast Cycle", ""))
            fill = row_fills["DELTA"] if "DELTA" in cy.upper() else (
                   row_fills["M1"] if any(x in cy.upper() for x in ["FEB","MAR","APR","MAY","JUN","M-1","M1"]) and "DELTA" not in cy.upper() else
                   row_fills["M0"])
            for c, val in enumerate(row, 1):
                cell = ws1.cell(row=r, column=c)
                cell.value = round(val, 6) if isinstance(val, float) else val
                cell.fill = fill
                if c > 8:
                    _num(cell)
                    if isinstance(val, (int, float)) and val < 0:
                        cell.font = Font(color="C00000", name="Calibri", size=10)
        ws1.auto_filter.ref = f"A1:{get_column_letter(len(hdrs))}1"

    # ── Changes by Platform ────────────────────────────────
    ws2 = wb.create_sheet("Changes by Platform")
    ws2.freeze_panes = "B4"
    ws2["A1"] = "PACKS – Forecast Delta by Platform"
    ws2["A1"].font = Font(bold=True, size=12, color=CLR_DARK, name="Calibri")
    if not platform_df.empty:
        hdr_row = 3
        for c, h in enumerate(platform_df.columns, 1):
            cell = ws2.cell(row=hdr_row, column=c, value=str(h) if h else "")
            _hdr(cell, bg="203864")
            ws2.column_dimensions[get_column_letter(c)].width = 28 if c == 1 else 13
        for r, (_, row) in enumerate(platform_df.iterrows(), hdr_row + 1):
            for c, val in enumerate(row, 1):
                cell = ws2.cell(row=r, column=c)
                cell.value = round(val, 2) if isinstance(val, float) else val
                if c > 1:
                    _num(cell)
                    if isinstance(val, (int, float)) and val < 0:
                        cell.font = Font(color="C00000", name="Calibri")

    # ── Pivot Data ─────────────────────────────────────────
    ws3 = wb.create_sheet("Pivot Data")
    ws3.freeze_panes = "A2"
    if not changes_df.empty:
        keep = [c for c in changes_df.columns if c != "Factor"]
        for c, h in enumerate(keep, 1):
            cell = ws3.cell(row=1, column=c, value=str(h) if h else "")
            _hdr(cell, bg=CLR_MID)
            ws3.column_dimensions[get_column_letter(c)].width = 20 if c <= 7 else 13
        for r, (_, row) in enumerate(changes_df[keep].iterrows(), 2):
            for c, val in enumerate(row, 1):
                cell = ws3.cell(row=r, column=c)
                cell.value = round(val, 6) if isinstance(val, float) else val

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ════════════════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ════════════════════════════════════════════════════════════════════════════

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>📊 IBP Forecast Compare Tool</h1>
  <p>上传文件 → 填写设置 → 点击生成 → 下载 Excel 输出</p>
</div>
""", unsafe_allow_html=True)

if not XLSB_OK:
    st.error("⚠️  缺少依赖库 pyxlsb。请确认 requirements.txt 已包含 pyxlsb，并重新部署。")
    st.stop()

# ── Layout: 2 columns ────────────────────────────────────────────────────────
left, right = st.columns([1.1, 1])

with left:
    st.markdown("### 📂 上传文件")

    ahmed_file = st.file_uploader(
        "① Ahmed File（当前 cycle）",
        type=["xlsb"],
        help="WW_Validation_xx-xx-xxxx.xlsb",
        key="ahmed"
    )
    prev_file = st.file_uploader(
        "② Previous Working File（上一个 cycle）",
        type=["xlsx", "xlsm"],
        help="上一个 cycle 的 working.xlsx，其 PACKS-M0 Fcst tab 将作为本次 M-1",
        key="prev"
    )
    master_file = st.file_uploader(
        "③ Master Vlookup（Table-Nov25）",
        type=["xlsx", "xlsm"],
        help="从 SharePoint 下载的最新 Master Vlookup 文件",
        key="master"
    )

with right:
    st.markdown("### ⚙️ Cycle 设置")

    c1, c2 = st.columns(2)
    with c1:
        cycle_m0 = st.text_input("M0 Cycle（当前）", value="202604", max_chars=6)
        label_m0 = st.text_input("M0 Label", value="Apr Forecast")
    with c2:
        cycle_m1 = st.text_input("M-1 Cycle（上一个）", value="202603", max_chars=6)
        label_m1 = st.text_input("M-1 Label", value="Mar Forecast")

    st.markdown("---")
    st.markdown("### 📅 Actuals Column  ⚠️ 每次必填")

    actuals_col = st.text_input(
        "Ahmed L10 表中 M-1 Actuals 所在列（Excel 字母）",
        value="M",
        max_chars=2,
        help="例：March cycle → M（26-Feb）｜April cycle → N（26-Mar）"
    ).strip().upper()

    # Quick reference table
    st.markdown("""
| Cycle | 月份 | 列字母 |
|-------|------|--------|
| 202603 March | 26-Feb | **M** |
| 202604 April | 26-Mar | **N** |
| 202605 May   | 26-Apr | **O** |
| 202606 June  | 26-May | **P** |
""")

# ── Run button ───────────────────────────────────────────────────────────────
st.markdown("---")
run = st.button("▶  生成 IBP Forecast Compare", use_container_width=True)

if run:
    # Validate inputs
    missing = []
    if not ahmed_file:  missing.append("Ahmed File (.xlsb)")
    if not prev_file:   missing.append("Previous Working File (.xlsx)")
    if not master_file: missing.append("Master Vlookup (.xlsx)")
    if not actuals_col: missing.append("Actuals Column 字母")

    if missing:
        st.warning("请先上传以下文件：\n" + "\n".join(f"• {m}" for m in missing))
        st.stop()

    try:
        actuals_col_idx = excel_col_to_index(actuals_col)
    except Exception:
        st.error(f"Actuals Column 格式不正确：'{actuals_col}'，请输入 Excel 列字母，例如 M")
        st.stop()

    progress = st.progress(0)
    status   = st.empty()

    try:
        # 1. Master
        status.info("📖  加载 Master Vlookup...")
        master = load_master(master_file.read())
        progress.progress(15)
        status.info(f"✅  Master Vlookup: {len(master):,} 个 SKU")

        # 2. M0
        status.info("📊  读取 Ahmed File → PACKS-M0 Fcst...")
        ahmed_bytes = ahmed_file.read()
        m0_df, month_labels = build_m0_fcst(ahmed_bytes, master)
        progress.progress(35)
        status.info(f"✅  M0 Fcst: {len(m0_df):,} 行 | {len(month_labels)} 个月")

        # 3. Actuals
        status.info("📦  读取 Trade Actuals (L10)...")
        actuals_df = build_trade_actuals(ahmed_bytes, master, actuals_col_idx)
        act_lkp    = actuals_df.set_index("Conc")["Sum_SHIP_QT"]
        m0_df["M1_Actuals"] = m0_df["ConC"].map(act_lkp).fillna(0)
        m1_act_label = month_labels[0]
        fcst_months  = month_labels[1:]
        progress.progress(55)

        # 4. M-1
        status.info("📁  加载 M-1 Fcst（上一个 cycle）...")
        m1_df = load_m1_fcst(prev_file.read())
        progress.progress(65)

        # 5. Changes
        status.info("🔢  计算 PACKS-Fcst Changes...")
        changes_df = compute_changes(m0_df, m1_df, fcst_months)
        progress.progress(75)

        # 6. Compare Packs
        status.info("📋  构建 Compare Packs...")
        compare_df = build_compare_packs(
            m0_df, m1_df, changes_df,
            month_labels, label_m0, label_m1, m1_act_label
        )
        progress.progress(85)

        # 7. Platform summary
        platform_df = build_platform_summary(changes_df, month_labels)

        # 8. Write Excel
        status.info("💾  生成 Excel 输出文件...")
        excel_bytes = write_excel(compare_df, platform_df, changes_df, month_labels)
        progress.progress(100)

        # ── Success + Download ──────────────────────────────
        status.empty()
        st.success(f"✅  完成！共生成 {len(compare_df):,} 行数据")

        filename = f"{cycle_m0}_vs_{cycle_m1}_IBP_Forecast_Compare.xlsx"

        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.download_button(
                label=f"⬇️  下载  {filename}",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # ── Preview ────────────────────────────────────────
        with st.expander("📊 预览：Changes by Platform（前 10 行）"):
            if not platform_df.empty:
                preview_cols = ["Platform"] + [c for c in platform_df.columns if c != "Platform"][:6]
                st.dataframe(platform_df[preview_cols].head(10), use_container_width=True)

        with st.expander("📋 预览：Compare Packs（前 15 行）"):
            if not compare_df.empty:
                preview_cols = list(compare_df.columns[:9])
                st.dataframe(compare_df[preview_cols].head(15), use_container_width=True)

    except Exception as e:
        progress.empty()
        st.error(f"❌  运行出错：{e}")
        with st.expander("查看详细错误信息（发给维护人员）"):
            st.code(traceback.format_exc())

# ── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("IBP Forecast Compare Tool  v1.0  |  数据仅在运行时处理，不做任何储存")
