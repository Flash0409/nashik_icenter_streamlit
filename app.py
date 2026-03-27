import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
import json
import os
import re
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────
# CONFIG & PAGE SETUP
# ─────────────────────────────────────────────────
st.set_page_config(
    page_title="Material Shortage Forecaster",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

SAVED_LOTS_FILE = "saved_lots.json"

# ─────────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

    :root {
        --bg-dark: #0a0e17;
        --bg-card: #111827;
        --bg-card-hover: #1a2235;
        --accent-blue: #3b82f6;
        --accent-cyan: #06b6d4;
        --accent-emerald: #10b981;
        --accent-amber: #f59e0b;
        --accent-red: #ef4444;
        --accent-purple: #8b5cf6;
        --text-primary: #f1f5f9;
        --text-secondary: #94a3b8;
        --border-color: #1e293b;
    }

    .stApp {
        font-family: 'DM Sans', sans-serif;
    }

    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #0f172a 100%);
        border: 1px solid #2d2b55;
        border-radius: 16px;
        padding: 2rem 2.5rem;
        margin-bottom: 2rem;
        position: relative;
        overflow: hidden;
    }
    .main-header::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; bottom: 0;
        background: radial-gradient(ellipse at 20% 50%, rgba(59,130,246,0.08) 0%, transparent 60%),
                    radial-gradient(ellipse at 80% 50%, rgba(139,92,246,0.08) 0%, transparent 60%);
        pointer-events: none;
    }
    .main-header h1 {
        margin: 0; font-size: 1.8rem; font-weight: 700;
        background: linear-gradient(90deg, #e2e8f0, #93c5fd);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .main-header p {
        margin: 0.3rem 0 0 0; color: #94a3b8; font-size: 0.95rem;
    }

    .metric-card {
        background: linear-gradient(145deg, #111827, #1a2235);
        border: 1px solid #1e293b;
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        text-align: center;
        transition: all 0.2s ease;
    }
    .metric-card:hover { border-color: #3b82f6; transform: translateY(-2px); }
    .metric-value {
        font-size: 2rem; font-weight: 700;
        font-family: 'JetBrains Mono', monospace;
    }
    .metric-label {
        font-size: 0.8rem; color: #94a3b8; text-transform: uppercase;
        letter-spacing: 0.05em; margin-top: 0.25rem;
    }
    .metric-blue .metric-value { color: #3b82f6; }
    .metric-emerald .metric-value { color: #10b981; }
    .metric-amber .metric-value { color: #f59e0b; }
    .metric-red .metric-value { color: #ef4444; }
    .metric-purple .metric-value { color: #8b5cf6; }
    .metric-cyan .metric-value { color: #06b6d4; }

    .section-header {
        font-size: 1.15rem; font-weight: 600; color: #e2e8f0;
        border-bottom: 2px solid #3b82f6; padding-bottom: 0.5rem;
        margin: 1.5rem 0 1rem 0;
    }

    .status-badge {
        display: inline-block; padding: 0.2rem 0.7rem;
        border-radius: 20px; font-size: 0.75rem; font-weight: 600;
    }
    .status-ok { background: rgba(16,185,129,0.15); color: #34d399; border: 1px solid rgba(16,185,129,0.3); }
    .status-warn { background: rgba(245,158,11,0.15); color: #fbbf24; border: 1px solid rgba(245,158,11,0.3); }
    .status-critical { background: rgba(239,68,68,0.15); color: #f87171; border: 1px solid rgba(239,68,68,0.3); }

    .lot-chip {
        display: inline-block;
        background: rgba(139,92,246,0.15);
        color: #c4b5fd;
        border: 1px solid rgba(139,92,246,0.3);
        border-radius: 6px;
        padding: 0.15rem 0.5rem;
        font-size: 0.72rem;
        font-weight: 600;
        margin: 1px 2px;
    }

    div[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #1e1b4b 100%);
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────
def load_saved_lots() -> dict:
    if os.path.exists(SAVED_LOTS_FILE):
        with open(SAVED_LOTS_FILE, "r") as f:
            return json.load(f)
    return {}


def save_lots(lots: dict):
    with open(SAVED_LOTS_FILE, "w") as f:
        json.dump(lots, f, indent=2, default=str)


def robust_to_datetime(series: pd.Series) -> pd.Series:
    if series.dropna().empty:
        return pd.to_datetime(series, errors="coerce")
    best_result = None
    best_count = -1
    strategies = [dict(dayfirst=False), dict(dayfirst=True)]
    for kwargs in strategies:
        result = pd.to_datetime(series, errors="coerce", **kwargs)
        count = result.notna().sum()
        if count > best_count:
            best_result = result
            best_count = count
    for dayfirst in [False, True]:
        try:
            result = pd.to_datetime(series, errors="coerce", format="mixed", dayfirst=dayfirst)
            count = result.notna().sum()
            if count > best_count:
                best_result = result
                best_count = count
        except Exception:
            pass
    return best_result if best_result is not None else pd.to_datetime(series, errors="coerce")


def parse_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def render_metric(label: str, value, color_class: str = "blue"):
    st.markdown(f"""
    <div class="metric-card metric-{color_class}">
        <div class="metric-value">{value}</div>
        <div class="metric-label">{label}</div>
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────
# COLUMN MATCHING
# ─────────────────────────────────────────────────
def find_col(df: pd.DataFrame, candidates: list[str], required: bool = False) -> str | None:
    col_map = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in col_map:
            return col_map[key]
        for k, v in col_map.items():
            if key in k or k in key:
                return v
    if required:
        st.error(f"Could not find column matching any of: {candidates}")
    return None


def clean_columns(df):
    df.columns = [str(c).strip().replace("\n", " ").replace("\r", " ") for c in df.columns]
    return df


def score_sheet_format(df, expected_alias_groups):
    if not expected_alias_groups:
        return 0
    return sum(1 for aliases in expected_alias_groups if find_col(df, aliases) is not None)


@st.cache_data(show_spinner=False)
def read_upload_from_bytes(file_bytes, file_name, expected_alias_groups, preferred_sheet_keywords):
    name = file_name.lower()
    expected_groups = [list(g) for g in expected_alias_groups]
    preferred_keywords = [k.lower() for k in preferred_sheet_keywords]

    if name.endswith(".csv"):
        df_csv = clean_columns(pd.read_csv(BytesIO(file_bytes)))
        score = score_sheet_format(df_csv, expected_groups)
        return df_csv, "CSV", score, ["CSV"]

    try:
        xls = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
        sheet_names = xls.sheet_names
    except Exception:
        df_single = clean_columns(pd.read_excel(BytesIO(file_bytes)))
        score = score_sheet_format(df_single, expected_groups)
        return df_single, "Sheet1", score, []

    best_sheet = None
    best_score = -1
    best_header_row = 0

    for sheet in sheet_names:
        try:
            best_score_for_sheet = -1
            best_header_for_sheet = 0
            for header_row in [0, 1, 2]:
                try:
                    header_df = clean_columns(pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, header=header_row, nrows=5, engine="openpyxl"))
                    score = score_sheet_format(header_df, expected_groups)
                    if score > best_score_for_sheet:
                        best_score_for_sheet = score
                        best_header_for_sheet = header_row
                except Exception:
                    continue
            sheet_name_norm = str(sheet).strip().lower().replace("_", " ")
            if any(k in sheet_name_norm for k in preferred_keywords):
                best_score_for_sheet += 1
            if best_score_for_sheet > best_score:
                best_sheet = sheet
                best_score = best_score_for_sheet
                best_header_row = best_header_for_sheet
        except Exception:
            continue

    if best_sheet is None:
        return None, None, 0, sheet_names
    df_selected = clean_columns(pd.read_excel(BytesIO(file_bytes), sheet_name=best_sheet, header=best_header_row, engine="openpyxl"))
    return df_selected, best_sheet, best_score, sheet_names


# ─────────────────────────────────────────────────
# COLUMN ALIAS DEFINITIONS
# ─────────────────────────────────────────────────
FC_PROJECT_NAME = ["project name", "project"]
FC_PROJECT_ID = ["project id (icenter)", "project id", "project_id"]
FC_CABINETS_QTY = ["cabinets qty", "cabinet qty", "cabinets quantity", "cab qty"]
FC_NEED_BY_DATE = ["project need by date for shipment (exw)", "need by date", "shipment date"]
FC_FAT_START = ["project need by date for fat start", "fat start date", "fat start"]
FC_ASSEMBLY_START = ["assembly start date", "assembly start"]
FC_ASSEMBLY_END = ["assembly completion date", "assembly end", "assembly completion"]
FC_MATERIAL_AVAIL = ["material availability", "material avail"]
FC_BOM_AVAIL = ["bom availability planned (in oracle)", "bom availability", "bom avail"]
FC_STATUS = ["project status", "status"]
FC_TYPE = ["type (confirmed / forecasted)", "type confirmed forecasted", "confirmed/forecasted", "type"]
FC_SR = ["sr no", "sr", "sr.no", "sr no."]
FC_SHIP_TO = ["ship to"]
FC_BU = ["bu"]
FC_BUILT_BY = ["built by"]
FC_BUILD_PERIOD = ["build period"]
FC_PROMISED_BUILD_DATE = ["promised build date", "promised build datre", "promised build"]
FC_SHIP_PERIOD = ["ship period"]
FC_PROMISED_SHIP = ["promised shipped date", "promised ship date"]
FC_ACTUAL_SHIP = ["actual shipped dates", "actual ship date"]
FC_REMARKS = ["remarks"]

BOM_PROJECT_NUM = ["project num", "project number", "project_num"]
BOM_PROJECT_NAME = ["project name", "project"]
BOM_ORDER_NUM = ["order number", "order num", "order_number"]
BOM_WO_NUM = ["work order number", "work order", "wo number"]
BOM_ITEM = ["item"]
BOM_ITEM_DESC = ["item desc", "item description", "item_desc"]
BOM_COMPONENT_CODE = ["component code", "comp code", "component_code"]
BOM_COMPONENT_DESC = ["component desc", "component description", "comp desc", "component_desc"]
BOM_REQUIRED_QTY = ["required quantity", "required qty", "req qty"]
BOM_ISSUED_QTY = ["quantity issued", "qty issued", "issued qty"]
BOM_OPEN_QTY = ["open quantity2", "open quantity", "open qty"]
BOM_ON_HAND = ["on hand quantity", "on hand qty", "on_hand"]
BOM_SCHEDULE_SHIP = ["schedule ship date", "scheduled ship date"]
BOM_COMP_MAKE_BUY = ["comp make or buy", "comp make/buy", "component make or buy"]
BOM_MAKE_BUY = ["make or buy", "make/buy"]
BOM_WO_STATUS = ["work order status", "wo status"]
BOM_CABINET_TYPE = ["cabinet/buyout/mro", "cabinet buyout mro", "cab/buyout"]
BOM_CUST_PO = ["cust po number", "customer po", "cust po"]
BOM_SALES_STATUS = ["sales status"]
BOM_VARIANCE = ["variance"]
BOM_AVAILABILITY = ["availability"]
BOM_TOTAL_AVAIL = ["total available", "total avail"]
BOM_PO_RECEIVING = ["po in receiving", "po receiving"]
BOM_SUPPLIER = ["supplier"]
BOM_BUYER = ["buyer"]
BOM_MFG_LEAD = ["manufacturing lead time", "mfg lead time", "lead time"]
BOM_JOB_START = ["job start date", "job start"]
BOM_TOTAL_DEMAND = ["total demand"]
BOM_NET_EXT_AVAIL = ["net extended available qty", "net extended avail"]

STK_ITEM_NUM = ["item number", "item num", "item_number", "item no"]
STK_ITEM_DESC = ["item description", "item desc"]
STK_ON_HAND = ["on hand quantity", "on hand qty", "onhand"]
STK_SUB_INV = ["sub inventory", "subinventory", "sub inv"]
STK_PROJECT = ["project number", "project num", "project"]
STK_ORG = ["organization code", "org code"]
STK_UOM = ["uom", "unit"]
STK_STATUS = ["status"]
STK_LOCATOR = ["stock locator", "locator"]


# ─────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📦 Material Shortage Forecaster</h1>
    <p>Upload Forecasting, Open Orders BOM &amp; Stock sheets to predict material shortages across projects</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📂 Data Upload")
    fc_file = st.file_uploader("Forecasting Sheet", type=["xlsx", "xls", "xlsm", "csv"], key="fc")
    bom_file = st.file_uploader("Open Orders BOM Sheet", type=["xlsx", "xls", "xlsm", "csv"], key="bom")
    stock_file = st.file_uploader("Stock Sheet", type=["xlsx", "xls", "xlsm", "csv"], key="stock")

    st.markdown("---")
    st.markdown("### 💾 Saved Lots")
    saved_lots = load_saved_lots()

    if saved_lots:
        lot_names = list(saved_lots.keys())
        selected_saved = st.selectbox("Load a saved lot", ["-- Select --"] + lot_names)
        if selected_saved != "-- Select --":
            st.success(f"Lot **{selected_saved}** loaded")
        col_del1, col_del2 = st.columns(2)
        with col_del1:
            del_lot = st.selectbox("Delete lot", ["-- Select --"] + lot_names, key="del_lot")
        with col_del2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️ Delete", use_container_width=True):
                if del_lot != "-- Select --":
                    del saved_lots[del_lot]
                    save_lots(saved_lots)
                    st.rerun()
    else:
        st.info("No saved lots yet. Create lots in Project Drill-Down tab.")


# ─────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────
def read_upload(file_obj, expected_alias_groups, preferred_sheet_keywords=()):
    if file_obj is None:
        return None, None, 0, []
    frozen_aliases = tuple(tuple(g) for g in expected_alias_groups)
    frozen_keywords = tuple(preferred_sheet_keywords)
    return read_upload_from_bytes(file_obj.getvalue(), file_obj.name, frozen_aliases, frozen_keywords)


FC_FORMAT_ALIASES = [FC_PROJECT_NAME, FC_CABINETS_QTY, FC_NEED_BY_DATE, FC_STATUS, FC_BUILD_PERIOD]
BOM_FORMAT_ALIASES = [BOM_PROJECT_NAME, BOM_WO_NUM, BOM_ITEM, BOM_COMPONENT_CODE, BOM_REQUIRED_QTY, BOM_OPEN_QTY]
STK_FORMAT_ALIASES = [STK_ITEM_NUM, STK_ITEM_DESC, STK_ON_HAND, STK_SUB_INV]

df_fc, fc_sheet_used, fc_format_score, fc_sheets = read_upload(fc_file, FC_FORMAT_ALIASES, ["forecast", "planning", "schedule", "fy"])
df_bom, bom_sheet_used, bom_format_score, bom_sheets = read_upload(bom_file, BOM_FORMAT_ALIASES, ["openordersbom", "open orders bom", "bom", "openorders"])
df_stock, stock_sheet_used, stock_format_score, stock_sheets = read_upload(stock_file, STK_FORMAT_ALIASES, ["stock", "onhand", "inventory"])

all_loaded = df_fc is not None and df_bom is not None and df_stock is not None

if all_loaded:
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 🧭 Sheet Detection")
        st.caption(f"Forecasting: {fc_sheet_used} ({fc_format_score}/{len(FC_FORMAT_ALIASES)} cols)")
        st.caption(f"BOM: {bom_sheet_used} ({bom_format_score}/{len(BOM_FORMAT_ALIASES)} cols)")
        st.caption(f"Stock: {stock_sheet_used} ({stock_format_score}/{len(STK_FORMAT_ALIASES)} cols)")

if not all_loaded:
    st.info("⬅️ Upload all three sheets from the sidebar to begin analysis.")
    with st.expander("ℹ️  Expected Sheet Formats", expanded=True):
        st.markdown("""
**Forecasting Sheet** – Key columns: `Project Name`, `Cabinets Qty`, `Project Need by date for Shipment (EXW)`, `Assembly Start Date`, `Project Status`

**Open Orders BOM Sheet** – Key columns: `Project Name`, `Work Order Number`, `ITEM`, `Item Desc`, `Component Code`, `Required Quantity`, `Open Quantity2`

**Stock Sheet** – Key columns: `Item Number`, `Item Description`, `On Hand Quantity`, `Sub Inventory`
        """)
    st.stop()


# ─────────────────────────────────────────────────
# RESOLVE COLUMNS
# ─────────────────────────────────────────────────
fc_proj_name_col = find_col(df_fc, FC_PROJECT_NAME)
fc_proj_id_col = find_col(df_fc, FC_PROJECT_ID)
fc_cab_qty_col = find_col(df_fc, FC_CABINETS_QTY)
fc_need_by_col = find_col(df_fc, FC_NEED_BY_DATE)
fc_fat_start_col = find_col(df_fc, FC_FAT_START)
fc_asm_start_col = find_col(df_fc, FC_ASSEMBLY_START)
fc_asm_end_col = find_col(df_fc, FC_ASSEMBLY_END)
fc_mat_avail_col = find_col(df_fc, FC_MATERIAL_AVAIL)
fc_bom_avail_col = find_col(df_fc, FC_BOM_AVAIL)
fc_status_col = find_col(df_fc, FC_STATUS)
fc_type_col = find_col(df_fc, FC_TYPE)
fc_sr_col = find_col(df_fc, FC_SR)
fc_ship_to_col = find_col(df_fc, FC_SHIP_TO)
fc_bu_col = find_col(df_fc, FC_BU)
fc_built_by_col = find_col(df_fc, FC_BUILT_BY)
fc_build_period_col = find_col(df_fc, FC_BUILD_PERIOD)
fc_promised_build_col = find_col(df_fc, FC_PROMISED_BUILD_DATE)
fc_ship_period_col = find_col(df_fc, FC_SHIP_PERIOD)
fc_promised_ship_col = find_col(df_fc, FC_PROMISED_SHIP)
fc_actual_ship_col = find_col(df_fc, FC_ACTUAL_SHIP)
fc_remarks_col = find_col(df_fc, FC_REMARKS)

if not fc_promised_build_col:
    fc_promised_build_col = fc_promised_ship_col

bom_proj_num_col = find_col(df_bom, BOM_PROJECT_NUM)
bom_proj_name_col = find_col(df_bom, BOM_PROJECT_NAME)
bom_order_col = find_col(df_bom, BOM_ORDER_NUM)
bom_wo_col = find_col(df_bom, BOM_WO_NUM)
bom_item_col = find_col(df_bom, BOM_ITEM)
bom_item_desc_col = find_col(df_bom, BOM_ITEM_DESC)
bom_comp_code_col = find_col(df_bom, BOM_COMPONENT_CODE)
bom_comp_desc_col = find_col(df_bom, BOM_COMPONENT_DESC)
bom_req_qty_col = find_col(df_bom, BOM_REQUIRED_QTY)
bom_issued_qty_col = find_col(df_bom, BOM_ISSUED_QTY)
bom_open_qty_col = find_col(df_bom, BOM_OPEN_QTY)
bom_onhand_col = find_col(df_bom, BOM_ON_HAND)
bom_ship_col = find_col(df_bom, BOM_SCHEDULE_SHIP)
bom_comp_mb_col = find_col(df_bom, BOM_COMP_MAKE_BUY)
bom_mb_col = find_col(df_bom, BOM_MAKE_BUY)
bom_wo_status_col = find_col(df_bom, BOM_WO_STATUS)
bom_cab_type_col = find_col(df_bom, BOM_CABINET_TYPE)
bom_cust_po_col = find_col(df_bom, BOM_CUST_PO)
bom_sales_status_col = find_col(df_bom, BOM_SALES_STATUS)
bom_variance_col = find_col(df_bom, BOM_VARIANCE)
bom_availability_col = find_col(df_bom, BOM_AVAILABILITY)
bom_total_avail_col = find_col(df_bom, BOM_TOTAL_AVAIL)
bom_po_recv_col = find_col(df_bom, BOM_PO_RECEIVING)
bom_supplier_col = find_col(df_bom, BOM_SUPPLIER)
bom_buyer_col = find_col(df_bom, BOM_BUYER)
bom_mfg_lead_col = find_col(df_bom, BOM_MFG_LEAD)
bom_job_start_col = find_col(df_bom, BOM_JOB_START)
bom_total_demand_col = find_col(df_bom, BOM_TOTAL_DEMAND)
bom_net_ext_col = find_col(df_bom, BOM_NET_EXT_AVAIL)

stk_item_col = find_col(df_stock, STK_ITEM_NUM)
stk_desc_col = find_col(df_stock, STK_ITEM_DESC)
stk_onhand_col = find_col(df_stock, STK_ON_HAND)
stk_subinv_col = find_col(df_stock, STK_SUB_INV)
stk_proj_col = find_col(df_stock, STK_PROJECT)
stk_org_col = find_col(df_stock, STK_ORG)
stk_uom_col = find_col(df_stock, STK_UOM)
stk_status_col = find_col(df_stock, STK_STATUS)
stk_locator_col = find_col(df_stock, STK_LOCATOR)


# ─────────────────────────────────────────────────
# PARSE NUMERICS & DATES
# ─────────────────────────────────────────────────
numeric_cols_fc = [fc_cab_qty_col]
for nc in numeric_cols_fc:
    if nc and nc in df_fc.columns:
        df_fc[nc] = parse_numeric(df_fc[nc])

_fc_date_cols_to_parse = set()
for dc in [fc_need_by_col, fc_fat_start_col, fc_asm_start_col, fc_asm_end_col,
           fc_mat_avail_col, fc_bom_avail_col, fc_promised_build_col, fc_promised_ship_col, fc_actual_ship_col]:
    if dc and dc in df_fc.columns:
        _fc_date_cols_to_parse.add(dc)
for dc in _fc_date_cols_to_parse:
    df_fc[dc] = robust_to_datetime(df_fc[dc])

for nc in [bom_req_qty_col, bom_issued_qty_col, bom_open_qty_col, bom_onhand_col,
           bom_variance_col, bom_total_avail_col, bom_po_recv_col, bom_total_demand_col, bom_net_ext_col]:
    if nc and nc in df_bom.columns:
        df_bom[nc] = parse_numeric(df_bom[nc])

for dc in [bom_ship_col, bom_job_start_col]:
    if dc and dc in df_bom.columns:
        df_bom[dc] = robust_to_datetime(df_bom[dc])

if stk_onhand_col:
    df_stock[stk_onhand_col] = parse_numeric(df_stock[stk_onhand_col])


# ─────────────────────────────────────────────────
# AGGREGATE STOCK
# ─────────────────────────────────────────────────
if stk_item_col and stk_onhand_col:
    stock_agg = (
        df_stock.groupby(stk_item_col, dropna=False)[stk_onhand_col]
        .sum().reset_index()
        .rename(columns={stk_item_col: "_stk_item", stk_onhand_col: "_stk_total_onhand"})
    )
else:
    stock_agg = pd.DataFrame(columns=["_stk_item", "_stk_total_onhand"])


# ─────────────────────────────────────────────────
# ROW-LEVEL LOT ASSIGNMENT PERSISTENCE
# ─────────────────────────────────────────────────
ROW_LOT_ASSIGNMENTS_FILE = "row_lot_assignments.json"

def load_row_lot_assignments() -> dict:
    """Load row-level lot assignments. Key = row_key, Value = list of lot names."""
    if os.path.exists(ROW_LOT_ASSIGNMENTS_FILE):
        with open(ROW_LOT_ASSIGNMENTS_FILE, "r") as f:
            return json.load(f)
    return {}

def save_row_lot_assignments(assignments: dict):
    with open(ROW_LOT_ASSIGNMENTS_FILE, "w") as f:
        json.dump(assignments, f, indent=2, default=str)


# ─────────────────────────────────────────────────
# ROW KEY HELPERS
# ─────────────────────────────────────────────────
ROW_KEY_SEP = "|||"

def make_fc_row_key(row, proj_col, bp_col, asm_col, idx=None):
    """Create a unique key for a forecasting schedule row using project + build period + assembly start + index."""
    pname = str(row.get(proj_col, "")).strip() if pd.notna(row.get(proj_col)) else ""
    bp = str(row.get(bp_col, "")).strip() if bp_col and pd.notna(row.get(bp_col)) else ""
    asm = ""
    if asm_col and pd.notna(row.get(asm_col)):
        val = row[asm_col]
        asm = val.strftime("%Y-%m-%d") if isinstance(val, pd.Timestamp) else str(val)
    idx_str = str(idx) if idx is not None else ""
    return f"{pname}{ROW_KEY_SEP}{bp}{ROW_KEY_SEP}{asm}{ROW_KEY_SEP}{idx_str}"


def make_fc_row_label(row, proj_col, bp_col, asm_col):
    """Create a display label for a forecasting row."""
    pname = str(row.get(proj_col, "")).strip() if pd.notna(row.get(proj_col)) else "Unknown"
    bp = str(row.get(bp_col, "")).strip() if bp_col and pd.notna(row.get(bp_col)) else "N/A"
    asm = "N/A"
    if asm_col and pd.notna(row.get(asm_col)):
        val = row[asm_col]
        asm = val.strftime("%d-%b-%Y") if isinstance(val, pd.Timestamp) else str(val)
    return f"{pname}  |  Build: {bp}  |  Asm Start: {asm}"


def format_date_col(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series.dt.strftime("%d-%b-%Y").fillna("")
    return series


# ─────────────────────────────────────────────────
# ADD ROW KEYS TO FORECASTING DF
# ─────────────────────────────────────────────────
df_fc["_row_key"] = [
    make_fc_row_key(row, fc_proj_name_col, fc_build_period_col, fc_asm_start_col, idx=i)
    for i, row in df_fc.iterrows()
]
df_fc["_row_label"] = [
    make_fc_row_label(row, fc_proj_name_col, fc_build_period_col, fc_asm_start_col)
    for _, row in df_fc.iterrows()
]


# ─────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────
tab_overview, tab_project, tab_forecasting, tab_shortage, tab_timeline, tab_raw = st.tabs([
    "📊 Overview", "🏗️ Project Drill-Down", "📈 Forecasting",
    "🚨 Shortage Forecast", "📅 Timeline", "📋 Raw Data"
])


# ═══════════════════════════════════════════════════
# TAB 1 – OVERVIEW
# ═══════════════════════════════════════════════════
with tab_overview:
    st.markdown('<div class="section-header">Portfolio Summary</div>', unsafe_allow_html=True)

    total_projects = df_fc[fc_proj_name_col].nunique() if fc_proj_name_col else len(df_fc)
    total_cabinets = int(df_fc[fc_cab_qty_col].sum()) if fc_cab_qty_col else "N/A"
    total_work_orders = df_bom[bom_wo_col].nunique() if bom_wo_col else "N/A"
    unique_components = df_bom[bom_comp_code_col].nunique() if bom_comp_code_col else "N/A"
    total_open_qty = int(df_bom[bom_open_qty_col].sum()) if bom_open_qty_col else "N/A"

    shortage_count = 0
    comp_vs_stock = pd.DataFrame()
    if bom_comp_code_col and bom_open_qty_col:
        comp_demand = (
            df_bom.groupby(bom_comp_code_col, dropna=False)[bom_open_qty_col]
            .sum().reset_index()
            .rename(columns={bom_comp_code_col: "_comp", bom_open_qty_col: "_demand"})
        )
        comp_vs_stock = comp_demand.merge(stock_agg, left_on="_comp", right_on="_stk_item", how="left")
        comp_vs_stock["_stk_total_onhand"] = comp_vs_stock["_stk_total_onhand"].fillna(0)
        comp_vs_stock["_gap"] = comp_vs_stock["_stk_total_onhand"] - comp_vs_stock["_demand"]
        shortage_count = int((comp_vs_stock["_gap"] < 0).sum())

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    with m1: render_metric("Projects", total_projects, "blue")
    with m2: render_metric("Total Cabinets", total_cabinets, "purple")
    with m3: render_metric("Work Orders", total_work_orders, "cyan")
    with m4: render_metric("Unique Components", unique_components, "emerald")
    with m5: render_metric("Total Open Qty", total_open_qty, "amber")
    with m6: render_metric("Components Short", shortage_count, "red")

    st.markdown("")
    ch1, ch2 = st.columns(2)

    with ch1:
        st.markdown('<div class="section-header">Cabinets by Project</div>', unsafe_allow_html=True)
        if fc_proj_name_col and fc_cab_qty_col:
            cab_data = df_fc.groupby(fc_proj_name_col, dropna=False)[fc_cab_qty_col].sum().reset_index().sort_values(fc_cab_qty_col, ascending=True).tail(15)
            fig = px.bar(cab_data, y=fc_proj_name_col, x=fc_cab_qty_col, orientation="h", color_discrete_sequence=["#3b82f6"])
            fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#94a3b8"), height=400, xaxis_title="Cabinets", yaxis_title="", margin=dict(l=10, r=10, t=10, b=30))
            st.plotly_chart(fig, use_container_width=True)

    with ch2:
        st.markdown('<div class="section-header">Shortage Distribution</div>', unsafe_allow_html=True)
        if not comp_vs_stock.empty:
            comp_vs_stock["_status"] = comp_vs_stock["_gap"].apply(lambda x: "Sufficient" if x >= 0 else "Short")
            status_counts = comp_vs_stock["_status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Count"]
            fig2 = px.pie(status_counts, names="Status", values="Count", color="Status", color_discrete_map={"Sufficient": "#10b981", "Short": "#ef4444"}, hole=0.55)
            fig2.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#94a3b8"), height=400, margin=dict(l=10, r=10, t=10, b=30), legend=dict(orientation="h", yanchor="bottom", y=-0.1))
            st.plotly_chart(fig2, use_container_width=True)

    if fc_status_col and fc_proj_name_col:
        st.markdown('<div class="section-header">Project Status Overview</div>', unsafe_allow_html=True)
        status_df = df_fc[[fc_proj_name_col, fc_status_col]].dropna(subset=[fc_status_col])
        if not status_df.empty:
            sc2 = status_df[fc_status_col].value_counts().reset_index()
            sc2.columns = ["Status", "Count"]
            fig3 = px.bar(sc2, x="Status", y="Count", color_discrete_sequence=["#8b5cf6"])
            fig3.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#94a3b8"), height=300, margin=dict(l=10, r=10, t=10, b=30))
            st.plotly_chart(fig3, use_container_width=True)


# ═══════════════════════════════════════════════════
# TAB 2 – PROJECT DRILL-DOWN
# ═══════════════════════════════════════════════════
with tab_project:
    st.markdown('<div class="section-header">Project → Work Order → Item → Component Drill-Down</div>', unsafe_allow_html=True)

    project_list = sorted(df_bom[bom_proj_name_col].dropna().unique().tolist()) if bom_proj_name_col else []
    if not project_list:
        project_list = sorted(df_bom[bom_proj_num_col].dropna().unique().tolist()) if bom_proj_num_col else []

    selected_project = st.selectbox("Select Project", ["-- All --"] + project_list, key="proj_sel")
    df_filtered = df_bom.copy()
    if selected_project != "-- All --":
        proj_col_used = bom_proj_name_col or bom_proj_num_col
        if proj_col_used:
            df_filtered = df_filtered[df_filtered[proj_col_used] == selected_project]

    selected_wos = []
    selected_item_codes = []
    selected_item_labels = []

    wo_list = sorted(df_filtered[bom_wo_col].dropna().unique().tolist()) if bom_wo_col else []
    wo_options = [str(w) for w in wo_list]
    selected_wos = st.multiselect("Select Work Orders", wo_options, default=[], key="wo_sel_multi", help="Leave empty = all")
    if selected_wos and bom_wo_col:
        df_filtered = df_filtered[df_filtered[bom_wo_col].astype(str).isin(selected_wos)]

    if bom_item_col and bom_item_desc_col:
        item_combos = df_filtered[[bom_item_col, bom_item_desc_col]].drop_duplicates().dropna(subset=[bom_item_col])
        item_combos["_label"] = item_combos[bom_item_col].astype(str) + " — " + item_combos[bom_item_desc_col].fillna("").astype(str)
        item_labels = sorted(item_combos["_label"].tolist())
        selected_item_labels = st.multiselect("Select ITEM Codes", item_labels, default=[], key="item_sel_multi", help="Leave empty = all")
        if selected_item_labels:
            selected_item_codes = [lbl.split(" — ")[0].strip() for lbl in selected_item_labels]
            df_filtered = df_filtered[df_filtered[bom_item_col].astype(str).isin(selected_item_codes)]
    elif bom_item_col:
        item_list = sorted(df_filtered[bom_item_col].dropna().unique().tolist())
        selected_item_codes = st.multiselect("Select ITEM Codes", [str(i) for i in item_list], default=[], key="item_sel_multi_plain")
        if selected_item_codes:
            df_filtered = df_filtered[df_filtered[bom_item_col].astype(str).isin(selected_item_codes)]

    scope_wo_count = len(selected_wos) if selected_wos else len(wo_options)
    scope_item_count = len(selected_item_codes) if selected_item_codes else (df_filtered[bom_item_col].dropna().astype(str).nunique() if bom_item_col else 0)
    st.caption(f"Scope: {scope_wo_count} Work Order(s), {scope_item_count} ITEM(s).")

    st.markdown('<div class="section-header">Component Detail</div>', unsafe_allow_html=True)

    display_cols = [c for c in [bom_comp_code_col, bom_comp_desc_col, bom_req_qty_col, bom_issued_qty_col, bom_open_qty_col, bom_onhand_col, bom_comp_mb_col, bom_variance_col, bom_availability_col, bom_total_avail_col, bom_po_recv_col, bom_supplier_col, bom_buyer_col, bom_mfg_lead_col, bom_wo_status_col] if c is not None]

    comp_detail = pd.DataFrame()
    if display_cols:
        comp_base = df_filtered[display_cols].copy()
        if bom_comp_code_col:
            group_keys = [bom_comp_code_col]
            if bom_comp_desc_col: group_keys.append(bom_comp_desc_col)
            agg_map = {}
            for sc in [bom_req_qty_col, bom_issued_qty_col, bom_open_qty_col, bom_total_avail_col, bom_po_recv_col]:
                if sc and sc in comp_base.columns and sc not in group_keys: agg_map[sc] = "sum"
            for mc in [bom_onhand_col, bom_variance_col]:
                if mc and mc in comp_base.columns and mc not in group_keys: agg_map[mc] = "max"
            for fc_c in [bom_comp_mb_col, bom_availability_col, bom_supplier_col, bom_buyer_col, bom_mfg_lead_col, bom_wo_status_col]:
                if fc_c and fc_c in comp_base.columns and fc_c not in group_keys and fc_c not in agg_map: agg_map[fc_c] = "first"
            comp_detail = comp_base.groupby(group_keys, dropna=False, as_index=False).agg(agg_map)
            if bom_wo_col and bom_wo_col in df_filtered.columns:
                wo_counts = df_filtered.groupby(group_keys, dropna=False)[bom_wo_col].nunique().reset_index(name="Work Orders")
                comp_detail = comp_detail.merge(wo_counts, on=group_keys, how="left")
        else:
            comp_detail = comp_base.copy()

        if bom_comp_code_col and not stock_agg.empty:
            comp_detail = comp_detail.merge(stock_agg, left_on=bom_comp_code_col, right_on="_stk_item", how="left")
            comp_detail["_stk_total_onhand"] = comp_detail["_stk_total_onhand"].fillna(0)
            if bom_open_qty_col:
                comp_detail["Stock vs Open Qty"] = comp_detail["_stk_total_onhand"] - comp_detail[bom_open_qty_col]
            comp_detail.rename(columns={"_stk_total_onhand": "Total Stock On Hand"}, inplace=True)
            comp_detail.drop(columns=["_stk_item"], errors="ignore", inplace=True)

        if bom_open_qty_col and bom_open_qty_col in comp_detail.columns:
            comp_detail = comp_detail.sort_values(by=bom_open_qty_col, ascending=False)

        st.dataframe(comp_detail, use_container_width=True, height=450)

        mc1, mc2, mc3, mc4 = st.columns(4)
        with mc1: render_metric("Components", len(comp_detail), "cyan")
        with mc2:
            if bom_open_qty_col: render_metric("Total Open Qty", int(comp_detail[bom_open_qty_col].sum()), "amber")
        with mc3:
            if "Stock vs Open Qty" in comp_detail.columns:
                render_metric("Short Components", int((comp_detail["Stock vs Open Qty"] < 0).sum()), "red")
        with mc4:
            if bom_req_qty_col: render_metric("Total Required", int(comp_detail[bom_req_qty_col].sum()), "purple")

    st.markdown("---")
    st.markdown("### 💾 Save Current Selection as Lot")
    lot_name = st.text_input("Lot Name", key="lot_name_input")
    if st.button("Save Lot", type="primary", key="save_lot_btn"):
        if lot_name.strip():
            lot_work_orders = selected_wos if selected_wos else ["-- All --"]
            lot_items = selected_item_codes if selected_item_codes else ["-- All --"]
            short_count = int((comp_detail["Stock vs Open Qty"] < 0).sum()) if "Stock vs Open Qty" in comp_detail.columns else 0
            total_open = int(comp_detail[bom_open_qty_col].sum()) if bom_open_qty_col and bom_open_qty_col in comp_detail.columns else 0
            total_req = int(comp_detail[bom_req_qty_col].sum()) if bom_req_qty_col and bom_req_qty_col in comp_detail.columns else 0

            lot_info = {
                "project": selected_project,
                "work_orders": lot_work_orders,
                "items": lot_items,
                "item_labels": selected_item_labels if selected_item_labels else lot_items,
                "work_order_count": scope_wo_count,
                "item_count": scope_item_count,
                "work_order": "All" if lot_work_orders == ["-- All --"] else ", ".join(lot_work_orders),
                "item": "All" if lot_items == ["-- All --"] else ", ".join(lot_items),
                "saved_at": datetime.now().isoformat(),
                "component_count": len(comp_detail),
                "short_components": short_count,
                "total_open_qty": total_open,
                "total_required_qty": total_req,
            }
            saved_lots[lot_name.strip()] = lot_info
            save_lots(saved_lots)
            st.success(f"Lot **{lot_name}** saved!")
            st.rerun()
        else:
            st.error("Enter a lot name.")


# ═══════════════════════════════════════════════════
# TAB 3 – FORECASTING (with Row-Level Lot Assignment)
# ═══════════════════════════════════════════════════
with tab_forecasting:
    st.markdown('<div class="section-header">📈 Forecasting Schedule with Row-Level Lot Assignment</div>', unsafe_allow_html=True)

    if not fc_proj_name_col:
        st.error("❌ Project Name column not found in Forecasting sheet.")
        st.stop()
    else:
        # ── Build display table ──
        fc_display_cols = []
        seen_cols = set()
        for c in [fc_sr_col, fc_proj_name_col, fc_proj_id_col, fc_ship_to_col, fc_bu_col, fc_type_col,
                  fc_cab_qty_col, fc_built_by_col, fc_need_by_col, fc_fat_start_col, fc_build_period_col,
                  fc_promised_build_col, fc_bom_avail_col, fc_mat_avail_col, fc_asm_start_col, fc_asm_end_col,
                  fc_ship_period_col, fc_promised_ship_col, fc_actual_ship_col, fc_status_col, fc_remarks_col]:
            if c is not None and c in df_fc.columns and c not in seen_cols:
                fc_display_cols.append(c)
                seen_cols.add(c)

        if not fc_display_cols:
            fc_display_cols = [fc_proj_name_col]

        fc_view = df_fc[fc_display_cols + ["_row_key", "_row_label"]].copy()

        # Format date columns for display
        for col in fc_view.columns:
            if col not in ["_row_key", "_row_label"] and pd.api.types.is_datetime64_any_dtype(fc_view[col]):
                fc_view[col] = fc_view[col].dt.strftime("%d-%b-%Y").fillna("")

        fc_view = fc_view.dropna(how='all', subset=[c for c in fc_display_cols if c in fc_view.columns])

        # Summary metrics
        st.markdown('<div class="section-header">📊 Forecasting Data Summary</div>', unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            render_metric("Total Records", len(fc_view), "blue")
        with m2:
            if fc_cab_qty_col and fc_cab_qty_col in fc_view.columns:
                total_cabs = pd.to_numeric(fc_view[fc_cab_qty_col], errors='coerce').sum()
                render_metric("Total Cabinets", int(total_cabs) if not pd.isna(total_cabs) else "N/A", "purple")
        with m3:
            if fc_status_col and fc_status_col in fc_view.columns:
                status_counts = fc_view[fc_status_col].value_counts()
                render_metric("Status Types", len(status_counts), "emerald")
        with m4:
            if fc_type_col and fc_type_col in fc_view.columns:
                type_counts = fc_view[fc_type_col].value_counts()
                render_metric("Forecast Types", len(type_counts), "amber")

        # ── ROW-LEVEL Lot Assignment ──
        st.markdown('<div class="section-header">🔗 Assign Lots to Forecasting Schedule Rows</div>', unsafe_allow_html=True)
        st.caption("Assign lots to specific rows identified by Project Name + Build Period + Assembly Start Date")

        row_lot_assignments = load_row_lot_assignments()
        saved_lots = load_saved_lots()

        if not saved_lots:
            st.info("ℹ️  No saved lots yet. Create lots in the **Project Drill-Down** tab first.")
        else:
            lot_name_options = sorted(saved_lots.keys())

            # Build row selector options — show label but track key
            row_options_map = {}  # label -> key
            for _, row in fc_view.iterrows():
                label = row["_row_label"]
                key = row["_row_key"]
                row_options_map[label] = key

            row_labels_sorted = sorted(row_options_map.keys())

            ac1, ac2, ac3 = st.columns([3, 2, 1])
            with ac1:
                assign_row_label = st.selectbox(
                    "Select Forecasting Row (Project | Build Period | Asm Start)",
                    ["-- Select --"] + row_labels_sorted,
                    key="fc_assign_row"
                )
            with ac2:
                assign_lots = st.multiselect("Assign Lots", lot_name_options, key="fc_assign_lots")
            with ac3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("✅ Assign", use_container_width=True, key="fc_assign_btn"):
                    if assign_row_label != "-- Select --" and assign_lots:
                        row_key = row_options_map.get(assign_row_label, "")
                        if row_key:
                            existing = row_lot_assignments.get(row_key, [])
                            merged = list(dict.fromkeys(existing + assign_lots))
                            row_lot_assignments[row_key] = merged
                            save_row_lot_assignments(row_lot_assignments)
                            st.success(f"Assigned {len(assign_lots)} lot(s) to row")
                            st.rerun()
                    else:
                        st.error("Select both a row and at least one lot.")

            # Show current assignments
            if row_lot_assignments:
                with st.expander("📋 Current Row-Level Lot Assignments", expanded=False):
                    # Build reverse lookup: key -> label
                    key_to_label = {v: k for k, v in row_options_map.items()}
                    for rk, lots_list in sorted(row_lot_assignments.items(), key=lambda x: x[0]):
                        if lots_list:
                            display_label = key_to_label.get(rk, rk)
                            chips = " ".join([f'<span class="lot-chip">{l}</span>' for l in lots_list])
                            st.markdown(f"**{display_label}:** {chips}", unsafe_allow_html=True)

                    st.markdown("---")
                    rc1, rc2, rc3 = st.columns([3, 2, 1])
                    with rc1:
                        # Show only assigned rows for removal
                        assigned_labels = [key_to_label.get(k, k) for k in row_lot_assignments.keys() if row_lot_assignments[k]]
                        rm_row_label = st.selectbox("Remove from row", ["-- Select --"] + sorted(assigned_labels), key="fc_rm_row")
                    with rc2:
                        rm_row_key = row_options_map.get(rm_row_label, "") if rm_row_label != "-- Select --" else ""
                        rm_options = row_lot_assignments.get(rm_row_key, [])
                        rm_lots = st.multiselect("Lots to remove", rm_options, key="fc_rm_lots")
                    with rc3:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("🗑️ Remove", use_container_width=True, key="fc_rm_btn"):
                            if rm_row_key and rm_lots:
                                row_lot_assignments[rm_row_key] = [l for l in row_lot_assignments.get(rm_row_key, []) if l not in rm_lots]
                                if not row_lot_assignments[rm_row_key]:
                                    del row_lot_assignments[rm_row_key]
                                save_row_lot_assignments(row_lot_assignments)
                                st.rerun()

        # ── Enrich table with lot info per row ──
        lot_col_assigned = []
        lot_col_comp_count = []
        lot_col_short = []
        lot_col_open_qty = []

        for _, row in fc_view.iterrows():
            rk = row["_row_key"]
            assigned = [l for l in row_lot_assignments.get(rk, []) if l in saved_lots]

            if assigned:
                lot_col_assigned.append(", ".join(assigned))
                lot_col_comp_count.append(sum(saved_lots[l].get("component_count", 0) for l in assigned))
                lot_col_short.append(sum(saved_lots[l].get("short_components", 0) for l in assigned))
                lot_col_open_qty.append(sum(saved_lots[l].get("total_open_qty", 0) for l in assigned))
            else:
                lot_col_assigned.append("")
                lot_col_comp_count.append(0)
                lot_col_short.append(0)
                lot_col_open_qty.append(0)

        fc_view["Assigned Lots"] = lot_col_assigned
        fc_view["Lot Components"] = lot_col_comp_count
        fc_view["Lot Shortages"] = lot_col_short
        fc_view["Lot Open Qty"] = lot_col_open_qty

        # Sort by assembly start date
        if fc_asm_start_col and fc_asm_start_col in df_fc.columns:
            sort_order = df_fc[fc_asm_start_col].sort_values(na_position="last").index
            # Reindex fc_view to match sort order for rows that exist
            valid_idx = [i for i in sort_order if i in fc_view.index]
            fc_view = fc_view.loc[valid_idx].reset_index(drop=True)

        # ── Filters ──
        flt1, flt2, flt3 = st.columns(3)
        with flt1:
            fc_filter_status = []
            if fc_status_col and fc_status_col in fc_view.columns:
                statuses = sorted(fc_view[fc_status_col].dropna().unique().tolist())
                fc_filter_status = st.multiselect("Filter by Status", statuses, key="fc_flt_status")
        with flt2:
            fc_filter_lot = st.selectbox("Lot Filter", ["All", "With Lots Only", "Without Lots Only"], key="fc_flt_lot")
        with flt3:
            fc_search = st.text_input("Search Project", key="fc_search_proj")

        fc_display = fc_view.copy()
        if fc_filter_status and fc_status_col in fc_display.columns:
            fc_display = fc_display[fc_display[fc_status_col].isin(fc_filter_status)]
        if fc_filter_lot == "With Lots Only":
            fc_display = fc_display[fc_display["Assigned Lots"] != ""]
        elif fc_filter_lot == "Without Lots Only":
            fc_display = fc_display[fc_display["Assigned Lots"] == ""]
        if fc_search:
            fc_display = fc_display[fc_display[fc_proj_name_col].astype(str).str.contains(fc_search, case=False, na=False)]

        # ── Display metrics ──
        m1, m2, m3, m4 = st.columns(4)
        with m1: render_metric("Rows Shown", len(fc_display), "blue")
        with m2: render_metric("With Lots", int((fc_display["Assigned Lots"] != "").sum()), "purple")
        with m3: render_metric("Total Lot Shortages", int(fc_display["Lot Shortages"].sum()), "red")
        with m4:
            if fc_cab_qty_col and fc_cab_qty_col in fc_display.columns:
                render_metric("Total Cabinets", int(parse_numeric(fc_display[fc_cab_qty_col]).sum()), "cyan")

        # Drop internal columns for display
        display_table = fc_display.drop(columns=["_row_key", "_row_label"], errors="ignore")
        st.dataframe(display_table, use_container_width=True, height=550)
        st.caption(f"Showing {len(fc_display):,} of {len(fc_view):,} rows.")

        # ── Lot detail drill-down per row ──
        rows_with_lots = fc_display[fc_display["Assigned Lots"] != ""]
        if not rows_with_lots.empty:
            st.markdown('<div class="section-header">Lot Component Detail per Row</div>', unsafe_allow_html=True)
            detail_row_labels = rows_with_lots["_row_label"].tolist()
            detail_row_sel = st.selectbox("Select row to view lot details", ["-- Select --"] + detail_row_labels, key="fc_lot_detail_row")

            if detail_row_sel != "-- Select --":
                sel_row = rows_with_lots[rows_with_lots["_row_label"] == detail_row_sel].iloc[0]
                sel_rk = sel_row["_row_key"]
                assigned_lot_names = [l for l in row_lot_assignments.get(sel_rk, []) if l in saved_lots]

                for lot_nm in assigned_lot_names:
                    lot_info = saved_lots.get(lot_nm, {})
                    if not lot_info:
                        continue
                    with st.expander(f"📦 Lot: **{lot_nm}**", expanded=True):
                        d1, d2, d3, d4 = st.columns(4)
                        with d1: st.metric("BOM Project", lot_info.get("project", "N/A"))
                        with d2: st.metric("Work Orders", lot_info.get("work_order_count", "N/A"))
                        with d3: st.metric("Components", lot_info.get("component_count", 0))
                        with d4: st.metric("Short", lot_info.get("short_components", 0))
                        st.caption(f"Items: {lot_info.get('item', 'All')}  |  Saved: {lot_info.get('saved_at', 'N/A')}")

                        # Reconstruct BOM for this lot
                        lot_bom_project = lot_info.get("project", "")
                        lot_wo_list = lot_info.get("work_orders", ["-- All --"])
                        lot_item_list = lot_info.get("items", ["-- All --"])

                        df_lot_bom = df_bom.copy()
                        if lot_bom_project and lot_bom_project != "-- All --":
                            p_col = bom_proj_name_col or bom_proj_num_col
                            if p_col:
                                df_lot_bom = df_lot_bom[df_lot_bom[p_col] == lot_bom_project]
                        if lot_wo_list != ["-- All --"] and bom_wo_col:
                            df_lot_bom = df_lot_bom[df_lot_bom[bom_wo_col].astype(str).isin(lot_wo_list)]
                        if lot_item_list != ["-- All --"] and bom_item_col:
                            df_lot_bom = df_lot_bom[df_lot_bom[bom_item_col].astype(str).isin(lot_item_list)]

                        if not df_lot_bom.empty and bom_comp_code_col and bom_open_qty_col:
                            grp = [bom_comp_code_col]
                            if bom_comp_desc_col: grp.append(bom_comp_desc_col)
                            lot_comp = df_lot_bom.groupby(grp, dropna=False).agg(Open_Qty=(bom_open_qty_col, "sum")).reset_index()
                            lot_comp = lot_comp.merge(stock_agg, left_on=bom_comp_code_col, right_on="_stk_item", how="left")
                            lot_comp["_stk_total_onhand"] = lot_comp["_stk_total_onhand"].fillna(0)
                            lot_comp["Gap"] = lot_comp["_stk_total_onhand"] - lot_comp["Open_Qty"]
                            lot_comp.rename(columns={"_stk_total_onhand": "Stock"}, inplace=True)
                            lot_comp.drop(columns=["_stk_item"], errors="ignore", inplace=True)
                            lot_comp = lot_comp.sort_values("Gap", ascending=True)
                            st.dataframe(lot_comp, use_container_width=True, height=300)
                        else:
                            st.caption("No matching BOM data for this lot.")


# ═══════════════════════════════════════════════════
# TAB 4 – SHORTAGE FORECAST (Priority-based Stock Allocation)
# ═══════════════════════════════════════════════════
with tab_shortage:
    st.markdown('<div class="section-header">Priority-Based Material Shortage Forecast (Project + Lot)</div>', unsafe_allow_html=True)
    st.caption("Stock is allocated to projects in order of Assembly Start Date. Earlier projects get stock first; later projects see what remains.")

    row_lot_assignments = load_row_lot_assignments()
    saved_lots = load_saved_lots()

    if bom_comp_code_col and bom_open_qty_col:
        # ── Step 1: Build priority-ordered list of (row_key, lot_name, assembly_start_date) ──
        priority_rows = []
        for idx, row in df_fc.iterrows():
            rk = row["_row_key"]
            pname = str(row.get(fc_proj_name_col, "")).strip() if pd.notna(row.get(fc_proj_name_col)) else "Unknown"
            asm_date = row.get(fc_asm_start_col) if fc_asm_start_col else None
            bp = str(row.get(fc_build_period_col, "")).strip() if fc_build_period_col and pd.notna(row.get(fc_build_period_col)) else ""

            assigned_lots = [l for l in row_lot_assignments.get(rk, []) if l in saved_lots]
            if not assigned_lots:
                continue

            for lot_nm in assigned_lots:
                priority_rows.append({
                    "row_key": rk,
                    "project": pname,
                    "lot": lot_nm,
                    "build_period": bp,
                    "asm_date": asm_date,
                    "asm_date_sort": asm_date if pd.notna(asm_date) else pd.Timestamp("2099-12-31"),
                })

        if not priority_rows:
            st.info("ℹ️  No lots assigned to forecasting rows yet. Assign lots in the **Forecasting** tab first.")
        else:
            priority_df = pd.DataFrame(priority_rows).sort_values("asm_date_sort").reset_index(drop=True)

            # ── Step 2: For each project+lot in priority order, compute BOM demand and allocate stock ──
            # Track remaining stock globally
            remaining_stock = stock_agg.set_index("_stk_item")["_stk_total_onhand"].to_dict()

            all_shortage_rows = []

            for _, prow in priority_df.iterrows():
                lot_nm = prow["lot"]
                lot_info = saved_lots.get(lot_nm, {})
                if not lot_info:
                    continue

                # Get BOM for this lot
                lot_bom_project = lot_info.get("project", "")
                lot_wo_list = lot_info.get("work_orders", ["-- All --"])
                lot_item_list = lot_info.get("items", ["-- All --"])

                df_lot_bom = df_bom.copy()
                if lot_bom_project and lot_bom_project != "-- All --":
                    p_col = bom_proj_name_col or bom_proj_num_col
                    if p_col:
                        df_lot_bom = df_lot_bom[df_lot_bom[p_col] == lot_bom_project]
                if lot_wo_list != ["-- All --"] and bom_wo_col:
                    df_lot_bom = df_lot_bom[df_lot_bom[bom_wo_col].astype(str).isin(lot_wo_list)]
                if lot_item_list != ["-- All --"] and bom_item_col:
                    df_lot_bom = df_lot_bom[df_lot_bom[bom_item_col].astype(str).isin(lot_item_list)]

                if df_lot_bom.empty or not bom_comp_code_col:
                    continue

                # Aggregate demand per component
                grp = [bom_comp_code_col]
                if bom_comp_desc_col:
                    grp.append(bom_comp_desc_col)

                agg_dict = {bom_open_qty_col: "sum"}
                if bom_req_qty_col:
                    agg_dict[bom_req_qty_col] = "sum"

                lot_demand = df_lot_bom.groupby(grp, dropna=False).agg(agg_dict).reset_index()

                for _, comp_row in lot_demand.iterrows():
                    comp_code = comp_row[bom_comp_code_col]
                    comp_desc = comp_row.get(bom_comp_desc_col, "") if bom_comp_desc_col else ""
                    open_qty = comp_row[bom_open_qty_col]
                    req_qty = comp_row.get(bom_req_qty_col, open_qty) if bom_req_qty_col else open_qty

                    # Check available stock for this component
                    available = remaining_stock.get(comp_code, 0)
                    allocated = min(available, open_qty)
                    shortage = open_qty - allocated

                    # Deduct allocated stock
                    remaining_stock[comp_code] = available - allocated

                    risk = "🔴 Critical" if shortage > 50 else ("🟡 Warning" if shortage > 0 else "🟢 OK")

                    all_shortage_rows.append({
                        "Project": prow["project"],
                        "Lot": lot_nm,
                        "Build Period": prow["build_period"],
                        "Assembly Start": prow["asm_date"].strftime("%d-%b-%Y") if pd.notna(prow["asm_date"]) and isinstance(prow["asm_date"], pd.Timestamp) else "N/A",
                        "Component Code": comp_code,
                        "Component Desc": comp_desc,
                        "Required": int(req_qty),
                        "Open Qty": int(open_qty),
                        "Stock Available (Before)": int(available),
                        "Allocated": int(allocated),
                        "Shortage": int(shortage),
                        "Stock Remaining (After)": int(remaining_stock.get(comp_code, 0)),
                        "Risk": risk,
                        "_asm_sort": prow["asm_date_sort"],
                    })

            if all_shortage_rows:
                shortage_df = pd.DataFrame(all_shortage_rows)
                shortage_df = shortage_df.sort_values(["_asm_sort", "Project", "Lot", "Shortage"], ascending=[True, True, True, True])
                shortage_display = shortage_df.drop(columns=["_asm_sort"])

                # Filters
                f1, f2, f3 = st.columns(3)
                with f1:
                    risk_filter = st.multiselect("Risk Level", ["🔴 Critical", "🟡 Warning", "🟢 OK"],
                                                 default=["🔴 Critical", "🟡 Warning"], key="risk_flt")
                with f2:
                    proj_options = sorted(shortage_display["Project"].unique().tolist())
                    proj_flt = st.multiselect("Projects", proj_options, key="proj_flt")
                with f3:
                    lot_flt = st.multiselect("Lots", sorted(shortage_display["Lot"].unique().tolist()), key="lot_flt")

                filtered = shortage_display[shortage_display["Risk"].isin(risk_filter)].copy()
                if proj_flt:
                    filtered = filtered[filtered["Project"].isin(proj_flt)]
                if lot_flt:
                    filtered = filtered[filtered["Lot"].isin(lot_flt)]

                # Metrics
                mc1, mc2, mc3, mc4 = st.columns(4)
                with mc1:
                    render_metric("Critical", len(filtered[filtered["Risk"] == "🔴 Critical"]), "red")
                with mc2:
                    render_metric("Warning", len(filtered[filtered["Risk"] == "🟡 Warning"]), "amber")
                with mc3:
                    render_metric("OK", len(shortage_display[shortage_display["Risk"] == "🟢 OK"]), "emerald")
                with mc4:
                    render_metric("Total Shortage Qty", int(filtered["Shortage"].sum()), "purple")

                st.dataframe(filtered, use_container_width=True, height=500)

                # Project + Lot Risk Heatmap
                st.markdown('<div class="section-header">Project + Lot Risk Heatmap</div>', unsafe_allow_html=True)
                risk_summary = shortage_display.copy()
                risk_summary["Proj_Lot"] = risk_summary["Project"] + " → " + risk_summary["Lot"]
                risk_heatmap = risk_summary.groupby("Proj_Lot").agg(
                    Critical=("Risk", lambda x: (x == "🔴 Critical").sum()),
                    Warning=("Risk", lambda x: (x == "🟡 Warning").sum()),
                    OK=("Risk", lambda x: (x == "🟢 OK").sum()),
                ).reset_index().sort_values("Critical", ascending=False)

                fig_heat = go.Figure()
                fig_heat.add_trace(go.Bar(y=risk_heatmap["Proj_Lot"], x=risk_heatmap["Critical"], name="Critical", marker_color="#ef4444", orientation="h"))
                fig_heat.add_trace(go.Bar(y=risk_heatmap["Proj_Lot"], x=risk_heatmap["Warning"], name="Warning", marker_color="#f59e0b", orientation="h"))
                fig_heat.add_trace(go.Bar(y=risk_heatmap["Proj_Lot"], x=risk_heatmap["OK"], name="OK", marker_color="#10b981", orientation="h"))
                fig_heat.update_layout(barmode="stack", plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#94a3b8"), height=max(300, len(risk_heatmap) * 35),
                    margin=dict(l=10, r=10, t=10, b=30),
                    legend=dict(orientation="h", yanchor="bottom", y=-0.15))
                st.plotly_chart(fig_heat, use_container_width=True)

                st.markdown("---")
                st.download_button("⬇️ Download Priority Shortage Report (CSV)",
                    data=filtered.to_csv(index=False).encode("utf-8"),
                    file_name=f"priority_shortage_report_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv")
            else:
                st.info("No component data found for the assigned lots.")
    else:
        st.warning("Required columns not found for shortage analysis.")


# ═══════════════════════════════════════════════════
# TAB 5 – TIMELINE (with Project + Lot Schedule)
# ═══════════════════════════════════════════════════
with tab_timeline:
    st.markdown('<div class="section-header">Project Schedule Timeline</div>', unsafe_allow_html=True)

    if fc_proj_name_col:
        # ── Data labels toggle ──
        show_data_labels = st.checkbox("📌 Show date labels on chart", value=False, key="tl_data_labels")

        timeline_rows = []
        for _, row in df_fc.iterrows():
            pname = row.get(fc_proj_name_col, "Unknown")
            if pd.isna(pname): continue

            asm_s = row.get(fc_asm_start_col) if fc_asm_start_col else None
            asm_e = row.get(fc_asm_end_col) if fc_asm_end_col else None
            if pd.notna(asm_s) and pd.notna(asm_e):
                timeline_rows.append({"Project": str(pname), "Phase": "Assembly", "Start": asm_s, "End": asm_e})
            elif pd.notna(asm_s):
                timeline_rows.append({"Project": str(pname), "Phase": "Assembly", "Start": asm_s, "End": asm_s + timedelta(days=7)})

            fat_s = row.get(fc_fat_start_col) if fc_fat_start_col else None
            need_by = row.get(fc_need_by_col) if fc_need_by_col else None
            if pd.notna(fat_s):
                timeline_rows.append({"Project": str(pname), "Phase": "FAT", "Start": fat_s, "End": need_by if pd.notna(need_by) else fat_s + timedelta(days=5)})

            if pd.notna(need_by):
                timeline_rows.append({"Project": str(pname), "Phase": "Shipment", "Start": need_by, "End": need_by + timedelta(days=3)})

            mat = row.get(fc_mat_avail_col) if fc_mat_avail_col else None
            if pd.notna(mat):
                timeline_rows.append({"Project": str(pname), "Phase": "Material Avail", "Start": mat, "End": mat + timedelta(days=1)})

        if timeline_rows:
            tl_df = pd.DataFrame(timeline_rows)
            tl_df["Start"] = pd.to_datetime(tl_df["Start"], errors="coerce")
            tl_df["End"] = pd.to_datetime(tl_df["End"], errors="coerce")
            tl_df = tl_df.dropna(subset=["Start", "End"])

            fig_tl = px.timeline(tl_df, x_start="Start", x_end="End", y="Project", color="Phase",
                color_discrete_map={"Assembly": "#3b82f6", "FAT": "#8b5cf6", "Shipment": "#f59e0b", "Material Avail": "#10b981"})
            fig_tl.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#94a3b8"),
                height=max(400, len(tl_df["Project"].unique()) * 40), margin=dict(l=10, r=10, t=10, b=30),
                legend=dict(orientation="h", yanchor="bottom", y=-0.15), xaxis_title="", yaxis_title="")
            fig_tl.update_yaxes(autorange="reversed")

            # Add data labels if checkbox is checked
            if show_data_labels:
                for trace in fig_tl.data:
                    if hasattr(trace, 'base') and trace.base is not None:
                        start_labels = [pd.Timestamp(b).strftime("%d-%b") if pd.notna(b) else "" for b in trace.base]
                        end_dates = []
                        for b, x_val in zip(trace.base, trace.x):
                            try:
                                # x is duration in ms for timeline
                                end_dt = pd.Timestamp(b) + pd.Timedelta(milliseconds=x_val)
                                end_dates.append(end_dt.strftime("%d-%b"))
                            except Exception:
                                end_dates.append("")
                        text_labels = [f"{s} → {e}" for s, e in zip(start_labels, end_dates)]
                        trace.text = text_labels
                        trace.textposition = "inside"
                        trace.textfont = dict(size=9, color="white")

            st.plotly_chart(fig_tl, use_container_width=True)
        else:
            st.info("No date data available to build timeline.")

        # ─────────────────────────────────────────────────
        # NEW: Projects + Assigned Lots Schedule
        # ─────────────────────────────────────────────────
        st.markdown('<div class="section-header">📦 Projects + Assigned Lots Schedule (by Assembly Start Date)</div>', unsafe_allow_html=True)
        st.caption("Only projects with assigned lots are shown, sorted by Assembly Start Date.")

        row_lot_assignments = load_row_lot_assignments()
        saved_lots = load_saved_lots()

        lot_schedule_rows = []
        for idx, row in df_fc.iterrows():
            rk = row.get("_row_key", "")
            assigned = [l for l in row_lot_assignments.get(rk, []) if l in saved_lots]
            if not assigned:
                continue

            pname = str(row.get(fc_proj_name_col, "Unknown")).strip() if pd.notna(row.get(fc_proj_name_col)) else "Unknown"
            asm_start = row.get(fc_asm_start_col) if fc_asm_start_col else None
            asm_end = row.get(fc_asm_end_col) if fc_asm_end_col else None
            need_by = row.get(fc_need_by_col) if fc_need_by_col else None
            bp = str(row.get(fc_build_period_col, "")).strip() if fc_build_period_col and pd.notna(row.get(fc_build_period_col)) else ""
            status = str(row.get(fc_status_col, "")).strip() if fc_status_col and pd.notna(row.get(fc_status_col)) else ""
            cab_qty = row.get(fc_cab_qty_col, 0) if fc_cab_qty_col else 0

            for lot_nm in assigned:
                lot_info = saved_lots.get(lot_nm, {})
                lot_schedule_rows.append({
                    "Project": pname,
                    "Build Period": bp,
                    "Assembly Start": asm_start,
                    "Assembly End": asm_end,
                    "Need By Date": need_by,
                    "Status": status,
                    "Cabinets": int(cab_qty) if pd.notna(cab_qty) else 0,
                    "Assigned Lot": lot_nm,
                    "Lot BOM Project": lot_info.get("project", "N/A"),
                    "Lot Components": lot_info.get("component_count", 0),
                    "Lot Short": lot_info.get("short_components", 0),
                    "Lot Open Qty": lot_info.get("total_open_qty", 0),
                    "_asm_sort": asm_start if pd.notna(asm_start) else pd.Timestamp("2099-12-31"),
                })

        if lot_schedule_rows:
            lot_sched_df = pd.DataFrame(lot_schedule_rows)
            lot_sched_df = lot_sched_df.sort_values("_asm_sort").reset_index(drop=True)

            # Format dates for display
            lot_sched_display = lot_sched_df.drop(columns=["_asm_sort"]).copy()
            for dcol in ["Assembly Start", "Assembly End", "Need By Date"]:
                if dcol in lot_sched_display.columns:
                    lot_sched_display[dcol] = pd.to_datetime(lot_sched_display[dcol], errors="coerce").dt.strftime("%d-%b-%Y").fillna("")

            # Summary metrics
            lm1, lm2, lm3, lm4 = st.columns(4)
            with lm1:
                render_metric("Projects with Lots", lot_sched_display["Project"].nunique(), "blue")
            with lm2:
                render_metric("Total Lots Assigned", len(lot_sched_display), "purple")
            with lm3:
                render_metric("Total Short Components", int(lot_sched_display["Lot Short"].sum()), "red")
            with lm4:
                render_metric("Total Lot Open Qty", int(lot_sched_display["Lot Open Qty"].sum()), "amber")

            st.dataframe(lot_sched_display, use_container_width=True, height=400)

            # Gantt chart for lot schedule
            gantt_rows = []
            for _, lrow in lot_sched_df.iterrows():
                asm_s = lrow["Assembly Start"]
                asm_e = lrow["Assembly End"]
                if pd.notna(asm_s):
                    end = asm_e if pd.notna(asm_e) else (asm_s + timedelta(days=7))
                    gantt_rows.append({
                        "Label": f"{lrow['Project']} → {lrow['Assigned Lot']}",
                        "Phase": "Assembly",
                        "Start": asm_s,
                        "End": end,
                    })
                nbd = lrow["Need By Date"]
                if pd.notna(nbd):
                    gantt_rows.append({
                        "Label": f"{lrow['Project']} → {lrow['Assigned Lot']}",
                        "Phase": "Shipment",
                        "Start": nbd,
                        "End": nbd + timedelta(days=3),
                    })

            if gantt_rows:
                gantt_df = pd.DataFrame(gantt_rows)
                gantt_df["Start"] = pd.to_datetime(gantt_df["Start"], errors="coerce")
                gantt_df["End"] = pd.to_datetime(gantt_df["End"], errors="coerce")
                gantt_df = gantt_df.dropna(subset=["Start", "End"])

                fig_lot_tl = px.timeline(gantt_df, x_start="Start", x_end="End", y="Label", color="Phase",
                    color_discrete_map={"Assembly": "#3b82f6", "Shipment": "#f59e0b"})
                fig_lot_tl.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#94a3b8"),
                    height=max(350, len(gantt_df["Label"].unique()) * 45),
                    margin=dict(l=10, r=10, t=10, b=30),
                    legend=dict(orientation="h", yanchor="bottom", y=-0.15),
                    xaxis_title="", yaxis_title="")
                fig_lot_tl.update_yaxes(autorange="reversed")

                # Data labels on lot schedule gantt
                if show_data_labels:
                    for trace in fig_lot_tl.data:
                        if hasattr(trace, 'base') and trace.base is not None:
                            start_labels = [pd.Timestamp(b).strftime("%d-%b") if pd.notna(b) else "" for b in trace.base]
                            end_dates = []
                            for b, x_val in zip(trace.base, trace.x):
                                try:
                                    end_dt = pd.Timestamp(b) + pd.Timedelta(milliseconds=x_val)
                                    end_dates.append(end_dt.strftime("%d-%b"))
                                except Exception:
                                    end_dates.append("")
                            text_labels = [f"{s} → {e}" for s, e in zip(start_labels, end_dates)]
                            trace.text = text_labels
                            trace.textposition = "inside"
                            trace.textfont = dict(size=9, color="white")

                st.plotly_chart(fig_lot_tl, use_container_width=True)
        else:
            st.info("No lots assigned to any forecasting rows yet. Assign lots in the Forecasting tab.")

        # ── Upcoming Deadlines ──
        st.markdown('<div class="section-header">Upcoming Deadlines (Next 60 Days)</div>', unsafe_allow_html=True)
        today = pd.Timestamp.now().normalize()
        cutoff = today + timedelta(days=60)
        deadlines = []
        for _, row in df_fc.iterrows():
            pname = row.get(fc_proj_name_col, "Unknown")
            for col, label in [(fc_need_by_col, "Shipment"), (fc_fat_start_col, "FAT Start"), (fc_asm_start_col, "Assembly Start"), (fc_mat_avail_col, "Material Avail")]:
                if col:
                    dt = row.get(col)
                    if pd.notna(dt) and isinstance(dt, pd.Timestamp) and today <= dt <= cutoff:
                        deadlines.append({"Project": pname, "Milestone": label, "Date": dt, "Days Until": (dt - today).days})
        if deadlines:
            st.dataframe(pd.DataFrame(deadlines).sort_values("Days Until"), use_container_width=True, height=300)
        else:
            st.info("No upcoming deadlines within 60 days.")
    else:
        st.warning("Project Name column not found.")


# ═══════════════════════════════════════════════════
# TAB 6 – RAW DATA
# ═══════════════════════════════════════════════════
with tab_raw:
    st.markdown('<div class="section-header">Uploaded Raw Data Preview</div>', unsafe_allow_html=True)
    raw_tab1, raw_tab2, raw_tab3 = st.tabs(["Forecasting", "Open Orders BOM", "Stock"])
    with raw_tab1:
        st.markdown(f"**Rows:** {len(df_fc)}  |  **Columns:** {len(df_fc.columns)}")
        display_fc_raw = df_fc.drop(columns=["_row_key", "_row_label"], errors="ignore")
        st.dataframe(display_fc_raw, use_container_width=True, height=500)
    with raw_tab2:
        st.markdown(f"**Rows:** {len(df_bom)}  |  **Columns:** {len(df_bom.columns)}")
        st.dataframe(df_bom, use_container_width=True, height=500)
    with raw_tab3:
        st.markdown(f"**Rows:** {len(df_stock)}  |  **Columns:** {len(df_stock.columns)}")
        st.dataframe(df_stock, use_container_width=True, height=500)