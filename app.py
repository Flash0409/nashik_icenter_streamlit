from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Nashik iCenter Planning & Stock Analyzer", layout="wide")

ORDERBOOK_SHEET  = "OpenOrdersBOM"
OPEN_ORDER_SHEET = "Open Order"
STOCK_SHEET      = "Stock"

ORDERBOOK_ALIASES = {
    "project":          ["project name", "project"],
    "work_order":       ["order number", "work order", "workorder", "wo number", "wo"],
    "item":             ["component code", "component", "item number", "item no", "item"],
    "item_description": ["component desc", "component description", "item description", "item desc", "description"],
    "required_qty":     ["ordered quantity", "required qty", "quantity", "qty"],
    "open_qty":         ["open qty", "open quantity"],
}

FORECAST_ALIASES = {
    "project":        ["project name", "project"],
    "schedule":       ["build period", "ship period", "schedule", "period"],
    "no_of_cabinets": ["cabinets qty", "no of cabinets", "cabinets", "cab qty"],
    "work_order":     ["work order", "workorder", "wo"],
}

STOCK_ALIASES = {
    "item":        ["item number", "item", "component code", "component"],
    "description": ["component description", "item description", "description"],
    "on_hand_qty": ["on hand quantity", "on hand qty", "available stock", "stock", "quantity", "qty"],
}

OPEN_ORDER_ALIASES = {
    "item":     ["item number", "item", "component code", "component"],
    "open_qty": ["open quantity", "open qty", "qty"],
}


def normalise(val: str) -> str:
    return (
        str(val).strip().lower()
        .replace("\n", " ").replace("\r", " ")
        .replace("_", " ").replace("-", " ")
    )


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [
        str(c).strip().replace("\n", " ").replace("\r", " ")
        for c in df.columns
    ]
    return df


def find_col(columns, aliases):
    norm_map = {normalise(c): c for c in columns}
    for alias in aliases:
        hit = norm_map.get(normalise(alias))
        if hit:
            return hit
    for col in columns:
        cn = normalise(col)
        if any(normalise(a) in cn for a in aliases):
            return col
    return None


def map_cols(df_cols, aliases_dict):
    return {key: find_col(df_cols, aliases) for key, aliases in aliases_dict.items()}


def detect_header_row(file_bytes: bytes, sheet_name: str, terms, max_scan: int = 15) -> int:
    try:
        preview = pd.read_excel(
            BytesIO(file_bytes), sheet_name=sheet_name,
            header=None, nrows=max_scan, engine="openpyxl"
        )
    except Exception:
        return 0
    best_row, best_score = 0, -1
    for idx in range(len(preview)):
        row_text = " ".join(
            normalise(str(v)) for v in preview.iloc[idx]
            if pd.notna(v) and str(v) != "nan"
        )
        score = sum(1 for t in terms if normalise(t) in row_text)
        if score > best_score:
            best_score, best_row = score, idx
    return best_row


@st.cache_data(show_spinner=False)
def load_sheet(file_bytes: bytes, sheet_name: str, header_row: int = 0) -> pd.DataFrame:
    df = pd.read_excel(
        BytesIO(file_bytes), sheet_name=sheet_name,
        header=header_row, engine="openpyxl"
    )
    df = df.dropna(how="all").reset_index(drop=True)
    return clean_columns(df)


def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
              .str.replace("*", "", regex=False)
              .str.replace(",", "", regex=False)
              .str.strip(),
        errors="coerce"
    )


def status_tag(balance):
    if pd.isna(balance):
        return "Unknown"
    return "Covered" if balance >= 0 else "Short"


def get_projects(df, col):
    """Return sorted unique non-empty project names from a dataframe column."""
    bad = {"project name", "project", "nan", "none", ""}
    result = set()
    for v in df[col]:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            continue
        s = str(v).strip()
        if s.lower() not in bad:
            result.add(s)
    return sorted(result)


def main():
    st.title("Nashik iCenter — Planning & Stock Analyzer")
    st.caption("Upload Orderbook, Forecast, and Stock files to analyse component requirements vs availability.")

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("Upload Files")
        st.markdown("**1. Orderbook** (OpenOrdersBOM sheet)")
        orderbook_file = st.file_uploader("Orderbook (.xlsx/.xlsm)", type=["xlsx", "xlsm"], key="ob")
        st.markdown("**2. Forecast / Planning** (FY26 sheet)")
        forecast_file  = st.file_uploader("Forecast (.xlsx/.xlsm)",  type=["xlsx", "xlsm"], key="fc")
        st.markdown("**3. Stock** (Stock + Open Order sheets)")
        stock_file     = st.file_uploader("Stock (.xlsx/.xlsm)",     type=["xlsx", "xlsm"], key="st")

        st.divider()
        st.subheader("Sheet Names")
        for label, fobj in [("Orderbook", orderbook_file),
                             ("Forecast",  forecast_file),
                             ("Stock",     stock_file)]:
            if fobj:
                try:
                    names = pd.ExcelFile(BytesIO(fobj.getvalue()), engine="openpyxl").sheet_names
                    st.caption(f"{label}: {', '.join(names)}")
                except Exception:
                    pass

        ob_sh = st.text_input("OpenOrdersBOM sheet", value=ORDERBOOK_SHEET)
        fc_sh = st.text_input("Forecast sheet",      value="FY26")
        st_sh = st.text_input("Stock sheet",         value=STOCK_SHEET)

    missing_files = [n for n, f in [("Orderbook", orderbook_file),
                                     ("Forecast",  forecast_file),
                                     ("Stock",     stock_file)] if not f]
    if missing_files:
        st.info(f"Please upload: **{', '.join(missing_files)}**")
        return

    # ── Load sheets ───────────────────────────────────────────────────────────
    ob_bytes = orderbook_file.getvalue()
    try:
        ob_hdr = detect_header_row(ob_bytes, ob_sh,
            ["project name", "component", "component code", "component description", "ordered quantity"])
        ob_df = load_sheet(ob_bytes, ob_sh, ob_hdr)
    except Exception as exc:
        st.error(f"Cannot load '{ob_sh}': {exc}")
        try:
            st.info("Available sheets: " + str(pd.ExcelFile(BytesIO(ob_bytes), engine="openpyxl").sheet_names))
        except Exception: pass
        return

    fc_bytes = forecast_file.getvalue()
    try:
        fc_hdr = detect_header_row(fc_bytes, fc_sh,
            ["project name", "cabinets qty", "build period", "ship period", "sr no"])
        fc_df = load_sheet(fc_bytes, fc_sh, fc_hdr)
    except Exception as exc:
        st.error(f"Cannot load '{fc_sh}': {exc}")
        try:
            st.info("Available sheets: " + str(pd.ExcelFile(BytesIO(fc_bytes), engine="openpyxl").sheet_names))
        except Exception: pass
        return

    st_bytes = stock_file.getvalue()
    try:
        st_hdr = detect_header_row(st_bytes, st_sh,
            ["item number", "on hand quantity", "description"])
        stock_df = load_sheet(st_bytes, st_sh, st_hdr)
    except Exception as exc:
        st.error(f"Cannot load '{st_sh}': {exc}")
        try:
            st.info("Available sheets: " + str(pd.ExcelFile(BytesIO(st_bytes), engine="openpyxl").sheet_names))
        except Exception: pass
        return

    oo_df = pd.DataFrame()
    try:
        oo_hdr = detect_header_row(st_bytes, OPEN_ORDER_SHEET,
            ["item number", "open qty", "open quantity"])
        oo_df = load_sheet(st_bytes, OPEN_ORDER_SHEET, oo_hdr)
    except Exception:
        pass

    # ── Map columns ───────────────────────────────────────────────────────────
    ob_map = map_cols(ob_df.columns,    ORDERBOOK_ALIASES)
    fc_map = map_cols(fc_df.columns,    FORECAST_ALIASES)
    st_map = map_cols(stock_df.columns, STOCK_ALIASES)
    oo_map = map_cols(oo_df.columns,    OPEN_ORDER_ALIASES) if not oo_df.empty else {}

    # ── Validate critical columns ─────────────────────────────────────────────
    ob_miss = [k for k in ["project","work_order","item","item_description","required_qty"] if not ob_map.get(k)]
    fc_miss = [k for k in ["project","no_of_cabinets"] if not fc_map.get(k)]
    st_miss = [k for k in ["item","on_hand_qty"] if not st_map.get(k)]

    if ob_miss:
        st.error(f"OpenOrdersBOM missing columns for: {ob_miss}")
        st.info("Detected columns: " + str(list(ob_df.columns)))
        return
    if fc_miss:
        st.error(f"Forecast sheet missing columns for: {fc_miss}")
        st.info("Detected columns: " + str(list(fc_df.columns)))
        return
    if st_miss:
        st.error(f"Stock sheet missing columns for: {st_miss}")
        return

    ob_proj  = ob_map["project"];        ob_wo   = ob_map["work_order"]
    ob_item  = ob_map["item"];            ob_desc = ob_map["item_description"]
    ob_qty   = ob_map["required_qty"];    ob_open = ob_map.get("open_qty")

    fc_proj  = fc_map["project"];        fc_sched = fc_map.get("schedule")
    fc_cab   = fc_map["no_of_cabinets"]; fc_wo    = fc_map.get("work_order")

    st_item  = st_map["item"];            st_qty   = st_map["on_hand_qty"]

    # ── Pre-process ───────────────────────────────────────────────────────────
    ob_df[ob_proj] = ob_df[ob_proj].astype(str).str.strip()
    fc_df[fc_proj] = fc_df[fc_proj].astype(str).str.strip()
    fc_df[fc_cab]  = to_num(fc_df[fc_cab])

    bad_vals = {"project name", "project", "nan", "none", ""}
    fc_df = fc_df[~fc_df[fc_proj].str.lower().isin(bad_vals)].copy()
    ob_df = ob_df[~ob_df[ob_proj].str.lower().isin(bad_vals)].copy()

    ob_projects = get_projects(ob_df, ob_proj)
    fc_projects = get_projects(fc_df, fc_proj)

    # ══════════════════════════════════════════════════════════════════════════
    # PROJECT LIST VIEWER — always visible so user can inspect and manually map
    # ══════════════════════════════════════════════════════════════════════════
    with st.expander("📋 View All Projects from Both Files", expanded=True):
        st.markdown("Use this to find the correct matching names between files.")
        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown(f"### OpenOrdersBOM ({len(ob_projects)} projects)")
            ob_proj_df = pd.DataFrame({"#": range(1, len(ob_projects)+1),
                                        "Project Name (Orderbook)": ob_projects})
            st.dataframe(ob_proj_df, use_container_width=True, height=400)

        with col_b:
            st.markdown(f"### Forecast / {fc_sh} ({len(fc_projects)} projects)")
            fc_proj_df = pd.DataFrame({"#": range(1, len(fc_projects)+1),
                                        "Project Name (Forecast)": fc_projects})
            st.dataframe(fc_proj_df, use_container_width=True, height=400)

        # Download both lists as Excel for easy comparison
        list_buf = BytesIO()
        with pd.ExcelWriter(list_buf, engine="openpyxl") as writer:
            ob_proj_df.to_excel(writer, sheet_name="Orderbook Projects",  index=False)
            fc_proj_df.to_excel(writer, sheet_name="Forecast Projects",   index=False)
        list_buf.seek(0)
        st.download_button(
            "⬇️ Download Both Project Lists (.xlsx)",
            data=list_buf,
            file_name="Project_Lists.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.divider()

    # ══════════════════════════════════════════════════════════════════════════
    # MANUAL PROJECT MAPPING
    # ══════════════════════════════════════════════════════════════════════════
    st.subheader("🔗 Manual Project Mapping")
    st.caption("Pick the Orderbook project and the corresponding Forecast project, then proceed to analysis.")

    col1, col2 = st.columns(2)
    with col1:
        sel_ob_proj = st.selectbox("Orderbook project", ob_projects, key="ob_proj")
    with col2:
        # Pre-select forecast project if exact match exists, else first item
        fc_default_idx = fc_projects.index(sel_ob_proj) if sel_ob_proj in fc_projects else 0
        sel_fc_proj = st.selectbox("Matching Forecast project", fc_projects,
                                    index=fc_default_idx, key="fc_proj")

    if not sel_ob_proj or not sel_fc_proj:
        return

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 2 — Schedule
    # ══════════════════════════════════════════════════════════════════════════
    st.subheader("📅 Step 2 — Select Schedule Period(s)")
    proj_fc = fc_df[fc_df[fc_proj] == sel_fc_proj].copy()

    if fc_sched and fc_sched in proj_fc.columns:
        sched_opts = sorted(proj_fc[fc_sched].dropna().astype(str).str.strip().unique())
        sel_scheds = st.multiselect("Build / Ship Period", sched_opts, default=sched_opts, key="sched")
        scoped_fc  = proj_fc[proj_fc[fc_sched].astype(str).str.strip().isin(sel_scheds)].copy() \
                     if sel_scheds else proj_fc.iloc[0:0].copy()
    else:
        sel_scheds = []
        scoped_fc  = proj_fc.copy()

    planned_cabs = scoped_fc[fc_cab].fillna(0).sum()

    m1, m2, m3 = st.columns(3)
    m1.metric("Planned Cabinets", f"{planned_cabs:,.0f}")
    m2.metric("Schedule Rows",    f"{len(scoped_fc):,}")
    m3.metric("Qty Multiplier",   f"{planned_cabs:,.0f}")

    fc_show = [c for c in [fc_proj, fc_sched, fc_wo, fc_cab] if c and c in scoped_fc.columns]
    fc_show = list(dict.fromkeys(fc_show))
    with st.expander("Forecast / Planning Schedule", expanded=False):
        st.dataframe(scoped_fc[fc_show].reset_index(drop=True), use_container_width=True)

    if planned_cabs == 0:
        st.warning("No cabinets planned for the selected schedule(s). Please select at least one period.")
        return

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 3 — Work Orders
    # ══════════════════════════════════════════════════════════════════════════
    st.subheader("🔧 Step 3 — Select Work Orders")
    proj_ob = ob_df[ob_df[ob_proj] == sel_ob_proj].copy()
    wo_opts = sorted(proj_ob[ob_wo].dropna().astype(str).str.strip().unique())
    sel_wos = st.multiselect("Work Orders (leave empty = include all)", wo_opts, default=wo_opts, key="wo")
    scoped_ob = proj_ob[proj_ob[ob_wo].astype(str).str.strip().isin(sel_wos)].copy() \
                if sel_wos else proj_ob.copy()

    if scoped_ob.empty:
        st.info("No BOM rows found for the selected work orders.")
        return

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 4 — Component Analysis
    # ══════════════════════════════════════════════════════════════════════════
    st.subheader("📦 Step 4 — Component Requirement vs Stock")

    # Build stock lookup
    stock_df[st_item] = stock_df[st_item].astype(str).str.strip().str.upper()
    stock_df[st_qty]  = to_num(stock_df[st_qty]).fillna(0)
    stock_lookup = stock_df.groupby(st_item)[st_qty].sum().to_dict()

    # Build open order lookup
    oo_lookup = {}
    if not oo_df.empty and oo_map.get("item") and oo_map.get("open_qty"):
        oo_df[oo_map["item"]]     = oo_df[oo_map["item"]].astype(str).str.strip().str.upper()
        oo_df[oo_map["open_qty"]] = to_num(oo_df[oo_map["open_qty"]]).fillna(0)
        oo_lookup = oo_df.groupby(oo_map["item"])[oo_map["open_qty"]].sum().to_dict()

    scoped_ob[ob_qty] = to_num(scoped_ob[ob_qty]).fillna(0)
    if ob_open and ob_open in scoped_ob.columns:
        scoped_ob[ob_open] = to_num(scoped_ob[ob_open]).fillna(0)
        agg = (
            scoped_ob
            .groupby([ob_item, ob_desc], dropna=False, as_index=False)
            .agg(Req_Per_Cabinet=(ob_qty, "sum"), Open_Qty=(ob_open, "sum"))
        )
    else:
        agg = (
            scoped_ob
            .groupby([ob_item, ob_desc], dropna=False, as_index=False)
            .agg(Req_Per_Cabinet=(ob_qty, "sum"))
        )
    
    agg.rename(columns={ob_item: "Component Code", ob_desc: "Component Description"}, inplace=True)

    agg["Total Required"]    = agg["Req_Per_Cabinet"] * planned_cabs
    item_key                  = agg["Component Code"].astype(str).str.strip().str.upper()
    agg["Available Stock"]   = item_key.map(stock_lookup).fillna(0)
    
    # Use Open Qty from OpenOrdersBOM if available, otherwise use Open Order sheet
    if ob_open and "Open_Qty" in agg.columns:
        agg["Open Qty"] = agg["Open_Qty"]
    else:
        agg["Open Qty"] = item_key.map(oo_lookup).fillna(0)
    
    # Status based only on Available Stock + Open Qty vs Total Required
    agg["Status"] = ((agg["Available Stock"] + agg["Open Qty"]) >= agg["Total Required"]).apply(
        lambda x: "Covered" if x else "Short"
    )

    display_df = agg[[
        "Component Code", "Component Description",
        "Available Stock", "Open Qty",
        "Status"
    ]].reset_index(drop=True)

    def highlight(row):
        colour = "#ffe0e0" if row["Status"] == "Short" else "#e0ffe8"
        return ["background-color: " + colour if col == "Status" else "" for col in row.index]

    st.dataframe(display_df.style.apply(highlight, axis=1), use_container_width=True)

    total = len(display_df)
    short = (agg["Status"] == "Short").sum()
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Components", total)
    c2.metric("✅ Covered",        total - short)
    c3.metric("🔴 Short",          short)

    # ── Export ────────────────────────────────────────────────────────────────
    st.divider()
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        display_df.to_excel(writer, sheet_name="Component Analysis", index=False)
        scoped_fc[fc_show].reset_index(drop=True).to_excel(writer, sheet_name="Planning Scope", index=False)
    buf.seek(0)

    st.download_button(
        "⬇️ Download Analysis (.xlsx)",
        data=buf,
        file_name=f"StockAnalysis_{sel_ob_proj.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    main()