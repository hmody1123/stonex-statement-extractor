import streamlit as st
import pandas as pd
from parser import (
    extract,
    to_excel_bytes,
    grouped_positions_by_ref_month,
    grouped_positions_custom,
    grouped_positions_standard_view,
    grouped_positions_product_month_auto_standard_view,
    open_positions_standard_view,
    STANDARD_POSITION_COLUMNS,
    OPEN_POSITION_COLUMNS,
    GROUPED_POSITION_COLUMNS,
    merge_extracted_tables,
    statement_dates_by_account,
    realized_pnl_summary,
)

st.set_page_config(page_title="StoneX Statement Trade Extractor", layout="wide")
st.title("StoneX Statement Trade Extractor")
st.caption("Upload one or more StoneX statement PDFs, merge trades/positions, review grouped summaries, and download Excel. Version: v41 realized PNL closed positions detail.")

with st.sidebar:
    st.header("Options")
    include_open_positions = st.checkbox("Include open positions", value=True)
    show_source_lines = st.checkbox("Show source_line columns", value=False)
    st.markdown("---")
    st.subheader("Multi-PDF merge")
    st.caption("Open positions are always appended from all uploaded PDFs. No latest-snapshot filtering is applied.")
    st.markdown("---")
    st.subheader("Column display")
    st.caption("After upload, use the Customize table button above each table to show/hide columns.")
    st.markdown("---")
    st.caption("Tip: if uploading several month-end statements, keep latest snapshot per account to avoid double-counting open positions.")

uploaded_files = st.file_uploader("Upload one or more statement PDFs", type=["pdf"], accept_multiple_files=True)

def add_source_pdf(tables, pdf_name, pdf_index):
    for name, df in list(tables.items()):
        if isinstance(df, pd.DataFrame) and not df.empty:
            tables[name] = df.copy()
            tables[name]["source_pdf"] = pdf_name
            tables[name]["source_pdf_index"] = pdf_index
    return tables

def clean_for_display(df):
    if df is None:
        return pd.DataFrame()
    out = df.copy()
    if not show_source_lines and "source_line" in out.columns:
        out = out.drop(columns=["source_line"])
    return out

def _safe_key(text):
    return "".join(ch if ch.isalnum() else "_" for ch in str(text))

def display_custom_table(label, df, default_columns=None, default_only=False):
    df = clean_for_display(df)
    if df.empty:
        st.dataframe(df, use_container_width=True)
        return

    all_cols = list(df.columns)
    if default_columns:
        defaults = [c for c in default_columns if c in all_cols]
        if not default_only:
            defaults += [c for c in all_cols if c not in defaults]
    else:
        defaults = all_cols

    table_key = _safe_key(label)
    selected_key = f"selected_cols_{table_key}"

    # Initialize selected columns once per table. After that, users control the selection with toggles.
    if selected_key not in st.session_state:
        st.session_state[selected_key] = defaults

    selected_existing = [c for c in st.session_state[selected_key] if c in all_cols]
    st.session_state[selected_key] = selected_existing or defaults

    top_left, top_right = st.columns([0.72, 0.28])
    with top_left:
        st.caption(f"Showing {len(df):,} rows and {len(st.session_state[selected_key])}/{len(all_cols)} columns for {label}")
    with top_right:
        # Popover gives a compact column-customization panel similar to the screenshot.
        with st.popover("Customize table", use_container_width=True):
            st.markdown("**Show / hide columns**")
            b1, b2 = st.columns(2)
            if b1.button("Select all", key=f"select_all_{table_key}"):
                st.session_state[selected_key] = all_cols.copy()
                st.rerun()
            if b2.button("Reset", key=f"reset_{table_key}"):
                st.session_state[selected_key] = defaults.copy()
                st.rerun()

            st.divider()
            selected_set = set(st.session_state[selected_key])
            new_selection = []
            for col_name in all_cols:
                checked = st.checkbox(
                    col_name,
                    value=col_name in selected_set,
                    key=f"toggle_{table_key}_{_safe_key(col_name)}",
                )
                if checked:
                    new_selection.append(col_name)
            st.session_state[selected_key] = new_selection

    selected = [c for c in st.session_state[selected_key] if c in all_cols]
    if not selected:
        st.warning("No columns selected. Use Customize table to select columns.")
        return

    display_df = df[selected]

    def _quantity_color(value):
        try:
            numeric = float(str(value).replace(",", ""))
        except Exception:
            return ""
        if numeric < 0:
            return "color: red; font-weight: 600"
        if numeric > 0:
            return "color: green; font-weight: 600"
        return ""

    quantity_cols = [c for c in ["Quantity", "Net Quantity"] if c in display_df.columns]
    if quantity_cols:
        styled = display_df.style.map(_quantity_color, subset=quantity_cols)
        st.dataframe(styled, use_container_width=True)
    else:
        st.dataframe(display_df, use_container_width=True)

if uploaded_files:
    extracted_tables = []

    for idx, uploaded in enumerate(uploaded_files, start=1):
        pdf_bytes = uploaded.read()
        with st.spinner(f"Extracting {uploaded.name}..."):
            tables = extract(pdf_bytes, include_open_positions=include_open_positions)
            tables = add_source_pdf(tables, uploaded.name, idx)
            extracted_tables.append(tables)

    merged_tables = merge_extracted_tables(extracted_tables)
    merged_excel = to_excel_bytes(merged_tables)

    st.subheader("Merged output")
    counts = {name: len(df) for name, df in merged_tables.items() if name not in ["Summary"]}
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("PDFs", len(uploaded_files))
    c2.metric("Executed trades", counts.get("Executed Trades", 0))
    c3.metric("P&S rows", counts.get("Purchase & Sale", 0))
    c4.metric("Open positions", counts.get("Open Positions", 0))
    c5.metric("Exceptions", counts.get("Exceptions", 0))

    st.download_button(
        "Download merged Excel",
        data=merged_excel,
        file_name="stonex_merged_positions_trades.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_merged",
    )

    tabs = st.tabs([
        "Grouped Positions",
        "Realized PNL",
        "Statement Dates",
        "Executed Trades",
        "Open Positions",
        "All Sheets",
        "Per-PDF Audit",
    ])

    with tabs[0]:
        st.caption("Choose a position view. The selected view controls how raw open-position rows are aggregated.")
        grouping_presets = {
            "Product": {
                "description": "One row per product across all accounts and contract months.",
                "mode": "custom",
                "group_cols": ["product"],
            },
            "Product + Contract Month/Year": {
                "description": "Automatic risk view: futures group by product + contract month/year; options group by product + contract month/year + call/put + strike price.",
                "mode": "auto_futures_options",
                "group_cols": None,
            },
            "Account Grouping": {
                "description": "Account risk view: one row per account, product, and contract month/year.",
                "mode": "custom",
                "group_cols": ["account_number", "product", "ref_month"],
            },
        }
        selected_preset = st.radio(
            "Grouped Positions view",
            options=list(grouping_presets.keys()),
            horizontal=True,
            key="grouped_positions_preset_view",
        )
        preset = grouping_presets[selected_preset]
        if preset["mode"] == "auto_futures_options":
            grouped_pos_df = grouped_positions_product_month_auto_standard_view(merged_tables)
            st.write("Grouping by:", "Product + Contract Month/Year; options additionally by Call/Put + Strike Price")
        else:
            grouped_pos_df = grouped_positions_standard_view(merged_tables, preset["group_cols"])
            st.write("Grouping by:", selected_preset)
        st.caption(preset["description"] + " Weighted Avg Fill Price is calculated using absolute Net Quantity as the weight.")

        # Grouped Positions should stay risk-focused. Contract Description is intentionally
        # hidden here; use Open Positions for the exact PDF contract description.
        grouped_pos_df = grouped_pos_df.drop(columns=["Contract Description", "Full Name"], errors="ignore")
        grouped_default_columns = [c for c in GROUPED_POSITION_COLUMNS if c in grouped_pos_df.columns]

        display_custom_table(
            "Grouped Positions",
            grouped_pos_df,
            default_columns=grouped_default_columns,
            default_only=True,
        )
    with tabs[1]:
        st.caption("Realized PNL includes statement summary, closed-position GROSS PROFIT OR LOSS rows, and P&S trade audit detail when available.")
        pnl_df = realized_pnl_summary(merged_tables)
        if not pnl_df.empty and "realized_pnl" in pnl_df.columns:
            metric_df = pnl_df[pnl_df.get("pnl_view", "") == "Summary"] if "pnl_view" in pnl_df.columns else pnl_df
            st.metric("MTD realized PNL", f"{pd.to_numeric(metric_df['realized_pnl'], errors='coerce').sum():,.2f}")
        display_custom_table(
            "Realized PNL",
            pnl_df,
            default_columns=["pnl_view", "statement_date", "account_number", "currency", "realized_pnl", "mtd_realized_pnl", "ytd_realized_pnl", "close_date", "product", "exchange", "ref_month", "contract_description", "long", "short", "quantity", "price", "close_price", "trade_id", "source_sheet"],
            default_only=False,
        )
    with tabs[2]:
        st.caption("Use this sheet to confirm which statement date was used for each account/source file.")
        display_custom_table("Statement Dates", statement_dates_by_account(merged_tables))
    with tabs[3]:
        display_custom_table("Executed Trades", merged_tables.get("Executed Trades", pd.DataFrame()))
    with tabs[4]:
        display_custom_table(
            "Open Positions",
            open_positions_standard_view(merged_tables),
            default_columns=OPEN_POSITION_COLUMNS,
            default_only=True,
        )
    with tabs[5]:
        for name, df in merged_tables.items():
            with st.expander(f"{name} ({len(df)} rows)"):
                display_custom_table(f"All Sheets {name}", df)
    with tabs[6]:
        for idx, (uploaded, tables) in enumerate(zip(uploaded_files, extracted_tables), start=1):
            with st.expander(f"{idx}. {uploaded.name}"):
                p_counts = {name: len(df) for name, df in tables.items() if name != "Summary"}
                st.write(p_counts)
                st.download_button(
                    f"Download Excel for {uploaded.name}",
                    data=to_excel_bytes(tables),
                    file_name=uploaded.name.rsplit(".", 1)[0] + "_trades.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{idx}_{uploaded.name}",
                )
else:
    st.info("Upload one or more PDFs to begin.")
    st.markdown(
        """
        **Multi-PDF behavior:**
        - Trades are appended and de-duplicated when possible.
        - Open positions are always appended from all uploaded PDFs.
        - No latest-snapshot filtering is applied.
        """
    )
