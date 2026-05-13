import streamlit as st
import pandas as pd
from parser import (
    extract,
    to_excel_bytes,
    grouped_trades,
    grouped_positions_by_ref_month,
    grouped_positions_custom,
    merge_extracted_tables,
    statement_dates_by_account,
    realized_pnl_summary,
)

st.set_page_config(page_title="StoneX Statement Trade Extractor", layout="wide")
st.title("StoneX Statement Trade Extractor")
st.caption("Upload one or more StoneX statement PDFs, merge trades/positions, review grouped summaries, and download Excel.")

with st.sidebar:
    st.header("Options")
    include_open_positions = st.checkbox("Include open positions", value=True)
    show_source_lines = st.checkbox("Show source_line columns", value=False)
    st.markdown("---")
    st.subheader("Multi-PDF merge")
    position_mode = st.radio(
        "When multiple PDFs are uploaded, how should open positions be merged?",
        options=["latest_by_account", "append_all"],
        index=0,
        format_func=lambda x: "Keep latest position snapshot per account" if x == "latest_by_account" else "Append all position snapshots",
        help="Positions are snapshots. Use latest_by_account to avoid double-counting if you upload multiple statement dates for the same account.",
    )
    st.markdown("---")
    st.subheader("Column display")
    st.caption("After upload, use the column selector above each table to customize what is shown.")
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
    selected = st.multiselect(
        f"Columns for {label}",
        options=all_cols,
        default=defaults,
        key=f"cols_{label.replace(' ', '_').replace('&', 'and')}",
    )
    st.dataframe(df[selected] if selected else df, use_container_width=True)

if uploaded_files:
    extracted_tables = []

    for idx, uploaded in enumerate(uploaded_files, start=1):
        pdf_bytes = uploaded.read()
        with st.spinner(f"Extracting {uploaded.name}..."):
            tables = extract(pdf_bytes, include_open_positions=include_open_positions)
            tables = add_source_pdf(tables, uploaded.name, idx)
            extracted_tables.append(tables)

    merged_tables = merge_extracted_tables(extracted_tables, position_mode=position_mode)
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
        "Grouped Trades",
        "Grouped Positions",
        "Realized PNL",
        "Statement Dates",
        "Executed Trades",
        "Open Positions",
        "All Sheets",
        "Per-PDF Audit",
    ])

    with tabs[0]:
        display_custom_table("Grouped Trades", grouped_trades(merged_tables))
    with tabs[1]:
        st.caption("Select the fields below to dynamically group open positions. Example: select only Product to group all rows by product.")
        group_by_cols = st.multiselect(
            "Group rows by",
            options=[
                "source_pdf",
                "statement_date",
                "account_number",
                "product",
                "exchange",
                "product_name",
                "option_type",
                "ref_month",
                "strike",
                "unit",
                "currency",
                "contract_month",
                "contract_year",
                "delivery_date",
                "settlement_date",
                "source_section",
            ],
            default=["statement_date", "account_number", "product", "option_type", "ref_month", "strike", "unit"],
            key="grouped_positions_group_by_cols",
            help="These selected columns control the grouping, not just the display. Remove strike to aggregate across strikes; remove account_number to aggregate across accounts.",
        )
        grouped_pos_df = grouped_positions_custom(merged_tables, group_by_cols)
        st.write("Grouping by:", ", ".join(group_by_cols) if group_by_cols else "Default grouping")
        display_custom_table(
            "Grouped Positions",
            grouped_pos_df,
            default_columns=group_by_cols + ["long_qty", "short_qty", "net_qty", "position_rows", "market_value", "avg_trade_price"],
            default_only=True,
        )
    with tabs[2]:
        st.caption("Absolute-basic realized PNL view.")
        pnl_df = realized_pnl_summary(merged_tables)
        if not pnl_df.empty and "realized_pnl" in pnl_df.columns:
            st.metric("Total realized PNL", f"{pd.to_numeric(pnl_df['realized_pnl'], errors='coerce').sum():,.2f}")
        display_custom_table(
            "Realized PNL",
            pnl_df,
            default_columns=["statement_date", "account_number", "trade_date", "trade_id", "contract_description", "quantity", "trade_price", "cash_flow", "realized_pnl"],
            default_only=True,
        )
    with tabs[3]:
        st.caption("Use this sheet to confirm which statement date was used for each account/source file.")
        display_custom_table("Statement Dates", statement_dates_by_account(merged_tables))
    with tabs[4]:
        display_custom_table("Executed Trades", merged_tables.get("Executed Trades", pd.DataFrame()))
    with tabs[5]:
        display_custom_table(
            "Open Positions",
            merged_tables.get("Open Positions", pd.DataFrame()),
            default_columns=[
                "card", "account_type", "quantity", "contract_month", "contract_year",
                "delivery_date", "contract_description", "price", "trade_price",
                "start_date", "end_date", "settlement_date", "ref_month",
                "currency", "market_value", "market_value_signed",
                "statement_date", "account_number", "trade_date_iso", "source_section"
            ],
            default_only=True,
        )
    with tabs[6]:
        for name, df in merged_tables.items():
            with st.expander(f"{name} ({len(df)} rows)"):
                display_custom_table(f"All Sheets {name}", df)
    with tabs[7]:
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
        - Open positions are snapshots. The default keeps only the latest statement date per account.
        - Use **Append all position snapshots** only when you want a time-series/audit file.
        """
    )
