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
    prepare_grouped_positions_display,
)

st.set_page_config(page_title="StoneX Statement Trade Extractor", layout="wide")
st.title("StoneX Statement Trade Extractor")
st.caption("Upload one or more StoneX statement PDFs, merge trades/positions, review grouped summaries, and download Excel. Version: v67 OTC trigger/ref/original quantity mapping.")

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
    # Settlement Date was removed from user-facing views. Keep raw parser fields in
    # downloaded Excel, but hide them in the Streamlit UI.
    out = out.drop(columns=["Settlement Date", "settlement_date"], errors="ignore")
    if not show_source_lines and "source_line" in out.columns:
        out = out.drop(columns=["source_line"])
    return out


def _display_series_nonblank(series):
    if series is None:
        return pd.Series(dtype=bool)
    text = series.astype(str).str.strip().str.lower()
    return series.notna() & ~text.isin(["", "none", "nan", "nat", "unknown", "other", "multiple"])


def _open_positions_default_columns(df):
    defaults = [c for c in OPEN_POSITION_COLUMNS if c in df.columns]
    # Keep expiryDate available in Customize table even when blank, but do not
    # show an all-blank expiryDate column by default.
    if "expiryDate" in df.columns and not _display_series_nonblank(df["expiryDate"]).any():
        defaults = [c for c in defaults if c != "expiryDate"]
    return defaults


OPTION_DISPLAY_COLUMNS = ["Call/Put", "strikePrice"]


def _has_option_positions(df):
    """Return True when the uploaded statement has real option rows.

    Non-option rows usually carry blank/None Call/Put and strikePrice values.
    When neither field is populated anywhere, hide those option-only columns from
    Open Positions and Grouped Positions so futures/FX-only PDFs stay cleaner.
    """
    if df is None or df.empty:
        return False

    call_put_cols = [c for c in ["Call/Put", "option_type", "C/P"] if c in df.columns]
    for col in call_put_cols:
        text = df[col].astype(str).str.strip().str.upper()
        if text.isin(["CALL", "PUT", "C", "P"]).any():
            return True
        # Keep a fallback for parser formats that already normalize to words but
        # may use mixed case or unexpected spacing.
        if _display_series_nonblank(df[col]).any() and not text.isin(["NONE", "NAN", "NAT", "UNKNOWN", "OTHER", "MULTIPLE", ""]).all():
            return True

    # strikePrice can also be used for OTC accumulator strike levels, so a
    # populated strike alone should not make the statement an options statement.
    for desc_col in ["Contract Description", "contract_description"]:
        if desc_col in df.columns:
            if df[desc_col].astype(str).str.upper().str.contains(r"\b(?:CALL|PUT|OPTION)\b", regex=True, na=False).any():
                return True

    return False


def _drop_option_columns_when_no_options(df, has_options):
    """Remove option-only columns when no options exist.

    strikePrice is also used for OTC accumulator strike levels, so keep it when
    it has values even if the PDF has no listed options.
    """
    if df is None or has_options:
        return df
    drop_cols = ["Call/Put"]
    if "strikePrice" in df.columns and not _display_series_nonblank(df["strikePrice"]).any():
        drop_cols.append("strikePrice")
    return df.drop(columns=drop_cols, errors="ignore")

def _safe_key(text):
    return "".join(ch if ch.isalnum() else "_" for ch in str(text))

def display_custom_table(label, df, default_columns=None, default_only=False, selectable=False, selection_key=None, table_key_override=None):
    df = clean_for_display(df)
    if df.empty:
        st.dataframe(df, use_container_width=True)
        return None

    all_cols = list(df.columns)
    if default_columns:
        defaults = [c for c in default_columns if c in all_cols]
        if not default_only:
            defaults += [c for c in all_cols if c not in defaults]
    else:
        defaults = all_cols

    table_key = _safe_key(table_key_override or label)
    selected_key = f"selected_cols_{table_key}"

    # Initialize selected columns once per table. If defaults change between app
    # versions or statement formats, reset to the new defaults so newly important
    # columns like expiryDate are not accidentally hidden by stale browser state.
    defaults_signature_key = f"{selected_key}_defaults_signature"
    defaults_signature = tuple(defaults)
    if selected_key not in st.session_state or st.session_state.get(defaults_signature_key) != defaults_signature:
        st.session_state[selected_key] = defaults.copy()
        st.session_state[defaults_signature_key] = defaults_signature

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
        return None

    display_df = df[selected]

    if selectable:
        # Streamlit dataframe row selection powers the grouped-position drill-down.
        # Selection row indexes refer to the displayed dataframe, which is built from df
        # without reordering, so they can be used to fetch the selected source row.
        return st.dataframe(
            display_df,
            use_container_width=True,
            on_select="rerun",
            selection_mode="single-row",
            key=selection_key or f"df_select_{table_key}",
        )

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
    return None


def _event_selected_rows(event):
    """Return selected row indexes from Streamlit's dataframe selection event."""
    if event is None:
        return []
    try:
        return list(event.selection.rows)
    except Exception:
        pass
    try:
        return list(event.get("selection", {}).get("rows", []))
    except Exception:
        return []


def _is_blank_drill_value(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    text = str(value).strip()
    return text == "" or text.lower() in {"none", "nan", "nat", "multiple", "other"}


def _blank_series_mask(series):
    text = series.astype(str).str.strip().str.lower()
    return series.isna() | text.isin(["", "none", "nan", "nat", "multiple", "other"])


def _apply_text_filter(mask, df, column, value):
    if column not in df.columns or _is_blank_drill_value(value):
        return mask
    lhs = df[column].astype(str).str.strip().str.upper()
    rhs = str(value).strip().upper()
    return mask & (lhs == rhs)


def _apply_numeric_or_text_filter(mask, df, column, value):
    if column not in df.columns or _is_blank_drill_value(value):
        return mask
    target_num = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    series_num = pd.to_numeric(df[column], errors="coerce")
    if pd.notna(target_num):
        return mask & (series_num == target_num)
    return _apply_text_filter(mask, df, column, value)


def _apply_fx_drilldown_filters(mask, df, row):
    """When an FX grouped row is selected, keep drill-down aligned to FX keys."""
    # FX grouped rows may contain aggregated currency amounts and/or hidden trade prices.
    # Drill-down should therefore filter by stable identity fields only: product, currencies,
    # and value/settlement date when that date is present in the selected group.
    for col in ["CCY 1", "CCY 2"]:
        mask = _apply_text_filter(mask, df, col, row.get(col))
    if not _is_blank_drill_value(row.get("expiryDate")) and "expiryDate" in df.columns:
        mask = _apply_text_filter(mask, df, "expiryDate", row.get("expiryDate"))
    return mask




def _apply_contract_or_expiry_drill_filter(mask, df, row):
    """Filter by expiryDate when the grouped row has one; otherwise by Contract Month/Year."""
    expiry = row.get("expiryDate") if hasattr(row, "get") else None
    if not _is_blank_drill_value(expiry) and "expiryDate" in df.columns:
        return _apply_text_filter(mask, df, "expiryDate", expiry)
    return _apply_text_filter(mask, df, "Contract Month/Year", row.get("Contract Month/Year"))

def _open_positions_for_group(open_df, selected_group_row, selected_preset):
    """Filter open-position rows to the rows that make up a selected grouped row."""
    if open_df is None or open_df.empty or selected_group_row is None:
        return pd.DataFrame()

    df = open_df.copy()
    row = selected_group_row
    mask = pd.Series(True, index=df.index)

    # Always anchor drill-down on Product when available.
    mask = _apply_text_filter(mask, df, "Product", row.get("Product"))
    mask = _apply_fx_drilldown_filters(mask, df, row)

    if selected_preset == "Product":
        pass
    elif selected_preset == "Account Grouping":
        mask = _apply_text_filter(mask, df, "Trigger/Barrier", row.get("Trigger/Barrier"))
        mask = _apply_text_filter(mask, df, "Account Number", row.get("Account Number"))
        mask = _apply_contract_or_expiry_drill_filter(mask, df, row)
        # Account Grouping keeps option detail. Futures/non-options have blank option keys,
        # while options filter by Call/Put + strike.
        call_put = row.get("Call/Put")
        strike = row.get("strikePrice")
        has_option_key = (not _is_blank_drill_value(call_put)) or (not _is_blank_drill_value(strike))
        if has_option_key:
            mask = _apply_text_filter(mask, df, "Call/Put", call_put)
            mask = _apply_numeric_or_text_filter(mask, df, "strikePrice", strike)
        else:
            if "Call/Put" in df.columns:
                mask = mask & _blank_series_mask(df["Call/Put"])
            if "strikePrice" in df.columns:
                mask = mask & _blank_series_mask(df["strikePrice"])
    else:
        # Product + Contract Month/Year view. If expiryDate is present, the group is expiry-aware;
        # otherwise it falls back to Product + Contract Month/Year.
        mask = _apply_contract_or_expiry_drill_filter(mask, df, row)
        mask = _apply_text_filter(mask, df, "Trigger/Barrier", row.get("Trigger/Barrier"))
        call_put = row.get("Call/Put")
        strike = row.get("strikePrice")
        has_option_key = (not _is_blank_drill_value(call_put)) or (not _is_blank_drill_value(strike))
        if has_option_key:
            mask = _apply_text_filter(mask, df, "Call/Put", call_put)
            mask = _apply_numeric_or_text_filter(mask, df, "strikePrice", strike)
        else:
            # This selected grouped row is the futures/non-option bucket for the product+month.
            # Keep option rows with the same product/month out of the futures drill-down.
            if "Call/Put" in df.columns:
                mask = mask & _blank_series_mask(df["Call/Put"])
            if "strikePrice" in df.columns:
                mask = mask & _blank_series_mask(df["strikePrice"])

    return df[mask].reset_index(drop=True)


def _selected_group_description(selected_group_row, selected_preset):
    if selected_group_row is None:
        return ""
    parts = []
    for col in ["Account Number", "Product", "Contract Month/Year", "expiryDate", "Trigger/Barrier", "CCY 1", "CCY 2", "Call/Put", "strikePrice"]:
        if col in selected_group_row.index and not _is_blank_drill_value(selected_group_row.get(col)):
            parts.append(f"{col}: {selected_group_row.get(col)}")
    return f"{selected_preset} — " + ", ".join(parts) if parts else selected_preset

def _drilldown_signature(selected_group_row, selected_preset):
    """Stable identifier for the selected grouped-position row."""
    if selected_group_row is None:
        return None
    pieces = [str(selected_preset)]
    for col in ["Account Number", "Product", "Contract Month/Year", "expiryDate", "Trigger/Barrier", "CCY 1", "CCY 2", "Call/Put", "strikePrice"]:
        value = selected_group_row.get(col) if hasattr(selected_group_row, "get") else None
        pieces.append(f"{col}={'' if _is_blank_drill_value(value) else str(value).strip()}")
    return "|".join(pieces)


def _close_drilldown(sig=None):
    """Close the drill-down widget and avoid immediately reopening the same selected row."""
    if sig is None:
        sig = st.session_state.get("drilldown_signature")
    if sig:
        st.session_state["drilldown_closed_signature"] = sig
    for key in ["drilldown_row", "drilldown_preset", "drilldown_signature"]:
        st.session_state.pop(key, None)


def _render_drilldown_content(drill_df, selected_group_row, selected_preset, sig):
    """Render the open-position drill-down contents inside a dialog or fallback panel."""
    header_cols = st.columns([0.78, 0.22])
    with header_cols[0]:
        st.caption(_selected_group_description(selected_group_row, selected_preset))
    with header_cols[1]:
        if st.button("Close", key=f"close_drilldown_{_safe_key(sig)}", use_container_width=True):
            _close_drilldown(sig)
            st.rerun()

    if drill_df.empty:
        st.warning("No matching open-position rows were found for the selected group. This usually means one of the grouping fields is blank or formatted differently in the raw rows.")
        return

    total_qty = pd.to_numeric(drill_df.get("Quantity"), errors="coerce").sum() if "Quantity" in drill_df.columns else None
    metric_cols = st.columns(2)
    metric_cols[0].metric("Matched open-position rows", f"{len(drill_df):,}")
    if total_qty is not None:
        metric_cols[1].metric("Matched quantity", f"{total_qty:,.0f}")

    display_custom_table(
        "Selected Group Open Positions",
        drill_df,
        default_columns=OPEN_POSITION_COLUMNS,
        default_only=False,
    )


def _show_drilldown_widget(drill_df, selected_group_row, selected_preset, sig):
    """Show drill-down as a separate closable widget. Uses st.dialog when available."""
    title = "Open Positions making up selected group"
    if hasattr(st, "dialog"):
        try:
            dialog_decorator = st.dialog(title, width="large")
        except TypeError:
            dialog_decorator = st.dialog(title)

        @dialog_decorator
        def _dialog():
            _render_drilldown_content(drill_df, selected_group_row, selected_preset, sig)

        _dialog()
    else:
        # Fallback for older Streamlit versions: still separate and closable, just inline.
        with st.container(border=True):
            st.subheader(title)
            _render_drilldown_content(drill_df, selected_group_row, selected_preset, sig)


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
    open_positions_base_view_df = open_positions_standard_view(merged_tables)
    has_option_positions = _has_option_positions(open_positions_base_view_df)

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
                "description": "Automatic risk view: rows group by product + contract month/year, but use expiryDate/value date instead of month when expiryDate is available; OTC/swap rows keep Trigger/Barrier when available; options also keep Call/Put + Strike Price.",
                "mode": "auto_futures_options",
                "group_cols": None,
            },
            "Account Grouping": {
                "description": "Account risk view: account + product + contract month/year, using expiryDate/value date when available; Trigger/Barrier price when available; options additionally keep Call/Put + Strike Price.",
                "mode": "custom",
                "group_cols": ["account_number", "product", "ref_month", "trigger_barrier", "option_type", "strike"],
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
            st.write("Grouping by:", "Product + Contract Month/Year; expiryDate/value date when available; Trigger/Barrier when available; options additionally by Call/Put + Strike Price")
        else:
            grouped_pos_df = grouped_positions_standard_view(merged_tables, preset["group_cols"])
            st.write("Grouping by:", selected_preset)
        st.caption(preset["description"] + " Avg Fill Price is weighted by absolute Net Quantity and rounded to 2 decimals. NOV carries option OTE; option rows leave OTE blank. Trade Price remains available in Open Positions, not grouped views.")

        # Grouped Positions should stay risk-focused. Contract Description is intentionally
        # hidden here; use Open Positions for the exact PDF contract description.
        grouped_pos_df = grouped_pos_df.drop(columns=["Contract Description", "Full Name"], errors="ignore")
        # Finalize grouped layout:
        # - Product grouping hides Call/Put and strikePrice.
        # - NOV receives option OTE; OTE is reserved for futures/forwards.
        # - Realised PNL and Day PNL are present on every grouping view.
        grouped_pos_df = prepare_grouped_positions_display(
            grouped_pos_df,
            tables=merged_tables,
            selected_preset=selected_preset,
            drop_option_columns=(selected_preset == "Product"),
        )
        grouped_pos_df = _drop_option_columns_when_no_options(grouped_pos_df, has_option_positions)
        if not has_option_positions:
            st.caption("No option rows detected in the uploaded PDF(s), so Call/Put and strikePrice are hidden from position views.")
        # Only Account Grouping should display Account Number. Product-level risk views
        # intentionally aggregate across accounts, so hide Account Number there.
        if selected_preset in {"Product", "Product + Contract Month/Year"}:
            grouped_pos_df = grouped_pos_df.drop(columns=["Account Number"], errors="ignore")
        grouped_default_columns = [c for c in GROUPED_POSITION_COLUMNS if c in grouped_pos_df.columns]

        st.caption("Click a row in the Grouped Positions table to view the Open Positions rows that make up that group.")
        group_selection = display_custom_table(
            f"Grouped Positions - {selected_preset}",
            grouped_pos_df,
            default_columns=grouped_default_columns,
            default_only=True,
            selectable=True,
            selection_key=f"grouped_positions_row_selection_{_safe_key(selected_preset)}_v67",
            # Keep separate column-toggle state per grouping preset. Without this,
            # hiding Call/Put and strikePrice in the Product preset can make those
            # columns appear hidden when the user switches to Account Grouping.
            table_key_override=f"Grouped Positions {selected_preset} v67",
        )

        selected_rows = _event_selected_rows(group_selection)
        selected_group = None
        selected_sig = None
        if selected_rows:
            selected_pos = selected_rows[0]
            if 0 <= selected_pos < len(grouped_pos_df):
                selected_group = grouped_pos_df.iloc[selected_pos]
                selected_sig = _drilldown_signature(selected_group, selected_preset)
                if st.session_state.get("drilldown_closed_signature") != selected_sig:
                    st.session_state["drilldown_row"] = selected_group.to_dict()
                    st.session_state["drilldown_preset"] = selected_preset
                    st.session_state["drilldown_signature"] = selected_sig
                else:
                    if st.button("Open selected group details", key=f"reopen_drilldown_{_safe_key(selected_sig)}"):
                        st.session_state.pop("drilldown_closed_signature", None)
                        st.session_state["drilldown_row"] = selected_group.to_dict()
                        st.session_state["drilldown_preset"] = selected_preset
                        st.session_state["drilldown_signature"] = selected_sig
                        st.rerun()
        else:
            st.info("Select a grouped-position row to open a closable drill-down widget with the underlying Open Positions.")

        if st.session_state.get("drilldown_row") and st.session_state.get("drilldown_preset"):
            selected_group = pd.Series(st.session_state["drilldown_row"])
            selected_preset_for_drill = st.session_state["drilldown_preset"]
            drill_sig = st.session_state.get("drilldown_signature") or _drilldown_signature(selected_group, selected_preset_for_drill)
            open_df_for_drill = open_positions_base_view_df
            drill_df = _open_positions_for_group(open_df_for_drill, selected_group, selected_preset_for_drill)
            drill_df = _drop_option_columns_when_no_options(drill_df, has_option_positions)
            _show_drilldown_widget(drill_df, selected_group, selected_preset_for_drill, drill_sig)
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
        open_positions_view_df = _drop_option_columns_when_no_options(open_positions_base_view_df, has_option_positions)
        display_custom_table(
            "Open Positions",
            open_positions_view_df,
            default_columns=_open_positions_default_columns(open_positions_view_df),
            default_only=True,
            table_key_override="Open Positions v67 otc fields",
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
