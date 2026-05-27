import json

import streamlit as st
import pandas as pd
import html

import data_source as ds
from parser import (
    extract,
    to_excel_bytes,
    grouped_positions_standard_view,
    grouped_positions_product_month_auto_standard_view,
    open_positions_standard_view,
    OPEN_POSITION_COLUMNS,
    GROUPED_POSITION_COLUMNS,
    merge_extracted_tables,
    prepare_grouped_positions_display,
    realized_pnl_summary,
    standard_position_view_from_df,
    closed_positions_standard_view,
    grouped_realized_pnl_view,
    grouped_fx_trades_view,
    _apply_conditional_position_columns,
)

APP_VERSION_TAG = "v93"
APP_VERSION_DESCRIPTION = "Drop CCY columns from Aggregated Positions in all views, not just FX"

st.set_page_config(page_title="MyStoneX Positions", layout="wide")
st.title("MyStoneX Positions")
st.caption(f"Upload one or more StoneX statement PDFs, merge trades, review aggregated positions, drill into details, and export. Version: {APP_VERSION_TAG} {APP_VERSION_DESCRIPTION}.")

with st.sidebar:
    st.header("Options")
    source = st.radio(
        "Data source",
        options=["PDF upload", "Internal API"],
        index=0,
        key=f"data_source_{APP_VERSION_TAG}",
    )
    include_trades = st.checkbox("Include trades", value=True)
    show_source_lines = st.checkbox("Show source_line columns", value=False)
    st.markdown("---")
    st.subheader("Multi-PDF merge")
    st.caption("Trades are always appended from all uploaded PDFs. No latest-snapshot filtering is applied. API mode loads a single response.")
    st.markdown("---")
    st.subheader("Column display")
    st.caption("After upload, use the Customize table button above each table to show/hide columns.")
    st.markdown("---")
    st.caption("Tip: use filters and aggregation views to reconcile trade-level rows back to position exposure.")

def add_source_pdf(tables, pdf_name, pdf_index):
    """Return a new dict with source_pdf columns added; never mutate the input.

    The input may come from a Streamlit cache, so mutating it would corrupt the
    cached value across reruns and across different uploaded PDFs that hash to
    the same bytes.
    """
    out = {}
    for name, df in tables.items():
        if isinstance(df, pd.DataFrame) and not df.empty:
            new_df = df.copy()
            new_df["source_pdf"] = pdf_name
            new_df["source_pdf_index"] = pdf_index
            out[name] = new_df
        else:
            out[name] = df
    return out


@st.cache_data(show_spinner=False)
def _cached_extract(pdf_bytes: bytes, include_open_positions: bool):
    """Cache PDF parsing across reruns so drill-down row clicks don't re-parse.

    Keyed by the raw PDF bytes — Streamlit hashes them automatically. Returns a
    Dict[str, DataFrame] that callers must not mutate (see add_source_pdf).

    Cache version: v93 (Drop CCY columns from all Aggregated Positions views)
    """
    return extract(pdf_bytes, include_open_positions=include_open_positions)


@st.cache_data(show_spinner=False)
def _cached_to_excel_bytes(_tables, cache_key: str):
    """Cache Excel-export byte generation across reruns.

    cache_key carries the identity (typically derived from pdf_bytes hashes and
    include_open_positions) so Streamlit can invalidate when inputs change.
    """
    return to_excel_bytes(_tables)

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
    Trades and Aggregated Positions so futures/FX-only PDFs stay cleaner.
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
    drop_cols = ["Call/Put", "NOV"]
    if "strikePrice" in df.columns and not _display_series_nonblank(df["strikePrice"]).any():
        drop_cols.append("strikePrice")
    return df.drop(columns=drop_cols, errors="ignore")


def _numeric_series(df, cols):
    """Return the first available numeric series from a list of candidate columns."""
    if df is None or df.empty:
        return pd.Series(dtype=float)
    for col in cols:
        if col in df.columns:
            return pd.to_numeric(
                df[col].astype(str).str.replace("$", "", regex=False).str.replace(",", "", regex=False).str.replace("(", "-", regex=False).str.replace(")", "", regex=False),
                errors="coerce",
            )
    return pd.Series([0] * len(df), index=df.index, dtype=float)


def _sum_numeric(df, cols):
    series = _numeric_series(df, cols)
    if series.empty:
        return 0.0
    return float(series.fillna(0).sum())




def _option_rows_for_summary(df):
    if df is None or df.empty:
        return pd.Series(dtype=bool)
    mask = pd.Series(False, index=df.index)
    if "Type" in df.columns:
        mask = mask | df["Type"].astype(str).str.strip().str.upper().isin(["OPTION", "NDO"])
    if "Call/Put" in df.columns:
        cp = df["Call/Put"].astype(str).str.strip().str.upper()
        mask = mask | cp.isin(["CALL", "PUT", "C", "P"])
    if "Contract Description" in df.columns:
        mask = mask | df["Contract Description"].astype(str).str.upper().str.contains(r"\b(?:CALL|PUT|OPTION)\b", regex=True, na=False)
    return mask.fillna(False)


def _total_nov_for_summary(df):
    """Sum option/NDO NOV using Market Value as the source of truth when present."""
    if df is None or df.empty:
        return 0.0
    option_mask = _option_rows_for_summary(df)
    if not option_mask.any():
        return 0.0
    option_rows = df.loc[option_mask]
    mv_cols = [c for c in ["Market Value", "market_value", "MarketValue", "market_value_signed"] if c in option_rows.columns]
    if mv_cols:
        mv = _numeric_series(option_rows, mv_cols)
        return float(mv.fillna(0).sum())
    return _sum_numeric(option_rows, ["NOV"])



def _total_ote_for_summary(df):
    """Sum non-option OTE using Market Value as the source of truth when present."""
    if df is None or df.empty:
        return 0.0
    option_mask = _option_rows_for_summary(df)
    non_option = df.loc[~option_mask].copy()
    if non_option.empty:
        return 0.0
    mv_cols = [c for c in ["Market Value", "market_value", "MarketValue", "market_value_signed"] if c in non_option.columns]
    if mv_cols:
        mv = _numeric_series(non_option, mv_cols)
        return float(mv.fillna(0).sum())
    return _sum_numeric(non_option, ["Unrealised PNL (OTE)"])


def _metric_money(value):
    try:
        value = float(value)
    except Exception:
        value = 0.0
    return f"{value:,.2f}"


def _distinct_count(df, column):
    if df is None or df.empty or column not in df.columns:
        return 0
    vals = df[column].dropna().astype(str).str.strip()
    vals = vals[~vals.str.lower().isin(["", "none", "nan", "nat", "unknown", "multiple"])]
    return int(vals.nunique())


def _best_last_update(trades_df, merged_tables):
    candidates = []
    for df in [trades_df, merged_tables.get("Statement Dates", pd.DataFrame()) if isinstance(merged_tables, dict) else pd.DataFrame()]:
        if df is None or df.empty:
            continue
        for col in ["Last Update", "last_update", "statement_date", "Statement Date", "Trade Date"]:
            if col in df.columns:
                parsed = pd.to_datetime(df[col], errors="coerce")
                if parsed.notna().any():
                    candidates.append(parsed.max())
    if not candidates:
        return "N/A"
    return max(candidates).strftime("%Y-%m-%d")


def _summary_metrics(trades_df, standard_positions_df, merged_tables):
    """Return high-level metrics used by the Home page."""
    non_option_positions = standard_positions_df
    if standard_positions_df is not None and not standard_positions_df.empty and "Type" in standard_positions_df.columns:
        non_option_positions = standard_positions_df[
            standard_positions_df["Type"].astype(str).str.strip().str.upper() != "OPTION"
        ]

    return [
        ("Total trades", f"{len(trades_df):,}", "Row-level trade records in the current upload."),
        ("Total Products", f"{_distinct_count(trades_df, 'Product'):,}", "Distinct normalized products."),
        ("Total Accounts", f"{_distinct_count(trades_df, 'Account Number'):,}", "Distinct account numbers."),
        ("Total Unrealised PNL", _metric_money(_total_ote_for_summary(standard_positions_df)), "Non-option OTE / MTM using Market Value when available."),
        ("Total NOV", _metric_money(_total_nov_for_summary(standard_positions_df)), "Option value from option Market Value."),
        ("Total Market Value", _metric_money(_sum_numeric(standard_positions_df, ["Market Value", "market_value", "market_value_signed"])), "Total market value when supplied."),
        ("Data Last Updated", _best_last_update(trades_df, merged_tables), "Latest statement/update date in the upload."),
    ]


def _render_metric_cards(metrics):
    """Render large, non-truncated cards for the Home page."""
    st.markdown(
        """
        <style>
        .sx-card {
            border: 1px solid rgba(49, 51, 63, 0.15);
            border-radius: 14px;
            padding: 1rem 1.1rem;
            min-height: 118px;
            background: rgba(248, 249, 251, 0.75);
            box-shadow: 0 1px 2px rgba(49, 51, 63, 0.05);
        }
        .sx-card-label {
            color: rgba(49, 51, 63, 0.72);
            font-size: 0.95rem;
            font-weight: 600;
            margin-bottom: 0.55rem;
        }
        .sx-card-value {
            color: rgb(49, 51, 63);
            font-size: 1.85rem;
            font-weight: 700;
            line-height: 1.15;
            overflow-wrap: anywhere;
            word-break: break-word;
        }
        .sx-card-help {
            color: rgba(49, 51, 63, 0.55);
            font-size: 0.78rem;
            margin-top: 0.45rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    rows = [metrics[:4], metrics[4:]]
    for row_idx, row in enumerate(rows):
        cols = st.columns(len(row))
        for col, (label, value, help_text) in zip(cols, row):
            with col:
                st.markdown(
                    f"""
                    <div class="sx-card">
                      <div class="sx-card-label">{html.escape(str(label))}</div>
                      <div class="sx-card-value">{html.escape(str(value))}</div>
                      <div class="sx-card-help">{html.escape(str(help_text))}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        if row_idx == 0:
            st.write("")


def _render_home_page(trades_df, standard_positions_df, merged_tables, extracted_tables):
    """Render a Home page so summary values are readable and navigation is obvious."""
    st.subheader("Home")
    st.caption("A quick operating summary for the uploaded trade data. Values are shown in full so large numbers are not truncated.")
    _render_metric_cards(_summary_metrics(trades_df, standard_positions_df, merged_tables))

    st.markdown("---")
    left, right = st.columns([0.55, 0.45])
    with left:
        st.markdown("**What you can do from here**")
        st.markdown(
            """
            - **Aggregated Positions**: group trades by Product, Product + expiry/contract month, or Account.
            - **Trades**: inspect the row-level trade records and apply filters.
            - **Position Details / Drill-down**: review the trades behind the selected aggregated row.
            - **Exceptions / Data Quality**: check missing fields and parser/data quality messages.
            - **Export**: download Excel or CSV outputs.
            """
        )
    with right:
        st.markdown("**Current upload**")
        file_count = len(extracted_tables or [])
        exception_count = len(merged_tables.get("Exceptions", pd.DataFrame())) if isinstance(merged_tables, dict) else 0
        st.write(f"Files loaded: **{file_count:,}**")
        st.write(f"Exceptions / data-quality rows: **{exception_count:,}**")
        if trades_df is not None and not trades_df.empty and "Type" in trades_df.columns:
            type_counts = trades_df["Type"].fillna("Unknown").astype(str).replace({"": "Unknown"}).value_counts().reset_index()
            type_counts.columns = ["Type", "Trades"]
            st.dataframe(type_counts, use_container_width=True, hide_index=True)


def _download_df_csv(label, df, key):
    if df is None or df.empty:
        return
    st.download_button(
        label,
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{_safe_key(label).lower()}.csv",
        mime="text/csv",
        key=key,
    )


def _date_candidates(df):
    if df is None or df.empty:
        return []
    preferred = ["expiryDate", "Trade Date", "Last Update", "statement_date", "Statement Date"]
    return [c for c in preferred if c in df.columns]


def _date_bounds(df, column):
    if df is None or df.empty or column not in df.columns:
        return None, None
    parsed = pd.to_datetime(df[column], errors="coerce")
    if not parsed.notna().any():
        return None, None
    return parsed.min().date(), parsed.max().date()


def _data_quality_summary(df):
    checks = []
    if df is None or df.empty:
        return pd.DataFrame(columns=["Check", "Issue Count", "Description"])
    required_checks = [
        ("Missing Account Number", "Account Number", "Rows where account number is blank."),
        ("Missing Product", "Product", "Rows where normalized product is blank."),
        ("Missing Type", "Type", "Rows where trade type is blank."),
        ("Missing Quantity", "Quantity", "Rows where trade quantity is blank or not numeric."),
        ("Missing Currency", "Currency", "Rows where currency is blank."),
        ("Missing expiryDate", "expiryDate", "Rows where expiry/value/end date is blank."),
    ]
    for name, col, desc in required_checks:
        if col not in df.columns:
            checks.append({"Check": name, "Issue Count": "N/A", "Description": f"Column {col} is not available in this dataset."})
            continue
        if col == "Quantity":
            count = int(pd.to_numeric(df[col], errors="coerce").isna().sum())
        else:
            count = int((~_display_series_nonblank(df[col])).sum())
        checks.append({"Check": name, "Issue Count": count, "Description": desc})
    return pd.DataFrame(checks)


def _clean_filter_values(series):
    if series is None:
        return []
    vals = []
    for v in series.dropna().astype(str).str.strip().tolist():
        if v and v.lower() not in {"none", "nan", "nat", "unknown", "other", "multiple"}:
            vals.append(v)
    return sorted(set(vals), key=lambda x: x.upper())


def _render_position_search_controls(base_view_df, key_prefix):
    """Render account/product/type/exchange/currency/date filters for Trades and Aggregated Positions.

    Exchange renders before Product so selecting an exchange dynamically narrows the product list.
    """
    filters = {
        "account": "",
        "broker": "",
        "products": [],
        "types": [],
        "exchanges": [],
        "currencies": [],
        "date_enabled": False,
        "date_field": None,
        "date_from": None,
        "date_to": None,
    }
    if base_view_df is None or base_view_df.empty:
        return filters

    key_prefix_safe = _safe_key(key_prefix)
    exchange_key = f"{key_prefix_safe}_exchange_filter_{APP_VERSION_TAG}"

    with st.container(border=True):
        st.caption("Filters")
        # Row 1: Account | Exchange | Type  (Exchange first so product list can react to it)
        r1c1, r1c2, r1c3 = st.columns(3)
        filters["account"] = r1c1.text_input(
            "Account Number contains",
            value="",
            key=f"{key_prefix_safe}_account_search_{APP_VERSION_TAG}",
            placeholder="e.g. LME11630",
        )
        exchange_options = _clean_filter_values(base_view_df["Exchange"]) if "Exchange" in base_view_df.columns else []
        filters["exchanges"] = r1c2.multiselect(
            "Exchange",
            options=exchange_options,
            key=exchange_key,
        ) if exchange_options else []

        type_options = _clean_filter_values(base_view_df["Type"]) if "Type" in base_view_df.columns else []
        filters["types"] = r1c3.multiselect(
            "Type",
            options=type_options,
            key=f"{key_prefix_safe}_type_filter_{APP_VERSION_TAG}",
        ) if type_options else []

        # Row 2: Product (filtered by selected exchanges) | Currency | Date
        r2c1, r2c2, r2c3 = st.columns(3)

        # Narrow the product list to exchanges already selected (reads session_state from the
        # widget rendered above — updated on every Streamlit rerun).
        selected_exchanges_now = st.session_state.get(exchange_key) or []
        if selected_exchanges_now and "Exchange" in base_view_df.columns:
            wanted_ex = {str(x).strip().upper() for x in selected_exchanges_now}
            product_source_df = base_view_df[
                base_view_df["Exchange"].astype(str).str.strip().str.upper().isin(wanted_ex)
            ]
        else:
            product_source_df = base_view_df
        product_options = _clean_filter_values(product_source_df["Product"]) if "Product" in product_source_df.columns else []
        filters["products"] = r2c1.multiselect(
            "Product",
            options=product_options,
            key=f"{key_prefix_safe}_product_filter_{APP_VERSION_TAG}",
        ) if product_options else []

        currency_values = []
        for ccy_col in ["Currency", "CCY 1", "CCY 2"]:
            if ccy_col in base_view_df.columns:
                currency_values.extend(_clean_filter_values(base_view_df[ccy_col]))
        currency_options = sorted(set(currency_values), key=lambda x: x.upper())
        filters["currencies"] = r2c2.multiselect(
            "Currency",
            options=currency_options,
            key=f"{key_prefix_safe}_currency_filter_{APP_VERSION_TAG}",
        ) if currency_options else []

        date_fields = _date_candidates(base_view_df)
        if date_fields:
            filters["date_enabled"] = r2c3.checkbox("Filter by date range", key=f"{key_prefix_safe}_date_enabled_{APP_VERSION_TAG}")
            if filters["date_enabled"]:
                filters["date_field"] = r2c3.selectbox("Date field", date_fields, key=f"{key_prefix_safe}_date_field_{APP_VERSION_TAG}")
                min_date, max_date = _date_bounds(base_view_df, filters["date_field"])
                if min_date and max_date:
                    d1, d2 = r2c3.columns(2)
                    filters["date_from"] = d1.date_input("From", value=min_date, key=f"{key_prefix_safe}_date_from_{APP_VERSION_TAG}")
                    filters["date_to"] = d2.date_input("To", value=max_date, key=f"{key_prefix_safe}_date_to_{APP_VERSION_TAG}")
        else:
            r2c3.caption("Date range unavailable")

        r3c1, _, _ = st.columns(3)
        filters["broker"] = r3c1.text_input(
            "Broker Code contains",
            value="",
            key=f"{key_prefix_safe}_broker_search_{APP_VERSION_TAG}",
            placeholder="e.g. DP132",
        )
    return filters

def _position_filter_mask(view_df, filters):
    if view_df is None or view_df.empty:
        return pd.Series(dtype=bool)
    mask = pd.Series(True, index=view_df.index)

    account_q = str(filters.get("account") or "").strip()
    if account_q and "Account Number" in view_df.columns:
        mask &= view_df["Account Number"].astype(str).str.contains(account_q, case=False, na=False, regex=False)

    broker_q = str(filters.get("broker") or "").strip()
    if broker_q and "Broker Code" in view_df.columns:
        mask &= view_df["Broker Code"].astype(str).str.contains(broker_q, case=False, na=False, regex=False)

    selected_products = filters.get("products") or []
    if selected_products and "Product" in view_df.columns:
        wanted = {str(x).strip().upper() for x in selected_products}
        mask &= view_df["Product"].astype(str).str.strip().str.upper().isin(wanted)

    selected_types = filters.get("types") or []
    if selected_types and "Type" in view_df.columns:
        wanted = {str(x).strip().upper() for x in selected_types}
        mask &= view_df["Type"].astype(str).str.strip().str.upper().isin(wanted)

    selected_exchanges = filters.get("exchanges") or []
    if selected_exchanges and "Exchange" in view_df.columns:
        wanted = {str(x).strip().upper() for x in selected_exchanges}
        mask &= view_df["Exchange"].astype(str).str.strip().str.upper().isin(wanted)

    selected_currencies = filters.get("currencies") or []
    if selected_currencies:
        wanted = {str(x).strip().upper() for x in selected_currencies}
        ccy_mask = pd.Series(False, index=view_df.index)
        for ccy_col in ["Currency", "CCY 1", "CCY 2"]:
            if ccy_col in view_df.columns:
                ccy_mask |= view_df[ccy_col].astype(str).str.strip().str.upper().isin(wanted)
        mask &= ccy_mask

    if filters.get("date_enabled") and filters.get("date_field") in view_df.columns:
        date_field = filters.get("date_field")
        parsed = pd.to_datetime(view_df[date_field], errors="coerce")
        if filters.get("date_from") is not None:
            mask &= parsed >= pd.to_datetime(filters.get("date_from"))
        if filters.get("date_to") is not None:
            mask &= parsed <= pd.to_datetime(filters.get("date_to"))

    return mask.fillna(False)

def _apply_position_filters_to_view(view_df, filters):
    if view_df is None or view_df.empty:
        return view_df
    mask = _position_filter_mask(view_df, filters)
    if mask.empty:
        return view_df.iloc[0:0]
    return view_df.loc[mask].copy()


def _apply_position_filters_to_tables(tables, filters):
    """Filter raw Trades before grouping so hidden Account Number still works."""
    raw = tables.get("Open Positions", pd.DataFrame())
    if raw is None or raw.empty:
        return tables
    has_filter = bool(
        str(filters.get("account") or "").strip()
        or str(filters.get("broker") or "").strip()
        or filters.get("products")
        or filters.get("types")
        or filters.get("exchanges")
        or filters.get("currencies")
        or filters.get("date_enabled")
    )
    if not has_filter:
        return tables

    base_view = open_positions_standard_view({"Open Positions": raw})
    mask = _position_filter_mask(base_view, filters)
    filtered = raw.iloc[0:0].copy()
    if len(mask) == len(raw):
        filtered = raw.loc[mask.to_numpy()].copy()
    else:
        # Defensive fallback if indexes were transformed unexpectedly.
        filtered = raw.iloc[list(mask[mask].index)].copy()
    out = dict(tables)
    out["Open Positions"] = filtered
    return out

def _safe_key(text):
    return "".join(ch if ch.isalnum() else "_" for ch in str(text))

def display_custom_table(label, df, default_columns=None, default_only=False, selectable=False, selection_key=None, table_key_override=None, multi_select=False, height=None):
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
        # Streamlit dataframe row selection powers the aggregated-position drill-down.
        # Selection row indexes refer to the displayed dataframe, which is built from df
        # without reordering, so they can be used to fetch the selected source row.
        kwargs = dict(
            use_container_width=True,
            on_select="rerun",
            selection_mode="multi-row" if multi_select else "single-row",
            key=selection_key or f"df_select_{table_key}",
        )
        if height is not None:
            kwargs["height"] = height
        return st.dataframe(display_df, **kwargs)

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

    height_kwargs = {"height": height} if height is not None else {}
    quantity_cols = [c for c in ["Quantity", "Net Quantity"] if c in display_df.columns]
    if quantity_cols:
        styled = display_df.style.map(_quantity_color, subset=quantity_cols)
        st.dataframe(styled, use_container_width=True, **height_kwargs)
    else:
        st.dataframe(display_df, use_container_width=True, **height_kwargs)
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
    """When an FX aggregated row is selected, keep drill-down aligned to FX keys."""
    # FX aggregated rows may contain aggregated currency amounts and/or hidden trade prices.
    # Drill-down should therefore filter by stable identity fields only: product, currencies,
    # and value/settlement date when that date is present in the selected group.
    for col in ["CCY 1", "CCY 2"]:
        mask = _apply_text_filter(mask, df, col, row.get(col))
    if not _is_blank_drill_value(row.get("expiryDate")) and "expiryDate" in df.columns:
        mask = _apply_text_filter(mask, df, "expiryDate", row.get("expiryDate"))
    return mask




def _apply_contract_or_expiry_drill_filter(mask, df, row):
    """Filter by expiryDate when the aggregated row has one; otherwise by Contract Month/Year."""
    expiry = row.get("expiryDate") if hasattr(row, "get") else None
    if not _is_blank_drill_value(expiry) and "expiryDate" in df.columns:
        return _apply_text_filter(mask, df, "expiryDate", expiry)
    return _apply_text_filter(mask, df, "Contract Month/Year", row.get("Contract Month/Year"))

def _open_positions_for_group(open_df, selected_group_row, selected_preset):
    """Filter trade rows to the rows that make up a selected aggregated row."""
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
            # This selected aggregated row is the futures/non-option bucket for the product+month.
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
    for col in ["Account Number", "Product", "Type", "Contract Month/Year", "expiryDate", "Trigger/Barrier", "CCY 1", "CCY 2", "Call/Put", "strikePrice"]:
        if col in selected_group_row.index and not _is_blank_drill_value(selected_group_row.get(col)):
            parts.append(f"{col}: {selected_group_row.get(col)}")
    return f"{selected_preset} — " + ", ".join(parts) if parts else selected_preset

def _drilldown_signature(selected_group_row, selected_preset):
    """Stable identifier for the selected aggregated-position row."""
    if selected_group_row is None:
        return None
    pieces = [str(selected_preset)]
    for col in ["Account Number", "Product", "Type", "Contract Month/Year", "expiryDate", "Trigger/Barrier", "CCY 1", "CCY 2", "Call/Put", "strikePrice"]:
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
    """Render the trade drill-down contents inside a dialog or fallback panel."""
    header_cols = st.columns([0.78, 0.22])
    with header_cols[0]:
        st.caption(_selected_group_description(selected_group_row, selected_preset))
    with header_cols[1]:
        if st.button("Close", key=f"close_drilldown_{_safe_key(sig)}", use_container_width=True):
            _close_drilldown(sig)
            st.rerun()

    if drill_df.empty:
        st.warning("No matching trade rows were found for the selected group. This usually means one of the grouping fields is blank or formatted differently in the raw rows.")
        return

    total_qty = pd.to_numeric(drill_df.get("Quantity"), errors="coerce").sum() if "Quantity" in drill_df.columns else None
    metric_cols = st.columns(2)
    metric_cols[0].metric("Matched trade rows", f"{len(drill_df):,}")
    if total_qty is not None:
        metric_cols[1].metric("Matched quantity", f"{total_qty:,.0f}")

    display_custom_table(
        "Selected Group Trades",
        drill_df,
        default_columns=OPEN_POSITION_COLUMNS,
        default_only=False,
    )


def _show_drilldown_widget(drill_df, selected_group_row, selected_preset, sig):
    """Show drill-down as a separate closable widget. Uses st.dialog when available."""
    title = "Trades making up selected group"
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


def _unwrap_envelope(parsed):
    """Some API responses wrap rows under 'accountDetails'. Accept both shapes."""
    if isinstance(parsed, dict) and "accountDetails" in parsed:
        return parsed["accountDetails"]
    return parsed if isinstance(parsed, list) else []


extracted_tables = []
pdf_bytes_list: list[bytes] = []

if source == "PDF upload":
    uploaded_files = st.file_uploader(
        "Upload one or more statement PDFs",
        type=["pdf"],
        accept_multiple_files=True,
    )
    if uploaded_files:
        for idx, uploaded in enumerate(uploaded_files, start=1):
            pdf_bytes = uploaded.read()
            pdf_bytes_list.append(pdf_bytes)
            with st.spinner(f"Extracting {uploaded.name}..."):
                # Cached call — on the first run for these PDF bytes this
                # actually parses; subsequent reruns (e.g. drill-down row
                # clicks) are near-instant.
                tables = _cached_extract(pdf_bytes, include_trades)
                tables = add_source_pdf(tables, uploaded.name, idx)
                extracted_tables.append(tables)
else:
    st.info(
        "Upload JSON responses saved from the Swagger 'Try it out' panel. "
        "Trades JSON is required; positions JSON is optional and only enriches "
        "futures/options with product and exchange names. The HTTP fetcher comes next."
    )
    trades_file = st.file_uploader(
        "Trades response (JSON)",
        type=["json"],
        key=f"api_trades_{APP_VERSION_TAG}",
    )
    positions_file = st.file_uploader(
        "Positions response (JSON, optional)",
        type=["json"],
        key=f"api_positions_{APP_VERSION_TAG}",
    )
    if trades_file is not None:
        try:
            trades_payload = json.loads(trades_file.read().decode("utf-8"))
            positions_payload = []
            if positions_file is not None:
                positions_payload = json.loads(positions_file.read().decode("utf-8"))
            trades_json = _unwrap_envelope(trades_payload)
            positions_json = _unwrap_envelope(positions_payload)
            with st.spinner("Mapping API response..."):
                tables = ds.extract_from_api(
                    positions_json=positions_json,
                    trades_json=trades_json,
                )
                tables = add_source_pdf(tables, f"API:{trades_file.name}", 1)
                extracted_tables.append(tables)
        except json.JSONDecodeError as exc:
            st.error(f"Invalid JSON: {exc}")


if extracted_tables:

    merged_tables = merge_extracted_tables(extracted_tables)
    # Excel export is expensive to rebuild on every Streamlit rerun. Key the
    # cache by the combined identity of the inputs that actually drive output.
    if source == "PDF upload" and pdf_bytes_list:
        _excel_cache_key = f"pdf|{include_trades}|{hash(tuple(pdf_bytes_list))}"
    else:
        _excel_cache_key = f"api|{include_trades}|{id(merged_tables)}"
    merged_excel = _cached_to_excel_bytes(merged_tables, _excel_cache_key)
    open_positions_base_view_df = open_positions_standard_view(merged_tables)
    has_option_positions = _has_option_positions(open_positions_base_view_df)

    standard_positions_summary_df = standard_position_view_from_df(merged_tables.get("Open Positions", pd.DataFrame()))

    latest_aggregated_df = pd.DataFrame()
    latest_trades_df = pd.DataFrame()
    latest_drill_df = pd.DataFrame()

    tabs = st.tabs([
        "Home",
        "Aggregated Positions",
        "Trades",
        "Realised PNL",
        "Position Details / Drill-down",
        "Exceptions / Data Quality",
        "Export",
    ])

    with tabs[0]:
        _render_home_page(open_positions_base_view_df, standard_positions_summary_df, merged_tables, extracted_tables)

    with tabs[1]:
        st.caption("Choose an aggregation view. Filters apply before aggregation, so account/type/exchange/currency/date filters work even when those columns are hidden in the selected view.")
        grouping_presets = {
            "Product": {
                "description": "One row per product and exchange, across all accounts, contract months, expiry dates, strikes, and trade prices.",
                "mode": "custom",
                "group_cols": ["product", "exchange"],
            },
            "Product + Contract Month/Year": {
                "description": "Risk view by product and expiry/value/end date when available, otherwise contract month/year. OTC rows keep Trigger/Barrier when available; options also keep Call/Put + Strike Price.",
                "mode": "auto_futures_options",
                "group_cols": None,
            },
            "Account Grouping": {
                "description": "Ownership view by account + product and expiry/value/end date when available, otherwise contract month/year. OTC rows keep Trigger/Barrier; options also keep Call/Put + Strike Price.",
                "mode": "custom",
                "group_cols": ["account_number", "product", "ref_month", "trigger_barrier", "option_type", "strike"],
            },
        }
        grouped_filters = _render_position_search_controls(open_positions_base_view_df, "aggregated_positions")
        st.session_state["last_aggregated_filters"] = grouped_filters
        grouped_source_tables = _apply_position_filters_to_tables(merged_tables, grouped_filters)
        grouped_open_base_view_df = open_positions_standard_view(grouped_source_tables)

        selected_preset = st.radio(
            "Aggregated Positions view",
            options=list(grouping_presets.keys()),
            horizontal=True,
            key=f"aggregated_positions_preset_view_{APP_VERSION_TAG}",
        )
        preset = grouping_presets[selected_preset]
        if preset["mode"] == "auto_futures_options":
            grouped_pos_df = grouped_positions_product_month_auto_standard_view(grouped_source_tables)
            st.write("Grouping by:", "Product + expiryDate/value date when available, otherwise Contract Month/Year; Trigger/Barrier when available; options additionally by Call/Put + Strike Price")
        else:
            grouped_pos_df = grouped_positions_standard_view(grouped_source_tables, preset["group_cols"])
            st.write("Grouping by:", selected_preset)
        st.caption(preset["description"] + " NOV carries option OTE; option rows leave OTE blank. Trade Price, Ref Price, and Original Quantity remain available in Trades, not Aggregated Positions.")

        grouped_pos_df = grouped_pos_df.drop(columns=["Contract Description", "Full Name"], errors="ignore")
        grouped_pos_df = prepare_grouped_positions_display(
            grouped_pos_df,
            tables=grouped_source_tables,
            selected_preset=selected_preset,
            drop_option_columns=(selected_preset == "Product"),
        )
        if "Call/Put" not in grouped_pos_df.columns:
            st.caption("No option rows detected, so Call/Put and strikePrice are hidden from position views.")
        if selected_preset in {"Product", "Product + Contract Month/Year"}:
            grouped_pos_df = grouped_pos_df.drop(columns=["Account Number"], errors="ignore")
        # Pre-selected columns: use conditional logic to check only relevant columns by default.
        # The picker always shows ALL GROUPED_POSITION_COLUMNS (ensured by prepare_grouped_positions_display).
        _conditional_cols = set(_apply_conditional_position_columns(grouped_pos_df.copy()).columns)
        # Detect what's in the data to drive smart pre-selection.
        _grouped_has_options      = _has_option_positions(grouped_pos_df)
        _grouped_has_accumulators = "Trigger/Barrier" in _conditional_cols  # non-blank TB → accumulator

        # Product preset: Contract Month/Year and expiryDate are not grouping keys — exclude.
        # Exception: expiryDate is kept for options (strike/expiry pair) and accumulators.
        _preset_exclude = {"Contract Month/Year", "settlementPrice"} if selected_preset == "Product" else set()
        if selected_preset == "Product" and not _grouped_has_options and not _grouped_has_accumulators:
            _preset_exclude.add("expiryDate")

        # CCY 1/2 and their amounts are removed from the Aggregated Positions table
        # in all views — they are trade-level FX detail and not useful in a grouped
        # summary regardless of whether the view is purely FX or mixed.
        grouped_pos_df = grouped_pos_df.drop(
            columns=["CCY 1", "CCY 1 Amount", "CCY 2", "CCY 2 Amount"],
            errors="ignore",
        )
        _preset_exclude.update({"CCY 1", "CCY 1 Amount", "CCY 2", "CCY 2 Amount"})

        # For purely FX views also hide Exchange and Currency from defaults.
        if "Type" in grouped_pos_df.columns:
            _type_vals = grouped_pos_df["Type"].dropna().astype(str).str.strip().str.upper()
            _type_vals = _type_vals[_type_vals.str.lower().ne("multiple") & _type_vals.ne("")]
            _all_fx = _type_vals.str.startswith("FX").all() | _type_vals.isin(["NDO"]).all() if not _type_vals.empty else False
        else:
            _all_fx = False
        if _all_fx:
            _preset_exclude.update({"Exchange", "Currency"})

        # expiryDate IS a grouping key for Product + Contract Month/Year and Account Grouping
        # (replaces ref_month when available; explicit key for FX rows) — always force-include.
        # Trigger/Barrier is a grouping key when accumulators are present — always force-include.
        _force_include = set()
        if selected_preset in {"Product + Contract Month/Year", "Account Grouping"}:
            _force_include.add("expiryDate")
        if _grouped_has_accumulators:
            _force_include.update({"Trigger/Barrier", "expiryDate"})

        grouped_default_columns = [
            c for c in GROUPED_POSITION_COLUMNS
            if c != "Type"
            and c not in _preset_exclude
            and (c in _conditional_cols or c in _force_include)
        ]

        st.caption("Click a row in the Aggregated Positions table to view the trade rows that make up that group.")
        group_selection = display_custom_table(
            f"Aggregated Positions - {selected_preset}",
            grouped_pos_df,
            default_columns=grouped_default_columns,
            default_only=True,
            selectable=True,
            selection_key=f"aggregated_positions_row_selection_{_safe_key(selected_preset)}_{APP_VERSION_TAG}",
            table_key_override=f"Aggregated Positions {selected_preset} {APP_VERSION_TAG}",
        )
        latest_aggregated_df = grouped_pos_df.copy()
        _download_df_csv("Download aggregated positions CSV", grouped_pos_df, f"download_aggregated_positions_{_safe_key(selected_preset)}_{APP_VERSION_TAG}")

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
            st.info("Select an aggregated-position row to open a closeable drill-down widget with the underlying trades.")

        if st.session_state.get("drilldown_row") and st.session_state.get("drilldown_preset"):
            selected_group = pd.Series(st.session_state["drilldown_row"])
            selected_preset_for_drill = st.session_state["drilldown_preset"]
            drill_sig = st.session_state.get("drilldown_signature") or _drilldown_signature(selected_group, selected_preset_for_drill)
            open_df_for_drill = grouped_open_base_view_df
            drill_df = _open_positions_for_group(open_df_for_drill, selected_group, selected_preset_for_drill)
            latest_drill_df = drill_df.copy()
            _show_drilldown_widget(drill_df, selected_group, selected_preset_for_drill, drill_sig)

    with tabs[2]:
        st.caption("Trades are the row-level open trade records that make up aggregated exposure.")
        trades_filters = _render_position_search_controls(open_positions_base_view_df, "trades")
        trades_filtered_df = _apply_position_filters_to_view(open_positions_base_view_df, trades_filters)
        trades_has_option_positions = _has_option_positions(trades_filtered_df)
        trades_view_df = _drop_option_columns_when_no_options(trades_filtered_df, trades_has_option_positions)
        latest_trades_df = trades_view_df.copy()

        trade_view_mode = st.radio(
            "View",
            ["Individual Trades", "Aggregated FX"],
            horizontal=True,
            key=f"trade_view_mode_{APP_VERSION_TAG}",
        )

        if trade_view_mode == "Aggregated FX":
            fx_agg_df = grouped_fx_trades_view(trades_view_df)
            if fx_agg_df.empty:
                st.info("No FX trades found. Check your filters or upload a statement with FX positions.")
            else:
                _FX_AGG_DEFAULT_COLS = [
                    c for c in [
                        "Account Number", "Product", "Type", "CCY 1", "CCY 1 Amount",
                        "CCY 2", "CCY 2 Amount", "expiryDate", "Quantity", "Net Quantity",
                        "Unrealised PNL (OTE)", "Trade Count",
                    ] if c in fx_agg_df.columns
                ]

                # Split screen: aggregated table on the left, drill-down on the right.
                # Fixed height keeps both panels visible on screen simultaneously.
                _FX_TABLE_HEIGHT = 480
                agg_col, drill_col = st.columns([0.4, 0.6], gap="large")

                with agg_col:
                    st.caption("Select one or more CCY pairs to see their trades →")
                    fx_agg_selection = display_custom_table(
                        "Aggregated FX Trades",
                        fx_agg_df,
                        default_columns=_FX_AGG_DEFAULT_COLS,
                        default_only=True,
                        selectable=True,
                        multi_select=True,
                        selection_key=f"fx_agg_select_{APP_VERSION_TAG}",
                        table_key_override=f"FX Agg {APP_VERSION_TAG}",
                        height=_FX_TABLE_HEIGHT,
                    )
                    _download_df_csv(
                        "Download aggregated FX CSV",
                        fx_agg_df,
                        f"download_fx_agg_csv_{APP_VERSION_TAG}",
                    )

                # Collect all selected aggregated rows and union their matching trades.
                selected_fx_row_idxs = (
                    fx_agg_selection.selection.rows
                    if fx_agg_selection and hasattr(fx_agg_selection, "selection")
                    else []
                )

                with drill_col:
                    if not selected_fx_row_idxs:
                        st.info("Select one or more CCY pairs on the left to see the underlying trades here.")
                    else:
                        # Build a union mask across all selected aggregated rows.
                        union_mask = pd.Series(False, index=trades_view_df.index)
                        selected_labels = []
                        for row_idx in selected_fx_row_idxs:
                            sel = fx_agg_df.iloc[row_idx]
                            row_mask = pd.Series(True, index=trades_view_df.index)
                            for col in ["Account Number", "Product", "Type", "CCY 1", "CCY 2", "expiryDate"]:
                                try:
                                    val = sel[col]
                                except (KeyError, IndexError):
                                    val = None
                                is_blank = (
                                    val is None
                                    or (isinstance(val, float) and pd.isna(val))
                                    or str(val).strip() in ("", "nan", "None", "NaT", "NaN")
                                )
                                if not is_blank and col in trades_view_df.columns:
                                    row_mask &= trades_view_df[col].astype(str).str.strip() == str(val).strip()
                            union_mask |= row_mask

                            # Build a readable label for this CCY pair.
                            ccy1 = str(sel.get("CCY 1", "") or "").strip()
                            ccy2 = str(sel.get("CCY 2", "") or "").strip()
                            expiry = str(sel.get("expiryDate", "") or "").strip()
                            pair = f"{ccy1}/{ccy2}" if ccy1 or ccy2 else "—"
                            if expiry and expiry not in ("nan", "None", "NaT"):
                                pair += f" exp {expiry}"
                            selected_labels.append(pair)

                        fx_drill_df = trades_view_df[union_mask].reset_index(drop=True)
                        n = len(fx_drill_df)
                        pairs_str = ", ".join(selected_labels)
                        st.caption(
                            f"{n} trade{'s' if n != 1 else ''} for: {pairs_str}"
                        )
                        display_custom_table(
                            "FX Trades",
                            fx_drill_df,
                            default_columns=_open_positions_default_columns(fx_drill_df),
                            default_only=True,
                            table_key_override=f"FX Drill {APP_VERSION_TAG}",
                            height=_FX_TABLE_HEIGHT,
                        )
                        _download_df_csv(
                            "Download trades CSV",
                            fx_drill_df,
                            f"download_fx_drill_csv_{APP_VERSION_TAG}",
                        )
        else:
            display_custom_table(
                "Trades",
                trades_view_df,
                default_columns=_open_positions_default_columns(trades_view_df),
                default_only=True,
                table_key_override=f"Trades {APP_VERSION_TAG}",
            )
            _download_df_csv("Download trades CSV", trades_view_df, f"download_trades_csv_{APP_VERSION_TAG}")

    with tabs[3]:
        st.caption("Realised P&L from Purchase & Sale (Gross Profit or Loss) rows — parsed using the same column logic as Trades and aggregated using the same presets as Aggregated Positions.")

        realised_source_tables = merged_tables

        # Trade-level view: same parsing pipeline as Trades
        realised_trade_df = closed_positions_standard_view(realised_source_tables)

        if realised_trade_df is None or realised_trade_df.empty:
            st.info("No realised P&L rows found in the uploaded statements.")
        else:
            total_realised = pd.to_numeric(realised_trade_df.get("Realised PNL", pd.Series(dtype=float)), errors="coerce").sum()
            r1, r2 = st.columns(2)
            r1.metric("Realised PNL rows", f"{len(realised_trade_df):,}")
            r2.metric("Total realised PNL", f"{total_realised:,.2f}")

            # Aggregation presets: same as Aggregated Positions
            realised_preset_options = {
                "Product": {"group_cols": ["product", "exchange"], "mode": "custom"},
                "Product + Contract Month/Year": {"group_cols": None, "mode": "auto_futures_options"},
                "Account Grouping": {"group_cols": ["account_number", "product", "ref_month", "trigger_barrier", "option_type", "strike"], "mode": "custom"},
            }
            selected_realised_preset = st.radio(
                "Realised PNL view",
                options=list(realised_preset_options.keys()),
                horizontal=True,
                key=f"realised_pnl_preset_{APP_VERSION_TAG}",
            )
            realised_preset = realised_preset_options[selected_realised_preset]

            # Aggregated view: same grouping rules as Aggregated Positions
            realised_grouped_df = grouped_realized_pnl_view(
                realised_source_tables,
                realised_preset["group_cols"],
                mode=realised_preset.get("mode", "custom"),
            )

            # Same column order as Aggregated Positions, with "Realised PNL" in place of
            # "Unrealised PNL (OTE)" and an extra "Debit/Credit" column.
            _realised_grouped_default_cols = [c for c in (
                [col for col in GROUPED_POSITION_COLUMNS if col != "Unrealised PNL (OTE)"]
                + ["Realised PNL", "Debit/Credit"]
            ) if c in realised_grouped_df.columns]

            display_custom_table(
                f"Realised PNL — {selected_realised_preset}",
                realised_grouped_df,
                default_columns=_realised_grouped_default_cols,
                default_only=True,
                table_key_override=f"Realised PNL {selected_realised_preset} {APP_VERSION_TAG}",
            )
            _download_df_csv("Download realised PNL CSV", realised_grouped_df, f"download_realised_pnl_csv_{APP_VERSION_TAG}")

    with tabs[4]:
        st.caption("This section shows the trade-level rows behind the last selected Aggregated Positions row.")
        if st.session_state.get("drilldown_row") and st.session_state.get("drilldown_preset"):
            selected_group = pd.Series(st.session_state["drilldown_row"])
            selected_preset_for_drill = st.session_state["drilldown_preset"]
            filters_for_drill = st.session_state.get("last_aggregated_filters", {})
            drill_source = _apply_position_filters_to_view(open_positions_base_view_df, filters_for_drill)
            drill_df = _open_positions_for_group(drill_source, selected_group, selected_preset_for_drill)
            drill_df = _drop_option_columns_when_no_options(drill_df, _has_option_positions(drill_source))
            latest_drill_df = drill_df.copy()
            st.subheader("Selected group")
            st.caption(_selected_group_description(selected_group, selected_preset_for_drill))
            display_custom_table(
                "Position Details / Drill-down",
                drill_df,
                default_columns=_open_positions_default_columns(drill_df),
                default_only=False,
                table_key_override=f"Position Details Drilldown {APP_VERSION_TAG}",
            )
            _download_df_csv("Download drill-down trades CSV", drill_df, f"download_drilldown_csv_{APP_VERSION_TAG}")
            if st.button("Clear selected group", key=f"clear_drilldown_tab_{APP_VERSION_TAG}"):
                _close_drilldown()
                st.rerun()
        else:
            st.info("Select a row in Aggregated Positions to see the underlying trades here.")

    with tabs[5]:
        st.caption("Review parser/data-quality issues and completeness checks.")
        exceptions_df = merged_tables.get("Exceptions", pd.DataFrame())
        q1, q2, q3 = st.columns(3)
        q1.metric("Exceptions", f"{len(exceptions_df):,}")
        quality_df = _data_quality_summary(open_positions_base_view_df)
        numeric_issue_counts = pd.to_numeric(quality_df.get("Issue Count", pd.Series(dtype=float)), errors="coerce")
        q2.metric("Data quality issues", f"{int(numeric_issue_counts.fillna(0).sum()):,}")
        q3.metric("Filtered trade rows", f"{len(open_positions_base_view_df):,}")
        with st.expander("Data Quality Checks", expanded=True):
            display_custom_table("Data Quality Checks", quality_df, default_only=False)
        with st.expander(f"Exceptions ({len(exceptions_df):,} rows)", expanded=not exceptions_df.empty):
            display_custom_table("Exceptions / Data Quality", exceptions_df, default_only=False)
            _download_df_csv("Download exceptions CSV", exceptions_df, f"download_exceptions_csv_{APP_VERSION_TAG}")

    with tabs[6]:
        st.caption("Download full or view-specific outputs.")
        st.download_button(
            "Download merged Excel",
            data=merged_excel,
            file_name="stonex_merged_trades_positions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_merged_{APP_VERSION_TAG}",
        )
        st.markdown("**View exports**")
        export_cols = st.columns(3)
        with export_cols[0]:
            _download_df_csv("Download all trades CSV", open_positions_base_view_df, f"download_all_trades_csv_{APP_VERSION_TAG}")
        with export_cols[1]:
            if 'latest_aggregated_df' in locals() and not latest_aggregated_df.empty:
                _download_df_csv("Download current aggregated CSV", latest_aggregated_df, f"download_current_aggregated_csv_{APP_VERSION_TAG}")
        with export_cols[2]:
            if 'latest_drill_df' in locals() and not latest_drill_df.empty:
                _download_df_csv("Download current drill-down CSV", latest_drill_df, f"download_current_drilldown_csv_{APP_VERSION_TAG}")

        if source == "PDF upload":
            with st.expander("Per-file exports"):
                for idx, (uploaded, tables) in enumerate(zip(uploaded_files, extracted_tables), start=1):
                    st.download_button(
                        f"Download Excel for {uploaded.name}",
                        data=to_excel_bytes(tables),
                        file_name=uploaded.name.rsplit(".", 1)[0] + "_trades.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{idx}_{uploaded.name}_{APP_VERSION_TAG}",
                    )


else:
    st.subheader("Home")
    if source == "PDF upload":
        st.info("Upload one or more PDFs to begin.")
    else:
        st.info("Upload a Trades response JSON to begin (positions JSON is optional).")
    st.markdown(
        """
        This prototype is organized around five working sections after upload:

        - **Aggregated Positions** for exposure-level views.
        - **Trades** for row-level trade records.
        - **Position Details / Drill-down** for the trades behind a selected aggregate.
        - **Exceptions / Data Quality** for completeness checks.
        - **Export** for Excel and CSV outputs.

        **Sources:** PDF statements (multi-upload merge) or Internal API responses (single response per load). The HTTP fetcher for the API path is coming next; for now upload the saved JSON.
        """
    )