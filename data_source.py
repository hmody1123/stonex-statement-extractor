"""API-source adapter for MyStoneX Positions.

Maps the StoneX positions-API and trades-API responses into the same
``Dict[str, pd.DataFrame]`` shape that ``parser.extract()`` returns from a PDF.
Drop-in for the rest of the app: grouping, drill-down, exceptions, and Excel
export work unchanged because they consume the same dict.

This is a pure mapping module. No HTTP, no auth, no network. It takes
already-fetched JSON (``list[dict]`` per endpoint) and returns parser-shaped
DataFrames. That keeps tokens, rate limits, and retries in the caller, and
makes the mapping unit-testable from sample payloads alone.

Endpoint coverage:
  - ``GET /account/{id}/positions``  -> listed futures/options metadata
                                        (productGroup, exchange) used to
                                        enrich trade rows via positionId join.
  - ``GET /account/{id}/trades``     -> populates the "Open Positions" table
                                        for every instrument type
                                        (FX, OTC, Futures, Options).

Not yet wired (central work pending):
  - Realized Gain and Loss / Realized PNL Summary — server-side central
    calculation not yet available.
  - Contract multiplier / lot size — needed for Market Value and option NOV.
    Deferred until an instrument-metadata endpoint exists.
"""
from __future__ import annotations

from typing import Any, Dict, List

import pandas as pd


# ---------------------------------------------------------------------------
# Instrument-type classification
# ---------------------------------------------------------------------------

# Maps the API's ``instrumentType`` discriminator to the parser's ``Type``
# column. "Outright" and "Forward" are aliases (FX dealer terminology).
# "Swap" is handled separately because FX swaps and OTC commodity swaps
# share the enum value and are disambiguated by ccy1 population.
_INSTRUMENT_TYPE_MAP: Dict[str, str] = {
    "spot":             "FX SPOT",
    "outright":         "FX FWD",
    "forward":          "FX FWD",
    "futures":          "FUT",
    "option":           "OPT",
    "optionsonfutures": "OPT",
    "accumulator":      "OTC ACCUMULATOR",
    # Non-deliverable variants — both word orders the API has used historically.
    # The parser-side normalizer collapses these to the canonical "FX NDF" /
    # "NDO" display labels so the Type filter only shows one entry each.
    "ndf":              "FX NDF",
    "ndffx":            "FX NDF",
    "fxndf":            "FX NDF",
    "nondeliverableforward":  "FX NDF",
    "ndo":              "NDO",
    "ndofx":            "NDO",
    "fxndo":            "NDO",
    "nondeliverableoption":   "NDO",
}


def _classify_type(row: Dict[str, Any]) -> str:
    """Return the parser-side ``Type`` for a trades-endpoint row."""
    raw = str(row.get("instrumentType", "")).strip().casefold()
    if raw == "swap":
        return "FX SWAP" if (row.get("ccy1") or "").strip() else "OTC SWAP"
    # Collapse whitespace and dashes so "NDF FX", "ndf-fx", "Non-Deliverable
    # Forward" all hit the same map entry.
    compact = "".join(ch for ch in raw if ch.isalnum())
    return _INSTRUMENT_TYPE_MAP.get(compact, _INSTRUMENT_TYPE_MAP.get(raw, "OTHER"))


# ---------------------------------------------------------------------------
# Field-level helpers
# ---------------------------------------------------------------------------

def _ms_to_iso_date(value: Any) -> str | None:
    """Convert a Unix-millisecond field to an ISO date string (``YYYY-MM-DD``).

    Strings are what the parser's date-normalization layer expects, so emitting
    them here lets API rows flow through the same view pipeline as PDF rows.
    """
    if value in (None, 0, "0", ""):
        return None
    try:
        ms = int(value)
    except (TypeError, ValueError):
        return None
    if ms == 0:
        return None
    ts = pd.to_datetime(ms, unit="ms", errors="coerce")
    return None if pd.isna(ts) else ts.strftime("%Y-%m-%d")


def _source_system(account_id: Any) -> str | None:
    """``accountId`` is prefixed by source system, e.g. ``Murex-12139``."""
    if not account_id:
        return None
    text = str(account_id)
    return text.split("-", 1)[0] if "-" in text else None


def _position_id_stem(position_id: Any) -> str | None:
    """Strip the leg suffix (e.g. ``-NEARLEG-COUNTERPART``) from a positionId."""
    if not position_id:
        return None
    return str(position_id).split("-", 1)[0]


def _normalize_call_put(value: Any) -> str | None:
    """Return ``"Call"``/``"Put"`` or ``None``. ``"Unknown"``/blank means not an option."""
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.casefold() == "unknown":
        return None
    return text.title()


def _coerce_zero_to_none(value: Any) -> float | None:
    """Turn API zeros into ``None`` for fields that mean 'not applicable'."""
    if value in (None, "", 0, 0.0):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _contract_month_year(row: Dict[str, Any]) -> str | None:
    """Prefer API ``contract`` (e.g. ``"AUG-26"``); else format ``contractDate`` YYYYMMDD."""
    contract = str(row.get("contract") or "").strip()
    if contract:
        return contract
    contract_date = str(row.get("contractDate") or "").strip()
    if len(contract_date) == 8 and contract_date.isdigit():
        parsed = pd.to_datetime(contract_date, format="%Y%m%d", errors="coerce")
        if pd.notna(parsed):
            return parsed.strftime("%b-%y").upper()
    return None


def _construct_contract_description(row: Dict[str, Any], parser_type: str) -> str:
    """Build a Contract Description when the API doesn't carry one.

    Mirrors the PDF parser's presentation: product + month + option detail.
    """
    pieces: List[str] = []
    if parser_type.startswith("FX"):
        ccy1 = str(row.get("ccy1") or "").strip()
        ccy2 = str(row.get("ccy2") or "").strip()
        if ccy1 and ccy2:
            pieces.append(f"{ccy1}/{ccy2}")
        pieces.append(parser_type)
    else:
        pieces.append(parser_type)

    month = _contract_month_year(row)
    if month:
        pieces.append(month)

    strike = _coerce_zero_to_none(row.get("strikePrice"))
    call_put = _normalize_call_put(row.get("callPut"))
    if call_put and strike is not None:
        pieces.append(f"{call_put} {strike:g}")
    elif strike is not None:
        pieces.append(f"Strike {strike:g}")

    return " ".join(pieces).strip()


# ---------------------------------------------------------------------------
# Row mappers
# ---------------------------------------------------------------------------

def _map_trade_row(row: Dict[str, Any]) -> Dict[str, Any]:
    """Map one trades-endpoint row -> parser ``Open Positions`` row.

    Output uses snake_case keys to match what ``parser.extract()`` emits, so
    ``standard_position_view_from_df`` and the grouping helpers work without
    changes. Eight redundant API fields are collapsed at this layer:

      endDate          -> dropped (== expiryDate or zero)
      counterCurrency  -> dropped (== ccy2)
      rate             -> dropped (== tradePrice)
      pay              -> dropped (== ccy1Amount)
      receivedQuantity -> dropped (== ccy2Amount)
      pnlAmount        -> dropped (== ote)
      marketPrice      -> dropped (== settlementPrice for now)
      globalId         -> kept but mapped to a single column
    """
    parser_type = _classify_type(row)
    long_qty = float(row.get("long") or 0.0)
    short_qty = float(row.get("short") or 0.0)
    net_qty = long_qty - short_qty

    return {
        # Identity
        "trade_id":             str(row.get("tradeId") or ""),
        "position_id":          str(row.get("positionId") or ""),
        "global_id":            str(row.get("globalId") or ""),
        "account_number":       row.get("accountNumber"),
        "account_name":         row.get("accountName"),
        "entity_name":          row.get("entityName"),
        "source_system":        _source_system(row.get("accountId")),

        # Classification
        "position_type":        parser_type,
        "Type":                 parser_type,
        "product":              None,  # joined from positions endpoint (listed)
        "exchange":             None,  # joined from positions endpoint (listed)
        "contract_description": _construct_contract_description(row, parser_type),
        "contract_month":       _contract_month_year(row),
        "ref_month":            _contract_month_year(row),

        # Quantities (snake_case + display-name for the view layer)
        "long":                 long_qty,
        "short":                short_qty,
        "long_qty":             long_qty,
        "short_qty":            short_qty,
        "net_qty":              net_qty,
        "quantity":             net_qty,

        # Pricing
        "trade_price":          _coerce_zero_to_none(row.get("tradePrice")),
        "settlement_price":     _coerce_zero_to_none(row.get("settlementPrice")),
        "avg_fill_price":       _coerce_zero_to_none(row.get("tradePrice")),

        # Option fields
        "option_type":          _normalize_call_put(row.get("callPut")),
        "strike":               _coerce_zero_to_none(row.get("strikePrice")),
        "strikePrice":          _coerce_zero_to_none(row.get("strikePrice")),
        "delta":                _coerce_zero_to_none(row.get("delta")),

        # FX fields
        "ccy_1":                (str(row.get("ccy1") or "").strip() or None),
        "ccy_2":                (str(row.get("ccy2") or "").strip() or None),
        "ccy_1_amount":         _coerce_zero_to_none(row.get("ccy1Amount")),
        "ccy_2_amount":         _coerce_zero_to_none(row.get("ccy2Amount")),
        "primary_currency":     (str(row.get("ccy1") or "").strip() or None),
        "secondary_currency":   (str(row.get("ccy2") or "").strip() or None),
        "primary_amount":       _coerce_zero_to_none(row.get("ccy1Amount")),
        "secondary_amount":     _coerce_zero_to_none(row.get("ccy2Amount")),
        "currency":             row.get("tradeCurrency"),
        "base_currency":        row.get("baseCurrency"),

        # OTC fields
        "trigger_barrier":      (str(row.get("triggerBarrier") or "").strip() or None),
        "settlement_frequency": (str(row.get("settlementFrequency") or "").strip() or None),

        # P&L
        "ote":                  _coerce_zero_to_none(row.get("ote")),
        "native_pnl":           _coerce_zero_to_none(row.get("nativeProfitLoss")),

        # Dates (canonical + per-instrument secondaries; see date cheat sheet)
        "trade_date":           _ms_to_iso_date(row.get("tradeDate")),
        "trade_date_iso":       _ms_to_iso_date(row.get("tradeDate")),
        "expiryDate":           _ms_to_iso_date(row.get("expiryDate")),
        "end_date":             _ms_to_iso_date(row.get("endDate")),
        "fixing_date":          _ms_to_iso_date(row.get("fixingDate")),
        "value_date":           _ms_to_iso_date(row.get("valueDate")),
        "maturity_date":        _ms_to_iso_date(row.get("maturityDate")),
        "statement_date":       _ms_to_iso_date(row.get("businessDate")),

        # Status
        "position_status":      row.get("positionStatus"),

        # Audit
        "source_section":       "API Trades",
        "source_line":          f"API trade {row.get('tradeId')} {parser_type}",
    }


def _map_position_metadata(row: Dict[str, Any]) -> Dict[str, Any]:
    """Map one positions-endpoint row -> enrichment metadata for listed trades.

    The trades endpoint omits ``productGroup``, ``masterMarketAcronym``, and
    ``masterInstrumentExchangeSymbol``. We pull them from the positions endpoint
    and join on positionId at the caller.
    """
    return {
        "position_id_key":      row.get("id"),
        "product":              row.get("productGroup"),
        "exchange":             row.get("masterMarketAcronym"),
        "contract_symbol":      row.get("masterInstrumentExchangeSymbol"),
        "master_market_name":   row.get("masterMarketName"),
        "underlier":            row.get("productUnderlyer"),
    }


# ---------------------------------------------------------------------------
# Listed trade enrichment
# ---------------------------------------------------------------------------

def _enrich_listed_trades(
    open_positions: pd.DataFrame,
    positions_metadata: pd.DataFrame,
) -> pd.DataFrame:
    """Fill product/exchange on listed trade rows by joining on positionId.

    Trades carry positionIds with leg suffixes (``-NEARLEG-COUNTERPART``); the
    positions endpoint carries the bare id. We match on the stripped stem.
    """
    if open_positions.empty or positions_metadata.empty:
        return open_positions

    # Build a position_id_stem -> metadata lookup.
    metadata = positions_metadata.dropna(subset=["position_id_key"]).copy()
    metadata["position_id_key"] = metadata["position_id_key"].astype(str)
    metadata = metadata.drop_duplicates(subset=["position_id_key"]).set_index("position_id_key")

    listed_mask = open_positions["Type"].isin(["FUT", "OPT"])
    if not listed_mask.any():
        return open_positions

    stems = open_positions.loc[listed_mask, "position_id"].map(_position_id_stem)
    for src_col, dest_col in [
        ("product",            "product"),
        ("exchange",           "exchange"),
        ("contract_symbol",    "contract_symbol"),
        ("master_market_name", "master_market_name"),
        ("underlier",          "underlier"),
    ]:
        if src_col not in metadata.columns:
            continue
        looked_up = stems.map(metadata[src_col])
        if dest_col in open_positions.columns:
            existing = open_positions.loc[listed_mask, dest_col]
            open_positions.loc[listed_mask, dest_col] = existing.where(
                existing.notna() & (existing.astype(str).str.strip() != ""),
                looked_up,
            )
        else:
            open_positions[dest_col] = pd.NA
            open_positions.loc[listed_mask, dest_col] = looked_up

    return open_positions


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

# Match parser.extract() output keys exactly so merge_extracted_tables,
# open_positions_standard_view, and grouped_positions_* work unchanged.
_PARSER_TABLE_KEYS: List[str] = [
    "Executed Trades",
    "Purchase & Sale",
    "Closed Positions",
    "Receives Delivers",
    "Journal Entries",
    "Realized Gain and Loss",
    "Realized PNL Summary",
    "Open Positions",
    "Notes",
    "Exceptions",
]


def extract_from_api(
    positions_json: List[Dict[str, Any]] | None = None,
    trades_json: List[Dict[str, Any]] | None = None,
) -> Dict[str, pd.DataFrame]:
    """Map API responses into a parser-shaped ``Dict[str, DataFrame]``.

    Args:
        positions_json: response body from ``/account/{id}/positions``
            (listed futures/options aggregated rows).
        trades_json: response body from ``/account/{id}/trades``
            (trade-level rows across FX, OTC, Futures, Options).

    Returns:
        Dict keyed exactly like ``parser.extract()``. ``"Open Positions"``
        is populated from ``trades_json``; listed rows are enriched with
        product/exchange via positionId join. Realized P&L tables remain
        empty until the central calculation endpoint exists.
    """
    positions_json = positions_json or []
    trades_json = trades_json or []

    tables: Dict[str, pd.DataFrame] = {key: pd.DataFrame() for key in _PARSER_TABLE_KEYS}

    if not trades_json:
        return tables

    open_positions = pd.DataFrame([_map_trade_row(row) for row in trades_json])

    if positions_json:
        positions_metadata = pd.DataFrame(
            [_map_position_metadata(row) for row in positions_json]
        )
        open_positions = _enrich_listed_trades(open_positions, positions_metadata)

    tables["Open Positions"] = open_positions

    # TODO: Realized Gain and Loss / Realized PNL Summary —
    #       central calculation endpoint pending.
    # TODO: Contract multiplier / lot size for Market Value and option NOV —
    #       instrument-metadata endpoint pending.

    return tables
