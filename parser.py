"""StoneX Daily Statement trade extractor prototype v2.

Extracts trade-related sections from StoneX Daily Statement PDFs into pandas DataFrames.
This is a local/offline parser using PyMuPDF text extraction and regex rules tuned to the
sample statement layout. It keeps source lines and page numbers for audit/reconciliation.
"""
from __future__ import annotations

import io
import re
from typing import Dict, List, Tuple

import fitz  # PyMuPDF
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment



def _daily_open_positions_backstop(pdf_bytes: bytes, include_open_positions: bool = True) -> list[dict]:
    """Parse all classic daily-statement OPEN POSITION rows across all pages.

    This is a safety net for continuation pages that repeat only the column header
    ("CONTRACT DESCRIPTION-OPEN") and not the full spaced OPEN POSITIONS title.
    """
    if not include_open_positions:
        return []
    rows: list[dict] = []
    for page_no, text in pdf_text(pdf_bytes):
        if "CONTRACT DESCRIPTION-OPEN" not in text:
            continue
        stmt_date = _statement_date(text)
        account = _account_number(text)
        for raw in text.splitlines():
            line = " ".join(raw.strip().split())
            if not line or line.startswith("-------") or line.startswith("TRADE CARD"):
                continue
            parsed = _parse_daily_open_position_line(line)
            if parsed:
                row = _base_row(parsed, stmt_date, account, page_no, line)
                row["market_value_signed"] = _signed(row.get("market_value"), row.get("drcr"))
                rows.append(row)
    return rows

MONTHS = {m: i for i, m in enumerate(["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"], 1)}
DATE_RE = re.compile(r"(?P<m>JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(?P<d>\d{1,2}),\s+(?P<y>\d{4})")
STATEMENT_DATE_RE = re.compile(r"STATEMENT DATE:\s+" + DATE_RE.pattern, re.I)
STATEMENT_DATE_DDMMM_RE = re.compile(r"Statement Date:\s+(?P<d>\d{1,2})-(?P<m>[A-Za-z]{3})-(?P<y>\d{4})", re.I)
ACCOUNT_RE = re.compile(r"ACCOUNT NUMBER:\s+(?P<account>\S+)", re.I)

TRADE_LINE_RE = re.compile(
    r"^(?P<trade_date>\d{1,2}/\d{2}/\d)\s+"
    r"(?P<card>[A-Z0-9]+)\s+"
    r"(?P<account_type>U\d)\s+"
    r"(?P<quantity>\d[\d,]*)\s+"
    r"(?P<contract_month>[A-Z]{3})\s+"
    r"(?P<contract_year>\d{2})\s+"
    r"(?P<contract_description>.+?)\s+"
    r"(?P<price>\d+(?:\.\d+)?)\s+"
    r"(?P<currency>[A-Z]{2})(?:\s+(?P<amount>[\d,]+\.\d{2})(?P<drcr>DR|CR))?\s*$"
)

RD_LINE_RE = re.compile(
    r"^(?P<trade_date>\d{1,2}/\d{2}/\d)\s+(?:(?P<card>[A-Z0-9]+)\s+)?(?P<account_type>U\d)\s+"
    r"(?P<quantity>[\d,]+)\s+"
    r"(?P<description>USTB\s+DUE\s+\d{1,2}/\d{1,2}/\d{4}(?:\s+U)?)\s+"
    r"(?P<legend>PURCHASE|REDEMPTION)\s+"
    r"(?P<currency>[A-Z]{2})\s+(?P<amount>[\d,]+\.\d{2})(?P<drcr>DR|CR)\s*$"
)

JOURNAL_RE = re.compile(
    r"^(?P<trade_date>\d{1,2}/\d{2}/\d)\s+(?P<account_type>U\d)\s+(?P<description>.+?)\s+"
    r"(?P<currency>[A-Z]{2})\s+(?P<amount>[\d,]+\.\d{2})(?P<drcr>DR|CR)\s*$"
)

OPEN_POSITION_LINE_RE = re.compile(
    r"^(?P<trade_date>\d{1,2}/\d{2}/\d)\s+(?:(?P<card>[A-Z0-9]+)\s+)?(?P<account_type>U\d)\s+"
    r"(?P<quantity>[\d,]+)\s+"
    r"(?P<contract_month>[A-Z]{3})\s+(?P<contract_year>\d{2})\s+"
    r"(?P<contract_description>.+?)\s+(?P<price>\d+(?:\.\d+)?)\s+"
    r"(?P<currency>[A-Z]{2})\s+(?P<market_value>[\d,]+\.\d{2})(?P<drcr>DR|CR)\s*$"
)

FEE_OR_AVG_KEYWORDS = ["COMMISSION", "CLEARING FEE", "AVG LONG", "AVG SHORT", "GROSS PROFIT OR LOSS"]

FRACTION_RE = re.compile(r"^\d+/\d+$")

def _parse_price_token(tokens: list[str]) -> float | None:
    """Parse StoneX prices like 4.31 1/2, .43 1/4, 5.0275, or 4.75."""
    if not tokens:
        return None
    try:
        base = float(tokens[0])
    except ValueError:
        return None
    if len(tokens) > 1 and FRACTION_RE.match(tokens[1]):
        num, den = tokens[1].split('/')
        try:
            base += float(num) / float(den) / 100.0
        except Exception:
            pass
    return base


def _split_amount_drcr(token: str | None) -> tuple[str | None, str | None]:
    if token is None:
        return None, None
    t = token.strip()
    m = re.match(r"^(?P<amount>(?:[\d,]+)?\.\d{2}|[\d,]+\.\d{2})(?P<drcr>DR|CR)?$", t)
    if not m:
        return None, None
    return m.group('amount'), m.group('drcr')


def _parse_daily_open_position_line(line: str) -> dict | None:
    toks = line.split()
    if len(toks) < 9 or not re.match(r"^\d{1,2}/\d{2}/\d$", toks[0]):
        return None
    acct_idx = None
    for i in range(1, min(len(toks), 5)):
        if re.match(r"^U\d$", toks[i]):
            acct_idx = i
            break
    if acct_idx is None or acct_idx + 2 >= len(toks):
        return None
    qty_tok = toks[acct_idx + 1]
    if not re.match(r"^[\d,]+$", qty_tok):
        return None
    currency_idx = None
    for i in range(len(toks)-2, acct_idx+2, -1):
        if re.match(r"^[A-Z]{2}$", toks[i]):
            currency_idx = i
            break
    if currency_idx is None or currency_idx + 1 >= len(toks):
        return None
    amount, drcr = _split_amount_drcr(toks[currency_idx + 1])
    if amount is None:
        return None
    if currency_idx >= 2 and FRACTION_RE.match(toks[currency_idx - 1]):
        price_tokens = toks[currency_idx-2:currency_idx]
        price_start = currency_idx - 2
    else:
        price_tokens = [toks[currency_idx - 1]]
        price_start = currency_idx - 1
    price = _parse_price_token(price_tokens)
    if price is None:
        return None
    # Some statements have a status token immediately before price (e.g., SE), but
    # others put the product symbol there (e.g., SCM TSR20RUBBR 210.0 US 1,020.00DR).
    # Treat the pre-price token as status only when it looks like a short code.
    pre_price = toks[price_start - 1] if price_start - 1 > acct_idx else None
    if pre_price and re.match(r"^[A-Z]{1,3}$", pre_price):
        st = pre_price
        desc_tokens = toks[acct_idx + 2:price_start-1]
    else:
        st = None
        desc_tokens = toks[acct_idx + 2:price_start]
    if not desc_tokens:
        return None
    contract_month = contract_year = option_type = None
    if desc_tokens[0] in {"CALL", "PUT"} and len(desc_tokens) >= 4:
        option_type = desc_tokens[0]
        contract_month = desc_tokens[1]
        contract_year = desc_tokens[2]
        contract_description = " ".join(desc_tokens)
    elif len(desc_tokens) >= 3 and re.match(r"^[A-Z]{3}$", desc_tokens[0]) and re.match(r"^\d{2}$", desc_tokens[1]):
        contract_month = desc_tokens[0]
        contract_year = desc_tokens[1]
        contract_description = " ".join(desc_tokens[2:])
    else:
        contract_description = " ".join(desc_tokens)
    return {
        "trade_date": toks[0], "card": " ".join(toks[1:acct_idx]) or None,
        "account_type": toks[acct_idx], "quantity": _to_num(qty_tok),
        "contract_month": contract_month, "contract_year": contract_year,
        "contract_description": contract_description, "option_type_raw": option_type,
        "st": st, "price": price, "price_text": " ".join(price_tokens),
        "currency": toks[currency_idx], "market_value": amount, "drcr": drcr,
    }


def _to_num(value: str | float | int | None) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return float(str(value).replace(",", ""))


def _signed(amount: str | None, drcr: str | None) -> float | None:
    value = _to_num(amount)
    if value is None:
        return None
    return -value if drcr == "DR" else value


def _statement_date(page_text: str) -> str | None:
    match = STATEMENT_DATE_RE.search(page_text)
    if match:
        return f"{int(match.group('y')):04d}-{MONTHS[match.group('m').upper()]:02d}-{int(match.group('d')):02d}"
    match = STATEMENT_DATE_DDMMM_RE.search(page_text)
    if match:
        return f"{int(match.group('y')):04d}-{MONTHS[match.group('m').upper()]:02d}-{int(match.group('d')):02d}"
    return None


def _account_number(page_text: str) -> str | None:
    match = ACCOUNT_RE.search(page_text)
    return match.group("account") if match else None


def _normalize_trade_date(short_date: str, statement_date: str | None) -> str:
    mm, dd, y1 = short_date.split("/")
    year = int(statement_date[:3] + y1) if statement_date else 2020 + int(y1)
    return f"{year:04d}-{int(mm):02d}-{int(dd):02d}"




def _normalize_any_date(date_text: str | None, statement_date: str | None = None) -> str | None:
    """Normalize either M/DD/Y or DD-Mon-YYYY to ISO."""
    if not date_text:
        return None
    s = str(date_text).strip()
    if "/" in s:
        try:
            return _normalize_trade_date(s, statement_date)
        except Exception:
            return None
    m = re.match(r"^(?P<d>\d{1,2})-(?P<m>[A-Za-z]{3})-(?P<y>\d{4})$", s)
    if m:
        return f"{int(m.group('y')):04d}-{MONTHS[m.group('m').upper()]:02d}-{int(m.group('d')):02d}"
    m = re.match(r"^(?P<d>\d{1,2})-(?P<m>[A-Za-z]{3})-(?P<y>\d{2})$", s)
    if m:
        yy = int(m.group('y'))
        year = 2000 + yy if yy < 70 else 1900 + yy
        return f"{year:04d}-{MONTHS[m.group('m').upper()]:02d}-{int(m.group('d')):02d}"
    return None


def _ref_month_from_date(date_text: str | None) -> str | None:
    iso = _normalize_any_date(date_text)
    if not iso:
        return None
    y, m, _d = iso.split('-')
    mon = list(MONTHS.keys())[int(m) - 1]
    return f"{mon}-{y[-2:]}"


LME_AMOUNT_RE = r"\(?[\d,]+(?:\.\d+)?\)?"
LME_AVG_OPEN_RE = re.compile(
    rf"^(?P<trade_date>\d{{1,2}}-[A-Za-z]{{3}}-\d{{4}})\s+"
    rf"(?P<quantity>[\d,]+)\s+"
    rf"(?P<price>ASP|3MS|(?:ASP|3MS)[+-]\d+(?:\.\d+)?)\s+"
    rf"(?P<start_date>\d{{1,2}}-[A-Za-z]{{3}}-\d{{4}})\s+"
    rf"(?P<end_date>\d{{1,2}}-[A-Za-z]{{3}}-\d{{4}})\s+"
    rf"(?P<settlement_date>\d{{1,2}}-[A-Za-z]{{3}}-\d{{4}})\s+"
    rf"(?P<currency>[A-Z]{{3}})\s+(?P<market_value>{LME_AMOUNT_RE})$"
)
LME_PRODUCT_RE = re.compile(r"^(?P<delivery_date>\d{1,2}-[A-Za-z]{3}-\d{4})\s+(?P<product>LME\s+.+)$")
LME_FUT_OPT_RE = re.compile(
    rf"^(?P<trade_date>\d{{1,2}}-[A-Za-z]{{3}}-\d{{4}})\s+"
    rf"(?P<quantity>[\d,]+)\s+"
    rf"(?P<delivery_date>\d{{1,2}}-[A-Za-z]{{3}}-\d{{2}})\s+"
    rf"(?P<product>LME\s+.+?)\s+"
    rf"(?P<trade_price>\d+(?:\.\d+)?)\s+"
    rf"(?P<price_type>TradeWhite|Trade|White)?\s*"
    rf"(?P<currency>[A-Z]{{3}})\s+(?P<market_value>{LME_AMOUNT_RE})$"
)


def _parse_lme_average_row(line: str) -> dict | None:
    m = LME_AVG_OPEN_RE.match(line)
    if not m:
        return None
    row = m.groupdict()
    row["source_section"] = "LME Average Open Positions"
    row["exchange"] = "LME"
    row["contract_description"] = None
    row["contract_month"] = None
    row["contract_year"] = None
    row["ref_month"] = _ref_month_from_date(row.get("settlement_date"))
    row["market_value_signed"] = _num_any(row.get("market_value"))
    row["trade_date_iso"] = _normalize_any_date(row.get("trade_date"))
    return row


def _parse_lme_fut_opt_row(line: str) -> dict | None:
    m = LME_FUT_OPT_RE.match(line)
    if not m:
        return None
    row = m.groupdict()
    row["source_section"] = "Futures / Options Open Positions"
    row["exchange"] = "LME"
    row["contract_description"] = row.pop("product")
    row["ref_month"] = _ref_month_from_date(row.get("delivery_date"))
    row["trade_date_iso"] = _normalize_any_date(row.get("trade_date"))
    row["market_value_signed"] = _num_any(row.get("market_value"))
    return row

def _side_from_section(section: str | None, quantity: float | None) -> str | None:
    # Statement layout does not print BUY/SELL text directly for each row; the row appears under BUY or SELL columns.
    # The parser stores source_section and leaves side blank unless a future layout rule identifies it reliably.
    return None


def pdf_text(pdf_bytes: bytes) -> List[Tuple[int, str]]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    return [(i + 1, page.get_text("text")) for i, page in enumerate(doc)]


def _base_row(groupdict: dict, stmt_date: str | None, account: str | None, page_no: int, raw_line: str) -> dict:
    row = dict(groupdict)
    row["statement_date"] = stmt_date
    row["account_number"] = account
    row["trade_date_iso"] = _normalize_any_date(row.get("trade_date"), stmt_date)
    row["quantity"] = _to_num(row.get("quantity"))
    if "price" in row:
        row["price"] = _to_num(row.get("price"))
    row["amount_signed"] = _signed(row.get("amount"), row.get("drcr"))
    row["page"] = page_no
    row["source_line"] = raw_line
    return row


# ---------------- Monthly statement parser (landscape Activity format) ----------------
DATE_DDMMMYYYY_RE = re.compile(r"^\d{2}-[A-Za-z]{3}-\d{4}$")
MONEY_RE = re.compile(r"^\(?\$?[\d,]+(?:\.\d+)?\)?$|^\$?\d+(?:\.\d+)?$|^\([\d,]+(?:\.\d+)?\)$")


def _num_any(value: str | None) -> float | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    neg = (s.startswith('(') and s.endswith(')')) or s.endswith('-') or s.startswith('-')
    s = s.replace('$','').replace(',','').replace('(','').replace(')','').strip()
    if s.endswith('-'):
        s = s[:-1].strip()
    if s.startswith('-'):
        s = s[1:].strip()
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def _page_header_info(text: str) -> tuple[str | None, str | None]:
    stmt = None
    acct = None
    m = re.search(r"Monthly Statement\s+(\d{2}-[A-Za-z]{3}-\d{4})", text)
    if not m:
        m = re.search(r"(\d{2}-[A-Za-z]{3}-\d{4})", text)
    if m:
        from datetime import datetime
        try:
            stmt = datetime.strptime(m.group(1), "%d-%b-%Y").strftime("%Y-%m-%d")
        except Exception:
            stmt = m.group(1)
    m = re.search(r"Account Number:\s*(\S+)", text)
    if m:
        acct = m.group(1)
    return stmt, acct


def _rows_from_words(page):
    # This statement is landscape/rotated: x0 is the visual row position and y0 is the visual column position.
    words = page.get_text('words')
    buckets = {}
    for x0, y0, x1, y1, w, *_ in words:
        key = round(x0 / 3.0) * 3.0
        buckets.setdefault(key, []).append((y0, w))
    out = []
    for key in sorted(buckets):
        items = sorted(buckets[key], key=lambda t: -t[0])  # visual left -> right
        out.append((key, items, ' '.join(w for _, w in items)))
    return out


def _col(items, lo, hi):
    vals = [w for y, w in items if lo <= y <= hi]
    return ' '.join(vals).strip() or None


def _desc(items):
    vals = [w for y, w in items if 385 <= y <= 545]
    return ' '.join(vals).replace(' - ', ' - ').strip() or None


def _section_limits(rows, names):
    out = {}
    for key, items, line in rows:
        for name in names:
            if name in line and name not in out:
                out[name] = key
    return out


def extract_monthly(pdf_bytes: bytes, include_open_positions: bool = True) -> Dict[str, pd.DataFrame]:
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    trades, open_pos, cash_settlements, realized, notes, exceptions = [], [], [], [], [], []
    for pno, page in enumerate(doc, start=1):
        text = page.get_text('text')
        stmt_date, account = _page_header_info(text)
        rows = _rows_from_words(page)
        limits = _section_limits(rows, ['Commodity New Trades', 'Cash Settlements', 'Commodity Cash Settlements', 'Realized Gain and Loss', 'Commodity Open Positions', 'Account Information'])
        new_start = limits.get('Commodity New Trades', -1)
        cash_start = limits.get('Cash Settlements', 10**9)
        realized_start = limits.get('Realized Gain and Loss', 10**9)
        pos_start = limits.get('Commodity Open Positions', -1)
        account_start = limits.get('Account Information', 10**9)

        for key, items, line in rows:
            date = _col(items, 735, 778)
            if not date or not DATE_DDMMMYYYY_RE.match(date):
                continue
            section = None
            if ((new_start != -1 and key > new_start) or (pno == 2 and key < cash_start)) and key < min(cash_start, realized_start, pos_start if pos_start != -1 else 10**9):
                section = 'Executed Trades'
            elif 'Trade Id Long Short Type Description' in text or (limits.get('Commodity Cash Settlements', -1) != -1 and key > limits.get('Commodity Cash Settlements', -1) and key < realized_start):
                section = 'Cash Settlements'
            elif realized_start != 10**9 and key > realized_start and (pos_start == -1 or key < pos_start):
                section = 'Realized Gain and Loss'
            elif include_open_positions and ((pos_start != -1 and key > pos_start and key < account_start) or (pno >= 5 and pno <= 11)):
                section = 'Open Positions'
            else:
                continue

            common = {
                'statement_date': stmt_date,
                'account_number': account,
                'page': pno,
                'trade_date': date,
                'trade_id': _col(items, 695, 733),
                'global_id': _col(items, 645, 682),
                'long': _num_any(_col(items, 590, 620)),
                'short': _num_any(_col(items, 545, 575)),
                'contract_description': _desc(items),
                'end_date': _col(items, 298, 340),
                'ref_month': _col(items, 258, 296),
                'source_line': line,
            }
            common['side'] = 'Long' if common['long'] is not None else ('Short' if common['short'] is not None else None)
            common['quantity'] = common['long'] if common['long'] is not None else common['short']

            if section == 'Executed Trades':
                common.update({
                    'trade_price': _num_any(_col(items, 138, 178)),
                    'commission': _num_any(_col(items, 78, 125)),
                    'premium': _num_any(_col(items, 20, 65)),
                    'source_section': section,
                })
                trades.append(common)
            elif section == 'Open Positions':
                common.update({
                    'trade_price': _num_any(_col(items, 178, 210)),
                    'market_price': _num_any(_col(items, 125, 160)),
                    'native_mv': _num_any(_col(items, 72, 105)),
                    'market_value': _num_any(_col(items, 15, 65)),
                    'source_section': section,
                })
                open_pos.append(common)
            elif section == 'Cash Settlements':
                common.update({'source_section': section, 'trade_price': _num_any(_col(items, 138, 178)), 'cash_amount': _num_any(_col(items, 80, 130))})
                cash_settlements.append(common)
            elif section == 'Realized Gain and Loss':
                common.update({'source_section': section, 'trade_price': _num_any(_col(items, 138, 178)), 'cash_flow': _num_any(_col(items, 20, 65))})
                realized.append(common)

    tables: Dict[str, pd.DataFrame] = {
        'Executed Trades': pd.DataFrame(trades),
        'Purchase & Sale': pd.DataFrame(),
        'Receives Delivers': pd.DataFrame(cash_settlements),
        'Journal Entries': pd.DataFrame(),
        'Realized Gain and Loss': pd.DataFrame(realized),
        'Open Positions': pd.DataFrame(open_pos),
        'Notes': pd.DataFrame(notes),
        'Exceptions': pd.DataFrame(exceptions),
    }
    tables['Summary'] = build_summary(tables)
    return tables


def looks_like_monthly_statement(pdf_bytes: bytes) -> bool:
    try:
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        text = doc[0].get_text('text')
        return ('Commodity New Trades' in text or 'Open Positions and Market Values' in text or 'Monthly Statement' in text)
    except Exception:
        return False




def _is_dd_mmm_yyyy(s: str) -> bool:
    return bool(re.match(r"^\d{1,2}-[A-Za-z]{3}-\d{4}$", str(s).strip()))


def _is_dd_mmm_yy(s: str) -> bool:
    return bool(re.match(r"^\d{1,2}-[A-Za-z]{3}-\d{2}$", str(s).strip()))


def _is_amount_line(s: str) -> bool:
    return bool(re.match(r"^\(?[\d,]+(?:\.\d+)?\)?$", str(s).strip()))


def _is_ccy_line(s: str) -> bool:
    return bool(re.match(r"^[A-Z]{3}$", str(s).strip()))


def _parse_lme_open_positions_from_lines(lines: list[str], stmt_date: str | None, account: str | None, page_no: int) -> list[dict]:
    """Parse metal/LME open positions where PyMuPDF splits each visual row into multiple text lines."""
    rows: list[dict] = []
    section: str | None = None
    i = 0
    while i < len(lines):
        line = lines[i]
        if line == "FUTURES / OPTIONS OPEN POSITIONS":
            section = "fut_opt"
            i += 1
            continue
        if line == "LME AVERAGE OPEN POSITIONS":
            section = "lme_avg"
            i += 1
            continue
        if line.startswith("Diagnostic Key:") and "OPEN POSITIONS" in line:
            # wait for the actual table title
            i += 1
            continue
        if line.startswith("Diagnostic Key:") and section is not None:
            section = None
            i += 1
            continue

        if section == "fut_opt" and _is_dd_mmm_yyyy(line):
            try:
                qty = lines[i + 1]
                delivery_product = lines[i + 2]
                trade_price = lines[i + 3]
                j = i + 4
                price_type = None
                if j < len(lines) and not _is_ccy_line(lines[j]):
                    price_type = lines[j]
                    j += 1
                ccy = lines[j]
                amount = lines[j + 1]
                dm = re.match(r"^(?P<delivery_date>\d{1,2}-[A-Za-z]{3}-\d{2})\s+(?P<product>LME\s+.+)$", delivery_product)
                if dm and re.match(r"^[\d,]+$", qty) and _is_ccy_line(ccy) and _is_amount_line(amount):
                    raw_line = " ".join(lines[i:j+2])
                    row = {
                        "trade_date": line,
                        "trade_date_iso": _normalize_any_date(line, stmt_date),
                        "quantity": _num_any(qty),
                        "delivery_date": dm.group("delivery_date"),
                        "contract_description": dm.group("product"),
                        "exchange": "LME",
                        "ref_month": _ref_month_from_date(dm.group("delivery_date")),
                        "trade_price": _num_any(trade_price),
                        "price_type": price_type,
                        "currency": ccy,
                        "market_value": _num_any(amount),
                        "market_value_signed": _num_any(amount),
                        "statement_date": stmt_date,
                        "account_number": account,
                        "page": page_no,
                        "source_section": "Futures / Options Open Positions",
                        "source_line": raw_line,
                    }
                    rows.append(row)
                    i = j + 2
                    continue
            except Exception:
                pass

        if section == "lme_avg" and _is_dd_mmm_yyyy(line):
            try:
                qty = lines[i + 1]
                price_start = lines[i + 2]
                end_date = lines[i + 3]
                settlement_date = lines[i + 4]
                if not re.match(r"^[\d,]+$", qty) or not _is_dd_mmm_yyyy(end_date) or not _is_dd_mmm_yyyy(settlement_date):
                    i += 1
                    continue
                ps = price_start.split()
                if len(ps) < 2 or not _is_dd_mmm_yyyy(ps[-1]):
                    i += 1
                    continue
                price = " ".join(ps[:-1])
                start_date = ps[-1]
                # Amount/currency are usually the next two text lines, but the order can be amount then ccy.
                amount = None
                ccy = None
                for k in range(i + 5, min(i + 9, len(lines))):
                    if amount is None and _is_amount_line(lines[k]):
                        amount = lines[k]
                    if ccy is None and _is_ccy_line(lines[k]):
                        ccy = lines[k]
                product = None
                delivery_date = None
                for k in range(i + 5, min(i + 22, len(lines))):
                    pm = LME_PRODUCT_RE.match(lines[k])
                    if pm and pm.group("delivery_date") == settlement_date:
                        delivery_date = pm.group("delivery_date")
                        product = pm.group("product")
                        break
                if amount is not None and ccy is not None:
                    raw_line = " ".join(lines[i:min(i+8, len(lines))])
                    row = {
                        "trade_date": line,
                        "trade_date_iso": _normalize_any_date(line, stmt_date),
                        "quantity": _num_any(qty),
                        "delivery_date": delivery_date or settlement_date,
                        "contract_description": product,
                        "exchange": "LME",
                        "price": price,
                        "trade_price": price,
                        "start_date": start_date,
                        "end_date": end_date,
                        "settlement_date": settlement_date,
                        "ref_month": _ref_month_from_date(settlement_date),
                        "currency": ccy,
                        "market_value": _num_any(amount),
                        "market_value_signed": _num_any(amount),
                        "statement_date": stmt_date,
                        "account_number": account,
                        "page": page_no,
                        "source_section": "LME Average Open Positions",
                        "source_line": raw_line,
                    }
                    rows.append(row)
                    i += 5
                    continue
            except Exception:
                pass

        i += 1
    return rows

def extract(pdf_bytes: bytes, include_open_positions: bool = True) -> Dict[str, pd.DataFrame]:
    if looks_like_monthly_statement(pdf_bytes):
        return extract_monthly(pdf_bytes, include_open_positions=include_open_positions)

    executed: List[dict] = []
    purchase_sale: List[dict] = []
    receives_delivers: List[dict] = []
    journals: List[dict] = []
    open_positions: List[dict] = []
    notes: List[dict] = []
    exceptions: List[dict] = []

    pending_lme_avg: dict | None = None

    for page_no, text in pdf_text(pdf_bytes):
        stmt_date = _statement_date(text)
        account = _account_number(text)
        section: str | None = None
        saw_stmt = bool(stmt_date)
        if include_open_positions and "CONTRACT DESCRIPTION-OPEN" in text:
            section = "open_positions"

        normalized_lines = [" ".join(raw.strip().split()) for raw in text.splitlines() if " ".join(raw.strip().split())]
        if include_open_positions and ("LME AVERAGE OPEN POSITIONS" in text or "FUTURES / OPTIONS OPEN POSITIONS" in text):
            open_positions.extend(_parse_lme_open_positions_from_lines(normalized_lines, stmt_date, account, page_no))

        for raw in text.splitlines():
            line = " ".join(raw.strip().split())
            if not line:
                continue

            if "THE FOLLOWING TRADES HAVE BEEN MADE" in line:
                section = "executed"
                continue
            if "P U R C H A S E" in line and "S A L E" in line:
                section = "purchase_sale"
                continue
            if "THE FOLLOWING ITEMS HAVE BEEN RECEIVED/DELIVERED" in line:
                section = "receives_delivers"
                continue
            if "THE FOLLOWING JOURNAL ENTRIES" in line:
                section = "journal"
                continue
            if "FUTURES / OPTIONS OPEN POSITIONS" in line:
                section = "lme_fut_opt_open" if include_open_positions else None
                continue
            if "LME AVERAGE OPEN POSITIONS" in line:
                section = "lme_average_open" if include_open_positions else None
                continue
            if "O P E N P O S I T I O N S" in line:
                section = "open_positions" if include_open_positions else None
                continue
            if line.startswith("*USD") or line.startswith("**TOTAL") or line.startswith("MARGIN CALL") or line.startswith("SUMMARY"):
                section = None
                continue
            if line.startswith("-------") or line.startswith("TRADE CARD") or line.startswith("TRADE AT"):
                continue

            if section in {"executed", "purchase_sale"}:
                match = TRADE_LINE_RE.match(line)
                if match:
                    row = _base_row(match.groupdict(), stmt_date, account, page_no, line)
                    row["source_section"] = "Executed Trades" if section == "executed" else "Purchase & Sale"
                    row["side"] = _side_from_section(section, row.get("quantity"))
                    (executed if section == "executed" else purchase_sale).append(row)
                elif any(keyword in line for keyword in FEE_OR_AVG_KEYWORDS):
                    notes.append({"statement_date": stmt_date, "account_number": account, "page": page_no, "section": section, "text": line})
                elif re.match(r"^\d{1,2}/\d{2}/\d\s+", line):
                    exceptions.append({"statement_date": stmt_date, "account_number": account, "page": page_no, "section": section, "reason": "Unmatched trade-like line", "source_line": line})

            elif section == "receives_delivers":
                match = RD_LINE_RE.match(line)
                if match:
                    row = _base_row(match.groupdict(), stmt_date, account, page_no, line)
                    receives_delivers.append(row)
                elif re.match(r"^\d{1,2}/\d{2}/\d\s+", line):
                    exceptions.append({"statement_date": stmt_date, "account_number": account, "page": page_no, "section": section, "reason": "Unmatched receive/deliver line", "source_line": line})

            elif section == "journal":
                match = JOURNAL_RE.match(line)
                if match:
                    row = _base_row(match.groupdict(), stmt_date, account, page_no, line)
                    journals.append(row)
                elif re.match(r"^\d{1,2}/\d{2}/\d\s+", line):
                    exceptions.append({"statement_date": stmt_date, "account_number": account, "page": page_no, "section": section, "reason": "Unmatched journal line", "source_line": line})

            elif section == "lme_fut_opt_open" and include_open_positions:
                parsed = _parse_lme_fut_opt_row(line)
                if parsed:
                    row = _base_row(parsed, stmt_date, account, page_no, line)
                    row["quantity"] = _num_any(parsed.get("quantity"))
                    row["market_value"] = _num_any(parsed.get("market_value"))
                    row["market_value_signed"] = _num_any(parsed.get("market_value"))
                    row["trade_price"] = _num_any(parsed.get("trade_price"))
                    open_positions.append(row)
                elif line.startswith("Total") or line.startswith("Last Traded Date") or line.startswith("Trade Date"):
                    pass

            elif section == "lme_average_open" and include_open_positions:
                prod_match = LME_PRODUCT_RE.match(line)
                if prod_match and pending_lme_avg is not None:
                    pending_lme_avg["delivery_date"] = prod_match.group("delivery_date")
                    pending_lme_avg["contract_description"] = prod_match.group("product")
                    row = _base_row(pending_lme_avg, stmt_date, account, page_no, pending_lme_avg.get("source_line", line))
                    row["quantity"] = _num_any(pending_lme_avg.get("quantity"))
                    row["market_value"] = _num_any(pending_lme_avg.get("market_value"))
                    row["market_value_signed"] = _num_any(pending_lme_avg.get("market_value"))
                    row["trade_price"] = pending_lme_avg.get("price")
                    open_positions.append(row)
                    pending_lme_avg = None
                else:
                    parsed = _parse_lme_average_row(line)
                    if parsed:
                        parsed["source_line"] = line
                        pending_lme_avg = parsed
                    elif line.startswith("Total") or line.startswith("Priced Lots") or line.startswith("Unpriced Lots") or "P&L" in line or line.startswith("Trade Date"):
                        pass

            elif section == "open_positions" and include_open_positions:
                parsed = _parse_daily_open_position_line(line)
                if parsed:
                    row = _base_row(parsed, stmt_date, account, page_no, line)
                    row["market_value_signed"] = _signed(row.get("market_value"), row.get("drcr"))
                    open_positions.append(row)
                else:
                    match = OPEN_POSITION_LINE_RE.match(line)
                    if match:
                        row = _base_row(match.groupdict(), stmt_date, account, page_no, line)
                        row["market_value_signed"] = _signed(row.get("market_value"), row.get("drcr"))
                        open_positions.append(row)
                    elif any(keyword in line for keyword in ["AVG LONG", "AVG SHORT", "CLOSE", "S.P."]):
                        notes.append({"statement_date": stmt_date, "account_number": account, "page": page_no, "section": section, "text": line})
                    elif re.match(r"^\d{1,2}/\d{2}/\d\s+", line):
                        exceptions.append({"statement_date": stmt_date, "account_number": account, "page": page_no, "section": section, "reason": "Unmatched open position line", "source_line": line})

        if not saw_stmt:
            exceptions.append({"statement_date": None, "account_number": account, "page": page_no, "section": None, "reason": "Statement date not found", "source_line": ""})

    # Backstop parse ensures continuation pages with only the classic open-position
    # column header are included. Avoid duplicates by source line + page.
    backstop_positions = _daily_open_positions_backstop(pdf_bytes, include_open_positions=include_open_positions)
    if backstop_positions:
        existing_keys = {(r.get("page"), r.get("source_line")) for r in open_positions}
        for r in backstop_positions:
            key = (r.get("page"), r.get("source_line"))
            if key not in existing_keys:
                open_positions.append(r)
                existing_keys.add(key)

    tables: Dict[str, pd.DataFrame] = {
        "Executed Trades": pd.DataFrame(executed),
        "Purchase & Sale": pd.DataFrame(purchase_sale),
        "Receives Delivers": pd.DataFrame(receives_delivers),
        "Journal Entries": pd.DataFrame(journals),
        "Realized Gain and Loss": pd.DataFrame(),
        "Realized PNL Summary": _parse_realized_pnl_summary_from_text(pdf_bytes),
        "Open Positions": pd.DataFrame(open_positions),
        "Notes": pd.DataFrame(notes),
        "Exceptions": pd.DataFrame(exceptions),
    }
    tables = enrich_open_positions_metadata(tables)
    tables["Summary"] = build_summary(tables)
    return tables


def enrich_open_positions_metadata(tables: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    """Add exchange/product metadata to raw Open Positions without changing row count."""
    pos = tables.get("Open Positions", pd.DataFrame())
    if pos is None or pos.empty or "contract_description" not in pos.columns:
        return tables
    df = pos.copy()
    parsed = df["contract_description"].apply(lambda x: pd.Series(parse_contract_product(x)))
    for col in parsed.columns:
        if col not in df.columns:
            df[col] = parsed[col]
        else:
            current = df[col]
            missing = current.isna() | (current.astype(str).str.strip() == "") | (current.astype(str).str.lower() == "none")
            df[col] = current.where(~missing, parsed[col])
    tables["Open Positions"] = df
    return tables




def _parse_realized_pnl_summary_from_text(pdf_bytes: bytes) -> pd.DataFrame:
    """Parse summary-level realized P&L rows when no row-level realized detail exists.

    Classic daily statements can show only a summary line such as:
      REALIZED PROFIT & LOSS
       USD 431,445.00- 440,805.00-
    where values are M-T-D and Y-T-D. This function captures those rows.
    """
    rows: list[dict] = []
    amount_pat = r"(?:[\d,]+(?:\.\d+)?-?|\([\d,]+(?:\.\d+)?\))"
    for page_no, text in pdf_text(pdf_bytes):
        stmt_date = _statement_date(text)
        account = _account_number(text)
        lines = [" ".join(x.strip().split()) for x in text.splitlines() if x.strip()]
        for i, line in enumerate(lines):
            if "REALIZED PROFIT" in line.upper() and "LOSS" in line.upper():
                # Usually the next non-empty line is: USD 431,445.00- 440,805.00-
                for j in range(i + 1, min(i + 6, len(lines))):
                    nxt = lines[j]
                    m = re.match(rf"^(?P<currency>[A-Z]{{3}})\s+(?P<mtd>{amount_pat})\s+(?P<ytd>{amount_pat})\s*$", nxt)
                    if m:
                        rows.append({
                            "statement_date": stmt_date,
                            "account_number": account,
                            "currency": m.group("currency"),
                            "mtd_realized_pnl": _num_any(m.group("mtd")),
                            "ytd_realized_pnl": _num_any(m.group("ytd")),
                            "source_sheet": "Statement Summary",
                            "source_section": "REALIZED PROFIT & LOSS",
                            "page": page_no,
                            "source_line": nxt,
                        })
                        break
    return pd.DataFrame(rows)

def build_summary(tables: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = []
    for name, df in tables.items():
        if name == "Summary":
            continue
        rows.append({"sheet": name, "rows": len(df), "columns": len(df.columns) if not df.empty else 0})

    executed = tables.get("Executed Trades", pd.DataFrame())
    if not executed.empty and "quantity" in executed.columns:
        group_cols = [c for c in ["statement_date", "account_type", "contract_month", "contract_year", "contract_description"] if c in executed.columns]
        grouped = executed.groupby(group_cols, dropna=False)["quantity"].sum().reset_index() if group_cols else pd.DataFrame()
        rows.append({"sheet": "Executed Trades total quantity", "rows": float(executed["quantity"].sum()), "columns": ""})
        rows.append({"sheet": "Executed Trades grouped rows", "rows": len(grouped), "columns": len(grouped.columns) if not grouped.empty else 0})
    return pd.DataFrame(rows)


def grouped_trades(tables: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    executed = tables.get("Executed Trades", pd.DataFrame())
    if executed.empty:
        return pd.DataFrame()
    group_cols = [c for c in ["statement_date", "trade_date_iso", "account_number", "account_type", "contract_month", "contract_year", "contract_description", "price", "currency"] if c in executed.columns]
    agg = {"quantity": "sum"}
    if "amount_signed" in executed.columns:
        agg["amount_signed"] = "sum"
    return executed.groupby(group_cols, dropna=False).agg(agg).reset_index().sort_values(group_cols)



def parse_contract_product(desc: str | None) -> dict:
    """Parse a StoneX contract description into grouping fields."""
    desc = "" if desc is None else str(desc).strip()
    exchanges = "LME|SCM|CBOT|CBT|NYMEX|CME|ICE|MGEX"
    if desc.upper().startswith("LME "):
        parts = desc.split()
        exchange = "LME"
        # Last token is usually metal currency (USD/EUR); keep it as unit/currency and exclude from product name.
        unit = parts[-1] if len(parts) > 2 and re.match(r"^[A-Z]{3}|\$$", parts[-1]) else None
        product_name = " ".join(parts[1:-1] if unit else parts[1:])
        return {"product": f"LME {product_name}".strip(), "exchange": exchange, "product_name": product_name, "strike": None, "option_type": "Other", "unit": unit}
    if desc.upper().startswith("SCM "):
        parts = desc.split(maxsplit=1)
        product_name = parts[1] if len(parts) > 1 else ""
        return {"product": f"SCM {product_name}".strip(), "exchange": "SCM", "product_name": product_name, "strike": None, "option_type": "Other", "unit": None}
    m = re.search(rf"\b({exchanges})\s+(.+?)(?=\s+\d+(?:\.\d+)?)", desc)
    if m:
        exchange = m.group(1)
        product_name = " ".join(m.group(2).split())
        product = f"{exchange} {product_name}".strip()
    else:
        m2 = re.search(rf"\b({exchanges})\s+([A-Za-z ]+)", desc)
        exchange = m2.group(1) if m2 else None
        product_name = " ".join(m2.group(2).split()) if m2 else None
        product = f"{exchange} {product_name}".strip() if m2 else "Unknown"

    strike_value = None
    strike_match = re.search(r"\b(\d+(?:\.\d+)?)\s+Euro\s+Option\b", desc)
    if strike_match:
        strike_value = _num_any(strike_match.group(1))
    elif re.match(r"^(CALL|PUT)\b", desc, re.I):
        mstrike = re.search(r"\b(CBT|CBOT|NYMEX|CME|ICE|MGEX)\s+[A-Z ]+\s+(\d+(?:\.\d+)?)\b", desc, re.I)
        if mstrike:
            strike_value = _num_any(mstrike.group(2))

    option_type = "Call" if re.search(r"\bCall\b", desc, re.I) else ("Put" if re.search(r"\bPut\b", desc, re.I) else "Other")
    unit_match = re.search(r"\bUSD/([A-Z]+)\b", desc)

    return {
        "product": product,
        "exchange": exchange,
        "product_name": product_name,
        "strike": strike_value,
        "option_type": option_type,
        "unit": unit_match.group(1) if unit_match else None,
    }


def _prepared_positions_for_grouping(tables: Dict[str, pd.DataFrame]) -> tuple[pd.DataFrame, str | None, str | None]:
    """Normalize open positions before any grouping so app selections are truly dynamic."""
    pos = tables.get("Open Positions", pd.DataFrame())
    if pos is None or pos.empty:
        return pd.DataFrame(), None, None

    df = pos.copy()
    if "contract_description" not in df.columns:
        return pd.DataFrame(), None, None

    parsed = df["contract_description"].apply(lambda x: pd.Series(parse_contract_product(x)))
    for col in parsed.columns:
        if col not in df.columns:
            df[col] = parsed[col]
        else:
            current = df[col]
            missing = current.isna() | (current.astype(str).str.strip() == "") | (current.astype(str).str.lower() == "none")
            df[col] = current.where(~missing, parsed[col])

    if "long" in df.columns or "short" in df.columns:
        long_qty = pd.to_numeric(df.get("long", 0), errors="coerce").fillna(0)
        short_qty = pd.to_numeric(df.get("short", 0), errors="coerce").fillna(0)
        df["net_qty"] = long_qty - short_qty
        df["long_qty"] = long_qty
        df["short_qty"] = short_qty
    elif "quantity" in df.columns:
        qty = pd.to_numeric(df["quantity"], errors="coerce").fillna(0)
        if "side" in df.columns:
            side = df["side"].fillna("").astype(str).str.lower()
            df["net_qty"] = qty.where(~side.str.contains("short|sell"), -qty)
        else:
            df["net_qty"] = qty
        df["long_qty"] = df["net_qty"].where(df["net_qty"] > 0, 0)
        df["short_qty"] = (-df["net_qty"]).where(df["net_qty"] < 0, 0)
    else:
        df["net_qty"] = 0.0
        df["long_qty"] = 0.0
        df["short_qty"] = 0.0

    if "ref_month" not in df.columns:
        if "contract_month" in df.columns and "contract_year" in df.columns:
            df["ref_month"] = df["contract_month"].astype(str) + "-" + df["contract_year"].astype(str)
        elif "contract_month" in df.columns:
            df["ref_month"] = df["contract_month"].astype(str)
        else:
            df["ref_month"] = None

    if "end_date" in df.columns:
        extracted_ref = df["end_date"].astype(str).str.extract(r"\b([A-Z]{3}-\d{2})\b", expand=False)
        df["ref_month"] = df["ref_month"].where(df["ref_month"].notna() & (df["ref_month"].astype(str) != "None"), extracted_ref)
        df["end_date"] = df["end_date"].astype(str).str.replace(r"\s+[A-Z]{3}-\d{2}\b", "", regex=True)

    price_col = "trade_price" if "trade_price" in df.columns else ("price" if "price" in df.columns else None)
    mv_col = "market_value_signed" if "market_value_signed" in df.columns else ("market_value" if "market_value" in df.columns else None)
    if price_col:
        df[price_col] = pd.to_numeric(df[price_col], errors="coerce")
    if mv_col:
        df[mv_col] = df[mv_col].apply(_num_any)

    return df, price_col, mv_col


def _group_prepared_positions(df: pd.DataFrame, group_cols: list[str], price_col: str | None, mv_col: str | None) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    group_cols = [c for c in group_cols if c in df.columns]
    if not group_cols:
        group_cols = [c for c in ["account_number", "product", "option_type", "ref_month"] if c in df.columns]
    if not group_cols:
        return df

    df = df.copy()
    df["_position_rows"] = 1

    agg = {
        "long_qty": "sum",
        "short_qty": "sum",
        "net_qty": "sum",
        "_position_rows": "sum",
    }
    if price_col:
        df["_weighted_price_qty"] = df[price_col].fillna(0) * df["net_qty"].abs()
        agg["_weighted_price_qty"] = "sum"
    if mv_col:
        agg[mv_col] = "sum"

    grouped = df.groupby(group_cols, dropna=False).agg(agg).reset_index()
    rename_map = {"_position_rows": "position_rows"}
    if mv_col:
        rename_map[mv_col] = "market_value"
    grouped = grouped.rename(columns=rename_map)

    if price_col:
        qty_abs = df.groupby(group_cols, dropna=False)["net_qty"].apply(lambda s: s.abs().sum()).reset_index(name="_abs_qty")
        grouped = grouped.merge(qty_abs, on=group_cols, how="left")
        grouped["avg_trade_price"] = grouped.apply(lambda r: r["_weighted_price_qty"] / r["_abs_qty"] if r.get("_abs_qty", 0) else None, axis=1)
        grouped = grouped.drop(columns=["_weighted_price_qty", "_abs_qty"], errors="ignore")

    return grouped.sort_values(group_cols).reset_index(drop=True)


def grouped_positions(tables: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Default strike-level grouped open positions."""
    df, price_col, mv_col = _prepared_positions_for_grouping(tables)
    group_cols = ["statement_date", "account_number", "exchange", "product", "option_type", "ref_month", "strike", "unit"]
    return _group_prepared_positions(df, group_cols, price_col, mv_col)


def grouped_positions_by_ref_month(tables: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Grouped positions by ref month, including strike price."""
    df, price_col, mv_col = _prepared_positions_for_grouping(tables)
    group_cols = ["statement_date", "account_number", "exchange", "product", "option_type", "ref_month", "strike", "unit"]
    return _group_prepared_positions(df, group_cols, price_col, mv_col)


def grouped_positions_by_account(tables: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Account-first grouped open positions without strike detail."""
    df, price_col, mv_col = _prepared_positions_for_grouping(tables)
    group_cols = ["statement_date", "account_number", "exchange", "product", "option_type", "ref_month", "unit"]
    return _group_prepared_positions(df, group_cols, price_col, mv_col)


def grouped_positions_custom(tables: Dict[str, pd.DataFrame], group_cols: list[str] | None = None) -> pd.DataFrame:
    """Group positions using caller-selected dimensions from the sidebar."""
    df, price_col, mv_col = _prepared_positions_for_grouping(tables)
    if not group_cols:
        group_cols = ["account_number", "product", "option_type", "ref_month"]
    return _group_prepared_positions(df, group_cols, price_col, mv_col)


def merge_extracted_tables(tables_list: list[Dict[str, pd.DataFrame]], position_mode: str = "append_all") -> Dict[str, pd.DataFrame]:
    """Merge extracted tables from multiple PDFs.

    position_mode:
      - append_all: keep every open-position row from every PDF. This is now the only UI behavior.
    """
    sheet_names = [
        "Executed Trades", "Purchase & Sale", "Receives Delivers", "Journal Entries",
        "Realized Gain and Loss", "Realized PNL Summary", "Open Positions", "Notes", "Exceptions"
    ]
    merged: Dict[str, pd.DataFrame] = {}
    for name in sheet_names:
        frames = []
        for i, tables in enumerate(tables_list, start=1):
            df = tables.get(name, pd.DataFrame())
            if df is not None and not df.empty:
                temp = df.copy()
                if "source_pdf_index" not in temp.columns:
                    temp["source_pdf_index"] = i
                frames.append(temp)
        merged[name] = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()

    # De-duplicate trades where the same PDF/statement is uploaded twice.
    for name in ["Executed Trades", "Purchase & Sale", "Receives Delivers", "Journal Entries", "Realized Gain and Loss"]:
        df = merged.get(name, pd.DataFrame())
        if df.empty:
            continue
        keys = [c for c in ["account_number", "statement_date", "trade_date", "trade_date_iso", "trade_id", "global_id", "contract_description", "ref_month", "quantity", "trade_price", "price"] if c in df.columns]
        if keys:
            merged[name] = df.drop_duplicates(subset=keys, keep="first").reset_index(drop=True)

    pos = merged.get("Open Positions", pd.DataFrame())
    if pos is not None and not pos.empty:
        # Open-position rows are card/trade-line level. Do NOT de-duplicate only on
        # product/qty/price/value because statements can legitimately contain many
        # separate cards with identical qty/price/value. The rubber/SCM statement is
        # a good example: many one-lot rows at the same price.
        #
        # This strict key removes true duplicate uploads/duplicate extracted lines,
        # while preserving unique cards/rows on continuation pages.
        strict_keys = [c for c in [
            "account_number", "statement_date", "trade_date", "trade_date_iso",
            "card", "account_type", "contract_month", "contract_year",
            "contract_description", "ref_month", "delivery_date", "settlement_date",
            "long", "short", "quantity", "trade_price", "price",
            "currency", "market_value", "market_value_signed", "drcr",
            "page", "source_line"
        ] if c in pos.columns]
        if strict_keys:
            merged["Open Positions"] = pos.drop_duplicates(subset=strict_keys, keep="first").reset_index(drop=True)
            pos = merged["Open Positions"]

    # Add a merge summary/notes sheet.
    note_rows = []
    for i, tables in enumerate(tables_list, start=1):
        # The app adds source_pdf_name to tables in app.py; fall back to index.
        summary = tables.get("Summary", pd.DataFrame())
        source_name = None
        for df in tables.values():
            if isinstance(df, pd.DataFrame) and not df.empty and "source_pdf" in df.columns:
                source_name = str(df["source_pdf"].iloc[0])
                break
        note_rows.append({"source_pdf_index": i, "source_pdf": source_name or f"PDF {i}", "position_mode": position_mode})
    merged["Merge Notes"] = pd.DataFrame(note_rows)
    merged["Summary"] = build_summary(merged)
    return merged


def realized_pnl_summary(tables: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Absolute-basic realized PNL view.

    Uses row-level Realized Gain and Loss / Purchase & Sale where available.
    If the statement only has a summary-level realized P&L section, includes MTD/YTD rows.
    """
    frames = []
    realized = tables.get("Realized Gain and Loss", pd.DataFrame())
    if realized is not None and not realized.empty:
        df = realized.copy()
        df["source_sheet"] = "Realized Gain and Loss"
        if "cash_flow" in df.columns:
            df["realized_pnl"] = df["cash_flow"].apply(_num_any)
        elif "amount_signed" in df.columns:
            df["realized_pnl"] = df["amount_signed"].apply(_num_any)
        df["pnl_view"] = "Detail"
        frames.append(df)

    ps = tables.get("Purchase & Sale", pd.DataFrame())
    if ps is not None and not ps.empty:
        df = ps.copy()
        df["source_sheet"] = "Purchase & Sale"
        if "amount_signed" in df.columns:
            df["realized_pnl"] = df["amount_signed"].apply(_num_any)
        elif "amount" in df.columns:
            df["realized_pnl"] = [_signed(a, d) for a, d in zip(df.get("amount"), df.get("drcr"))]
        df["pnl_view"] = "Detail"
        frames.append(df)

    summary = tables.get("Realized PNL Summary", pd.DataFrame())
    if summary is not None and not summary.empty:
        df = summary.copy()
        df["source_sheet"] = "Statement Summary"
        df["pnl_view"] = "Summary"
        # Use MTD as the primary realized_pnl displayed in total metric.
        if "mtd_realized_pnl" in df.columns:
            df["realized_pnl"] = df["mtd_realized_pnl"].apply(_num_any)
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    detail = pd.concat(frames, ignore_index=True, sort=False)
    if "contract_description" in detail.columns:
        parsed = detail["contract_description"].apply(parse_contract_product).apply(pd.Series)
        for col in parsed.columns:
            if col not in detail.columns or detail[col].isna().all():
                detail[col] = parsed[col]

    wanted = [
        "pnl_view", "statement_date", "account_number", "currency",
        "mtd_realized_pnl", "ytd_realized_pnl", "realized_pnl",
        "trade_date", "trade_date_iso", "trade_id", "contract_description",
        "quantity", "long", "short", "trade_price", "price", "cash_flow", "amount_signed",
        "source_sheet", "source_section", "source_pdf", "page", "source_line"
    ]
    cols = [c for c in wanted if c in detail.columns]
    return detail[cols].copy()


def statement_dates_by_account(tables: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Return statement-date coverage by account for audit after multi-PDF merge."""
    frames = []
    for name in ["Executed Trades", "Open Positions", "Purchase & Sale", "Receives Delivers", "Realized Gain and Loss", "Realized PNL Summary"]:
        df = tables.get(name, pd.DataFrame())
        if df is not None and not df.empty:
            cols = [c for c in ["source_pdf", "account_number", "statement_date"] if c in df.columns]
            if cols:
                temp = df[cols].drop_duplicates().copy()
                temp["sheet"] = name
                frames.append(temp)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True, sort=False).drop_duplicates()
    return out.sort_values([c for c in ["account_number", "statement_date", "source_pdf", "sheet"] if c in out.columns]).reset_index(drop=True)

def to_excel_bytes(tables: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    ordered = {
        "Summary": tables.get("Summary", pd.DataFrame()),
        "Grouped Trades": grouped_trades(tables),
        "Grouped Positions": grouped_positions_by_ref_month(tables),
        "Realized PNL": realized_pnl_summary(tables),
        "Statement Dates": statement_dates_by_account(tables),
    }
    for name, df in tables.items():
        if name != "Summary":
            ordered[name] = df

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in ordered.items():
            if df is None or df.empty:
                df = pd.DataFrame([{"message": "No rows extracted"}])
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
            ws = writer.book[sheet_name[:31]]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="4F81BD")
                cell.alignment = Alignment(horizontal="center")
            for col in ws.columns:
                width = min(max(len(str(cell.value or "")) for cell in col) + 2, 60)
                ws.column_dimensions[col[0].column_letter].width = width
    output.seek(0)
    return output.getvalue()


def extract_to_excel(pdf_bytes: bytes) -> bytes:
    return to_excel_bytes(extract(pdf_bytes))
