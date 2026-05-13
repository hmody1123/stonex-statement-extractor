# StoneX Statement Trade Extractor v6

Local Streamlit prototype for extracting trades and positions from StoneX statement PDFs.

## Run on Windows

1. Unzip this folder.
2. Double-click `run_windows.bat`.
3. Upload one or more PDF statements.
4. Download the merged Excel workbook.

## Multi-PDF merge behavior

- **Trades** are appended across uploaded PDFs and de-duplicated where possible.
- **Open positions** are statement-date snapshots.
- Default mode: **Keep latest position snapshot per account** to avoid double-counting positions when multiple statement dates are uploaded.
- Optional mode: **Append all position snapshots** for audit/time-series review.

## Workbook sheets

- Summary
- Grouped Trades
- Grouped Positions
- Grouped Positions Acct
- Statement Dates
- Executed Trades
- Purchase & Sale
- Receives Delivers
- Journal Entries
- Realized Gain and Loss
- Open Positions
- Notes
- Exceptions
- Merge Notes

## Requirements

Install dependencies with:

```bash
pip install -r requirements.txt
```

Then run:

```bash
streamlit run app.py
```


## v7 update
- Added **By Ref Month** tab and **Grouped by Ref Month** Excel sheet.
- This groups positions by account, product, option type, ref month, and unit without splitting by strike.
- The existing **Grouped Positions** tab remains strike-level detail.

## v8 dynamic grouping

The sidebar Position grouping controls now drive the Dynamic Group tab directly. Use the Quick grouping preset dropdown for common views, then add/remove columns in Group positions by. For example:

- Ref month summary: statement_date, account_number, product, option_type, ref_month, unit
- Strike detail: statement_date, account_number, product, option_type, ref_month, strike, unit
- Account summary: statement_date, account_number, product, ref_month

The Fixed tabs are still available for audit and reconciliation.


## v16 metal/LME positions update

- Added parsing for LME metal statements.
- Supports FUTURES / OPTIONS OPEN POSITIONS rows.
- Supports LME AVERAGE OPEN POSITIONS rows split across multiple PDF text lines.
- Adds delivery_date, start_date, end_date, settlement_date, ref_month, source_section, and signed market value fields where available.
- Open Positions default view now includes metal-specific fields.

## v18 update

- Adds and preserves an `exchange` field on Open Positions.
- LME/metal positions are tagged with exchange = `LME`.
- Grouped Positions defaults now include `exchange` and can dynamically group by it.
- Open Positions default display includes exchange/product/product_name fields.
