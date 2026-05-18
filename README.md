# MyStoneX Positions

Version: v76 FX additional types parser

This Streamlit prototype parses StoneX statement PDFs into trade-level rows, aggregated positions, drill-down details, data quality checks, and exports.

## v76 change

- Added support/routing for StoneX Markets LLC FX-only statements where Account Summary and Account Information pages appear before the FX Open Positions section.
- Parses FX SPOT, FX FWD, and FX Swap rows from FX Spot/Forward Open Positions.
- Uses the existing FX grouping logic for these additional FX types.
- Retains v75 NDO option-aware grouping.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

