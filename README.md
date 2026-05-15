# StoneX Statement Trade Extractor

Version v67: OTC Trigger/Barrier, Ref Price, Original Quantity, and accumulator strike mapping.

Updates in this version:
- Trigger/Barrier now stores only the trigger/barrier price.
- Added Ref Price from BP values.
- Added Original Quantity from OQ values.
- Maps OTC accumulator base level, for example `ICE Cotton 0.8456 Daily Consumer Accum...`, to `strikePrice`.
- Adds Trigger/Barrier price to Account Grouping keys when available.
- Keeps prior commodity, LME, FX, and grouped drill-down support.
