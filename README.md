# Metals & Mining Research — Yalchin Nijat

A public equity-research project on silver miners, anchored on First Majestic Silver (NYSE: AG).
The site grades my November 2024 buy call against what actually happened, and maintains a
self-updating peer comparables dashboard.

**Live site:** served from `docs/` via GitHub Pages.

## Structure

- `01-retrospective/` — the Nov 2024 call graded against outcomes, with sources
- `02-data-engine/metals_engine.py` — pulls prices + financial statements (Yahoo Finance) for
  10 silver miners, recomputes the ratio framework, writes `comps_workbook.xlsx` and `docs/data.js`
- `docs/` — the website (single self-contained page + generated data file)
- `.github/workflows/refresh.yml` — re-runs the engine every weekday after U.S. market close
  and republishes the site automatically

## Refresh manually

```
pip install -r requirements.txt
cd 02-data-engine
python metals_engine.py
```

## Disclaimer

Independent student research for educational purposes only — not investment advice.
Data from Yahoo Finance and company disclosures; unaudited and may contain errors.
