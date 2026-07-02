# Metals & Mining Research (Yalchin Nijat)

A public equity research project on silver miners, built around First Majestic Silver (NYSE: AG).
The site grades my November 2024 buy call against what actually happened, hosts my current research
note, and runs a peer comparables dashboard that refreshes itself.

**Live site:** served from `docs/` via GitHub Pages.

## What's in here

- `01-retrospective/`: the Nov 2024 call graded against outcomes, with sources
- `02-data-engine/metals_engine.py`: pulls prices and financial statements (Yahoo Finance) for
  10 silver miners, recomputes the ratio framework, writes `comps_workbook.xlsx` and `docs/data.js`
- `04-nav-model/`: mine-by-mine NAV model behind the July 2026 research note
- `docs/`: the website (dashboard, research note, generated data file)
- `.github/workflows/refresh.yml`: reruns the engine every weekday after U.S. market close
  and republishes the site

I wrote the code with AI coding tools. The assumptions, the ratings and the mistakes are mine.

## Refresh manually

```
pip install -r requirements.txt
cd 02-data-engine
python metals_engine.py
```

## Disclaimer

Independent student research for educational purposes only, not investment advice.
Data comes from Yahoo Finance and company disclosures. It is unaudited and may contain errors.
