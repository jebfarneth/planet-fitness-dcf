#!/usr/bin/env bash
# Reproduce the model end to end. Downloads are cached in data/raw/ (re-runs do not refetch).
# Requires Python 3.11+ ; Excel is optional (only for the Excel-vs-Python cell check).
set -euo pipefail
cd "$(dirname "$0")"
PY=.venv/bin/python
[ -x "$PY" ] || { python3 -m venv .venv && .venv/bin/pip install -q -r requirements.txt; }
$PY src/fetch.py tickers all      # SEC (User-Agent set in src/fetch.py), Treasury, Damodaran, Nasdaq, FRED
$PY src/to_text.py                # filings -> text for source checks
$PY src/peers.py > /dev/null      # peer LTM data -> data/inputs/peers_raw.csv
$PY src/market.py                 # weekly / monthly price panels for beta
$PY src/build_model.py            # -> output/PLNT_DCF_Model.xlsx
$PY src/verify.py --refresh --all # independent Python rebuild vs. recalculated workbook (3 scenarios)
$PY src/check_sources.py          # every hard input vs. its downloaded source
$PY src/tornado.py
$PY src/charts.py                 # docs/football_field_{light,dark}.png
$PY src/memo.py                   # output/PLNT_Investment_Memo.pdf
