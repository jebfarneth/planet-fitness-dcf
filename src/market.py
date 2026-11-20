"""Weekly / monthly price panels for the beta regressions.

Writes the price panels that the workbook's Beta tab regresses with SLOPE(),
plus a Python cross-check of each beta.
"""
import json
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
INP = ROOT / "data" / "inputs"
TICKERS = ["PLNT", "LTH", "XPOF", "DPZ", "WING", "WEN", "CHH", "VVV"]
WEEK_END = "2026-11-13"   # last complete Friday before the valuation date
MONTH_END = "2026-10-31"  # last complete month


def stock(t):
    rows = json.load(open(RAW / f"prices_{t}.json"))["data"]["tradesTable"]["rows"]
    s = pd.Series({pd.Timestamp(r["date"]): float(r["close"].replace("$", "").replace(",", ""))
                   for r in rows}).sort_index()
    return s.rename(t)


def spx():
    d = pd.read_csv(RAW / "fred_SP500.csv", parse_dates=["observation_date"])
    d["SP500"] = pd.to_numeric(d["SP500"], errors="coerce")
    return d.dropna().set_index("observation_date")["SP500"].rename("SPX")


def panels():
    daily = pd.concat([spx()] + [stock(t) for t in TICKERS], axis=1).sort_index()
    daily = daily[daily.index >= "2021-09-01"].ffill()
    weekly = daily.resample("W-FRI").last()
    weekly = weekly[weekly.index <= WEEK_END].tail(105)          # 104 weekly returns
    monthly = daily.resample("ME").last()
    monthly = monthly[monthly.index <= MONTH_END].tail(61)       # 60 monthly returns
    return weekly, monthly


def beta(p, t):
    r = p.pct_change().dropna()
    return np.polyfit(r["SPX"], r[t], 1)[0]


if __name__ == "__main__":
    w, m = panels()
    w.to_csv(INP / "weekly_prices.csv", index_label="week_ending")
    m.to_csv(INP / "monthly_prices.csv", index_label="month_end")
    print(f"weekly {w.index[0].date()}..{w.index[-1].date()} n={len(w)}; "
          f"monthly {m.index[0].date()}..{m.index[-1].date()} n={len(m)}")
    for t in TICKERS:
        print(t, "2y weekly beta", round(beta(w, t), 3), "| 5y monthly", round(beta(m, t), 3))
