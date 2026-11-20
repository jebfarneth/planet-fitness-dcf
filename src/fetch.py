"""Download and checkpoint every external input to data/raw/.

Re-running skips any file already on disk. SEC requests are throttled well
under the SEC fair-access limit of 10 requests/second.
"""
import json
import os
import sys
import time
from pathlib import Path

import subprocess

import requests

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
RAW.mkdir(parents=True, exist_ok=True)

def _sec_user_agent():
    """SEC asks for a name + contact email in the User-Agent. Set SEC_USER_AGENT or
    put it in private/sec_user_agent.txt (git-ignored); it is never committed."""
    ua = os.environ.get("SEC_USER_AGENT")
    f = ROOT / "private" / "sec_user_agent.txt"
    if not ua and f.exists():
        ua = f.read_text().strip()
    if not ua:
        raise SystemExit("Set SEC_USER_AGENT='Your Name your@email' (SEC fair-access policy) before fetching.")
    return ua
_last_sec = [0.0]


def _get(url, headers=None, tries=3):
    for i in range(tries):
        try:
            r = requests.get(url, headers=headers or {}, timeout=120)
            r.raise_for_status()
            return r.content
        except requests.RequestException:
            if i == tries - 1:
                raise
            time.sleep(3 * (i + 1))


def cached(name, url, sec=False, force=False, use_curl=False):
    path = RAW / name
    if path.exists() and not force:
        return path
    if use_curl:  # FRED times out python-requests clients but serves curl
        subprocess.run(["curl", "-sf", "-m", "120", "--retry", "3", "-o", str(path), url], check=True)
        print(f"downloaded {name}")
        return path
    headers = {"User-Agent": _sec_user_agent() if sec else "dcf-valuation-script/1.0"}
    if sec:
        wait = 0.25 - (time.time() - _last_sec[0])  # <= 4 req/s
        if wait > 0:
            time.sleep(wait)
        _last_sec[0] = time.time()
    path.write_bytes(_get(url, headers))
    print(f"downloaded {name}")
    return path


# Subject company and peers: ticker -> zero-padded CIK
COMPANIES = {
    "PLNT": "0001637207",  # Planet Fitness
}


def sec_company(ticker, cik):
    cached(f"companyfacts_{ticker}.json",
           f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json", sec=True)
    cached(f"submissions_{ticker}.json",
           f"https://data.sec.gov/submissions/CIK{cik}.json", sec=True)


if __name__ == "__main__":
    for t, c in COMPANIES.items():
        sec_company(t, c)
    if "tickers" in sys.argv:
        cached("company_tickers.json", "https://www.sec.gov/files/company_tickers.json", sec=True)


PLNT_DOCS = {  # form/period -> (accession, primary document)
    "10K_FY2021": ("0001637207-22-000011", "plnt-20211231.htm"),
    "10K_FY2022": ("0001637207-23-000005", "plnt-20221231.htm"),
    "10K_FY2023": ("0001637207-24-000020", "plnt-20231231.htm"),
    "10K_FY2024": ("0001637207-25-000016", "plnt-20241231.htm"),
    "10K_FY2025": ("0001637207-26-000011", "plnt-20251231.htm"),
    "10Q_Q3_2025": ("0001637207-25-000077", "plnt-20250930.htm"),
    "10Q_Q3_2026": ("0001637207-26-000051", "plnt-20260930.htm"),
}


def filing_url(cik, acc, doc):
    return f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{acc.replace('-', '')}/{doc}"


def plnt_docs():
    for k, (acc, doc) in PLNT_DOCS.items():
        cached(f"PLNT_{k}.htm", filing_url(COMPANIES["PLNT"], acc, doc), sec=True)


# ---- Market data -----------------------------------------------------------
PEERS = {  # ticker -> CIK (fitness operators/franchisors + asset-light franchisors)
    "LTH": "0001869198", "XPOF": "0001802156", "DPZ": "0001286681", "WING": "0001636222",
    "WEN": "0000030697", "CHH": "0001046311", "DRVN": "0001804745", "VVV": "0001674910",
}
PRICE_FROM, PRICE_TO = "2021-09-01", "2026-11-19"


def nasdaq_prices(ticker):
    """Daily closes from Nasdaq.com's public historical-quotes endpoint (Stooq now
    serves a JavaScript bot challenge, so it cannot be fetched as a CSV)."""
    url = (f"https://api.nasdaq.com/api/quote/{ticker}/historical?assetclass=stocks"
           f"&fromdate={PRICE_FROM}&todate={PRICE_TO}&limit=9999")
    return cached(f"prices_{ticker}.json", url)


def market():
    # S&P 500 price index, daily (FRED series SP500, source: S&P Dow Jones Indices)
    cached("fred_SP500.csv", "https://fred.stlouisfed.org/graph/fredgraph.csv?id=SP500", use_curl=True)
    # US Treasury daily par yield curve, 2026 (treasury.gov)
    cached("treasury_par_yield_2026.csv",
           "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/"
           "daily-treasury-rates.csv/2026/all?type=daily_treasury_yield_curve"
           "&field_tdr_date_value=2026&page&_format=csv")
    # Damodaran implied ERP history (annual) and monthly updates
    cached("damodaran_histimpl.xls", "https://pages.stern.nyu.edu/~adamodar/pc/datasets/histimpl.xls")
    cached("damodaran_ERPbymonth.xlsx", "https://pages.stern.nyu.edu/~adamodar/pc/implprem/ERPbymonth.xlsx")
    for t in ["PLNT", *PEERS]:
        nasdaq_prices(t)
    for t, c in PEERS.items():
        sec_company(t, c)


def treasury_2025():
    cached("treasury_par_yield_2025.csv",
           "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/"
           "daily-treasury-rates.csv/2025/all?type=daily_treasury_yield_curve"
           "&field_tdr_date_value=2025&page&_format=csv")


if __name__ == "__main__" and "all" in sys.argv:
    plnt_docs()
    plnt_8k()
    market()
    treasury_2025()


def plnt_8k():
    """Q3 2026 earnings press release (8-K Exhibit 99.1, filed 2026-11-05) - 2026 outlook."""
    cached("PLNT_8K_2026-11-05_ex991.htm",
           filing_url(COMPANIES["PLNT"], "0001637207-26-000047", "plntq32026pressreleaseex991.htm"), sec=True)
