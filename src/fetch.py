"""Download and checkpoint every external input to data/raw/.

Re-running skips any file already on disk. SEC requests are throttled well
under the SEC fair-access limit of 10 requests/second.
"""
import json
import sys
import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
RAW.mkdir(parents=True, exist_ok=True)

SEC_UA = "Jeb Farneth jebwfarneth@outlook.com"
_last_sec = [0.0]


def _get(url, headers=None):
    r = requests.get(url, headers=headers or {}, timeout=60)
    r.raise_for_status()
    return r.content


def cached(name, url, sec=False, force=False):
    path = RAW / name
    if path.exists() and not force:
        return path
    headers = {"User-Agent": SEC_UA if sec else "Mozilla/5.0 (educational DCF model)"}
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
