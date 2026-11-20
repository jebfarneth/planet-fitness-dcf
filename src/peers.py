"""Build last-twelve-months (LTM) peer data from SEC companyfacts + Nasdaq prices.

LTM = latest fiscal year + current year-to-date - prior year-to-date, using the
facts reported in each company's most recent 10-Q/10-K. Every value written to
data/inputs/peers.csv carries its XBRL tag and filing accession number.
"""
import csv
import json
from datetime import date
from pathlib import Path

import xbrl

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
PRICE_DATE = "11/19/2026"

NAMES = {
    "PLNT": "Planet Fitness", "LTH": "Life Time Group", "XPOF": "Xponential Fitness",
    "DPZ": "Domino's Pizza", "WING": "Wingstop", "WEN": "Wendy's", "CHH": "Choice Hotels",
    "VVV": "Valvoline",
}
# Driven Brands (DRVN) was screened out: it filed an NT 10-Q on 2026-11-06 (late
# Q3 filing) and reports discontinued operations, so no clean LTM is available.

# Ticker-specific tag choices where a company's tagging differs from the default
TAG_OVERRIDE = {
    "XPOF": {"revenue": ["RevenueFromContractWithCustomerExcludingAssessedTax"],  # 'Revenues' is a sub-line
             "cash": ["CashAndCashEquivalentsAtCarryingValue",
                      "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"]},
}

FLOW = {
    "revenue": ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"],
    "ebit": ["OperatingIncomeLoss"],
    "da": ["DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
           "DepreciationAmortizationAndAccretionNet"],
    "net_income": ["NetIncomeLoss", "NetIncomeLossAvailableToCommonStockholdersBasic"],
}
STOCK = {
    "cash": ["CashAndCashEquivalentsAtCarryingValue"],
    "debt_lt": ["LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations",
                "LongTermNotesPayable", "LongTermDebtAndFinanceLeaseObligationsNoncurrent"],
    "debt_cur": ["LongTermDebtCurrent", "LongTermDebtAndCapitalLeaseObligationsCurrent",
                 "DebtCurrent", "LongTermDebtAndFinanceLeaseObligationsCurrent"],
    "st_borrow": ["ShortTermBorrowings", "LineOfCredit"],
    "fin_lease": ["FinanceLeaseLiabilityNoncurrent"],
    "fin_lease_cur": ["FinanceLeaseLiabilityCurrent"],
    "nci": ["MinorityInterest"],
}


def latest_period_end(cf):
    """Most recent period end reported in a 10-Q/10-K for operating income."""
    fs = [f for t in FLOW["ebit"] for f in xbrl.facts(cf, t) if f.get("form") in ("10-Q", "10-K")]
    f = max(fs, key=lambda x: (x["end"], x["filed"]))
    return f["end"], f["fy"], f["fp"], f["form"]


def _match(cf, tags, end, lo, hi, unit="USD"):
    """Fact for the first tag whose duration (days) is within [lo, hi] ending at end."""
    for tag in tags:
        c = [f for f in xbrl.facts(cf, tag, unit) if f["end"] == end and "start" in f
             and lo <= xbrl._days(f) <= hi and f.get("form") in ("10-Q", "10-K")]
        if c:
            f = max(c, key=lambda x: x["filed"])
            return f["val"], tag, f["accn"], f["start"]
    return None


def ltm(cf, tags, end):
    """Return (value, citation) for the twelve months ending `end`."""
    full = _match(cf, tags, end, 350, 380)
    if full:
        return full[0], f"{full[1]} FY ending {end} ({full[2]})"
    # find YTD fact ending at `end`
    for tag in tags:
        ytd = _match(cf, [tag], end, 80, 300)
        if not ytd:
            continue
        val, _, accn, start = ytd
        months = round(xbrl._days({"start": start, "end": end}) / 30.4)
        # prior fiscal year end = day before current YTD start
        fy_end = date.fromordinal(date.fromisoformat(start).toordinal() - 1).isoformat()
        fy = _match(cf, [tag], fy_end, 350, 380)
        # prior-year YTD: same length, ending ~1 year before
        prior = None
        for f in xbrl.facts(cf, tag):
            if "start" in f and abs(xbrl._days({"start": f["end"], "end": end}) - 364) <= 8 \
                    and abs(xbrl._days(f) - xbrl._days({"start": start, "end": end})) <= 8:
                if prior is None or f["filed"] > prior["filed"]:
                    prior = f
        if fy and prior:
            v = fy[0] + val - prior["val"]
            return v, (f"{tag}: FY to {fy_end} ({fy[2]}) + {months}M YTD to {end} ({accn}) "
                       f"- {months}M YTD to {prior['end']}")
    return None, "not found"


def instant(cf, tags, end):
    for tag in tags:
        r = xbrl.value(cf, tag, end, kind="instant")
        if r:
            return r[0], f"{tag} at {end} ({r[3]})"
    return 0, "not reported (0)"


def price(ticker):
    rows = json.load(open(RAW / f"prices_{ticker}.json"))["data"]["tradesTable"]["rows"]
    r = next(r for r in rows if r["date"] == PRICE_DATE)
    return float(r["close"].replace("$", "").replace(",", ""))


# Companies whose cover-page share counts are not in companyfacts (custom tags)
SHARE_OVERRIDE = {
    "PLNT": (74_252_489 + 266_128,
             "10-Q Q3 2026 cover page: 74,252,489 Class A + 266,128 Class B "
             "(exchangeable 1:1) as of Oct 30, 2026 (0001637207-26-000051)"),
}


def shares(cf, end, ticker=None):
    if ticker in SHARE_OVERRIDE:
        return SHARE_OVERRIDE[ticker]
    """Latest cover-page share count (all classes summed), else diluted weighted average."""
    dei = cf["facts"].get("dei", {}).get("EntityCommonStockSharesOutstanding")
    if dei:
        fs = dei["units"]["shares"]
        last = max(f["filed"] for f in fs)
        tot = sum(f["val"] for f in fs if f["filed"] == last)
        accn = next(f["accn"] for f in fs if f["filed"] == last)
        return tot, f"dei:EntityCommonStockSharesOutstanding, cover page ({accn})"
    r = _match(cf, ["WeightedAverageNumberOfDilutedSharesOutstanding"], end, 80, 100, "shares")
    if r is None:
        return None, "not found"
    return r[0], f"WeightedAverageNumberOfDilutedSharesOutstanding, quarter to {end} ({r[2]})"


def build(tickers):
    out = []
    for t in tickers:
        cf = xbrl.load(t)
        end, fy, fp, form = latest_period_end(cf)
        row = {"ticker": t, "name": NAMES[t], "period_end": end, "latest_form": form,
               "price": price(t), "price_date": "2026-11-19",
               "price_source": "Nasdaq.com historical quotes, close"}
        ov = TAG_OVERRIDE.get(t, {})
        for k, tags in FLOW.items():
            row[k], row[k + "_src"] = ltm(cf, ov.get(k, tags), end)
        if row["da"] is None:  # e.g. VVV tags depreciation and amortization separately
            d, ds = ltm(cf, ["Depreciation"], end)
            a, as_ = ltm(cf, ["AmortizationOfIntangibleAssets"], end)
            row["da"], row["da_src"] = d + a, f"{ds} + {as_}"
        for k, tags in STOCK.items():
            row[k], row[k + "_src"] = instant(cf, ov.get(k, tags), end)
        if "CapitalLease" in row["debt_lt_src"]:  # finance leases already inside debt tag
            row["fin_lease"] = row["fin_lease_cur"] = 0
            row["fin_lease_src"] = row["fin_lease_cur_src"] = "included in debt tag"
        row["shares"], row["shares_src"] = shares(cf, end, t)
        out.append(row)
    return out


if __name__ == "__main__":
    import sys
    rows = build(sys.argv[1:] or list(NAMES))
    for r in rows:
        print({k: v for k, v in r.items() if not k.endswith("_src")})
        for k, v in r.items():
            if k.endswith("_src"):
                print("   ", k, v)
    with open(ROOT / "data" / "inputs" / "peers_raw.csv", "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)
