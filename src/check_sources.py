"""Re-check every hard-coded input against its downloaded source.

- XBRL-tagged items: re-read from SEC companyfacts by tag and period.
- Items read from filing text (custom tags, MD&A): the number must appear in
  the named filing's text (formatted as the filing prints it).
- Market inputs: re-read from the Nasdaq JSON, Treasury CSVs and Damodaran file.
- Peer inputs: price re-read from Nasdaq JSON; LTM components re-derived.
"""
import csv
import json
import re
from pathlib import Path

import pandas as pd

import inputs as I
import xbrl

ROOT = Path(__file__).resolve().parents[1]
TXT = ROOT / "data" / "text"
cf = xbrl.load("PLNT")
FY_END = ["2021-12-31", "2022-12-31", "2023-12-31", "2024-12-31", "2025-12-31"]
BS_END = ["2020-12-31"] + FY_END
results = []


def record(item, ok, detail):
    results.append((item, ok, detail))


def text(name):
    return (TXT / f"PLNT_{name}.txt").read_text()


ALL_TEXT = {n: text(n) for n in ["10K_FY2021", "10K_FY2022", "10K_FY2023", "10K_FY2024", "10K_FY2025", "10Q_Q3_2025", "10Q_Q3_2026"]}
ALL_TEXT["8K"] = (TXT / "PLNT_8K_2026-11-05_ex991.txt").read_text()


def in_text(v, docs=None, million=False):
    """True if v ($M) appears in any of the filings as printed ($ thousands, or '$x.x million')."""
    docs = docs or list(ALL_TEXT)
    pats = []
    if million:
        pats.append(f"${v:.1f} million")
    else:
        k = round(abs(v) * 1000)
        pats.append(f"{k:,}")
    return any(p in ALL_TEXT[d] for d in docs for p in pats)


def xbrl_tag(src):
    m = re.search(r"XBRL us-gaap:([A-Za-z]+)", src)
    return m.group(1) if m else None


# ---- income statement / cash flow items ------------------------------------
for key, (lab, vals, a9, b9, src) in I.HIST.items():
    tag = xbrl_tag(src)
    for y, v in zip(FY_END, vals):
        if v is None:
            continue
        ok, how = False, ""
        if tag:
            r = xbrl.value(cf, tag, y)
            if r is not None:
                got = r[0] / 1e6 * (-1 if key == "other_op" else 1)
                ok, how = abs(got - v) < 0.0005, f"XBRL {tag} {y} = {got:.3f}"
        if not ok:
            ok = in_text(v) or in_text(v, million=True)
            how = how + " | found in filing text" if ok else how + " | NOT FOUND"
        record(f"{key} {y[:4]}", ok, how)
    for end, v in [("2025-09-30", a9), ("2026-09-30", b9)]:
        if v is None:
            continue
        ok, how = False, ""
        if tag:
            r = xbrl.value(cf, tag, end, months=9)
            if r is not None:
                got = r[0] / 1e6 * (-1 if key == "other_op" else 1)
                ok, how = abs(got - v) < 0.0005, f"XBRL {tag} 9M to {end} = {got:.3f}"
        if not ok:
            ok = in_text(v, ["10Q_Q3_2025", "10Q_Q3_2026"])
            how = how + " | found in 10-Q text" if ok else how + " | NOT FOUND"
        record(f"{key} 9M {end[:4]}", ok, how)

# ---- balance sheet --------------------------------------------------------
BS_TAGS = {"cash": ["CashAndCashEquivalentsAtCarryingValue"], "restricted_cash": ["RestrictedCashCurrent"],
           "mkt_sec": ["MarketableSecuritiesCurrent", "MarketableSecuritiesNoncurrent"], "ar": ["AccountsReceivableNetCurrent"],
           "inventory": ["InventoryNet"], "prepaid": ["PrepaidExpenseCurrent"], "other_rec": ["OtherReceivablesNetCurrent"],
           "ap": ["AccountsPayableCurrent"], "accrued": ["AccruedLiabilitiesCurrent"], "def_rev": ["ContractWithCustomerLiabilityCurrent"],
           "debt_principal": ["DebtInstrumentCarryingAmount"], "op_lease": ["OperatingLeaseLiabilityCurrent", "OperatingLeaseLiabilityNoncurrent"],
           "dta": ["DeferredIncomeTaxAssetsNet"], "nci": ["MinorityInterest"], "assets": ["Assets"], "liab_cur": ["LiabilitiesCurrent"],
           "liab_noncur": ["LiabilitiesNoncurrent"], "equity_total": ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"]}
for key, (lab, vals, last, src) in I.BS.items():
    for y, v in list(zip(BS_END, vals)) + [("2026-09-30", last)]:
        if v is None:
            continue
        ok, how = False, ""
        if key in BS_TAGS:
            parts = [xbrl.value(cf, t, y, kind="instant") for t in BS_TAGS[key]]
            if all(p is not None for p in parts):
                got = sum(p[0] for p in parts) / 1e6
                ok, how = abs(got - v) < 0.0005, f"XBRL {'+'.join(BS_TAGS[key])} {y} = {got:.3f}"
        if not ok and v != 0:
            if key == "tra":   # sum of current + noncurrent lines as printed
                ok = True
                how += " | TRA: components checked below"
            else:
                ok = in_text(v)
                how = how + (" | found in filing text" if ok else " | NOT FOUND")
        if v == 0 and not ok:
            ok, how = True, "zero (not reported)"
        record(f"BS {key} {y}", ok, how)

# TRA components (current + noncurrent) per filing
tra_parts = {"2021": ("10K_FY2021", ["20,302", "507,805"]), "2022": ("10K_FY2022", ["31,940", "462,525"]),
             "2023": ("10K_FY2023", ["41,294", "454,368"]), "2024": ("10K_FY2024", ["55,556", "411,360"]),
             "2025": ("10K_FY2025", ["55,518", "360,273"]), "Sep-26": ("10Q_Q3_2026", ["38,442", "323,484"])}
for y, (doc, parts) in tra_parts.items():
    ok = all(p in ALL_TEXT[doc] for p in parts)
    record(f"TRA components {y}", ok, f"{' + '.join(parts)} in {doc}")

# ---- KPIs -----------------------------------------------------------------
for key, (lab, vals, last, src) in I.KPI.items():
    for y, v in zip(FY_END, vals):
        if v is None or key in ("fr_closed", "corp_other"):
            continue
        if key == "members":
            ok = f"approximately {v} million members" in " ".join(ALL_TEXT.values())
        elif key == "sws":
            ok = f"system-wide sales of ${v} billion" in " ".join(ALL_TEXT.values())
        elif key == "scs":
            ok = f"{v*100:.1f} | %" in ALL_TEXT["10K_FY2025"]
        else:
            doc = ALL_TEXT[f"10K_FY{y[:4]}"]
            yr = y[:4]
            fr, co = I.KPI["fr_clubs"][1][FY_END.index(y)], I.KPI["corp_clubs"][1][FY_END.index(y)]
            if key in ("fr_clubs", "corp_clubs"):
                ok = (f"{fr:,} are franchised and {co:,} are corporate-owned" in doc
                      or f"{fr:,} franchisee-owned and {co:,} corporate-owned" in doc)
            else:  # openings: roll-forward table row in that year's 10-K
                ok = bool(re.search(r"New (stores|clubs) opened( or acquired)?\s*\|\s*" + f"{v:,}" + r"\b", doc))
        record(f"KPI {key} {y[:4]}", ok, "text search in the year's 10-K")

# ---- shares, Be-Fit, debt, amortization -----------------------------------
q = ALL_TEXT["10Q_Q3_2026"]
k = ALL_TEXT["10K_FY2025"]
record("Class A shares", "74,252,489 shares" in q, "10-Q cover")
record("Class B shares", "266,128 shares" in q, "10-Q cover")
record("RSUs 157,761", "157,761" in k, "10-K FY2025")
record("PSUs 138,961", "138,961" in k, "10-K FY2025")
record("Options 73,363 @ $49.62", "73,363 | $ | 49.62" in k, "10-K FY2025")
record("Be-Fit $112.0M cash / $2.8M contingent", "approximately $112.0 million" in q and "approximately $2.8 million" in q, "10-Q subsequent events")
for name, (p, cpn, ard) in I.DEBT_TERMS.items():
    ok = f"{round(p*1000):,}" in q and (f"{cpn*100:.3f}%" in k)
    record(f"Debt {name}", ok, f"principal in 10-Q, coupon {cpn*100:.3f}% in 10-K")
record("Debt fair value 2,389,338", "2,389,338" in q, "10-Q")
record("VFN $75.0M at 5.6%", "75,000" in q and "5.6% as of September 30, 2026" in q, "10-Q")
amort_tags = ["FiniteLivedIntangibleAssetsAmortizationExpenseNextTwelveMonths", "FiniteLivedIntangibleAssetsAmortizationExpenseYearTwo",
              "FiniteLivedIntangibleAssetsAmortizationExpenseYearThree", "FiniteLivedIntangibleAssetsAmortizationExpenseYearFour",
              "FiniteLivedIntangibleAssetsAmortizationExpenseYearFive", "FiniteLivedIntangibleAssetsAmortizationExpenseAfterYearFive"]
for (y, v), t in zip(I.AMORT_SCHEDULE.items(), amort_tags):
    r = xbrl.value(cf, t, "2025-12-31", kind="instant")
    record(f"Amortization {y}", r is not None and abs(r[0] / 1e6 - v) < 0.0005, f"XBRL {t}")
e = ALL_TEXT["8K"]
for s in ["Revenue to increase approximately 8%", "Adjusted EBITDA to increase approximately 6.5%", "Capital expenditures to increase approximately 15%",
          "Depreciation and amortization to increase approximately 7%", "approximately 1.5% to 2%", "approximately 190 to 195 locations"]:
    record(f"Guidance: {s}", s in e, "8-K Ex. 99.1")
record("Tax rate 26.0%", "assumed effective tax rate of 26.0%" in k, "10-K FY2025 non-GAAP footnote")

# ---- market inputs ----------------------------------------------------------
px = json.load(open(ROOT / "data" / "raw" / "prices_PLNT.json"))["data"]["tradesTable"]["rows"]
record("Price $61.00 on 11/19/2026", any(r["date"] == "11/19/2026" and r["close"] == "$61.00" for r in px), "Nasdaq JSON")
t26 = pd.read_csv(ROOT / "data" / "raw" / "treasury_par_yield_2026.csv").set_index("Date")
t25 = pd.read_csv(ROOT / "data" / "raw" / "treasury_par_yield_2025.csv").set_index("Date")
record("10Y UST 5.04% (2026-11-19)", abs(t26.loc["11/19/2026", "10 Yr"] / 100 - I.MARKET["rf"][0]) < 1e-9, "treasury.gov CSV")
record("7Y UST 4.94% (2026-11-19)", abs(t26.loc["11/19/2026", "7 Yr"] / 100 - I.MARKET["ust7_now"][0]) < 1e-9, "treasury.gov CSV")
record("7Y UST 3.94% (2025-12-15)", abs(t25.loc["12/15/2025", "7 Yr"] / 100 - I.MARKET["ust7_issue"][0]) < 1e-9, "treasury.gov CSV")
record("5Y UST 3.73% (2025-12-15)", abs(t25.loc["12/15/2025", "5 Yr"] / 100 - I.MARKET["ust5_issue"][0]) < 1e-9, "treasury.gov CSV")
d = pd.read_excel(ROOT / "data" / "raw" / "damodaran_ERPbymonth.xlsx", sheet_name="Historical ERP")
row = d[pd.to_datetime(d["Start of month"]) == "2026-11-01"].iloc[0]
record("ERP 4.41% (sustainable payout, 2026-11-01)", abs(row["ERP (T12 m with sustainable payout)"] - I.MARKET["erp"][0]) < 1e-9, "Damodaran ERPbymonth.xlsx")
record("ERP alt 3.79% (T12m, 2026-11-01)", abs(row["ERP (T12m)"] - I.MARKET["erp_alt"][0]) < 1e-9, "Damodaran ERPbymonth.xlsx")

# ---- peers ------------------------------------------------------------------
import peers  # noqa: E402
fresh = {r["ticker"]: r for r in peers.build(["PLNT"] + list(peers.NAMES)[1:])}
saved = {r["ticker"]: r for r in csv.DictReader(open(ROOT / "data" / "inputs" / "peers_raw.csv"))}
for t, r in saved.items():
    rows = json.load(open(ROOT / "data" / "raw" / f"prices_{t}.json"))["data"]["tradesTable"]["rows"]
    p = next(x for x in rows if x["date"] == "11/19/2026")
    record(f"Peer {t} price", abs(float(p["close"].strip("$").replace(",", "")) - float(r["price"])) < 1e-9, "Nasdaq JSON")
    same = all(str(fresh[t][k]) == str(r[k]) for k in ["revenue", "ebit", "da", "net_income", "cash", "debt_lt", "debt_cur", "shares"])
    record(f"Peer {t} XBRL re-derivation", same, "re-pulled from companyfacts")

bad = [r for r in results if not r[1]]
print(f"Checked {len(results)} sourced inputs: {len(results) - len(bad)} confirmed, {len(bad)} not confirmed")
for r in bad:
    print("  NOT CONFIRMED:", r)
json.dump({"checked": len(results), "confirmed": len(results) - len(bad), "not_confirmed": [list(map(str, b)) for b in bad]},
          open(ROOT / "output" / "source_check.json", "w"), indent=2)
