"""Build the Planet Fitness DCF workbook (output/PLNT_DCF_Model.xlsx).

Every calculated cell is an Excel formula. Hard inputs come only from
src/inputs.py (sourced filing / market data), src/assumptions.py (forecast
judgement) and the downloaded price panels (data/inputs/*.csv).
"""
import csv
import json
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font

import inputs as I
from assumptions import DRIVERS, PROJ_YEARS
from book import BLUE, FMT, INPUT_FILL, Book, col

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "PLNT_DCF_Model.xlsx"
YC = ["D", "E", "F", "G", "H", "I"]          # projection columns FY2026E..FY2031E
HC = ["C", "D", "E", "F", "G"]               # historical columns FY2021..FY2025
PEERS = ["LTH", "XPOF", "DPZ", "WING", "WEN", "CHH", "VVV"]
ALL = ["PLNT"] + PEERS

B = Book()


# =============================================================================
# Assumptions
# =============================================================================
def build_assumptions():
    ws = B.sheet("Assumptions")
    B.title(ws, "Assumptions & scenario drivers", "Blue = hard-coded input, black = formula, green = link to another tab. $ in millions.", 10)
    B.widths(ws, {"A": 52, "B": 4, "C": 12, **{c: 11 for c in YC}, "J": 110})

    B.section(ws, 3, "Scenario", 10)
    B.label(ws, "A4", "Scenario selector (1 = Base, 2 = Upside, 3 = Downside)", bold=True)
    B.inp(ws, "C4", 1, "n", name="scen").fill = INPUT_FILL
    B.f(ws, "D4", '=CHOOSE({scen},"Base","Upside","Downside")', "gen", name="scen_name", bold=True)

    B.section(ws, 6, "Timing & tax", 10)
    rows = [
        ("Valuation date", datetime(2026, 11, 20), "date", "val_date", "Day after the last close used for the share price (2026-11-19)"),
        ("Fiscal year end (FY2026)", datetime(2026, 12, 31), "date", "fye", "PLNT fiscal year ends Dec 31"),
        ("Latest balance sheet date", datetime(2026, 9, 30), "date", "bs_date", "10-Q Q3 2026"),
    ]
    r = 7
    for lab, v, fmt, nm, src in rows:
        B.label(ws, f"A{r}", lab)
        B.inp(ws, f"C{r}", v, fmt, name=nm)
        B.label(ws, f"J{r}", src, italic=True)
        r += 1
    B.label(ws, "A10", "Stub fraction of FY2026 remaining after valuation date")
    B.f(ws, "C10", "=({fye}-{val_date})/365", "num4", name="stub")
    B.label(ws, "J10", "FY2026E unlevered FCF x stub = cash flow still to be received (41 of 365 days)", italic=True)
    B.label(ws, "A11", "Mid-year discounting convention (1 = on, 0 = off)")
    B.inp(ws, "C11", 1, "n", name="midyear")
    B.label(ws, "A12", "Tax rate (cash taxes on EBIT; also used to un/relever beta)")
    B.inp(ws, "C12", I.MARKET["tax"][0], "pct", name="tax")
    B.label(ws, "J12", I.MARKET["tax"][1] + ". FY2025 GAAP effective rate 27.8% (XBRL EffectiveIncomeTaxRateContinuingOperations).", italic=True)

    B.section(ws, 14, "Operating drivers by scenario (the 'Live' row feeds the model)", 10)
    B.header_row(ws, 15, 3, ["Case"] + PROJ_YEARS)
    B.label(ws, "J15", "Rationale / historical anchor", bold=True)
    r = 16
    for key, lab, fmt, base, up, down, why in DRIVERS:
        B.label(ws, f"A{r}", lab, bold=True)
        ws[f"J{r}"].value = why
        ws[f"J{r}"].font = Font(italic=True, size=9, name="Calibri")
        ws[f"J{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(f"J{r}:J{r+3}")
        for k, (case, vals) in enumerate([("Base", base), ("Upside", up), ("Downside", down)]):
            B.label(ws, f"C{r+k}", case)
            for c, v in zip(YC, vals):
                B.inp(ws, f"{c}{r+k}", v, fmt)
        B.label(ws, f"C{r+3}", "Live", bold=True)
        for c in YC:
            B.f(ws, f"{c}{r+3}", f"=CHOOSE({{scen}},{c}{r},{c}{r+1},{c}{r+2})", fmt, bold=True)
        B.name_row(f"drv_{key}", ws, r + 3)
        r += 5

    B.section(ws, r, "Amortization of existing intangibles (all scenarios)", 10)
    r += 1
    B.label(ws, f"A{r}", "Scheduled amortization ($M)")
    for c, y in zip(YC, PROJ_YEARS):
        B.inp(ws, f"{c}{r}", I.AMORT_SCHEDULE[y.replace("E", "")], "m3")
    B.name_row("amort_sched", ws, r)
    B.label(ws, f"J{r}", I.AMORT_SRC, italic=True)

    r += 2
    B.section(ws, r, "Management 2026 outlook (8-K filed 2026-11-05, Exhibit 99.1) - used to calibrate FY2026E", 10)
    g = I.GUIDANCE_2026
    for lab, v, nm in [("Revenue growth, FY2026 (approx.)", g["revenue_growth"], "g_rev"),
                       ("Adjusted EBITDA growth, FY2026 (approx.)", g["adj_ebitda_growth"], "g_ebitda"),
                       ("Capital expenditures growth, FY2026 (approx.)", g["capex_growth"], "g_capex"),
                       ("D&A growth, FY2026 (approx.)", g["da_growth"], "g_da"),
                       ("System-wide same club sales growth - low", g["scs_low"], "g_scs_lo"),
                       ("System-wide same club sales growth - high", g["scs_high"], "g_scs_hi")]:
        r += 1
        B.label(ws, f"A{r}", lab)
        B.inp(ws, f"C{r}", v, "pct", name=nm)
    r += 1
    B.label(ws, f"A{r}", "System-wide new club openings incl. 25 Be-Fit clubs (low / high)")
    B.inp(ws, f"C{r}", g["openings_low"], "n", name="g_open_lo")
    B.inp(ws, f"D{r}", g["openings_high"], "n", name="g_open_hi")
    ws.freeze_panes = "D16"


# =============================================================================
# Historicals
# =============================================================================
def build_historicals():
    ws = B.sheet("Historicals")
    B.title(ws, "Historical financials - Planet Fitness, Inc. (NYSE: PLNT)",
            "$ in millions (filing $ thousands / 1,000). LTM = FY2025 + 9M 2026 - 9M 2025. Sources in column K.", 11)
    B.widths(ws, {"A": 50, "B": 3, **{c: 11 for c in "CDEFGHIJ"}, "K": 120})
    B.header_row(ws, 4, 3, I.YEARS + ["9M 2025", "9M 2026", "LTM Sep-26"])

    def hist(key, row, fmt="m3"):
        lab, vals, a9, b9, src = I.HIST[key]
        B.label(ws, f"A{row}", lab)
        for c, v in zip(HC, vals):
            if v is not None:
                B.inp(ws, f"{c}{row}", v, fmt)
        if a9 is not None:
            B.inp(ws, f"H{row}", a9, fmt)
            B.inp(ws, f"I{row}", b9, fmt)
            B.f(ws, f"J{row}", f"=G{row}+I{row}-H{row}", fmt)
        B.label(ws, f"K{row}", src, italic=True)
        B.name_row(f"h_{key}", ws, row)

    def calc(name, row, label, tpl, cols="CDEFGHIJ", fmt="m3", bold=False):
        B.label(ws, f"A{row}", label, bold=bold)
        for c in cols:
            B.f(ws, f"{c}{row}", tpl.replace("#", c), fmt, bold=bold)
        if name:
            B.name_row(name, ws, row)

    r = 5
    B.section(ws, r, "Income statement", 11)
    hist("rev_franchise", 6); hist("rev_royalty", 7)
    calc("h_rev_other_fr", 8, "  of which: franchise & other fees, placement, HVAC", "={h_rev_franchise@#}-{h_rev_royalty@#}")
    hist("rev_naf", 9); hist("rev_corp", 10); hist("rev_equip", 11)
    calc("h_rev", 12, "Total revenue", "={h_rev_franchise@#}+{h_rev_naf@#}+{h_rev_corp@#}+{h_rev_equip@#}", bold=True)
    hist("rev_total_xbrl", 13)
    calc("h_rev_chk", 14, "  Check: build vs. reported (should be 0)", "={h_rev@#}-{h_rev_total_xbrl@#}")
    calc("h_rev_g", 15, "  Revenue growth", "=IF(ISNUMBER({h_rev@#prev}),{h_rev@#}/{h_rev@#prev}-1,\"\")", cols="", fmt="pct")
    for c, p in zip("DEFG", "CDEF"):
        B.f(ws, f"{c}15", f"={c}12/{p}12-1", "pct")
    B.f(ws, "I15", "=I12/H12-1", "pct")
    calc("h_rev_xnaf", 16, "Revenue excl. NAF (NAF is a pass-through fund)", "={h_rev@#}-{h_rev_naf@#}")
    hist("cogs", 18); hist("club_ops", 19); hist("sga", 20); hist("naf_exp", 21); hist("da", 22); hist("other_op", 23)
    calc("h_ebit", 24, "Income from operations (EBIT)", "={h_rev@#}-{h_cogs@#}-{h_club_ops@#}-{h_sga@#}-{h_naf_exp@#}-{h_da@#}-{h_other_op@#}", bold=True)
    hist("ebit_xbrl", 25)
    calc("h_ebit_chk", 26, "  Check: recomputed vs. reported (should be 0)", "={h_ebit@#}-{h_ebit_xbrl@#}")
    calc("h_ebitda", 27, "EBITDA (EBIT + D&A, GAAP)", "={h_ebit@#}+{h_da@#}", bold=True)
    hist("adj_ebitda", 28)
    calc("h_norm", 29, "  Normalization adjustments (Adj. EBITDA - GAAP EBITDA)", "={h_adj_ebitda@#}-{h_ebitda@#}")
    calc("h_adj_m", 30, "  Adj. EBITDA margin (% of revenue excl. NAF)", "={h_adj_ebitda@#}/{h_rev_xnaf@#}", fmt="pct")
    hist("int_exp", 32); hist("int_inc", 33); hist("tax", 34); hist("net_income", 35); hist("ni_parent", 36)
    hist("sbc", 37)
    calc("h_sbc_pct", 38, "  SBC as % of revenue excl. NAF", "={h_sbc@#}/{h_rev_xnaf@#}", fmt="pct")
    hist("amort", 39)
    calc("h_dep", 40, "Depreciation (D&A less amortization)", "={h_da@#}-{h_amort@#}", cols="CDEFG")
    calc("h_dep_pct", 41, "  Depreciation % of revenue excl. NAF", "={h_dep@#}/{h_rev_xnaf@#}", cols="CDEFG", fmt="pct")
    hist("lease_cost", 42)

    r = 44
    B.section(ws, r, "Segment Adjusted EBITDA (current definition, reported from FY2023)", 11)
    hist("seg_fr", 45); hist("seg_corp", 46); hist("seg_equip", 47)
    sc = "EFGHIJ"
    calc("h_seg_oh", 48, "Corporate & other (unallocated) = Adj. EBITDA - segments", "={h_adj_ebitda@#}-{h_seg_fr@#}-{h_seg_corp@#}-{h_seg_equip@#}", cols=sc)
    calc("h_fr_m", 49, "  Franchise margin excl. NAF ((segment - NAF surplus) / franchise revenue)",
         "=({h_seg_fr@#}-({h_rev_naf@#}-{h_naf_exp@#}))/{h_rev_franchise@#}", cols=sc, fmt="pct")
    calc("h_co_m", 50, "  Corporate-owned clubs margin", "={h_seg_corp@#}/{h_rev_corp@#}", cols=sc, fmt="pct")
    calc("h_eq_m", 51, "  Equipment margin", "={h_seg_equip@#}/{h_rev_equip@#}", cols=sc, fmt="pct")
    calc("h_oh_pct", 52, "  Unallocated corporate as % of revenue excl. NAF", "=-{h_seg_oh@#}/{h_rev_xnaf@#}", cols=sc, fmt="pct")
    B.label(ws, "K45", I.HIST["seg_fr"][4] + ". FY2021-22 were reported under the older 'Segment EBITDA' definition and are not comparable, so they are omitted.", italic=True)

    # ---- KPIs ---------------------------------------------------------------
    r = 54
    B.section(ws, r, "Operating KPIs and revenue drivers", 11)
    kpi_cols = HC + ["I"]   # FY2021..FY2025 and 9M 2026 (roll-forward to 30-Sep-2026)

    def kpi_inputs(key, row, opening=None):
        lab, vals, last, src = I.KPI[key]
        B.label(ws, f"A{row}", lab)
        for c, v in zip(HC, vals):
            if v is not None:
                B.inp(ws, f"{c}{row}", v, "n" if isinstance(v, int) else ("pct" if v < 1 else "n3"))
        if last is not None:
            B.inp(ws, f"I{row}", last, "n" if isinstance(last, int) else "pct")
        B.label(ws, f"K{row}", src, italic=True)
        B.name_row(f"k_{key}", ws, row)

    # Franchise club roll-forward
    B.label(ws, "A55", "Franchisee-owned clubs - beginning of period")
    B.inp(ws, "C55", 2021, "n")
    B.label(ws, "K55", "FY2021 opening = 2,021 franchise clubs at Dec 31, 2020 (10-K FY2021 roll-forward)", italic=True)
    for c, p in zip("DEFG", "CDEF"):
        B.f(ws, f"{c}55", f"={p}58", "n")
    B.f(ws, "I55", "=G58", "n")
    B.name_row("k_fr_bop", ws, 55)
    kpi_inputs("fr_open", 56)
    kpi_inputs("fr_closed", 57)
    calc("k_fr_eop", 58, "Franchisee-owned clubs - end of period", "={k_fr_bop@#}+{k_fr_open@#}+{k_fr_closed@#}", cols="CDEFGI", fmt="n", bold=True)
    kpi_inputs("fr_clubs", 59)
    calc("k_fr_chk", 60, "  Check: roll-forward vs. reported (should be 0)", "={k_fr_eop@#}-{k_fr_clubs@#}", cols="CDEFGI", fmt="n")
    B.label(ws, "A61", "Corporate-owned clubs - beginning of period")
    B.inp(ws, "C61", 103, "n")
    B.label(ws, "K61", "FY2021 opening = 103 corporate clubs at Dec 31, 2020 (10-K FY2021 roll-forward)", italic=True)
    for c, p in zip("DEFG", "CDEF"):
        B.f(ws, f"{c}61", f"={p}64", "n")
    B.f(ws, "I61", "=G64", "n")
    B.name_row("k_co_bop", ws, 61)
    kpi_inputs("corp_open", 62)
    kpi_inputs("corp_other", 63)
    calc("k_co_eop", 64, "Corporate-owned clubs - end of period", "={k_co_bop@#}+{k_corp_open@#}+{k_corp_other@#}", cols="CDEFGI", fmt="n", bold=True)
    kpi_inputs("corp_clubs", 65)
    calc("k_co_chk", 66, "  Check: roll-forward vs. reported (should be 0)", "={k_co_eop@#}-{k_corp_clubs@#}", cols="CDEFGI", fmt="n")
    calc("k_fr_avg", 67, "Average franchisee-owned clubs", "=AVERAGE({k_fr_bop@#},{k_fr_eop@#})", cols="CDEFG", fmt="m")
    calc("k_co_avg", 68, "Average corporate-owned clubs", "=AVERAGE({k_co_bop@#},{k_co_eop@#})", cols="CDEFG", fmt="m")
    calc("k_roy_pc", 69, "Royalty revenue per avg. franchise club ($M)", "={h_rev_royalty@#}/{k_fr_avg@#}", cols="CDEFG", fmt="num4")
    calc("k_roy_pc_g", 70, "  growth", "", cols="", fmt="pct")
    for c, p in zip("DEFG", "CDEF"):
        B.f(ws, f"{c}70", f"={c}69/{p}69-1", "pct")
    calc("k_oth_pc", 71, "Other franchise revenue per avg. franchise club ($M)", "={h_rev_other_fr@#}/{k_fr_avg@#}", cols="CDEFG", fmt="num4")
    calc("k_naf_ratio", 72, "NAF revenue / royalty revenue", "={h_rev_naf@#}/{h_rev_royalty@#}", cols="CDEFGHIJ", fmt="pct")
    calc("k_co_pc", 73, "Corporate club revenue per avg. corporate club ($M)", "={h_rev_corp@#}/{k_co_avg@#}", cols="CDEFG", fmt="num4")
    calc("k_eq_pc", 74, "Equipment revenue per avg. franchise club ($M)", "={h_rev_equip@#}/{k_fr_avg@#}", cols="CDEFG", fmt="num4")
    kpi_inputs("members", 75)
    kpi_inputs("sws", 76)
    kpi_inputs("scs", 77)

    # ---- Cash flow ---------------------------------------------------------
    B.section(ws, 79, "Cash flow", 11)
    hist("cfo", 80); hist("capex", 81)
    calc("h_fcf", 82, "Free cash flow (CFO - capex)", "={h_cfo@#}-{h_capex@#}", bold=True)
    calc("h_capex_pct", 83, "  Capex % of revenue excl. NAF", "={h_capex@#}/{h_rev_xnaf@#}", fmt="pct")
    calc("h_fcf_conv", 84, "  FCF / Adj. EBITDA", "={h_fcf@#}/{h_adj_ebitda@#}", fmt="pct")
    hist("acq", 85); hist("buyback", 86)

    # ---- Balance sheet -----------------------------------------------------
    B.section(ws, 88, "Balance sheet (FY columns = Dec 31; column J = Sep 30, 2026)", 11)
    bcols = HC + ["J"]
    r = 89
    for key, (lab, vals, last, src) in I.BS.items():
        B.label(ws, f"A{r}", lab)
        for c, v in zip(HC, vals[1:]):
            if v is not None:
                B.inp(ws, f"{c}{r}", v, "m3")
        B.inp(ws, f"J{r}", last, "m3")
        B.label(ws, f"K{r}", src, italic=True)
        B.name_row(f"b_{key}", ws, r)
        r += 1
    calc("b_nwc", r, "Operating net working capital", "={b_ar@#}+{b_inventory@#}+{b_prepaid@#}+{b_other_rec@#}-{b_ap@#}-{b_accrued@#}-{b_equip_dep@#}-{b_def_rev@#}", cols="CDEFGJ", bold=True)
    B.label(ws, f"K{r}", "AR + inventory + prepaid + other receivables - AP - accrued - equipment deposits - current deferred revenue. Excludes cash, restricted cash, tax items, lease liabilities and debt.", italic=True)
    r += 1
    calc("b_nwc_pct", r, "  NWC % of revenue excl. NAF (Sep-26 uses LTM revenue)", "={b_nwc@#}/{h_rev_xnaf@#}", cols="CDEFGJ", fmt="pct")
    r += 1
    calc("b_bal_chk", r, "Check: assets - (liabilities + equity) (should be 0)", "={b_assets@#}-{b_liab_cur@#}-{b_liab_noncur@#}-{b_equity_total@#}", cols="CDEFGJ")
    r += 2

    B.section(ws, r, "Normalizations and accounting treatment", 11)
    notes = [
        "1. One-off items: the model forecasts from company 'Adjusted EBITDA', which removes executive-transition costs, severance, gains on club / equity-stake sales (e.g. $12.5M Bravo Fit gain in 9M 2026), "
        "TRA remeasurement, credit-loss adjustments on the held-to-maturity investment and similar items (row 29 shows the size each year). These are non-recurring or non-operating.",
        "2. Stock-based compensation: NOT added back. Company Adjusted EBITDA is after SBC ($12.3M in FY2025), so the DCF treats SBC as a real (cash-equivalent) cost and does not double count dilution.",
        "3. Leases (ASC 842): operating lease cost (~$78M FY2025) sits inside club operations, so EBITDA is post-rent. Operating lease liabilities are therefore NOT treated as debt in EV, "
        "consistently for PLNT and every peer. Finance leases ($0.8M) are treated as debt.",
        "4. NAF (national advertising fund): restricted pass-through; revenue roughly equals expense. Margins, capex and NWC are expressed on revenue excl. NAF. The NAF contribution rose from 2% to 3% of dues in 2026.",
        "5. FY2022 includes the Feb-2022 Sunshine Fitness acquisition (114 clubs, $424.9M), which moved ~114 clubs from franchise to corporate and roughly doubled corporate revenue; per-club metrics use average clubs.",
        "6. FY2021 still reflects COVID recovery (clubs reopened through 2021), so it is shown for context but not used to anchor forecast margins.",
        "7. 9M figures are unaudited (10-Q). The latest 10-Q does not restate full equity-award counts, so dilutive securities use Dec 31, 2025 counts (10-K).",
    ]
    for n in notes:
        r += 1
        B.label(ws, f"A{r}", n, wrap=False)
    ws.freeze_panes = "C5"


# =============================================================================
# Comparable companies
# =============================================================================
def build_comps():
    ws = B.sheet("Comps")
    B.title(ws, "Comparable companies - LTM trading multiples",
            "Prices: Nasdaq.com close on 2026-11-19. Financials: SEC XBRL companyfacts, LTM to each company's latest 10-Q/10-K. $M except per share. NTM consensus estimates are not available from free sources, so LTM only.", 21)
    B.widths(ws, {"A": 8, "B": 20, "C": 11, **{c: 10 for c in "DEFGHIJKLMNOPQRST"}, "V": 3, **{c: 11 for c in "WXYZ"}})
    hdr = ["Ticker", "Company", "Period end", "Price", "Shares (m)", "Mkt cap", "Total debt", "Cash", "NCI", "EV",
           "LTM revenue", "LTM EBIT", "LTM D&A", "LTM EBITDA", "LTM net inc.", "EV/Revenue", "EV/EBITDA", "P/E", "Debt/(D+E)", "Debt/Equity"]
    B.header_row(ws, 4, 1, hdr)
    B.header_row(ws, 4, 23, ["LT debt", "Current debt", "ST borrow./VFN", "Fin. leases"])
    rows = list(csv.DictReader(open(ROOT / "data" / "inputs" / "peers_raw.csv")))
    data = {r["ticker"]: r for r in rows}
    r0 = 6
    srcs = []
    for i, t in enumerate(ALL):
        r = r0 + i if t == "PLNT" else r0 + i + 1   # blank row between PLNT and peers
        d = data[t]
        m = lambda k: float(d[k]) / 1e6
        B.label(ws, f"A{r}", t, bold=(t == "PLNT"))
        B.label(ws, f"B{r}", d["name"])
        B.label(ws, f"C{r}", d["period_end"])
        if t == "PLNT":
            B.f(ws, f"D{r}", "={price}", "ps")
            B.f(ws, f"E{r}", "={sh_class_a}+{sh_class_b}", "m3")
        else:
            B.inp(ws, f"D{r}", float(d["price"]), "ps")
            B.inp(ws, f"E{r}", float(d["shares"]) / 1e6, "m3")
        B.f(ws, f"F{r}", f"=D{r}*E{r}", "m")
        B.inp(ws, f"W{r}", m("debt_lt"), "m3")
        B.inp(ws, f"X{r}", m("debt_cur"), "m3")
        B.inp(ws, f"Y{r}", m("st_borrow") + (75.0 if t == "PLNT" and m("st_borrow") == 0 else 0), "m3")
        B.inp(ws, f"Z{r}", m("fin_lease") + m("fin_lease_cur"), "m3")
        B.f(ws, f"G{r}", f"=SUM(W{r}:Z{r})", "m")
        B.inp(ws, f"H{r}", m("cash"), "m3")
        B.inp(ws, f"I{r}", m("nci"), "m3")
        B.f(ws, f"J{r}", f"=F{r}+G{r}-H{r}+I{r}", "m")
        B.inp(ws, f"K{r}", m("revenue"), "m3")
        B.inp(ws, f"L{r}", m("ebit"), "m3")
        B.inp(ws, f"M{r}", m("da"), "m3")
        B.f(ws, f"N{r}", f"=L{r}+M{r}", "m")
        B.inp(ws, f"O{r}", m("net_income"), "m3")
        B.f(ws, f"P{r}", f"=J{r}/K{r}", "x")
        B.f(ws, f"Q{r}", f'=IF(N{r}>0,J{r}/N{r},"NM")', "x")
        B.f(ws, f"R{r}", f'=IF(O{r}>0,F{r}/O{r},"NM")', "x")
        B.f(ws, f"S{r}", f"=G{r}/(G{r}+F{r})", "pct")
        B.f(ws, f"T{r}", f"=G{r}/F{r}", "pct")
        B.name_cell(f"c_{t}_row", ws, "A", r)
        for k in ["shares", "revenue", "ebit", "da", "net_income", "cash", "debt_lt", "debt_cur", "st_borrow", "fin_lease", "fin_lease_cur", "nci"]:
            srcs.append((t, k, d[k + "_src"]))
    peer_first, peer_last = r0 + 2, r0 + 1 + len(PEERS)
    B.name_cell("c_peer_first", ws, "A", peer_first)
    rr = peer_last + 2
    stats = [("Peer median", "MEDIAN"), ("Peer mean", "AVERAGE"), ("Peer 25th percentile", "P25"), ("Peer 75th percentile", "P75")]
    for lab, fn in stats:
        B.label(ws, f"B{rr}", lab, bold=True)
        for c in "PQRST":
            rng = f"{c}{peer_first}:{c}{peer_last}"
            if fn == "P25":
                tpl = f"=PERCENTILE({rng},0.25)"
            elif fn == "P75":
                tpl = f"=PERCENTILE({rng},0.75)"
            else:
                tpl = f"={fn}({rng})"
            B.f(ws, f"{c}{rr}", tpl, "pct" if c in "ST" else "x", bold=True)
        B.name_row({"MEDIAN": "c_med", "AVERAGE": "c_mean", "P25": "c_p25", "P75": "c_p75"}[fn], ws, rr)
        rr += 1
    B.label(ws, f"B{rr}", "PLNT (current)", bold=True)
    for c in "PQRST":
        B.f(ws, f"{c}{rr}", f"={c}{r0}", "pct" if c in "ST" else "x")
    rr += 2

    # Implied valuation for PLNT
    B.section(ws, rr, "Implied PLNT value per share from peer multiples (25th / median / 75th percentile)", 21)
    rr += 1
    B.header_row(ws, rr, 2, ["Metric", "PLNT LTM", "", "25th", "Median", "75th"])
    plnt = r0
    specs = [("EV / LTM EBITDA", "Q", f"N{plnt}", "ev", "ebitda"),
             ("EV / LTM revenue", "P", f"K{plnt}", "ev", "rev"),
             ("P / LTM net income", "R", f"O{plnt}", "eq", "pe")]
    for lab, mc, metric, kind, nm in specs:
        rr += 1
        B.label(ws, f"B{rr}", lab, bold=True)
        B.f(ws, f"C{rr}", f"={metric}", "m")
        B.label(ws, f"D{rr}", "multiple")
        for c, statrow in zip("EFG", ["c_p25", "c_med", "c_p75"]):
            B.f(ws, f"{c}{rr}", f"={{{statrow}@{mc}}}", "x")
        rr += 1
        B.label(ws, f"D{rr}", "value / share")
        for c in "EFG":
            if kind == "ev":
                B.f(ws, f"{c}{rr}", f"=({c}{rr-1}*$C{rr-1}+{{br_adj_total}})/{{br_dil_shares}}", "ps", bold=True)
            else:
                B.f(ws, f"{c}{rr}", f"={c}{rr-1}*$C{rr-1}/{{br_dil_shares}}", "ps", bold=True)
        B.name_cell(f"cmp_{nm}_lo", ws, "E", rr)
        B.name_cell(f"cmp_{nm}_mid", ws, "F", rr)
        B.name_cell(f"cmp_{nm}_hi", ws, "G", rr)
    rr += 2
    notes = [
        "EBITDA = GAAP operating income + D&A (unadjusted) for every company, so PLNT is compared like-for-like on reported figures; PLNT LTM includes a $12.5M one-off gain (Adj. EBITDA would be $585.1M).",
        "EV = market cap + total debt (incl. finance leases, excl. operating leases) - cash & equivalents + NCI. Operating leases are excluded for all companies because EBITDA is post-rent under ASC 842.",
        "Peer set: fitness operators/franchisors (Life Time, Xponential) and asset-light, securitization-financed franchisors (Domino's, Wingstop, Wendy's, Choice Hotels, Valvoline).",
        "Driven Brands was screened out: it filed an NT 10-Q (late filing) on 2026-11-06 and reports discontinued operations. Xponential is loss-making (P/E NM).",
        "Share counts: cover-page shares outstanding (all classes) from the latest filing; Xponential uses diluted weighted-average shares (cover count not in XBRL). PLNT: 74.25M Class A + 0.27M Class B (10-Q cover).",
        "PLNT short-term borrowings = $75.0M Variable Funding Notes drawn (10-Q Q3 2026). DPZ debt tags already include finance leases.",
        "Domino's period is 36 weeks to 2026-09-06 (its fiscal Q3); Valvoline's latest filing is its FY to 2026-09-30 10-K.",
    ]
    for n in notes:
        B.label(ws, f"B{rr}", n, italic=True)
        rr += 1
    rr += 1
    B.section(ws, rr, "XBRL source for every input above (tag, period and accession number)", 21)
    for t, k, s in srcs:
        rr += 1
        B.label(ws, f"A{rr}", t)
        B.label(ws, f"B{rr}", k)
        B.label(ws, f"D{rr}", s, italic=True)
    ws.freeze_panes = "C5"
    return peer_first, peer_last


# =============================================================================
# Beta regressions
# =============================================================================
def build_beta():
    ws = B.sheet("Beta")
    B.title(ws, "Beta regressions vs. S&P 500 (price index)",
            "Weekly: 105 Friday closes, 2024-11-15 to 2026-11-13 (104 returns). Monthly: 61 month-end closes, Oct-2021 to Oct-2026 (60 returns). Stock prices: Nasdaq.com; S&P 500: FRED series SP500.", 12)
    B.widths(ws, {"A": 12, **{col(i): 10 for i in range(2, 60)}})
    w = pd.read_csv(ROOT / "data" / "inputs" / "weekly_prices.csv", parse_dates=["week_ending"])
    m = pd.read_csv(ROOT / "data" / "inputs" / "monthly_prices.csv", parse_dates=["month_end"])
    tick = ["SPX"] + ALL
    # summary
    B.header_row(ws, 4, 1, ["Ticker", "2y wkly beta", "R-squared", "# returns", "5y mthly beta", "R-squared", "# returns"])
    # weekly block: date col A at row 16+, prices B.., returns after a gap
    wr0 = 17
    B.label(ws, f"A{wr0-2}", "Weekly closes (inputs) and returns (formulas)", bold=True)
    B.header_row(ws, wr0 - 1, 1, ["Week ending"] + tick + [""] + [f"r {t}" for t in tick])
    nW = len(w)
    for i, row in w.iterrows():
        rr = wr0 + i
        ws[f"A{rr}"].value = row["week_ending"].to_pydatetime()
        ws[f"A{rr}"].number_format = FMT["date"]
        for j, t in enumerate(tick):
            B.inp(ws, f"{col(2+j)}{rr}", round(float(row[t]), 4), "#,##0.00")
            if i > 0:
                c = col(2 + j)
                B.f(ws, f"{col(2+len(tick)+1+j)}{rr}", f"={c}{rr}/{c}{rr-1}-1", "0.00%")
    wlast = wr0 + nW - 1
    # monthly block to the right
    mo = 2 + 2 * len(tick) + 2
    mr0 = wr0
    ws.cell(row=mr0 - 2, column=mo, value="Monthly closes (inputs) and returns (formulas)").font = Font(bold=True)
    B.header_row(ws, mr0 - 1, mo, ["Month end"] + tick + [""] + [f"r {t}" for t in tick])
    for i, row in m.iterrows():
        rr = mr0 + i
        cdate = ws.cell(row=rr, column=mo, value=row["month_end"].to_pydatetime())
        cdate.number_format = FMT["date"]
        for j, t in enumerate(tick):
            B.inp(ws, f"{col(mo+1+j)}{rr}", round(float(row[t]), 4), "#,##0.00")
            if i > 0:
                c = col(mo + 1 + j)
                B.f(ws, f"{col(mo+1+len(tick)+1+j)}{rr}", f"={c}{rr}/{c}{rr-1}-1", "0.00%")
    mlast = mr0 + len(m) - 1
    # summary formulas
    wret = lambda j: f"{col(2+len(tick)+1+j)}{wr0+1}:{col(2+len(tick)+1+j)}{wlast}"
    mret = lambda j: f"{col(mo+1+len(tick)+1+j)}{mr0+1}:{col(mo+1+len(tick)+1+j)}{mlast}"
    for k, t in enumerate(ALL):
        r = 5 + k
        j = k + 1
        B.label(ws, f"A{r}", t, bold=(t == "PLNT"))
        B.f(ws, f"B{r}", f"=SLOPE({wret(j)},{wret(0)})", "beta", name=f"bw_{t}")
        B.f(ws, f"C{r}", f"=RSQ({wret(j)},{wret(0)})", "beta", name=f"rw_{t}")
        B.f(ws, f"D{r}", f"=COUNT({wret(j)})", "n")
        B.f(ws, f"E{r}", f"=SLOPE({mret(j)},{mret(0)})", "beta", name=f"bm_{t}")
        B.f(ws, f"F{r}", f"=RSQ({mret(j)},{mret(0)})", "beta", name=f"rm_{t}")
        B.f(ws, f"G{r}", f"=COUNT({mret(j)})", "n")
    # 52-week trading range (daily PLNT closes)
    rows = json.load(open(ROOT / "data" / "raw" / "prices_PLNT.json"))["data"]["tradesTable"]["rows"]
    daily = sorted(((datetime.strptime(x["date"], "%m/%d/%Y"), float(x["close"].replace("$", "").replace(",", ""))) for x in rows))
    daily = [d for d in daily if datetime(2025, 11, 20) <= d[0] <= datetime(2026, 11, 19)]
    dc = mo + 2 * len(tick) + 3
    ws.cell(row=mr0 - 2, column=dc, value="PLNT daily closes, 52 weeks to 2026-11-19 (inputs)").font = Font(bold=True)
    B.header_row(ws, mr0 - 1, dc, ["Date", "Close"])
    for i, (d, px) in enumerate(daily):
        cdate = ws.cell(row=mr0 + i, column=dc, value=d)
        cdate.number_format = FMT["date"]
        B.inp(ws, f"{col(dc+1)}{mr0+i}", px, "ps")
    drng = f"{col(dc+1)}{mr0}:{col(dc+1)}{mr0+len(daily)-1}"
    B.label(ws, "I4", "PLNT 52-week trading range", bold=True)
    B.label(ws, "I5", "52-week low (close)")
    B.f(ws, "L5", f"=MIN({drng})", "ps", name="wk52_lo")
    B.label(ws, "I6", "52-week high (close)")
    B.f(ws, "L6", f"=MAX({drng})", "ps", name="wk52_hi")
    B.label(ws, "I8", "Read-across: PLNT's 2-year weekly beta is ~0 with R-squared ~0. The stock fell ~50% in 2026 on company-specific news", italic=True)
    B.label(ws, "I9", "(same-club sales slowed to 1.5% in Q3 2026; CFO transition), so the regression measures idiosyncratic risk, not market risk.", italic=True)
    B.label(ws, "I10", "The WACC therefore uses a bottom-up peer-median unlevered beta (5-year monthly), relevered at the target structure.", italic=True)
    ws.freeze_panes = "B5"


# =============================================================================
# WACC
# =============================================================================
def build_wacc(peer_first, peer_last):
    ws = B.sheet("WACC")
    B.title(ws, "Weighted average cost of capital", "CAPM cost of equity; after-tax cost of debt; target capital structure from peers.", 10)
    B.widths(ws, {"A": 58, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 90})
    r = 4
    B.section(ws, r, "Market inputs", 10)
    for lab, key, fmt, nm in [("Share price, PLNT ($)", "price", "ps", "price"),
                              ("Risk-free rate: 10-year US Treasury", "rf", "pct2", "rf"),
                              ("Equity risk premium (Damodaran implied, base)", "erp", "pct2", "erp"),
                              ("Equity risk premium (Damodaran implied, alternative - memo)", "erp_alt", "pct2", "erp_alt")]:
        r += 1
        B.label(ws, f"A{r}", lab)
        B.inp(ws, f"B{r}", I.MARKET[key][0], fmt, name=nm)
        B.label(ws, f"H{r}", I.MARKET[key][1], italic=True)
    r += 1
    B.label(ws, f"A{r}", "Size premium (optional)")
    B.inp(ws, f"B{r}", 0.0, "pct2", name="size_prem")
    B.label(ws, f"H{r}", "Set to 0. PLNT's equity value (~$4.5B) sits in CRSP deciles where Kroll's size premium is ~1%, but Kroll data is paid, and the ERP/beta already reflect a mid-cap peer set. Enter a value here to test it.", italic=True)

    r += 2
    B.section(ws, r, "Beta", 10)
    r += 1
    B.label(ws, f"A{r}", "PLNT raw beta - 2-year weekly (R-squared in C)")
    B.f(ws, f"B{r}", "={bw_PLNT}", "beta", name="b_plnt_w")
    B.f(ws, f"C{r}", "={rw_PLNT}", "beta")
    r += 1
    B.label(ws, f"A{r}", "PLNT raw beta - 5-year monthly (R-squared in C)")
    B.f(ws, f"B{r}", "={bm_PLNT}", "beta", name="b_plnt_m")
    B.f(ws, f"C{r}", "={rm_PLNT}", "beta")
    r += 2
    B.header_row(ws, r, 1, ["Peer", "Levered (5y m)", "Levered (2y w)", "Debt/Equity", "Tax rate", "Unlevered (5y m)", "Unlevered (2y w)"])
    first = r + 1
    for k, t in enumerate(PEERS):
        r += 1
        crow = peer_first + k
        B.label(ws, f"A{r}", t)
        B.f(ws, f"B{r}", f"={{bm_{t}}}", "beta")
        B.f(ws, f"C{r}", f"={{bw_{t}}}", "beta")
        B.f(ws, f"D{r}", f"=Comps!T{crow}", "pct")
        B.f(ws, f"E{r}", "={tax}", "pct")
        B.f(ws, f"F{r}", f"=B{r}/(1+(1-E{r})*D{r})", "beta")
        B.f(ws, f"G{r}", f"=C{r}/(1+(1-E{r})*D{r})", "beta")
    last = r
    r += 1
    B.label(ws, f"A{r}", "Peer median", bold=True)
    B.f(ws, f"D{r}", f"=MEDIAN(D{first}:D{last})", "pct", bold=True)
    B.f(ws, f"F{r}", f"=MEDIAN(F{first}:F{last})", "beta", name="bu_m", bold=True)
    B.f(ws, f"G{r}", f"=MEDIAN(G{first}:G{last})", "beta", name="bu_w", bold=True)
    B.label(ws, f"H{r}", "Unlevering: Hamada, beta_U = beta_L / (1 + (1 - t) x D/E), debt beta assumed 0. D/E at market values from Comps.", italic=True)
    r += 2
    B.label(ws, f"A{r}", "Target debt / total capital (peer median, market values)")
    B.f(ws, f"B{r}", f"=MEDIAN(Comps!S{peer_first}:S{peer_last})", "pct", name="dv_target")
    r += 1
    B.label(ws, f"A{r}", "PLNT current debt / total capital (market)")
    B.f(ws, f"B{r}", "=Comps!S6", "pct", name="dv_plnt")
    r += 1
    B.label(ws, f"A{r}", "Capital structure used (1 = peer target, 2 = PLNT current)")
    B.inp(ws, f"B{r}", 1, "n", name="wt_sel")
    r += 1
    B.label(ws, f"A{r}", "Debt / total capital used")
    B.f(ws, f"B{r}", "=CHOOSE({wt_sel},{dv_target},{dv_plnt})", "pct", name="wd")
    r += 1
    B.label(ws, f"A{r}", "Equity / total capital used")
    B.f(ws, f"B{r}", "=1-{wd}", "pct", name="we")
    r += 1
    B.label(ws, f"A{r}", "Debt / equity used")
    B.f(ws, f"B{r}", "={wd}/{we}", "pct", name="de_target")
    r += 1
    B.label(ws, f"A{r}", "Relevered beta - peer median 5y monthly")
    B.f(ws, f"B{r}", "={bu_m}*(1+(1-{tax})*{de_target})", "beta", name="b_relev_m")
    r += 1
    B.label(ws, f"A{r}", "Relevered beta - peer median 2y weekly")
    B.f(ws, f"B{r}", "={bu_w}*(1+(1-{tax})*{de_target})", "beta", name="b_relev_w")
    r += 1
    B.label(ws, f"A{r}", "Beta used (1 = peer 5y m relevered, 2 = peer 2y w relevered, 3 = PLNT 5y m, 4 = PLNT 2y w)")
    B.inp(ws, f"B{r}", 1, "n", name="beta_sel")
    B.label(ws, f"H{r}", "Base = 1. PLNT's own 2y weekly regression has R-squared ~0 (see Beta tab); its 5y monthly beta (option 3) is a cross-check.", italic=True)
    r += 1
    B.label(ws, f"A{r}", "Beta used", bold=True)
    B.f(ws, f"B{r}", "=CHOOSE({beta_sel},{b_relev_m},{b_relev_w},{b_plnt_m},{b_plnt_w})", "beta", name="beta", bold=True)
    r += 2
    B.section(ws, r, "Cost of equity", 10)
    r += 1
    B.label(ws, f"A{r}", "Cost of equity = Rf + beta x ERP + size premium", bold=True)
    B.f(ws, f"B{r}", "={rf}+{beta}*{erp}+{size_prem}", "pct2", name="ke", bold=True)

    r += 2
    B.section(ws, r, "Cost of debt", 10)
    r += 1
    B.label(ws, f"A{r}", "7-year UST today")
    B.inp(ws, f"B{r}", I.MARKET["ust7_now"][0], "pct2", name="ust7_now")
    B.label(ws, f"H{r}", I.MARKET["ust7_now"][1], italic=True)
    r += 1
    B.label(ws, f"A{r}", "Coupon, 2025-1 Class A-2-II notes (7-yr anticipated repayment, issued Dec-2025)")
    B.inp(ws, f"B{r}", 0.05649, "pct2", name="cpn_2025")
    B.label(ws, f"H{r}", "10-K FY2025 Note 10: Series 2025-1 5.649% Fixed Rate Senior Secured Notes, Class A-2-II; anticipated repayment Dec 2032", italic=True)
    r += 1
    B.label(ws, f"A{r}", "7-year UST on issue date (2025-12-15)")
    B.inp(ws, f"B{r}", I.MARKET["ust7_issue"][0], "pct2", name="ust7_issue")
    B.label(ws, f"H{r}", I.MARKET["ust7_issue"][1] + ". Closing date used as a proxy for pricing date.", italic=True)
    r += 1
    B.label(ws, f"A{r}", "New-issue credit spread")
    B.f(ws, f"B{r}", "={cpn_2025}-{ust7_issue}", "pct2", name="spread")
    r += 1
    B.label(ws, f"A{r}", "(1) Market-based: 7-yr UST today + new-issue spread")
    B.f(ws, f"B{r}", "={ust7_now}+{spread}", "pct2", name="kd1")
    r += 1
    B.label(ws, f"A{r}", "(2) LTM interest expense / average debt (Dec-25, Sep-26)")
    B.f(ws, f"B{r}", "={h_int_exp@J}/AVERAGE({b_debt_principal@G}+{b_vfn@G},{b_debt_principal@J}+{b_vfn@J})", "pct2", name="kd2")
    B.label(ws, f"H{r}", "Cross-check only: includes amortization of financing costs and reflects older, cheaper coupons.", italic=True)
    r += 1
    B.label(ws, f"A{r}", "(3) Weighted-average coupon of notes outstanding (debt table below)")
    B.f(ws, f"B{r}", "={debt_wac}", "pct2", name="kd3")
    r += 1
    B.label(ws, f"A{r}", "Cost of debt method used (1, 2 or 3)")
    B.inp(ws, f"B{r}", 1, "n", name="kd_sel")
    r += 1
    B.label(ws, f"A{r}", "Pre-tax cost of debt", bold=True)
    B.f(ws, f"B{r}", "=CHOOSE({kd_sel},{kd1},{kd2},{kd3})", "pct2", name="kd", bold=True)
    r += 1
    B.label(ws, f"A{r}", "After-tax cost of debt", bold=True)
    B.f(ws, f"B{r}", "={kd}*(1-{tax})", "pct2", name="kd_at", bold=True)
    r += 2
    B.section(ws, r, "WACC", 10)
    r += 1
    B.label(ws, f"A{r}", "WACC = E/V x Ke + D/V x Kd x (1 - t)", bold=True)
    B.f(ws, f"B{r}", "={we}*{ke}+{wd}*{kd_at}", "pct2", name="wacc", bold=True)
    ws[f"B{r}"].fill = INPUT_FILL

    r += 2
    B.section(ws, r, "Securitized debt outstanding at Sep 30, 2026", 10)
    r += 1
    B.header_row(ws, r, 1, ["Tranche", "Principal", "Coupon", "Anticipated repayment"])
    f0 = r + 1
    for name, (p, cpn, ard) in I.DEBT_TERMS.items():
        r += 1
        B.label(ws, f"A{r}", name)
        B.inp(ws, f"B{r}", p, "m3")
        B.inp(ws, f"C{r}", cpn, "pct2")
        B.label(ws, f"D{r}", ard)
    r += 1
    B.label(ws, f"A{r}", "2025 Variable Funding Notes (floating)")
    B.inp(ws, f"B{r}", 75.0, "m3")
    B.inp(ws, f"C{r}", 0.056, "pct2")
    B.label(ws, f"D{r}", "Revolving")
    l0 = r
    r += 1
    B.label(ws, f"A{r}", "Total / weighted-average coupon", bold=True)
    B.f(ws, f"B{r}", f"=SUM(B{f0}:B{l0})", "m", name="debt_tbl_total", bold=True)
    B.f(ws, f"C{r}", f"=SUMPRODUCT(B{f0}:B{l0},C{f0}:C{l0})/B{r}", "pct2", name="debt_wac", bold=True)
    r += 1
    B.label(ws, f"A{r}", "Fair value of fixed-rate notes (10-Q Q3 2026)")
    B.inp(ws, f"B{r}", I.DEBT_FAIR_VALUE, "m3", name="debt_fv")
    r += 1
    B.label(ws, f"A{r}", "Fair value as % of principal (below par => market yields above coupons)")
    B.f(ws, f"B{r}", f"={{debt_fv}}/SUM(B{f0}:B{l0-1})", "pct")
    B.label(ws, f"H{f0}", I.DEBT_SRC + " VFN rate 5.6% at 9/30/26 per 10-Q.", italic=True)


# =============================================================================
# Projections
# =============================================================================
def build_projections():
    ws = B.sheet("Projections")
    B.title(ws, "Projections - revenue build, EBITDA and unlevered free cash flow",
            "FY2025A from Historicals; FY2026E-FY2031E driven by the Live rows on Assumptions. $ in millions.", 10)
    B.widths(ws, {"A": 56, "B": 3, "C": 11, **{c: 11 for c in YC}, "J": 12, "K": 70})
    B.header_row(ws, 4, 3, ["FY2025A"] + PROJ_YEARS)
    rows = {}

    def line(name, row, label, c_tpl, d_tpl, fmt="m", bold=False, cols=YC):
        B.label(ws, f"A{row}", label, bold=bold)
        if c_tpl:
            B.f(ws, f"C{row}", c_tpl, fmt, bold=bold)
        for i, c in enumerate(cols):
            p = "C" if c == "D" else YC[YC.index(c) - 1]
            B.f(ws, f"{c}{row}", d_tpl.replace("#", c).replace("~", p), fmt, bold=bold)
        B.name_row(name, ws, row)

    r = 5
    B.section(ws, r, "Club build", 10)
    line("p_fr_bop", 6, "Franchisee-owned clubs - beginning", None, "={p_fr_eop@~}", "n")
    line("p_fr_adds", 7, "  net additions", None, "={drv_fr_adds@#}", "n")
    line("p_fr_eop", 8, "Franchisee-owned clubs - end", "={k_fr_eop@G}", "={p_fr_bop@#}+{p_fr_adds@#}", "n", True)
    line("p_fr_avg", 9, "  average", "={k_fr_avg@G}", "=AVERAGE({p_fr_bop@#},{p_fr_eop@#})", "m")
    line("p_co_bop", 10, "Corporate-owned clubs - beginning", None, "={p_co_eop@~}", "n")
    line("p_co_adds", 11, "  net additions", None, "={drv_corp_adds@#}", "n")
    line("p_co_eop", 12, "Corporate-owned clubs - end", "={k_co_eop@G}", "={p_co_bop@#}+{p_co_adds@#}", "n", True)
    line("p_co_avg", 13, "  average", "={k_co_avg@G}", "=AVERAGE({p_co_bop@#},{p_co_eop@#})", "m")
    line("p_clubs", 14, "Total clubs - end (excl. 25 Be-Fit clubs acquired Oct-2026)", "={p_fr_eop@C}+{p_co_eop@C}", "={p_fr_eop@#}+{p_co_eop@#}", "n", True)

    B.section(ws, 16, "Revenue build", 10)
    line("p_fr_scs", 17, "Franchise same-club sales growth", None, "={drv_fr_scs@#}", "pct")
    line("p_roy_adj", 18, "Royalty per club vs. SCS (ramp / rate mix)", None, "={drv_roy_adj@#}", "pct")
    line("p_roy_pc", 19, "Royalty revenue per avg. franchise club", "={k_roy_pc@G}", "={p_roy_pc@~}*(1+{p_fr_scs@#}+{p_roy_adj@#})", "num4")
    line("p_roy", 20, "Royalty revenue", "={h_rev_royalty@G}", "={p_roy_pc@#}*{p_fr_avg@#}")
    line("p_oth_g", 21, "Other franchise revenue per club - growth", None, "={drv_oth_fr_g@#}", "pct")
    line("p_oth_pc", 22, "Other franchise revenue per avg. franchise club", "={k_oth_pc@G}", "={p_oth_pc@~}*(1+{p_oth_g@#})", "num4")
    line("p_oth", 23, "Franchise & other fees, placement, HVAC", "={h_rev_other_fr@G}", "={p_oth_pc@#}*{p_fr_avg@#}")
    line("p_naf_ratio", 24, "NAF revenue / royalty revenue", "={k_naf_ratio@G}", "={drv_naf_ratio@#}", "pct")
    line("p_naf", 25, "NAF revenue (pass-through)", "={h_rev_naf@G}", "={p_naf_ratio@#}*{p_roy@#}")
    line("p_fr_seg", 26, "Franchise segment revenue", "={p_roy@C}+{p_oth@C}+{p_naf@C}", "={p_roy@#}+{p_oth@#}+{p_naf@#}", "m", True)
    line("p_co_scs", 27, "Corporate-owned club same-club sales growth", None, "={drv_corp_scs@#}", "pct")
    line("p_co_adj", 28, "Revenue per corporate club vs. SCS (ramp / mix)", None, "={drv_corp_adj@#}", "pct")
    line("p_co_pc", 29, "Revenue per avg. corporate club", "={k_co_pc@G}", "={p_co_pc@~}*(1+{p_co_scs@#}+{p_co_adj@#})", "num4")
    line("p_corp", 30, "Corporate-owned clubs segment revenue", "={h_rev_corp@G}", "={p_co_pc@#}*{p_co_avg@#}", "m", True)
    line("p_eq_g", 31, "Equipment revenue per franchise club - growth", None, "={drv_eq_g@#}", "pct")
    line("p_eq_pc", 32, "Equipment revenue per avg. franchise club", "={k_eq_pc@G}", "={p_eq_pc@~}*(1+{p_eq_g@#})", "num4")
    line("p_equip", 33, "Equipment segment revenue", "={h_rev_equip@G}", "={p_eq_pc@#}*{p_fr_avg@#}", "m", True)
    line("p_rev", 34, "Total revenue", "={p_fr_seg@C}+{p_corp@C}+{p_equip@C}", "={p_fr_seg@#}+{p_corp@#}+{p_equip@#}", "m", True)
    line("p_rev_g", 35, "  growth", None, "={p_rev@#}/{p_rev@~}-1", "pct")
    line("p_rev_xnaf", 36, "Revenue excl. NAF", "={p_rev@C}-{p_naf@C}", "={p_rev@#}-{p_naf@#}", "m", True)

    B.section(ws, 38, "Adjusted EBITDA build (segment margins)", 10)
    line("p_fr_m", 39, "Franchise margin (excl. NAF)", "={h_fr_m@G}", "={drv_fr_margin@#}", "pct")
    line("p_fr_e", 40, "Franchise segment Adj. EBITDA (NAF assumed to break even)", "={h_seg_fr@G}", "=({p_roy@#}+{p_oth@#})*{p_fr_m@#}")
    line("p_co_m", 41, "Corporate-owned clubs margin", "={h_co_m@G}", "={drv_corp_margin@#}", "pct")
    line("p_co_e", 42, "Corporate-owned clubs Adj. EBITDA", "={h_seg_corp@G}", "={p_corp@#}*{p_co_m@#}")
    line("p_eq_m", 43, "Equipment margin", "={h_eq_m@G}", "={drv_eq_margin@#}", "pct")
    line("p_eq_e", 44, "Equipment Adj. EBITDA", "={h_seg_equip@G}", "={p_equip@#}*{p_eq_m@#}")
    line("p_oh_pct", 45, "Unallocated corporate (% revenue excl. NAF)", "={h_oh_pct@G}", "={drv_overhead@#}", "pct")
    line("p_oh", 46, "Unallocated corporate costs", "={h_seg_oh@G}", "=-{p_oh_pct@#}*{p_rev_xnaf@#}")
    line("p_ebitda", 47, "Adjusted EBITDA", "={p_fr_e@C}+{p_co_e@C}+{p_eq_e@C}+{p_oh@C}", "={p_fr_e@#}+{p_co_e@#}+{p_eq_e@#}+{p_oh@#}", "m", True)
    line("p_ebitda_g", 48, "  growth", None, "={p_ebitda@#}/{p_ebitda@~}-1", "pct")
    line("p_ebitda_m", 49, "  margin (% revenue excl. NAF)", "={p_ebitda@C}/{p_rev_xnaf@C}", "={p_ebitda@#}/{p_rev_xnaf@#}", "pct")

    B.section(ws, 51, "Unlevered free cash flow", 10)
    line("p_dep_pct", 52, "Depreciation (% revenue excl. NAF)", "={h_dep_pct@G}", "={drv_dep_pct@#}", "pct")
    line("p_dep", 53, "Depreciation", "={h_dep@G}", "={p_dep_pct@#}*{p_rev_xnaf@#}")
    line("p_amort", 54, "Amortization of intangibles (10-K schedule)", "={h_amort@G}", "={amort_sched@#}")
    line("p_da", 55, "Total D&A", "={p_dep@C}+{p_amort@C}", "={p_dep@#}+{p_amort@#}", "m", True)
    line("p_ebit", 56, "EBIT (Adj. EBITDA - D&A)", "={p_ebitda@C}-{p_da@C}", "={p_ebitda@#}-{p_da@#}", "m", True)
    line("p_taxes", 57, "Less: cash taxes on EBIT", None, "=-{p_ebit@#}*{tax}")
    line("p_nopat", 58, "NOPAT", None, "={p_ebit@#}+{p_taxes@#}", "m", True)
    line("p_da_add", 59, "Plus: D&A", None, "={p_da@#}")
    line("p_capex_pct", 60, "Capex (% revenue excl. NAF)", "={h_capex_pct@G}", "={drv_capex_pct@#}", "pct")
    line("p_capex", 61, "Less: capital expenditures", "=-{h_capex@G}", "=-{p_capex_pct@#}*{p_rev_xnaf@#}")
    line("p_nwc_pct", 62, "NWC (% revenue excl. NAF)", "={b_nwc_pct@G}", "={drv_nwc_pct@#}", "pct")
    line("p_nwc", 63, "Operating net working capital", "={b_nwc@G}", "={p_nwc_pct@#}*{p_rev_xnaf@#}")
    line("p_dnwc", 64, "Less: increase in NWC", None, "=-({p_nwc@#}-{p_nwc@~})")
    line("p_ufcf", 65, "Unlevered free cash flow", None, "={p_nopat@#}+{p_da_add@#}+{p_capex@#}+{p_dnwc@#}", "m", True)
    line("p_ufcf_conv", 66, "  UFCF / Adj. EBITDA", None, "={p_ufcf@#}/{p_ebitda@#}", "pct")
    B.label(ws, "K57", "Unlevered: tax on EBIT at the marginal rate; ignores interest tax shield (captured in WACC) and the TRA (treated in the equity bridge).", italic=True)
    B.label(ws, "K59", "SBC is not added back: it is inside Adj. EBITDA as a real cost.", italic=True)

    B.section(ws, 68, "Discounting (mid-year convention, stub period)", 10)
    line("p_frac", 69, "Fraction of year's cash flow received after valuation date", None, "=IF(\"#\"=\"D\",{stub},1)", "num4")
    for c in YC:  # replace the placeholder IF with clean formulas
        ws[f"{c}69"].value = None
    B.pending = [p for p in B.pending if not (p[0].title == "Projections" and p[1] in [f"{c}69" for c in YC])]
    B.f(ws, "D69", "={stub}", "num4")
    for c in YC[1:]:
        B.inp(ws, f"{c}69", 1, "num4")
    line("p_tend", 70, "Years from valuation date to period end", None, "={p_tend@~}+{p_frac@#}", "num4")
    B.pending = [p for p in B.pending if not (p[0].title == "Projections" and p[1] == "D70")]
    B.f(ws, "D70", "={p_frac@D}", "num4")
    line("p_per", 71, "Discount period (mid-point if mid-year convention)", None, "={p_tend@#}-{midyear}*{p_frac@#}/2", "num4")
    line("p_cf", 72, "Cash flow to discount (UFCF x fraction)", None, "={p_ufcf@#}*{p_frac@#}")
    line("p_df", 73, "Discount factor at WACC", None, "=1/(1+{wacc})^{p_per@#}", "num4")
    line("p_pv", 74, "Present value of UFCF", None, "={p_cf@#}*{p_df@#}", "m", True)
    B.label(ws, "J73", "Sum of PV", bold=True)
    B.f(ws, "J74", "=SUM({p_pv@D:I})", "m", name="pv_sum", bold=True)

    B.section(ws, 76, "FY2026E vs. management outlook (8-K 2026-11-05)", 10)
    B.header_row(ws, 77, 3, ["Model", "Guidance", "Variance"])
    checks = [("Revenue growth", "={p_rev_g@D}", "={g_rev}"),
              ("Adjusted EBITDA growth", "={p_ebitda_g@D}", "={g_ebitda}"),
              ("Capex growth", "=-{p_capex@D}/{h_capex@G}-1", "={g_capex}"),
              ("D&A growth", "={p_da@D}/{h_da@G}-1", "={g_da}")]
    for i, (lab, mdl, gd) in enumerate(checks):
        rr = 78 + i
        B.label(ws, f"A{rr}", lab)
        B.f(ws, f"C{rr}", mdl, "pct")
        B.f(ws, f"D{rr}", gd, "pct")
        B.f(ws, f"E{rr}", f"=C{rr}-D{rr}", "pct")
    B.name_cell("guid_rev_var", ws, "E", 78)
    B.name_cell("guid_ebitda_var", ws, "E", 79)
    rr = 83
    B.label(ws, f"A{rr}", "Organic system openings FY2026E (model) vs. outlook excl. 25 Be-Fit clubs (low-high)")
    B.f(ws, f"C{rr}", "={p_fr_adds@D}+{p_co_adds@D}", "n")
    B.f(ws, f"D{rr}", "={g_open_lo}-25", "n")
    B.f(ws, f"E{rr}", "={g_open_hi}-25", "n")
    B.label(ws, f"K{rr}", "Model uses net additions (after closures); outlook is gross openings.", italic=True)

    B.section(ws, 85, "Revenue growth fade - drivers behind each year (revenue excl. NAF)", 10)
    B.header_row(ws, 86, 3, ["FY2025A"] + PROJ_YEARS)
    line("p_frclub_g", 87, "Average franchise club growth", None, "={p_fr_avg@#}/{p_fr_avg@~}-1", "pct")
    line("p_coclub_g", 88, "Average corporate club growth", None, "={p_co_avg@#}/{p_co_avg@~}-1", "pct")
    line("p_scs_show", 89, "Franchise same-club sales growth", None, "={p_fr_scs@#}", "pct")
    line("p_roypc_g", 90, "Royalty per avg. franchise club growth (SCS + ramp / rate mix)", None, "={p_roy_pc@#}/{p_roy_pc@~}-1", "pct")
    line("p_copc_g", 91, "Revenue per avg. corporate club growth", None, "={p_co_pc@#}/{p_co_pc@~}-1", "pct")
    line("p_rx_g", 92, "Revenue growth excl. NAF", None, "={p_rev_xnaf@#}/{p_rev_xnaf@~}-1", "pct", True)
    B.label(ws, "K87", "Club growth fades as the committed pipeline is worked down (net franchise adds 145 -> 80).", italic=True)
    B.label(ws, "K89", "Same-club sales: 3.0% in FY2027 (guidance-era pricing / Black Card mix) settling to ~2.5% (inflation-like).", italic=True)
    B.label(ws, "K92", "FY2026E is lower excl. NAF because the 2026 NAF rate step-up (2% -> 3%) inflates total revenue growth to ~8%.", italic=True)
    ws.freeze_panes = "C5"


# =============================================================================
# Terminal value
# =============================================================================
def build_tv():
    ws = B.sheet("Terminal Value")
    B.title(ws, "Terminal value - perpetuity growth and exit multiple", "Both terminal values are stated at the end of FY2031 and discounted with the same period.", 8)
    B.widths(ws, {"A": 62, "B": 13, "C": 13, "D": 4, "E": 100})
    rows = [
        ("Inputs", None, None, None, None),
        ("Terminal (perpetuity) growth rate", "inp", 0.030, "pct2", "tg",
         "Below the 10-yr Treasury (5.0%) and long-run US nominal GDP growth (~4%). Mature franchisor: ~2% inflation + ~1% real (pricing, mix, slow net club growth)."),
        ("Exit EV / EBITDA multiple (on FY2031E Adj. EBITDA)", "inp", 12.0, "x", "exit_mult",
         "PLNT trades at 11.5x LTM GAAP EBITDA; peer median is ~14.9x but includes higher-growth Wingstop. 12.0x assumes a mature, slower-growth PLNT in 2031."),
        ("WACC", "f", "={wacc}", "pct2", "tv_wacc", ""),
        ("FY2031E unlevered FCF", "f", "={p_ufcf@I}", "m", "tv_ufcf", ""),
        ("FY2031E Adjusted EBITDA", "f", "={p_ebitda@I}", "m", "tv_ebitda", ""),
        ("Discount period for terminal value (end of FY2031)", "f", "={p_tend@I}", "num4", "tv_n", ""),
        ("Perpetuity growth (Gordon) method", None, None, None, None),
        ("TV at end-FY2031 = UFCF x (1+g) / (WACC - g) x (1+WACC)^(0.5 x mid-year)", "f",
         "={tv_ufcf}*(1+{tg})/({tv_wacc}-{tg})*(1+{tv_wacc})^(0.5*{midyear})", "m", "tv_pgm",
         "The (1+WACC)^0.5 factor restates the Gordon value (which assumes mid-year flows) to a year-end basis, so both methods share one discount period."),
        ("PV of terminal value", "f", "={tv_pgm}/(1+{tv_wacc})^{tv_n}", "m", "pv_tv_pgm", ""),
        ("Enterprise value (PV of UFCF + PV of TV)", "f", "={pv_sum}+{pv_tv_pgm}", "m", "ev_pgm", ""),
        ("TV as % of EV", "f", "={pv_tv_pgm}/{ev_pgm}", "pct", "tvpct_pgm", "Flag in Checks if above 75%."),
        ("Implied exit multiple (TV / FY2031E EBITDA)", "f", "={tv_pgm}/{tv_ebitda}", "x", "impl_mult", "Cross-check against the exit multiple and peer multiples."),
        ("Exit multiple method", None, None, None, None),
        ("TV at end-FY2031 = FY2031E EBITDA x exit multiple", "f", "={tv_ebitda}*{exit_mult}", "m", "tv_exit", ""),
        ("PV of terminal value", "f", "={tv_exit}/(1+{tv_wacc})^{tv_n}", "m", "pv_tv_exit", ""),
        ("Enterprise value (PV of UFCF + PV of TV)", "f", "={pv_sum}+{pv_tv_exit}", "m", "ev_exit", ""),
        ("TV as % of EV", "f", "={pv_tv_exit}/{ev_exit}", "pct", "tvpct_exit", ""),
        ("Implied perpetuity growth rate from exit multiple", "f",
         "=({tv_exit}/(1+{tv_wacc})^(0.5*{midyear})*{tv_wacc}-{tv_ufcf})/({tv_exit}/(1+{tv_wacc})^(0.5*{midyear})+{tv_ufcf})", "pct2", "impl_g",
         "Solves TV = UFCF x (1+g)/(WACC-g) for g, after restating the exit TV to the Gordon (mid-year) basis."),
        ("Selected enterprise value", None, None, None, None),
        ("Method (1 = perpetuity, 2 = exit multiple, 3 = average of both)", "inp", 3, "n", "tv_sel", "Headline uses the average; both methods are shown on Cover and Equity Bridge."),
        ("Selected enterprise value", "f", "=CHOOSE({tv_sel},{ev_pgm},{ev_exit},AVERAGE({ev_pgm},{ev_exit}))", "m", "ev_sel", ""),
    ]
    r = 3
    for item in rows:
        r += 1
        if item[1] is None:
            B.section(ws, r, item[0], 8)
            continue
        lab, kind, v, fmt, nm, note = item
        B.label(ws, f"A{r}", lab, bold=nm in ("ev_pgm", "ev_exit", "ev_sel"))
        if kind == "inp":
            B.inp(ws, f"B{r}", v, fmt, name=nm).fill = INPUT_FILL
        else:
            B.f(ws, f"B{r}", v, fmt, name=nm, bold=nm in ("ev_pgm", "ev_exit", "ev_sel"))
        if note:
            B.label(ws, f"E{r}", note, italic=True)


# =============================================================================
# Equity bridge
# =============================================================================
def build_bridge():
    ws = B.sheet("Equity Bridge")
    B.title(ws, "Enterprise value to equity value per share",
            "Balance sheet items at Sep 30, 2026 (10-Q Q3 2026) adjusted for the Oct-30-2026 Be-Fit acquisition. Inputs in column G; signs are as they affect equity value.", 9)
    B.widths(ws, {"A": 66, "B": 3, "C": 13, "D": 13, "E": 13, "F": 3, "G": 13, "H": 110})
    B.header_row(ws, 4, 3, ["Perpetuity", "Exit multiple", "Selected", "", "Input"])
    B.label(ws, "A5", "Enterprise value", bold=True)
    B.f(ws, "C5", "={ev_pgm}", "m", bold=True)
    B.f(ws, "D5", "={ev_exit}", "m", bold=True)
    B.f(ws, "E5", "={ev_sel}", "m", bold=True)
    bs = I.BS
    items = [
        ("Plus: cash and cash equivalents", "b_cash", None, "10-Q Q3 2026 balance sheet. Restricted cash ($72.8M, securitization reserves) is excluded: it is not available to equity."),
        ("Plus: marketable securities (short- and long-term)", "b_mkt_sec", None, bs["mkt_sec"][3]),
        ("Plus: investments, net (equity-method stake in Planet Fitmex + held-to-maturity preferred)", "b_htm_inv", None,
         "Non-operating: their income is excluded from Adj. EBITDA. Carried at book value (10-Q: Fitmex carrying value $45.3M)."),
        ("Less: Be-Fit cash consideration paid Oct 30, 2026 (after balance sheet date)", None, -I.BEFIT["cash_paid"][0], I.BEFIT["cash_paid"][1]),
        ("Plus: Be-Fit business at purchase price (not in the projections)", None, "=-G9+G11*-1", "Acquired at an arm's-length price, so valued at cost; the projections exclude its 25 clubs. Net effect of the three Be-Fit lines = 0."),
        ("Less: Be-Fit contingent consideration", None, -I.BEFIT["contingent"][0], I.BEFIT["contingent"][1]),
        ("Less: securitized notes (principal)", "b_debt_principal", "neg", "Principal (not carrying value net of deferred financing costs): it is what must be repaid."),
        ("Less: Variable Funding Notes drawn", "b_vfn", "neg", bs["vfn"][3]),
        ("Less: finance lease liabilities", "b_fin_lease", "neg", bs["fin_lease"][3]),
        ("Less: tax receivable agreement (TRA) payable", "b_tra", "neg",
         "85% of tax savings from the IPO step-up are owed to pre-IPO owners: a debt-like claim. Book value (undiscounted)."),
        ("Plus: net deferred tax assets x toggle (value of the tax step-up retained by PLNT)", None, "dta", "The DTA is the source of the savings the TRA shares; including it offsets the TRA. Toggle below (1 = include)."),
        ("Less: non-controlling interest", None, 0, "Zero: NCI = Class B holders' share of Pla-Fit Holdings; those 0.27M units are counted in diluted shares (exchangeable 1:1) instead."),
    ]
    r = 6
    first = r
    for lab, key, v, note in items:
        B.label(ws, f"A{r}", lab)
        if key and v is None:
            B.f(ws, f"G{r}", f"={{{key}@J}}", "m3")
        elif key and v == "neg":
            B.f(ws, f"G{r}", f"=-{{{key}@J}}", "m3")
        elif v == "dta":
            B.f(ws, f"G{r}", "={dta_toggle}*({b_dta@J}-{b_dtl@J})", "m3")
        elif isinstance(v, str):
            B.f(ws, f"G{r}", v, "m3")
        else:
            B.inp(ws, f"G{r}", v, "m3")
        for c in "CDE":
            B.f(ws, f"{c}{r}", f"=$G{r}", "m")
        B.label(ws, f"H{r}", note, italic=True)
        r += 1
    last = r - 1
    # fix Be-Fit business line: = cash paid + contingent (both entered as negatives)
    B.pending = [p for p in B.pending if not (p[0].title == "Equity Bridge" and p[1] == "G10")]
    B.f(ws, "G10", "=-(G9+G11)", "m3")
    B.label(ws, f"A{r}", "Total adjustments", bold=True)
    B.f(ws, f"G{r}", f"=SUM(G{first}:G{last})", "m", name="br_adj_total", bold=True)
    for c in "CDE":
        B.f(ws, f"{c}{r}", f"=SUM({c}{first}:{c}{last})", "m", bold=True)
    r += 1
    B.label(ws, f"A{r}", "Equity value", bold=True)
    for c in "CDE":
        B.f(ws, f"{c}{r}", f"={c}5+{c}{r-1}", "m", bold=True)
    B.name_cell("eq_pgm", ws, "C", r); B.name_cell("eq_exit", ws, "D", r); B.name_cell("eq_sel", ws, "E", r)
    # toggle
    B.label(ws, f"A{r+2}", "Toggle: include net deferred tax assets (1 = yes, 0 = no)")
    B.inp(ws, f"G{r+2}", 1, "n", name="dta_toggle").fill = INPUT_FILL
    B.label(ws, f"A{r+3}", "Memo: restricted cash excluded from the bridge")
    B.f(ws, f"G{r+3}", "={b_restricted_cash@J}", "m3")
    B.label(ws, f"A{r+4}", "Memo: operating lease liabilities, not treated as debt (EBITDA is post-rent)")
    B.f(ws, f"G{r+4}", "={b_op_lease@J}", "m3")

    r = r + 6
    B.section(ws, r, "Diluted shares outstanding (treasury stock method)", 9)
    sh = I.SHARES
    spec = [("Class A shares outstanding (m)", "class_a", "n3"), ("Class B shares / Holdings units, exchangeable 1:1 (m)", "class_b", "n3"),
            ("Unvested RSUs (m)", "rsu", "n3"), ("Unvested PSUs at target (m)", "psu", "n3"),
            ("Options outstanding (m)", "options", "n3"), ("Options weighted-average strike ($)", "option_strike", "ps")]
    for lab, k, fmt in spec:
        r += 1
        B.label(ws, f"A{r}", lab)
        B.inp(ws, f"G{r}", sh[k][0], fmt, name=f"sh_{k}")
        B.label(ws, f"H{r}", sh[k][1], italic=True)
    r += 1
    B.label(ws, f"A{r}", "Share price used for treasury stock method (current price, avoids circularity)")
    B.f(ws, f"G{r}", "={price}", "ps")
    r += 1
    B.label(ws, f"A{r}", "Net new shares from options = options x (1 - strike / price) if in the money")
    B.f(ws, f"G{r}", f"=IF(G{r-1}>{{sh_option_strike}},{{sh_options}}*(1-{{sh_option_strike}}/G{r-1}),0)", "n3", name="sh_opt_net")
    r += 1
    B.label(ws, f"A{r}", "Diluted shares outstanding (m)", bold=True)
    B.f(ws, f"G{r}", "={sh_class_a}+{sh_class_b}+{sh_rsu}+{sh_psu}+{sh_opt_net}", "n3", name="br_dil_shares", bold=True)

    r += 2
    B.section(ws, r, "Value per share", 9)
    r += 1
    B.label(ws, f"A{r}", "Implied value per share ($)", bold=True)
    B.f(ws, f"C{r}", "={eq_pgm}/{br_dil_shares}", "ps", name="vps_pgm", bold=True)
    B.f(ws, f"D{r}", "={eq_exit}/{br_dil_shares}", "ps", name="vps_exit", bold=True)
    B.f(ws, f"E{r}", "={eq_sel}/{br_dil_shares}", "ps", name="vps", bold=True)
    r += 1
    B.label(ws, f"A{r}", "Current share price ($)")
    for c in "CDE":
        B.f(ws, f"{c}{r}", "={price}", "ps")
    r += 1
    B.label(ws, f"A{r}", "Upside / (downside) vs. current price")
    for c in "CDE":
        B.f(ws, f"{c}{r}", f"={c}{r-2}/{c}{r-1}-1", "pct")
    B.name_cell("upside", ws, "E", r)
    r += 2
    B.section(ws, r, "Implied multiples and market-implied enterprise value", 9)
    r += 1
    B.label(ws, f"A{r}", "EV / LTM Adj. EBITDA (Sep-26)")
    for c in "CDE":
        B.f(ws, f"{c}{r}", f"={c}5/{{h_adj_ebitda@J}}", "x")
    B.name_cell("impl_ltm_mult", ws, "E", r)
    r += 1
    B.label(ws, f"A{r}", "EV / FY2026E Adj. EBITDA")
    for c in "CDE":
        B.f(ws, f"{c}{r}", f"={c}5/{{p_ebitda@D}}", "x")
    B.name_cell("impl_fy26_mult", ws, "E", r)
    r += 1
    B.label(ws, f"A{r}", "EV / FY2027E Adj. EBITDA")
    for c in "CDE":
        B.f(ws, f"{c}{r}", f"={c}5/{{p_ebitda@E}}", "x")
    r += 1
    B.label(ws, f"A{r}", "Market-implied EV at current price = price x diluted shares - total adjustments")
    B.f(ws, f"E{r}", "={price}*{br_dil_shares}-{br_adj_total}", "m", name="mkt_ev")
    r += 1
    B.label(ws, f"A{r}", "Market EV / LTM Adj. EBITDA")
    B.f(ws, f"E{r}", "={mkt_ev}/{h_adj_ebitda@J}", "x", name="mkt_ltm_mult")


# =============================================================================
# Sensitivities
# =============================================================================
def build_sens():
    ws = B.sheet("Sensitivities")
    B.title(ws, "Sensitivity tables - implied value per share ($)",
            "Formula grids (not Excel Data Tables): each cell re-discounts the projected cash flows at its own WACC.", 9)
    B.widths(ws, {"A": 26, "B": 12, **{c: 12 for c in "CDEFGH"}, "I": 60})
    B.label(ws, "A4", "WACC step")
    B.inp(ws, "B4", 0.005, "pct2", name="st_w")
    B.label(ws, "C4", "g step")
    B.inp(ws, "D4", 0.005, "pct2", name="st_g")
    B.label(ws, "E4", "Multiple step")
    B.inp(ws, "F4", 1.0, "x", name="st_x")
    pv = "SUMPRODUCT({p_cf@D:I}/(1+$B#)^{p_per@D:I})"

    def grid(top, title, col_tpl, cell_tpl, colfmt, nm):
        B.section(ws, top, title, 9)
        B.label(ws, f"A{top+2}", "WACC (rows)", bold=True)
        for j, c in enumerate("CDEFG"):
            B.f(ws, f"{c}{top+1}", col_tpl.replace("@k", str(j - 2)), colfmt, bold=True)
        for i in range(5):
            rr = top + 2 + i
            B.f(ws, f"B{rr}", f"={{wacc}}+({i-2})*{{st_w}}", "pct2", bold=True)
            for c in "CDEFG":
                B.f(ws, f"{c}{rr}", cell_tpl.replace("#", str(rr)).replace("@C", c).replace("@T", str(top + 1)), "ps")
        B.name_cell(nm + "_center", ws, "E", top + 4)
        B.name_cell(nm + "_min3", ws, "D", top + 3)
        return top + 2, top + 6

    tpl1 = ("=(" + pv + "+{tv_ufcf}*(1+@C$@T)/($B#-@C$@T)*(1+$B#)^(0.5*{midyear})/(1+$B#)^{tv_n}+{br_adj_total})/{br_dil_shares}")
    a, b = grid(6, "Perpetuity growth method: WACC (rows) x terminal growth (columns)", "={tg}+(@k)*{st_g}", tpl1, "pct2", "s1")
    B.f(ws, "H8", "=MIN(D9:F11)", "ps", name="ff_pgm_lo")
    B.f(ws, "H9", "=MAX(D9:F11)", "ps", name="ff_pgm_hi")
    B.label(ws, "H7", "Inner 3x3 range", bold=True)
    tpl2 = ("=(" + pv + "+{tv_ebitda}*@C$@T/(1+$B#)^{tv_n}+{br_adj_total})/{br_dil_shares}")
    a, b = grid(15, "Exit multiple method: WACC (rows) x exit EV/EBITDA multiple (columns)", "={exit_mult}+(@k)*{st_x}", tpl2, "x", "s2")
    B.f(ws, "H17", "=MIN(D18:F20)", "ps", name="ff_exit_lo")
    B.f(ws, "H18", "=MAX(D18:F20)", "ps", name="ff_exit_hi")
    B.label(ws, "H16", "Inner 3x3 range", bold=True)
    # TV share of EV grid for the perpetuity method
    B.section(ws, 24, "Perpetuity method: terminal value as % of EV (WACC x g)", 9)
    for j, c in enumerate("CDEFG"):
        B.f(ws, f"{c}25", f"={{tg}}+({j-2})*{{st_g}}", "pct2", bold=True)
    for i in range(5):
        rr = 26 + i
        B.f(ws, f"B{rr}", f"={{wacc}}+({i-2})*{{st_w}}", "pct2", bold=True)
        for c in "CDEFG":
            tv = f"{{tv_ufcf}}*(1+{c}$25)/($B{rr}-{c}$25)*(1+$B{rr})^(0.5*{{midyear}})/(1+$B{rr})^{{tv_n}}"
            B.f(ws, f"{c}{rr}", f"={tv}/(SUMPRODUCT({{p_cf@D:I}}/(1+$B{rr})^{{p_per@D:I}})+{tv})", "pct")
    B.label(ws, "I8", "Rows: base WACC +/- steps. Columns: base g +/- steps. Centre cell = model output (checked on Checks tab).", italic=True)


# =============================================================================
# Reverse DCF
# =============================================================================
def build_reverse():
    ws = B.sheet("Reverse DCF")
    B.title(ws, "Reverse DCF - what the current share price implies",
            "Compact constant-growth version of the model (FY2027E-FY2031E) solved for the growth / margin / terminal assumptions that equal today's market EV.", 9)
    B.widths(ws, {"A": 66, "B": 13, "C": 13, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 70})
    r = 4
    B.section(ws, r, "Inputs (linked)", 9)
    lines = [
        ("Market-implied enterprise value at current price", "={mkt_ev}", "m", "rv_target"),
        ("FY2026E revenue excl. NAF (base year)", "={p_rev_xnaf@D}", "m", "rv_r0"),
        ("PV of FY2026E stub cash flow (from full model)", "={p_pv@D}", "m", "rv_pv0"),
        ("Base-case revenue CAGR FY2026E-FY2031E (excl. NAF)", "=({p_rev_xnaf@I}/{p_rev_xnaf@D})^(1/5)-1", "pct2", "rv_g0"),
        ("Base-case average Adj. EBITDA margin FY2027E-FY2031E (excl. NAF)", "=SUM({p_ebitda@E:I})/SUM({p_rev_xnaf@E:I})", "pct2", "rv_m0"),
        ("D&A % of revenue excl. NAF (FY2027E-31E average)", "=SUM({p_da@E:I})/SUM({p_rev_xnaf@E:I})", "pct2", "rv_da"),
        ("Capex % of revenue excl. NAF (FY2027E-31E average)", "=-SUM({p_capex@E:I})/SUM({p_rev_xnaf@E:I})", "pct2", "rv_cx"),
        ("NWC % of revenue excl. NAF", "={p_nwc_pct@I}", "pct2", "rv_nwc"),
        ("Tax rate", "={tax}", "pct2", "rv_tax"),
        ("WACC", "={wacc}", "pct2", "rv_w"),
        ("Terminal growth", "={tg}", "pct2", "rv_tg"),
        ("Terminal discount period", "={tv_n}", "num4", "rv_n"),
    ]
    for lab, tpl, fmt, nm in lines:
        r += 1
        B.label(ws, f"A{r}", lab)
        B.f(ws, f"B{r}", tpl, fmt, name=nm)
    r += 1
    B.label(ws, f"A{r}", "Year index t (FY2027E = 1 ... FY2031E = 5)")
    for j, c in enumerate("CDEFG"):
        B.inp(ws, f"{c}{r}", j + 1, "n")
    B.name_row("rv_t", ws, r)
    r += 1
    B.label(ws, f"A{r}", "Discount period (from Projections)")
    for c, pc in zip("CDEFG", "EFGHI"):
        B.f(ws, f"{c}{r}", f"={{p_per@{pc}}}", "num4")
    B.name_row("rv_p", ws, r)

    def ev(g, m):
        K = f"({m}*(1-{{rv_tax}})+{{rv_da}}*{{rv_tax}}-{{rv_cx}})"
        R = "{rv_r0}"
        T = "{rv_t@C:G}"
        P = "{rv_p@C:G}"
        flows = f"SUMPRODUCT(({R}*(1+{g})^{T}*{K}-{{rv_nwc}}*{R}*((1+{g})^{T}-(1+{g})^({T}-1)))/(1+{{rv_w}})^{P})"
        f5 = f"({R}*(1+{g})^5*{K}-{{rv_nwc}}*{R}*((1+{g})^5-(1+{g})^4))"
        tv = f"{f5}*(1+{{rv_tg}})/({{rv_w}}-{{rv_tg}})*(1+{{rv_w}})^(0.5*{{midyear}})/(1+{{rv_w}})^{{rv_n}}"
        return f"{{rv_pv0}}+{flows}+{tv}"

    r += 2
    B.section(ws, r, "Calibration: compact model at base-case growth and margin", 9)
    r += 1
    B.label(ws, f"A{r}", "Compact-model EV at base CAGR and base average margin")
    B.f(ws, f"B{r}", "=" + ev("{rv_g0}", "{rv_m0}"), "m", name="rv_ev0")
    r += 1
    B.label(ws, f"A{r}", "Full-model EV (perpetuity method)")
    B.f(ws, f"B{r}", "={ev_pgm}", "m")
    r += 1
    B.label(ws, f"A{r}", "Difference (compact uses constant growth and margin)")
    B.f(ws, f"B{r}", f"=B{r-2}/B{r-1}-1", "pct")

    r += 2
    B.section(ws, r, "Results", 9)
    r += 1
    B.label(ws, f"A{r}", "(a) Implied revenue CAGR FY2026E-31E, holding base-case margin", bold=True)
    res_a = r
    r += 1
    B.label(ws, f"A{r}", "(b) Implied average Adj. EBITDA margin, holding base-case growth", bold=True)
    ev1 = ev("{rv_g0}", "({rv_m0}+0.01)")
    B.f(ws, f"B{r}", f"={{rv_m0}}+({{rv_target}}-{{rv_ev0}})/(({ev1})-{{rv_ev0}})*0.01", "pct2", name="rv_impl_m", bold=True)
    B.label(ws, f"I{r}", "EV is linear in margin, so two evaluations solve it exactly.", italic=True)
    r += 1
    B.label(ws, f"A{r}", "(c) Implied terminal growth, holding the full base-case projections", bold=True)
    B.f(ws, f"B{r}", "=(({rv_target}-{pv_sum})*(1+{wacc})^{tv_n}/(1+{wacc})^(0.5*{midyear})*{wacc}-{tv_ufcf})/(({rv_target}-{pv_sum})*(1+{wacc})^{tv_n}/(1+{wacc})^(0.5*{midyear})+{tv_ufcf})",
        "pct2", name="rv_impl_tg", bold=True)
    r += 1
    B.label(ws, f"A{r}", "(d) Implied exit multiple on FY2031E EBITDA, holding the full base-case projections", bold=True)
    B.f(ws, f"B{r}", "=({rv_target}-{pv_sum})*(1+{wacc})^{tv_n}/{tv_ebitda}", "x", name="rv_impl_x", bold=True)
    r += 1
    B.label(ws, f"A{r}", "Base case for comparison: revenue CAGR / margin / terminal growth / exit multiple")
    B.f(ws, f"B{r}", "={rv_g0}", "pct2"); B.f(ws, f"C{r}", "={rv_m0}", "pct2")
    B.f(ws, f"D{r}", "={tg}", "pct2"); B.f(ws, f"E{r}", "={exit_mult}", "x")

    r += 2
    B.section(ws, r, "Bisection solver for (a): 40 iterations, bracket -15% to +30% (EV rises with growth)", 9)
    r += 1
    B.header_row(ws, r, 1, ["Iteration", "Low", "High", "Mid", "EV(mid) - target"])
    for k in range(40):
        r += 1
        ws[f"A{r}"].value = k + 1
        if k == 0:
            B.inp(ws, f"B{r}", -0.15, "pct2")
            B.inp(ws, f"C{r}", 0.30, "pct2")
        else:
            B.f(ws, f"B{r}", f"=IF(E{r-1}>0,B{r-1},D{r-1})", "pct2")
            B.f(ws, f"C{r}", f"=IF(E{r-1}>0,D{r-1},C{r-1})", "pct2")
        B.f(ws, f"D{r}", f"=(B{r}+C{r})/2", "pct2")
        B.f(ws, f"E{r}", "=" + ev(f"D{r}", "{rv_m0}") + "-{rv_target}", "m")
    B.f(ws, f"B{res_a}", f"=D{r}", "pct2", name="rv_impl_g", bold=True)
    B.f(ws, f"C{res_a}", f"=E{r}", "m")
    B.label(ws, f"I{res_a}", "Column C = residual EV error after 40 bisection steps ($M, ~0).", italic=True)


# =============================================================================
# Checks
# =============================================================================
def build_checks():
    ws = B.sheet("Checks")
    B.title(ws, "Model integrity checks", "ERROR if any hard check fails; warnings are judgement flags.", 6)
    B.widths(ws, {"A": 86, "B": 14, "C": 12, "D": 10, "E": 70})
    B.header_row(ws, 4, 1, ["Check", "Value", "Tolerance", "Status"])
    hard = [
        ("Historical balance sheets balance (max abs difference, FY2021-FY2025 & Sep-26)", "=MAX(ABS({b_bal_chk@C}),ABS({b_bal_chk@D}),ABS({b_bal_chk@E}),ABS({b_bal_chk@F}),ABS({b_bal_chk@G}),ABS({b_bal_chk@J}))", 0.0015),
        ("Revenue build = reported revenue (sum of abs differences, all periods)", "=SUMPRODUCT(ABS({h_rev_chk@C:I}))", 0.0015),
        ("Recomputed EBIT = reported EBIT (sum of abs differences)", "=SUMPRODUCT(ABS({h_ebit_chk@C:I}))", 0.005),
        ("Franchise club roll-forward ties to reported counts", "=SUMPRODUCT(ABS({k_fr_chk@C:G}))+ABS({k_fr_chk@I})", 0),
        ("Corporate club roll-forward ties to reported counts", "=SUMPRODUCT(ABS({k_co_chk@C:G}))+ABS({k_co_chk@I})", 0),
        ("LTM segment revenue = LTM revenue (as reported)", "={h_rev@J}-{h_rev_total_xbrl@J}", 0.0015),
        ("FY2025A segment revenue in Projections = Historicals", "={p_rev@C}-{h_rev@G}", 0.0005),
        ("FY2025A Adj. EBITDA build in Projections = company-reported", "={p_ebitda@C}-{h_adj_ebitda@G}", 0.0005),
        ("Projected revenue = sum of segments (all years)", "=SUMPRODUCT(ABS({p_rev@D:I}-{p_fr_seg@D:I}-{p_corp@D:I}-{p_equip@D:I}))", 1e-6),
        ("Projected EBITDA = sum of segment EBITDA (all years)", "=SUMPRODUCT(ABS({p_ebitda@D:I}-{p_fr_e@D:I}-{p_co_e@D:I}-{p_eq_e@D:I}-{p_oh@D:I}))", 1e-6),
        ("UFCF = EBIT x (1-t) + D&A - capex - increase in NWC (all years)", "=SUMPRODUCT(ABS({p_ufcf@D:I}-({p_ebit@D:I}*(1-{tax})+{p_da@D:I}+{p_capex@D:I}+{p_dnwc@D:I})))", 1e-6),
        ("Scenario selector is 1, 2 or 3", "=IF(OR({scen}=1,{scen}=2,{scen}=3),0,1)", 0),
        ("WACC > terminal growth", "=IF({wacc}>{tg},0,1)", 0),
        ("Terminal growth below the risk-free rate (long-run nominal growth cap)", "=IF({tg}<{rf},0,1)", 0),
        ("Equity value = EV + adjustments (selected)", "={eq_sel}-{ev_sel}-{br_adj_total}", 1e-6),
        ("Diluted shares >= basic shares", "=IF({br_dil_shares}>={sh_class_a},0,1)", 0),
        ("Sensitivity centre (perpetuity) = model value per share", "={s1_center}-{vps_pgm}", 1e-6),
        ("Sensitivity centre (exit multiple) = model value per share", "={s2_center}-{vps_exit}", 1e-6),
        ("Debt table total = bridge debt (notes + VFN)", "={debt_tbl_total}-({b_debt_principal@J}+{b_vfn@J})", 0.0015),
        ("Comps PLNT LTM revenue (XBRL pipeline) = Historicals LTM revenue", "=Comps!K6-{h_rev@J}", 0.0015),
        ("Comps PLNT LTM GAAP EBITDA (XBRL pipeline) = Historicals LTM EBITDA", "=Comps!N6-{h_ebitda@J}", 0.0015),
        ("Reverse DCF bisection converged (abs residual, $M)", "=ABS(Reverse DCF!C0)", 0.01),
    ]
    warn = [
        ("TV share of EV - perpetuity method <= 75%", "={tvpct_pgm}", 0.75),
        ("TV share of EV - exit multiple method <= 75%", "={tvpct_exit}", 0.75),
        ("WACC within 6%-12%", "=IF(AND({wacc}>=0.06,{wacc}<=0.12),0,1)", 0),
        ("FY2026E revenue growth within 1pp of guidance", "=ABS({guid_rev_var})", 0.01),
        ("FY2026E Adj. EBITDA growth within 1pp of guidance", "=ABS({guid_ebitda_var})", 0.01),
    ]
    r = 4
    first = r + 1
    for lab, tpl, tol in hard:
        r += 1
        if "Reverse DCF!C0" in tpl:
            tpl = "=ABS({rv_resid})"
        B.label(ws, f"A{r}", lab)
        B.f(ws, f"B{r}", tpl, "num4")
        B.inp(ws, f"C{r}", tol, "num4")
        B.f(ws, f"D{r}", f'=IF(ABS(B{r})<=C{r},"OK","ERROR")', "gen", bold=True)
    last = r
    r += 2
    B.label(ws, f"A{r}", "Warnings (judgement flags, not errors)", bold=True)
    wf = r + 1
    for lab, tpl, tol in warn:
        r += 1
        B.label(ws, f"A{r}", lab)
        B.f(ws, f"B{r}", tpl, "num4")
        B.inp(ws, f"C{r}", tol, "num4")
        B.f(ws, f"D{r}", f'=IF(B{r}<=C{r},"OK","FLAG")', "gen", bold=True)
    wl = r
    r += 2
    B.label(ws, f"A{r}", "OVERALL MODEL STATUS (all hard checks)", bold=True)
    B.f(ws, f"D{r}", f'=IF(COUNTIF(D{first}:D{last},"ERROR")=0,"OK","ERROR")', "gen", name="chk_all", bold=True)
    r += 1
    B.label(ws, f"A{r}", "Number of warnings flagged")
    B.f(ws, f"D{r}", f'=COUNTIF(D{wf}:D{wl},"FLAG")', "n", name="chk_warn")


# =============================================================================
# Cover
# =============================================================================
def build_cover():
    ws = B.sheet("Cover")
    B.wb.move_sheet("Cover", offset=-(len(B.wb.sheetnames) - 1))
    B.title(ws, "Planet Fitness, Inc. (NYSE: PLNT) - Discounted cash flow valuation",
            "Educational model built from public SEC filings and free market data; not investment advice. $ in millions except per share.", 10)
    B.widths(ws, {"A": 44, "B": 14, "C": 14, "D": 14, "E": 4, "F": 40, "G": 14, "H": 14, "I": 14, "J": 14})
    rows = [
        ("Valuation summary", None, None),
        ("Scenario", "={scen_name}", "gen"),
        ("Valuation date", "={val_date}", "date"),
        ("Current share price (close 2026-11-19)", "={price}", "ps"),
        ("Implied value per share - selected (average of methods)", "={vps}", "ps"),
        ("Implied value per share - perpetuity growth", "={vps_pgm}", "ps"),
        ("Implied value per share - exit multiple", "={vps_exit}", "ps"),
        ("Upside / (downside) vs. price", "={upside}", "pct"),
        ("Enterprise value - selected", "={ev_sel}", "m"),
        ("Equity value - selected", "={eq_sel}", "m"),
        ("Diluted shares (m)", "={br_dil_shares}", "n3"),
        ("Market-implied EV at current price", "={mkt_ev}", "m"),
        ("Key assumptions", None, None),
        ("WACC", "={wacc}", "pct2"),
        ("Cost of equity / beta", "={ke}", "pct2"),
        ("Pre-tax cost of debt", "={kd}", "pct2"),
        ("Terminal growth / implied exit multiple", "={tg}", "pct2"),
        ("Exit multiple / implied perpetuity growth", "={exit_mult}", "x"),
        ("Revenue CAGR FY2026E-31E (excl. NAF)", "={rv_g0}", "pct"),
        ("Avg. Adj. EBITDA margin FY2027E-31E (excl. NAF)", "={rv_m0}", "pct"),
        ("TV % of EV (perpetuity / exit)", "={tvpct_pgm}", "pct"),
        ("Implied EV / LTM Adj. EBITDA (selected) / market", "={impl_ltm_mult}", "x"),
        ("Reverse DCF", None, None),
        ("Price-implied revenue CAGR (at base margin)", "={rv_impl_g}", "pct"),
        ("Price-implied avg. EBITDA margin (at base growth)", "={rv_impl_m}", "pct"),
        ("Price-implied terminal growth (base projections)", "={rv_impl_tg}", "pct2"),
        ("Model checks", None, None),
        ("Overall status (all checks)", "={chk_all}", "gen"),
        ("Warnings flagged (see Checks tab)", "={chk_warn}", "n"),
    ]
    r = 3
    second = {"Cost of equity / beta": ("={beta}", "beta"), "Terminal growth / implied exit multiple": ("={impl_mult}", "x"),
              "Exit multiple / implied perpetuity growth": ("={impl_g}", "pct2"), "TV % of EV (perpetuity / exit)": ("={tvpct_exit}", "pct"),
              "Implied EV / LTM Adj. EBITDA (selected) / market": ("={mkt_ltm_mult}", "x")}
    for lab, tpl, fmt in rows:
        r += 1
        if tpl is None:
            B.section(ws, r, lab, 4)
            continue
        B.label(ws, f"A{r}", lab, bold=lab.startswith(("Implied value per share - selected", "Overall")))
        B.f(ws, f"B{r}", tpl, fmt, bold=lab.startswith(("Implied value per share - selected", "Overall")))
        if lab in second:
            B.f(ws, f"C{r}", second[lab][0], second[lab][1])
        if lab.startswith("Overall"):
            B.name_cell("cover_status", ws, "B", r)

    # Football field
    B.label(ws, "F3", "Football field - value per share ($)", bold=True)
    B.header_row(ws, 4, 6, ["Method", "Low", "High", "Spread", "Mid"])
    ff = [
        ("52-week trading range", "{wk52_lo}", "{wk52_hi}"),
        ("Comps: EV / LTM EBITDA (25th-75th)", "{cmp_ebitda_lo}", "{cmp_ebitda_hi}"),
        ("Comps: EV / LTM revenue (25th-75th)", "{cmp_rev_lo}", "{cmp_rev_hi}"),
        ("Comps: P / E (25th-75th)", "{cmp_pe_lo}", "{cmp_pe_hi}"),
        ("DCF: perpetuity (WACC +/-0.5%, g +/-0.5%)", "{ff_pgm_lo}", "{ff_pgm_hi}"),
        ("DCF: exit multiple (WACC +/-0.5%, +/-1.0x)", "{ff_exit_lo}", "{ff_exit_hi}"),
    ]
    for i, (lab, lo, hi) in enumerate(ff):
        rr = 5 + i
        B.label(ws, f"F{rr}", lab)
        B.f(ws, f"G{rr}", f"=MIN({lo},{hi})", "ps")
        B.f(ws, f"H{rr}", f"=MAX({lo},{hi})", "ps")
        B.f(ws, f"I{rr}", f"=H{rr}-G{rr}", "ps")
        B.f(ws, f"J{rr}", f"=AVERAGE(G{rr},H{rr})", "ps")
    B.name_cell("ff_first", ws, "F", 5)
    B.label(ws, "F11", "Current share price")
    B.f(ws, "G11", "={price}", "ps")
    B.label(ws, "F12", "DCF value per share (selected)")
    B.f(ws, "G12", "={vps}", "ps")

    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.title = "PLNT football field ($ per share)"
    ch.y_axis.title = "$ per share"
    ch.x_axis.scaling.orientation = "maxMin"
    data = Reference(ws, min_col=7, max_col=7, min_row=4, max_row=10)
    spread = Reference(ws, min_col=9, max_col=9, min_row=4, max_row=10)
    cats = Reference(ws, min_col=6, min_row=5, max_row=10)
    ch.add_data(data, titles_from_data=True)
    ch.add_data(spread, titles_from_data=True)
    ch.set_categories(cats)
    ch.series[0].graphicalProperties.noFill = True
    ch.series[0].graphicalProperties.line.noFill = True
    ch.series[1].graphicalProperties.solidFill = "1F3864"
    ch.legend = None
    ch.height, ch.width = 8, 18
    ws.add_chart(ch, "F14")
    B.label(ws, "F31", "Chart: invisible 'Low' bar + visible 'High - Low' bar. Vertical reference lines for price and DCF value are in the README PNG.", italic=True)
    B.label(ws, "A36", "Tab guide: Assumptions (scenario switch) > Historicals > Projections > WACC / Beta > Terminal Value > Equity Bridge > Sensitivities > Comps > Reverse DCF > Checks.", italic=True)
    B.label(ws, "A37", "Colour code: blue = hard-coded input (sourced), black = formula, green = link from another tab.", italic=True)


def main():
    build_assumptions()
    build_historicals()
    peer_first, peer_last = build_comps()
    build_beta()
    build_wacc(peer_first, peer_last)
    build_projections()
    build_tv()
    build_bridge()
    build_sens()
    build_reverse()
    # residual name for Checks
    ws = B.wb["Reverse DCF"]
    for row in ws.iter_rows(min_col=1, max_col=1):
        if row[0].value == "(a) Implied revenue CAGR FY2026E-31E, holding base-case margin":
            B.name_cell("rv_resid", ws, "C", row[0].row)
    build_checks()
    build_cover()
    B.resolve()
    OUT.parent.mkdir(exist_ok=True)
    B.wb.save(OUT)
    import pickle
    pickle.dump(B.names, open(OUT.parent / ".names.pkl", "wb"))
    print("saved", OUT)


if __name__ == "__main__":
    main()
