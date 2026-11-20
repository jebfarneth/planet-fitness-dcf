"""Independent re-implementation of the workbook in plain Python.

Reads the same raw inputs (inputs.py, assumptions.py, peers_raw.csv, price
panels) but none of the workbook's formulas, recomputes every output, and
compares it with the workbook as recalculated by the `formulas` engine.

Usage:  python src/verify.py            # base case
        python src/verify.py --all      # also recalculates Upside and Downside copies
"""
import csv
import json
import shutil
import sys
from datetime import date, datetime
from pathlib import Path
from statistics import median

import numpy as np
import pandas as pd
from openpyxl import load_workbook

import inputs as I
import recalc
from assumptions import DRIVERS

ROOT = Path(__file__).resolve().parents[1]
PEERS = ["LTH", "XPOF", "DPZ", "WING", "WEN", "CHH", "VVV"]
YC = ["D", "E", "F", "G", "H", "I"]


def H(key, i):            # historical value FY (i = 0..4) from inputs
    return I.HIST[key][1][i]


def LTM(key):
    _, v, a9, b9, _ = I.HIST[key]
    return v[4] + b9 - a9


def BSv(key, i):          # balance sheet FY2021..FY2025 -> i = 0..4, 'sep' for Sep-26
    lab, vals, last, _ = I.BS[key]
    return last if i == "sep" else vals[i + 1]


def percentile_inc(xs, p):
    xs = sorted(xs)
    k = (len(xs) - 1) * p
    f = int(np.floor(k))
    c = min(f + 1, len(xs) - 1)
    return xs[f] + (xs[c] - xs[f]) * (k - f)


def model(scen=1):
    out = {}
    drv = {k: [b, u, d][scen - 1] for k, _, _, b, u, d, _ in DRIVERS}

    # ---------------- historicals (FY2025 anchors) ----------------
    rev = [H("rev_franchise", i) + H("rev_naf", i) + H("rev_corp", i) + H("rev_equip", i) for i in range(5)]
    out["h_rev"] = rev
    rev_ltm = LTM("rev_franchise") + LTM("rev_naf") + LTM("rev_corp") + LTM("rev_equip")
    out["h_rev_ltm"] = rev_ltm
    rev_xnaf25 = rev[4] - H("rev_naf", 4)
    ebit = [rev[i] - H("cogs", i) - H("club_ops", i) - H("sga", i) - H("naf_exp", i) - H("da", i) - H("other_op", i) for i in range(5)]
    out["h_ebit"] = ebit
    ebit_ltm = rev_ltm - sum(LTM(k) for k in ["cogs", "club_ops", "sga", "naf_exp", "da", "other_op"])
    ebitda_ltm = ebit_ltm + LTM("da")
    out["h_ebitda_ltm"] = ebitda_ltm
    adj_ltm = LTM("adj_ebitda")
    # club roll-forward
    fr_eop, fr_bop = [], 2021
    co_eop, co_bop = [], 103
    for i in range(5):
        fr_eop.append(fr_bop + I.KPI["fr_open"][1][i] + I.KPI["fr_closed"][1][i])
        co_eop.append(co_bop + I.KPI["corp_open"][1][i] + I.KPI["corp_other"][1][i])
        fr_bop, co_bop = fr_eop[-1], co_eop[-1]
    assert fr_eop == I.KPI["fr_clubs"][1] and co_eop == I.KPI["corp_clubs"][1]
    fr_avg25 = (fr_eop[3] + fr_eop[4]) / 2
    co_avg25 = (co_eop[3] + co_eop[4]) / 2
    roy25, oth25 = H("rev_royalty", 4), H("rev_franchise", 4) - H("rev_royalty", 4)
    nwc = lambda i: (BSv("ar", i) + BSv("inventory", i) + BSv("prepaid", i) + BSv("other_rec", i)
                     - BSv("ap", i) - BSv("accrued", i) - BSv("equip_dep", i) - BSv("def_rev", i))
    seg_oh25 = H("adj_ebitda", 4) - H("seg_fr", 4) - H("seg_corp", 4) - H("seg_equip", 4)

    # ---------------- projections ----------------
    P = {k: [] for k in ["fr_eop", "co_eop", "roy", "oth", "naf", "corp", "equip", "rev", "rev_xnaf", "ebitda",
                         "da", "ebit", "capex", "nwc", "dnwc", "ufcf"]}
    fr_e, co_e = fr_eop[4], co_eop[4]
    roy_pc, oth_pc, co_pc, eq_pc = roy25 / fr_avg25, oth25 / fr_avg25, H("rev_corp", 4) / co_avg25, H("rev_equip", 4) / fr_avg25
    nwc_prev = nwc(4)
    for t in range(6):
        fr_b, co_b = fr_e, co_e
        fr_e, co_e = fr_b + drv["fr_adds"][t], co_b + drv["corp_adds"][t]
        fr_a, co_a = (fr_b + fr_e) / 2, (co_b + co_e) / 2
        roy_pc *= 1 + drv["fr_scs"][t] + drv["roy_adj"][t]
        oth_pc *= 1 + drv["oth_fr_g"][t]
        co_pc *= 1 + drv["corp_scs"][t] + drv["corp_adj"][t]
        eq_pc *= 1 + drv["eq_g"][t]
        roy, oth = roy_pc * fr_a, oth_pc * fr_a
        naf = drv["naf_ratio"][t] * roy
        corp, equip = co_pc * co_a, eq_pc * fr_a
        rv = roy + oth + naf + corp + equip
        rx = rv - naf
        ebitda = (roy + oth) * drv["fr_margin"][t] + corp * drv["corp_margin"][t] + equip * drv["eq_margin"][t] - drv["overhead"][t] * rx
        da = drv["dep_pct"][t] * rx + list(I.AMORT_SCHEDULE.values())[t]
        e = ebitda - da
        capex = drv["capex_pct"][t] * rx
        n = drv["nwc_pct"][t] * rx
        d = n - nwc_prev
        nwc_prev = n
        ufcf = e * (1 - I.MARKET["tax"][0]) + da - capex - d
        for k, v in zip(P, [fr_e, co_e, roy, oth, naf, corp, equip, rv, rx, ebitda, da, e, capex, n, d, ufcf]):
            P[k].append(v)
    out.update({f"p_{k}": v for k, v in P.items()})

    # ---------------- comps ----------------
    rows = {r["ticker"]: r for r in csv.DictReader(open(ROOT / "data" / "inputs" / "peers_raw.csv"))}
    comps = {}
    for t in ["PLNT"] + PEERS:
        r = rows[t]
        g = lambda k: float(r[k]) / 1e6
        price = I.MARKET["price"][0] if t == "PLNT" else float(r["price"])
        shares = I.SHARES["class_a"][0] + I.SHARES["class_b"][0] if t == "PLNT" else float(r["shares"]) / 1e6
        mcap = price * shares
        stb = g("st_borrow") + (75.0 if t == "PLNT" and g("st_borrow") == 0 else 0)
        debt = g("debt_lt") + g("debt_cur") + stb + g("fin_lease") + g("fin_lease_cur")
        ev = mcap + debt - g("cash") + g("nci")
        ebitda = g("ebit") + g("da")
        ni = g("net_income")
        comps[t] = dict(mcap=mcap, debt=debt, ev=ev, rev=g("revenue"), ebitda=ebitda, ni=ni,
                        ev_rev=ev / g("revenue"), ev_ebitda=ev / ebitda if ebitda > 0 else None,
                        pe=mcap / ni if ni > 0 else None, dv=debt / (debt + mcap), de=debt / mcap)
    out["comps"] = comps
    stat = {}
    for k in ["ev_rev", "ev_ebitda", "pe", "dv"]:
        xs = [comps[t][k] for t in PEERS if comps[t][k] is not None]
        stat[k] = dict(med=median(xs), p25=percentile_inc(xs, 0.25), p75=percentile_inc(xs, 0.75), mean=float(np.mean(xs)))
    out["comp_stats"] = stat

    # ---------------- beta ----------------
    w = pd.read_csv(ROOT / "data" / "inputs" / "weekly_prices.csv").drop(columns="week_ending").round(4)
    m = pd.read_csv(ROOT / "data" / "inputs" / "monthly_prices.csv").drop(columns="month_end").round(4)
    rw, rm = w.pct_change().dropna(), m.pct_change().dropna()
    bw = {t: np.polyfit(rw["SPX"], rw[t], 1)[0] for t in ["PLNT"] + PEERS}
    bm = {t: np.polyfit(rm["SPX"], rm[t], 1)[0] for t in ["PLNT"] + PEERS}
    out["beta_w"], out["beta_m"] = bw, bm
    tax = I.MARKET["tax"][0]
    bu_m = median([bm[t] / (1 + (1 - tax) * comps[t]["de"]) for t in PEERS])
    dv = stat["dv"]["med"]
    de = dv / (1 - dv)
    beta = bu_m * (1 + (1 - tax) * de)
    ke = I.MARKET["rf"][0] + beta * I.MARKET["erp"][0]
    kd = I.MARKET["ust7_now"][0] + (0.05649 - I.MARKET["ust7_issue"][0])
    wacc = (1 - dv) * ke + dv * kd * (1 - tax)
    out.update(beta=beta, ke=ke, kd=kd, wacc=wacc, bu_m=bu_m, dv_target=dv)

    # ---------------- discounting ----------------
    stub = (date(2026, 12, 31) - date(2026, 11, 20)).days / 365
    frac = [stub] + [1] * 5
    tend = np.cumsum(frac)
    per = tend - np.array(frac) / 2
    cf = np.array(P["ufcf"]) * np.array(frac)
    pv = cf / (1 + wacc) ** per
    pv_sum = pv.sum()
    N = tend[-1]
    tg, mult = 0.03, 12.0

    def ev_pgm_at(wc, g):
        pvs = (cf / (1 + wc) ** per).sum()
        tv = P["ufcf"][-1] * (1 + g) / (wc - g) * (1 + wc) ** 0.5
        return pvs + tv / (1 + wc) ** N, tv

    def ev_exit_at(wc, x):
        pvs = (cf / (1 + wc) ** per).sum()
        tv = P["ebitda"][-1] * x
        return pvs + tv / (1 + wc) ** N, tv

    ev_pgm, tv_pgm = ev_pgm_at(wacc, tg)
    ev_exit, tv_exit = ev_exit_at(wacc, mult)
    ev_sel = (ev_pgm + ev_exit) / 2
    out.update(pv_sum=pv_sum, ev_pgm=ev_pgm, ev_exit=ev_exit, ev_sel=ev_sel, tv_pgm=tv_pgm, tv_exit=tv_exit,
               tvpct_pgm=(tv_pgm / (1 + wacc) ** N) / ev_pgm, tvpct_exit=(tv_exit / (1 + wacc) ** N) / ev_exit,
               impl_mult=tv_pgm / P["ebitda"][-1])
    tvg = tv_exit / (1 + wacc) ** 0.5
    out["impl_g"] = (tvg * wacc - P["ufcf"][-1]) / (tvg + P["ufcf"][-1])

    # ---------------- bridge ----------------
    adj = (BSv("cash", "sep") + BSv("mkt_sec", "sep") + BSv("htm_inv", "sep")
           - I.BEFIT["cash_paid"][0] + (I.BEFIT["cash_paid"][0] + I.BEFIT["contingent"][0]) - I.BEFIT["contingent"][0]
           - BSv("debt_principal", "sep") - BSv("vfn", "sep") - BSv("fin_lease", "sep") - BSv("tra", "sep")
           + (BSv("dta", "sep") - BSv("dtl", "sep")))
    price = I.MARKET["price"][0]
    sh = I.SHARES
    opt = sh["options"][0] * (1 - sh["option_strike"][0] / price) if price > sh["option_strike"][0] else 0
    dil = sh["class_a"][0] + sh["class_b"][0] + sh["rsu"][0] + sh["psu"][0] + opt
    out.update(br_adj_total=adj, br_dil_shares=dil, vps_pgm=(ev_pgm + adj) / dil, vps_exit=(ev_exit + adj) / dil,
               vps=(ev_sel + adj) / dil, mkt_ev=price * dil - adj)

    # ---------------- sensitivities ----------------
    out["s1"] = [[(ev_pgm_at(wacc + i * 0.005, tg + j * 0.005)[0] + adj) / dil for j in range(-2, 3)] for i in range(-2, 3)]
    out["s2"] = [[(ev_exit_at(wacc + i * 0.005, mult + j)[0] + adj) / dil for j in range(-2, 3)] for i in range(-2, 3)]

    # ---------------- comps implied ----------------
    pl = comps["PLNT"]
    for nm, k, metric, kind in [("ebitda", "ev_ebitda", pl["ebitda"], "ev"), ("rev", "ev_rev", pl["rev"], "ev"), ("pe", "pe", pl["ni"], "eq")]:
        for lvl, s in [("lo", "p25"), ("mid", "med"), ("hi", "p75")]:
            x = stat[k][s] * metric
            out[f"cmp_{nm}_{lvl}"] = (x + adj) / dil if kind == "ev" else x / dil

    # ---------------- reverse DCF ----------------
    rx = np.array(P["rev_xnaf"])
    r0 = rx[0]
    g0 = (rx[5] / rx[0]) ** (1 / 5) - 1
    m0 = sum(P["ebitda"][1:]) / rx[1:].sum()
    da_p = sum(P["da"][1:]) / rx[1:].sum()
    cx_p = sum(P["capex"][1:]) / rx[1:].sum()
    nwc_p = drv["nwc_pct"][5]
    T = np.arange(1, 6)
    pv0 = pv[0]

    def ev_c(g, mm):
        K = mm * (1 - tax) + da_p * tax - cx_p
        f = r0 * (1 + g) ** T * K - nwc_p * r0 * ((1 + g) ** T - (1 + g) ** (T - 1))
        tv = f[-1] * (1 + tg) / (wacc - tg) * (1 + wacc) ** 0.5 / (1 + wacc) ** N
        return pv0 + (f / (1 + wacc) ** per[1:]).sum() + tv

    target = out["mkt_ev"]
    lo, hi = -0.15, 0.30
    for _ in range(40):
        mid = (lo + hi) / 2
        if ev_c(mid, m0) - target > 0:
            hi = mid
        else:
            lo = mid
    # the workbook reports the 40th midpoint; reproduce exactly
    lo2, hi2 = -0.15, 0.30
    for k in range(40):
        mid2 = (lo2 + hi2) / 2
        if k == 39:
            break
        if ev_c(mid2, m0) - target > 0:
            hi2 = mid2
        else:
            lo2 = mid2
    out["rv_impl_g"] = mid2
    e0, e1 = ev_c(g0, m0), ev_c(g0, m0 + 0.01)
    out["rv_impl_m"] = m0 + (target - e0) / (e1 - e0) * 0.01
    tvn = (target - pv_sum) * (1 + wacc) ** N / (1 + wacc) ** 0.5
    out["rv_impl_tg"] = (tvn * wacc - P["ufcf"][-1]) / (tvn + P["ufcf"][-1])
    out["rv_impl_x"] = (target - pv_sum) * (1 + wacc) ** N / P["ebitda"][-1]
    out.update(rv_g0=g0, rv_m0=m0, rv_ev0=e0)
    return out


# =============================================================================
def compare(vals, reg, py, label):
    rows = []

    def chk(name, wb, pyv, tol):
        ok = wb is not None and abs(wb - pyv) <= tol
        rows.append((name, wb, pyv, (wb - pyv) if wb is not None else None, ok))

    g = lambda n, c=None: recalc.get(vals, reg, n, c)
    # money outputs to the cent ($ per share) / to $0.01M ($M)
    for n in ["vps", "vps_pgm", "vps_exit"]:
        chk(n, g(n), py[n], 0.005)
    for n in ["ev_pgm", "ev_exit", "ev_sel", "pv_sum", "tv_pgm", "tv_exit", "br_adj_total", "mkt_ev"]:
        chk(n, g(n), py[n], 0.005)
    for n in ["wacc", "ke", "kd", "beta", "bu_m", "dv_target", "tvpct_pgm", "tvpct_exit", "impl_g",
              "rv_impl_m", "rv_impl_tg", "rv_g0", "rv_m0", "rv_impl_g"]:
        chk(n, g(n), py[n], 1e-9)
    for n in ["br_dil_shares", "impl_mult", "rv_impl_x"]:
        chk(n, g(n), py[n], 1e-6)
    chk("h_rev_ltm", g("h_rev", "J"), py["h_rev_ltm"], 1e-6)
    chk("h_ebitda_ltm", g("h_ebitda", "J"), py["h_ebitda_ltm"], 1e-6)
    for i, c in enumerate("CDEFG"):
        chk(f"h_ebit {c}", g("h_ebit", c), py["h_ebit"][i], 1e-6)
    for k in ["fr_eop", "co_eop", "roy", "oth", "naf", "corp", "equip", "rev", "rev_xnaf", "ebitda", "da", "ebit", "capex", "nwc", "ufcf"]:
        for i, c in enumerate(YC):
            wv = g(f"p_{k}", c)
            pv = py[f"p_{k}"][i] * (-1 if k == "capex" else 1)
            chk(f"p_{k} {c}", wv, pv, 0.005)
    for t in ["PLNT"] + PEERS:
        chk(f"beta_w {t}", g(f"bw_{t}"), py["beta_w"][t], 1e-9)
        chk(f"beta_m {t}", g(f"bm_{t}"), py["beta_m"][t], 1e-9)
    for nm in ["ebitda", "rev", "pe"]:
        for lvl in ["lo", "mid", "hi"]:
            chk(f"cmp_{nm}_{lvl}", g(f"cmp_{nm}_{lvl}"), py[f"cmp_{nm}_{lvl}"], 0.005)
    sh = "SENSITIVITIES"
    for gi, (top, grid) in enumerate([(8, py["s1"]), (17, py["s2"])]):
        for i in range(5):
            for j, c in enumerate("CDEFG"):
                chk(f"sens{gi+1} {c}{top+i}", vals.get((sh, f"{c}{top+i}")), grid[i][j], 0.005)
    bad = [r for r in rows if not r[4]]
    print(f"[{label}] compared {len(rows)} outputs: {len(rows)-len(bad)} match, {len(bad)} mismatch")
    for r in bad:
        print("   MISMATCH", r)
    return rows, bad


def scenario_copy(scen):
    src = recalc.XLSX
    dst = src.with_name(f"PLNT_DCF_Model_scen{scen}.xlsx")
    shutil.copy(src, dst)
    wb = load_workbook(dst)
    wb["Assumptions"]["C4"].value = scen
    wb.save(dst)
    return dst


if __name__ == "__main__":
    reg = recalc.names()
    results = {}
    vals = recalc.load(refresh="--refresh" in sys.argv)
    py = model(1)
    rows, bad = compare(vals, reg, py, "Base")
    results["Base"] = (py, recalc.get(vals, reg, "chk_all"), len(rows), len(bad))
    if "--all" in sys.argv:
        for scen, lab in [(2, "Upside"), (3, "Downside")]:
            path = scenario_copy(scen)
            v = recalc.calculate(path)
            p = model(scen)
            rows, bad = compare(v, reg, p, lab)
            results[lab] = (p, v.get(("COVER", recalc.names()["cover_status"][1] + str(recalc.names()["cover_status"][2]))), len(rows), len(bad))
            path.unlink()
    summary = {k: dict(vps=v[0]["vps"], vps_pgm=v[0]["vps_pgm"], vps_exit=v[0]["vps_exit"], ev_sel=v[0]["ev_sel"],
                       wacc=v[0]["wacc"], status=v[1], compared=v[2], mismatches=v[3],
                       rev_cagr=v[0]["rv_g0"], margin=v[0]["rv_m0"]) for k, v in results.items()}
    json.dump(summary, open(ROOT / "output" / "verification_summary.json", "w"), indent=2, default=float)
    print(json.dumps(summary, indent=2, default=float))
