"""Value-per-share swing for one assumption at a time (base case), using the
independent Python model; WACC / g / multiple swings come from the sensitivity
grids. Writes output/tornado.json."""
import copy
import json
from pathlib import Path

import assumptions as A
import verify

ROOT = Path(__file__).resolve().parents[1]


def run():
    m = verify.model(1)
    base, s1, s2 = m["vps"], m["s1"], m["s2"]
    orig = copy.deepcopy(A.DRIVERS)
    res = {}

    def bump(key, delta, label):
        out = []
        for sgn in (-1, 1):
            verify.DRIVERS = [(k, l, f, [v + (sgn * delta if (k == key and i >= 1) else 0) for i, v in enumerate(b)], u, d, w)
                              for k, l, f, b, u, d, w in orig]
            out.append(verify.model(1)["vps"] - base)
        verify.DRIVERS = orig
        res[label] = out

    bump("fr_scs", 0.01, "Franchise same-club sales +/-1pp (FY27-31)")
    bump("fr_adds", 20, "Franchise net adds +/-20 per year (FY27-31)")
    bump("capex_pct", 0.01, "Capex +/-1pp of revenue (FY27-31)")
    bump("corp_margin", 0.02, "Corporate club margin +/-2pp (FY27-31)")
    bump("fr_margin", 0.02, "Franchise margin +/-2pp (FY27-31)")
    # grids: selected value = average of the two methods; isolate each lever on its own method
    res["WACC +/-0.5% (perpetuity method)"] = [s1[1][2] - s1[2][2], s1[3][2] - s1[2][2]]
    res["Terminal growth +/-0.5% (perpetuity method)"] = [s1[2][1] - s1[2][2], s1[2][3] - s1[2][2]]
    res["Exit multiple +/-1.0x (exit method)"] = [s2[2][1] - s2[2][2], s2[2][3] - s2[2][2]]
    res["WACC +/-0.5% (exit method)"] = [s2[1][2] - s2[2][2], s2[3][2] - s2[2][2]]
    ranked = sorted(res.items(), key=lambda kv: -abs(kv[1][1] - kv[1][0]))
    json.dump({"base_vps": base, "tornado": [[k, v[0], v[1]] for k, v in ranked]}, open(ROOT / "output" / "tornado.json", "w"), indent=1)
    return base, ranked


if __name__ == "__main__":
    base, ranked = run()
    print("base", round(base, 2))
    for k, v in ranked:
        print(f"{k:48s} {v[0]:+7.2f} {v[1]:+7.2f}")
