"""Forecast drivers by scenario (FY2026E-FY2031E). These are judgement inputs;
the rationale and the historical anchor for each is written next to it in the
workbook.

Base case: FY2026E is calibrated to management's 2026 outlook (8-K 2026-11-05);
FY2027E revenue growth is held near that guidance (~7.5%) and then fades to
~4% by FY2031E as net club openings slow and same-club sales settle at ~2.5%.
Upside and Downside are symmetric: base +/- the same delta on every driver
(cost drivers move in the opposite direction), so the base is the midpoint.
"""

PROJ_YEARS = ["FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]

# Target base-case revenue growth (excl. NAF) the drivers were set to deliver
# (approximate; the exact path is computed in the workbook, Projections tab)
BASE_GROWTH_PATH = {"FY2027E": 0.076, "FY2028E": 0.067, "FY2029E": 0.054, "FY2030E": 0.046, "FY2031E": 0.042}

# key: (label, fmt, base, delta, sign, rationale)
#   upside = base + sign * delta ; downside = base - sign * delta
#   sign = +1 for revenue / margin drivers, -1 for cost drivers (overhead, capex)
_SPEC = [
    ("fr_adds", "Franchise clubs - net additions", "n",
     [142, 145, 125, 110, 95, 80],
     [5, 20, 20, 20, 20, 20], +1,
     "Base fades from 145 in FY2027E (FY2021-25 franchise openings 125-158/yr; 2026 outlook 190-195 system openings incl. 25 Be-Fit) to 80 by FY2031E as the "
     "~900 clubs franchisees had committed to open at end-2024 (10-K FY2024) is worked down and U.S. white space narrows: average franchise club growth slows from ~5.4% to ~2.8% a year."),
    ("corp_adds", "Corporate-owned clubs - net additions", "n",
     [20, 20, 16, 12, 10, 8],
     [2, 5, 5, 5, 5, 5], +1,
     "FY2023-25 corporate openings 18-23/yr; 9M 2026 +10. Fades with the announced strategic review of domestic corporate-owned clubs."),
    ("fr_scs", "Franchise same-club sales growth", "pct",
     [0.0175, 0.030, 0.0275, 0.025, 0.025, 0.025],
     [0.0025, 0.015, 0.015, 0.015, 0.015, 0.015], +1,
     "FY2024 5.2%, FY2025 6.8%, 9M 2026 2.3%, Q3 2026 1.5%; 2026 outlook 1.5-2.0%. Base: 3.0% in FY2027 (Black Card mix + dues pricing) settling to 2.5% (~inflation) from FY2029."),
    ("roy_adj", "Royalty per avg. franchise club vs. SCS (new-club ramp / rate mix)", "pct",
     [-0.020, -0.015, -0.0125, -0.010, -0.0075, -0.005],
     [0.005, 0.005, 0.005, 0.005, 0.005, 0.005], +1,
     "Royalty per average franchise club grew SCS -1.6pp (FY2024) and -3.1pp (FY2025): new clubs bill below mature clubs. Drag shrinks as openings slow; "
     "royalty-rate step-up to 7% on renewals is a partial offset."),
    ("oth_fr_g", "Other franchise revenue per avg. franchise club - growth", "pct",
     [0.060, 0.030, 0.025, 0.025, 0.020, 0.020],
     [0.010, 0.015, 0.015, 0.015, 0.015, 0.015], +1,
     "Franchise & other fees + equipment placement + HVAC. 9M 2026 +13% in total (ADA fees, PF Perks, replacement placements); fades to ~2% as new ADAs slow."),
    ("naf_ratio", "NAF revenue / royalty revenue", "pct",
     [0.385] * 6, [0.0] * 6, +1,
     "NAF contribution raised from 2% to 3% of dues in 2026 (10-Q Q3 2026). 9M 2026 NAF $97.0M / royalty $251.4M = 38.6%. Pass-through: EBITDA-neutral."),
    ("corp_scs", "Corporate-owned club same-club sales growth", "pct",
     [0.020, 0.030, 0.0275, 0.025, 0.025, 0.025],
     [0.0025, 0.015, 0.015, 0.015, 0.015, 0.015], +1,
     "FY2025 6.0%; 9M 2026 2.3%. Mirrors the franchise path."),
    ("corp_adj", "Revenue per avg. corporate club vs. SCS (ramp / mix)", "pct",
     [-0.040, -0.015, -0.0125, -0.010, -0.0075, -0.005],
     [0.005, 0.005, 0.005, 0.005, 0.005, 0.005], +1,
     "FY2025 revenue per avg. club +1.8% vs. SCS +6.0% (new Spain clubs ramping, 8 CA clubs sold Aug-2025). FY2026E calibrated to 9M 2026 segment revenue +4.1%; drag fades with fewer openings."),
    ("eq_g", "Equipment revenue per avg. franchise club - growth", "pct",
     [-0.010, 0.015, 0.010, 0.005, 0.000, 0.000],
     [0.020, 0.020, 0.020, 0.020, 0.020, 0.020], +1,
     "Re-equip cycle 5-9 yrs scales with the installed base; new-club equipment falls as openings slow. 2026 outlook cut new placements to 135-140 (from 150-160)."),
    ("fr_margin", "Franchise Adj. EBITDA margin (excl. NAF)", "pct",
     [0.895, 0.890, 0.890, 0.890, 0.890, 0.890],
     [0.005, 0.010, 0.010, 0.010, 0.010, 0.010], +1,
     "FY2023 85.9%, FY2024 87.5%, FY2025 88.5%, 9M 2026 90.2%, LTM 88.8% (segment Adj. EBITDA less NAF surplus, / franchise revenue)."),
    ("corp_margin", "Corporate-owned clubs Adj. EBITDA margin", "pct",
     [0.370, 0.370, 0.375, 0.375, 0.375, 0.375],
     [0.005, 0.010, 0.010, 0.010, 0.010, 0.010], +1,
     "FY2023 38.6%, FY2024 37.6%, FY2025 37.8%, 9M 2026 36.9% (new-club and Spain drag)."),
    ("eq_margin", "Equipment Adj. EBITDA margin", "pct",
     [0.295, 0.290, 0.290, 0.290, 0.290, 0.290],
     [0.005, 0.010, 0.010, 0.010, 0.010, 0.010], +1,
     "FY2023 24.1%, FY2024 28.0%, FY2025 30.5%, 9M 2026 29.5%."),
    ("overhead", "Unallocated corporate costs (% of revenue excl. NAF)", "pct",
     [0.064, 0.065, 0.065, 0.065, 0.065, 0.065],
     [0.001, 0.003, 0.003, 0.003, 0.003, 0.003], -1,
     "FY2023 6.7%, FY2024 6.7%, FY2025 6.9%, 9M 2026 6.3%, LTM 6.5%. Net of SBC: SBC is NOT added back (treated as a real cost)."),
    ("dep_pct", "Depreciation (% of revenue excl. NAF)", "pct",
     [0.1035, 0.100, 0.100, 0.100, 0.100, 0.100],
     [0.0] * 6, +1,
     "FY2023 9.8%, FY2024 10.1%, FY2025 9.6%. FY2026E set so total D&A rises ~7% (2026 outlook)."),
    ("capex_pct", "Capital expenditures (% of revenue excl. NAF)", "pct",
     [0.145, 0.135, 0.130, 0.125, 0.120, 0.115],
     [0.0, 0.005, 0.005, 0.005, 0.005, 0.005], -1,
     "FY2023 13.6%, FY2024 14.1%, FY2025 13.2%. 2026 outlook capex +15% (~14.5%). Fades as corporate openings slow; FY2031E 11.5% still exceeds D&A (~10.6%), funding terminal growth."),
    ("nwc_pct", "Net working capital (% of revenue excl. NAF)", "pct",
     [-0.040] * 6, [0.0] * 6, +1,
     "Operating NWC (AR + inventory + prepaid + other receivables - AP - accrued - equipment deposits - current deferred revenue): FY2022 -8.2%, FY2023 -8.3%, "
     "FY2024 -3.9%, FY2025 -3.8% of revenue excl. NAF; base uses the recent ~-4%. See Historicals."),
]


def _round(x):
    return round(x, 6) if isinstance(x, float) else x


DRIVERS = []
DELTAS = {}
for key, lab, fmt, base, delta, sign, why in _SPEC:
    up = [_round(b + sign * d) for b, d in zip(base, delta)]
    down = [_round(b - sign * d) for b, d in zip(base, delta)]
    DRIVERS.append((key, lab, fmt, base, up, down, why))
    DELTAS[key] = (delta, sign)
