"""Single source of truth for every hard-coded input in the model.

All dollar figures are in $ millions (filings report $ thousands; values are
the filing figures divided by 1,000, so no rounding is introduced). Each entry
carries its source. LTM = FY2025 + 9M 2026 - 9M 2025 (computed in the workbook
from the three inputs, not here).

Filing references
  10-K FY2021  0001637207-22-000011   10-K FY2022  0001637207-23-000005
  10-K FY2023  0001637207-24-000020   10-K FY2024  0001637207-25-000016
  10-K FY2025  0001637207-26-000011   10-Q Q3 2025 0001637207-25-000077
  10-Q Q3 2026 0001637207-26-000051   8-K 2026-11-05 Ex. 99.1 0001637207-26-000047
"""

YEARS = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"]
K21, K22, K23, K24, K25 = "10-K FY2021", "10-K FY2022", "10-K FY2023", "10-K FY2024", "10-K FY2025"
Q25, Q26 = "10-Q Q3 2025", "10-Q Q3 2026"

# key: (label, [FY2021..FY2025], 9M 2025, 9M 2026, source)
# None for 9M values means "not needed for LTM".
HIST = {
    # ---- Income statement -------------------------------------------------
    "rev_franchise": ("Franchise revenue (excl. NAF)", [238.349, 271.559, 317.917, 344.320, 380.971], 282.362, 302.934,
                      "Consolidated Statements of Operations, 'Franchise' line (FY2021 incl. $0.779M commission income, as recast in 10-K FY2023); 9M: 10-Q MD&A franchise revenue table"),
    "rev_royalty": ("  of which: royalty revenue", [205.9, 228.7, 260.7, 286.269, 314.647], 236.780, 251.449,
                    "MD&A franchise revenue discussion: FY2021-22 10-K FY2022 ($M, 1 decimal as disclosed); FY2023 10-K FY2024; FY2024-25 10-K FY2025 table; 9M 10-Q Q3 2026 table"),
    "rev_naf": ("National advertising fund (NAF) revenue", [52.361, 58.075, 70.012, 78.927, 86.987], 66.151, 96.956,
                "Statements of Operations; 9M = Franchise segment revenue less franchise revenue (10-Q Q3 2026 MD&A: $399.890M - $302.934M; 9M 2025 $348.513M - $282.362M)"),
    "rev_corp": ("Corporate-owned clubs revenue", [167.219, 379.393, 449.296, 502.287, 546.097], 410.491, 427.197,
                 "Statements of Operations / segment note"),
    "rev_equip": ("Equipment revenue", [129.094, 227.745, 234.101, 256.120, 310.089], 188.882, 235.349,
                  "Statements of Operations / segment note"),
    "rev_total_xbrl": ("Total revenue (as reported)", [587.023, 936.772, 1071.326, 1181.654, 1324.144], 947.886, 1062.436,
                       "XBRL us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax"),
    "cogs": ("Cost of revenue", [100.993, 177.200, 190.026, 197.122, 230.308], 140.063, 175.206, "XBRL us-gaap:CostOfRevenue"),
    "club_ops": ("Club operations", [110.716, 219.422, 253.619, 290.507, 318.545], 238.909, 254.231,
                 "XBRL us-gaap:OperatingCostsAndExpenses (FY2021 'Store operations' from 10-K FY2023 Statements of Operations)"),
    "sga": ("Selling, general & administrative", [94.540, 114.853, 124.930, 129.146, 137.634], 100.343, 103.495,
            "XBRL us-gaap:SellingGeneralAndAdministrativeExpense"),
    "naf_exp": ("NAF expense", [59.442, 66.116, 70.095, 79.009, 87.580], 66.150, 96.956, "Statements of Operations (custom tag; read from filing)"),
    "da": ("Depreciation & amortization", [62.800, 124.022, 149.413, 160.346, 155.785], 115.818, 121.965,
           "XBRL us-gaap:DepreciationAndAmortization"),
    "amort": ("  of which: amortization of intangibles", [16.677, 40.294, 51.500, 49.200, 36.800], None, None,
              "XBRL us-gaap:AmortizationOfIntangibleAssets"),
    "other_op": ("Other (gains) losses, net", [15.137, 5.081, 10.379, 1.326, -0.385], -2.069, -13.383,
                 "XBRL us-gaap:OtherOperatingIncomeExpenseNet (sign flipped to expense); 9M 2026 incl. $12.5M gain on sale of Bravo Fit equity stake"),
    "ebit_xbrl": ("Income from operations (as reported)", [143.395, 230.078, 272.864, 324.198, 394.677], 288.672, 323.966,
                  "XBRL us-gaap:OperatingIncomeLoss"),
    "int_exp": ("Interest expense", [81.211, 88.628, 86.576, 100.037, 108.244], 78.720, 100.508,
                "XBRL us-gaap:InterestExpenseNonoperating"),
    "int_inc": ("Interest income", [0.878, 5.005, 17.741, 23.115, 22.999], 17.438, 16.103, "XBRL us-gaap:InvestmentIncomeInterest"),
    "tax": ("Provision for income taxes", [5.659, 50.515, 58.512, 68.443, 85.874], 69.120, 67.803, "XBRL us-gaap:IncomeTaxExpenseBenefit"),
    "net_income": ("Net income (incl. NCI)", [46.122, 110.456, 147.035, 174.243, 220.264], 159.557, 171.447, "XBRL us-gaap:ProfitLoss"),
    "ni_parent": ("Net income attributable to PLNT", [42.774, 99.402, 138.313, 172.042, 219.104], 158.715, 170.637, "XBRL us-gaap:NetIncomeLoss"),
    "sbc": ("Stock-based compensation", [8.805, 8.068, 7.906, 8.913, 12.333], 9.132, 9.510, "XBRL us-gaap:ShareBasedCompensation (cash flow statement)"),
    "adj_ebitda": ("Adjusted EBITDA (company-reported)", [222.310, 365.834, 435.376, 487.710, 551.644], 405.380, 438.823,
                   "Non-GAAP reconciliation: FY2021 per 10-K FY2022 (current definition); FY2022 10-K FY2023; FY2023 10-K FY2024; FY2024-25 10-K FY2025; 9M 10-Q Q3 2026"),
    "lease_cost": ("Operating lease cost (in club ops / SG&A)", [None, None, None, None, 78.003], 58.027, 62.436,
                   "XBRL us-gaap:OperatingLeaseCost; 10-Q Q3 2026 Note 6"),
    # ---- Segment Adjusted EBITDA (FY2023+ definition) ------------------------
    "seg_fr": ("Franchise segment Adj. EBITDA", [None, None, 273.008, 301.122, 336.592], 253.734, 273.276,
               "Segment note: FY2023 per 10-K FY2024 (recast), FY2024-25 10-K FY2025; 9M 10-Q Q3 2026"),
    "seg_corp": ("Corporate-owned clubs segment Adj. EBITDA", [None, None, 173.322, 188.751, 206.347], 156.184, 157.483, "Segment note"),
    "seg_equip": ("Equipment segment Adj. EBITDA", [None, None, 56.362, 71.778, 94.478], 57.601, 69.340, "Segment note"),
    # ---- Cash flow ------------------------------------------------------
    "cfo": ("Cash flow from operations", [189.289, 240.207, 330.254, 343.873, 418.421], 309.361, 303.889,
            "XBRL us-gaap:NetCashProvidedByUsedInOperatingActivities"),
    "capex": ("Capital expenditures", [54.074, 100.057, 135.986, 155.061, 163.670], 113.577, 148.943,
              "XBRL us-gaap:PaymentsToAcquirePropertyPlantAndEquipment"),
    "acq": ("Acquisitions, net of cash acquired", [1.888, 424.940, 43.264, 0.0, 0.0], None, None,
            "XBRL us-gaap:PaymentsToAcquireBusinessesNetOfCashAcquired (FY2022 = Sunshine Fitness, 114 clubs)"),
    "buyback": ("Share repurchases", [0.0, 94.315, 125.030, 300.205, 500.373], 150.026, 302.431,
                "XBRL us-gaap:PaymentsForRepurchaseOfCommonStock; 9M from 10-Q cash flow statement"),
}

# Balance sheet: key: (label, [FY2020..FY2025], 30-Sep-2026, source)
BS = {
    "cash": ("Cash and cash equivalents", [439.478, 545.909, 409.840, 275.842, 293.150, 345.652], 310.185, "XBRL us-gaap:CashAndCashEquivalentsAtCarryingValue"),
    "restricted_cash": ("Restricted cash (securitization reserves)", [None, None, None, 46.279, 56.524, 66.304], 72.815, "XBRL us-gaap:RestrictedCashCurrent"),
    "mkt_sec": ("Marketable securities (ST + LT)", [0.0, 0.0, 0.0, 125.787, 179.831, 195.024], 132.044,
                "XBRL us-gaap:MarketableSecuritiesCurrent + MarketableSecuritiesNoncurrent; 9/30/26 = $90.290M + $41.754M (10-Q balance sheet)"),
    "htm_inv": ("Held-to-maturity investment, net", [None, None, None, None, None, 69.700], 56.000,
                "Balance sheet 'Investments, net of allowance' (10-K FY2025; 10-Q Q3 2026)"),
    "ar": ("Accounts receivable, net", [16.447, 27.257, 46.242, 41.890, 77.145, 70.431], 73.581, "XBRL us-gaap:AccountsReceivableNetCurrent"),
    "inventory": ("Inventory", [0.473, 1.155, 5.266, 4.677, 6.146, 7.581], 8.104, "XBRL us-gaap:InventoryNet"),
    "prepaid": ("Prepaid expenses", [11.881, 12.869, 11.078, 13.842, 21.499, 24.605], 22.968, "XBRL us-gaap:PrepaidExpenseCurrent"),
    "other_rec": ("Other receivables", [16.754, 13.519, 14.975, 11.072, 16.776, 34.094], 24.062, "XBRL us-gaap:OtherReceivablesNetCurrent"),
    "ap": ("Accounts payable", [19.388, 27.892, 20.578, 23.788, 32.887, 39.683], 36.581, "XBRL us-gaap:AccountsPayableCurrent"),
    "accrued": ("Accrued expenses", [22.042, 51.714, 66.993, 66.299, 67.895, 75.371], 77.332, "XBRL us-gaap:AccruedLiabilitiesCurrent"),
    "equip_dep": ("Equipment deposits", [0.795, 6.036, 8.443, 4.506, 1.851, 10.165], 23.021, "Balance sheet line (custom tag; read from 10-K / 10-Q)"),
    "def_rev": ("Deferred revenue, current", [26.691, 28.351, 53.759, 59.591, 62.111, 58.593], 68.745, "XBRL us-gaap:ContractWithCustomerLiabilityCurrent"),
    "debt_principal": ("Securitized notes, principal", [1792.501, 1775.000, 2025.187, 2004.438, 2195.750, 2514.188], 2496.750,
                       "XBRL us-gaap:DebtInstrumentCarryingAmount (principal before deferred financing costs); 9/30/26 10-Q Note 5"),
    "vfn": ("Variable Funding Notes drawn", [0, 0, 0, 0, 0, 0], 75.000, "10-Q Q3 2026 balance sheet / Note 5 (drawn Q2 2026)"),
    "fin_lease": ("Finance lease liabilities", [None, None, None, None, None, 0.976], 0.827, "10-Q Q3 2026 Note 6 ($0.200M current + $0.627M noncurrent)"),
    "op_lease": ("Operating lease liabilities", [187.454, 220.205, 375.076, 415.438, 442.355, 463.517], 555.746,
                 "XBRL us-gaap:OperatingLeaseLiabilityCurrent + OperatingLeaseLiabilityNoncurrent; 9/30/26 10-Q Note 6"),
    "tra": ("Payable pursuant to tax benefit arrangements (TRA)", [488.200, 528.107, 494.465, 495.662, 466.916, 415.791], 361.926,
            "Balance sheet, current + noncurrent (10-K FY2021-FY2025; 10-Q Q3 2026: $38.442M + $323.484M)"),
    "dta": ("Deferred income tax assets, net", [511.200, 539.264, 454.565, 504.188, 470.197, 406.724], 368.161, "XBRL us-gaap:DeferredIncomeTaxAssetsNet"),
    "dtl": ("Deferred tax liabilities", [None, None, None, None, None, 1.177], 0.451, "Balance sheet"),
    "nci": ("Non-controlling interests", [0.196, 2.510, -12.549, -3.342, 0.007, -0.601], -1.796, "XBRL us-gaap:MinorityInterest"),
    "assets": ("Total assets", [1849.737, 2015.983, 2854.589, 2969.693, 3069.708, 3103.395], 3097.862, "XBRL us-gaap:Assets"),
    "liab_cur": ("Total current liabilities", [111.895, 176.610, 244.530, 251.329, 282.495, 312.490], 395.439, "XBRL us-gaap:LiabilitiesCurrent"),
    "liab_noncur": ("Total noncurrent liabilities", [None, 2482.218, 2821.620, 2837.355, 3002.586, 3274.283], 3313.114, "XBRL us-gaap:LiabilitiesNoncurrent"),
    "equity_total": ("Total stockholders' equity (deficit) incl. NCI", [None, -642.845, -211.561, -118.991, -215.373, -483.378], -610.691,
                     "XBRL us-gaap:StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"),
}

# Operating KPIs: key: (label, [FY2021..FY2025], 9M-2026 / 30-Sep-2026 value, source)
KPI = {
    "fr_clubs": ("Franchisee-owned clubs (end of period)", [2142, 2176, 2319, 2445, 2604], 2660, "MD&A club roll-forward; 10-Q Q3 2026"),
    "corp_clubs": ("Corporate-owned clubs (end of period)", [112, 234, 256, 277, 292], 302, "MD&A club roll-forward; FY2022 incl. 114 Sunshine clubs acquired"),
    "fr_open": ("New franchisee-owned clubs opened", [125, 144, 147, 129, 158], 61, "MD&A club roll-forward ('opened or acquired' in FY2025)"),
    "corp_open": ("New corporate-owned clubs opened", [7, 14, 18, 21, 23], 10, "MD&A club roll-forward"),
    "fr_closed": ("Franchisee clubs debranded, sold, consolidated or transferred (net)", [-4, -110, -4, -3, 1], -5,
                  "Roll-forward: FY2022 = 6 acquired from Company - 116 debranded/sold (incl. 114 to Sunshine); FY2023 = 5 - 9; FY2025 = +8 refranchised - 7; 9M 2026 = derived (2,604 + 61 - 2,660)"),
    "corp_other": ("Corporate clubs acquired / (sold), net", [2, 108, 4, 0, -8], 0, "Roll-forward: FY2022 = 114 acquired - 6 sold; FY2023 = 9 - 5; FY2025 = 8 refranchised"),
    "members": ("Members (millions, end of period)", [15.2, 17.0, 18.7, 19.7, 20.8], None, "10-K Item 1 Business"),
    "sws": ("System-wide sales ($bn, rounded in filing)", [3.4, 3.9, 4.5, 4.8, 5.3], None, "10-K Item 1 Business"),
    "scs": ("System-wide same club sales growth", [None, None, None, 0.050, 0.067], 0.023, "MD&A; 9M 2026 from 10-Q Q3 2026"),
}

# Q4 2025 clubs used for opening FY2025 balance (already in KPI)

# ---- Capital structure / equity bridge inputs (as of 30-Sep-2026 unless noted)
SHARES = {
    "class_a": (74.252489, "Class A shares outstanding at Oct 30, 2026 (millions) - 10-Q Q3 2026 cover page"),
    "class_b": (0.266128, "Class B shares at Oct 30, 2026 (millions; exchangeable 1:1 with Holdings Units into Class A) - 10-Q Q3 2026 cover"),
    "rsu": (0.157761, "Unvested RSUs at Dec 31, 2025 (millions) - 10-K FY2025 Note 14 (latest full disclosure)"),
    "psu": (0.138961, "Unvested PSUs at Dec 31, 2025, at target (millions) - 10-K FY2025 Note 14"),
    "options": (0.073363, "Stock options outstanding at Dec 31, 2025 (millions) - XBRL ShareBasedCompensation...OptionsOutstandingNumber"),
    "option_strike": (49.62, "Weighted-average exercise price of options ($) - 10-K FY2025"),
}

BEFIT = {
    "cash_paid": (112.0, "Be-Fit (Portugal, 25 clubs) cash consideration paid at closing Oct 30, 2026 ($M) - 10-Q Q3 2026 Note 16 subsequent events"),
    "contingent": (2.8, "Be-Fit contingent consideration, up to EUR 2.5M (~$2.8M) - 10-Q Q3 2026 Note 16"),
}

DEBT_TERMS = {  # outstanding principal at 9/30/26 ($M), coupon, anticipated repayment date
    "2019-1 Class A-2": (512.875, 0.03858, "Dec 2029"),
    "2022-1 Class A-2-II": (453.625, 0.04008, "Dec 2031"),
    "2024-1 Class A-2-I": (416.500, 0.05765, "Jun 2029"),
    "2024-1 Class A-2-II": (367.500, 0.06237, "Jun 2034"),
    "2025-1 Class A-2-I": (398.000, 0.05274, "Dec 2030"),
    "2025-1 Class A-2-II": (348.250, 0.05649, "Dec 2032"),
}
DEBT_SRC = ("10-Q Q3 2026 Note 5 (principal); coupons and anticipated repayment dates from 10-K FY2025 Note 10. "
            "Fair value of notes at 9/30/26: $2,389.338M vs. $2,496.750M principal (95.7% of par).")
DEBT_FAIR_VALUE = 2389.338

# FY2025 10-K scheduled amortization of existing intangibles ($M) - XBRL FiniteLivedIntangibleAssetsAmortizationExpense*
AMORT_SCHEDULE = {"FY2026": 32.079, "FY2027": 27.956, "FY2028": 27.300, "FY2029": 23.675, "FY2030": 17.920, "FY2031": 11.075}
AMORT_SRC = ("10-K FY2025 intangible assets note: expected amortization 2026-2030 and 'thereafter' ($11.075M, "
             "conservatively placed entirely in FY2031)")

GUIDANCE_2026 = {  # 8-K 2026-11-05 Ex. 99.1 "2026 Outlook"
    "revenue_growth": 0.08, "adj_ebitda_growth": 0.065, "capex_growth": 0.15, "da_growth": 0.07,
    "scs_low": 0.015, "scs_high": 0.02, "openings_low": 190, "openings_high": 195, "befit_clubs": 25,
}

MARKET = {
    "price": (61.00, "PLNT closing price on 2026-11-19 (Nasdaq.com historical quotes)"),
    "rf": (0.0504, "10-year US Treasury par yield on 2026-11-19 (treasury.gov daily par yield curve)"),
    "ust7_now": (0.0494, "7-year US Treasury par yield on 2026-11-19 (treasury.gov)"),
    "ust7_issue": (0.0394, "7-year US Treasury par yield on 2025-12-15, the 2025-1 notes' closing date (treasury.gov)"),
    "ust5_issue": (0.0373, "5-year US Treasury par yield on 2025-12-15 (treasury.gov)"),
    "erp": (0.0441, "Damodaran implied ERP, 'T12m with sustainable payout', as of 2026-11-01 (ERPbymonth.xlsx, pages.stern.nyu.edu/~adamodar)"),
    "erp_alt": (0.0379, "Damodaran implied ERP, trailing-12-month cash yield, as of 2026-11-01 (same file) - shown as alternative"),
    "tax": (0.26, "PLNT's assumed corporate tax rate for adjusted net income, FY2025 (10-K FY2025 non-GAAP reconciliation, note 11)"),
    "wk52_high": (None, "52-week high close (Nasdaq.com), computed in build"),
    "wk52_low": (None, "52-week low close (Nasdaq.com), computed in build"),
}

VALUATION_DATE = "2026-11-20"
BALANCE_SHEET_DATE = "2026-09-30"
