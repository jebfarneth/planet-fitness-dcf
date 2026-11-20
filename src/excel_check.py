"""Compare a copy of the workbook recalculated by Microsoft Excel (data_only
values) with the `formulas`-engine recalculation, cell by cell."""
import datetime
import json
import sys
from pathlib import Path

from openpyxl import load_workbook

import recalc

ROOT = Path(__file__).resolve().parents[1]
excel_copy = sys.argv[1]
wbv = load_workbook(excel_copy, data_only=True)
wbf = load_workbook(recalc.XLSX)
vals = recalc.load()
n = empty = 0
bad, maxd, where = [], 0.0, None
for ws in wbf.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                n += 1
                ev = wbv[ws.title][c.coordinate].value
                fv = vals.get((ws.title.upper(), c.coordinate))
                if ev is None:
                    empty += 1
                    continue
                if isinstance(ev, datetime.datetime):   # Excel date vs serial number
                    ev = (ev - datetime.datetime(1899, 12, 30)).days
                if isinstance(ev, (int, float)) and isinstance(fv, (int, float)):
                    d = abs(ev - fv)
                    if d > maxd:
                        maxd, where = d, f"{ws.title}!{c.coordinate}"
                    if d > 1e-6 * max(1, abs(ev)):
                        bad.append((ws.title, c.coordinate, ev, fv))
                elif str(ev) != str(fv):
                    bad.append((ws.title, c.coordinate, ev, fv))
res = {"formula_cells": n, "excel_blank": empty, "mismatches": len(bad), "max_abs_diff": maxd, "at": where}
print(json.dumps(res), bad[:10])
json.dump(res, open(ROOT / "output" / "excel_check.json", "w"), indent=2)
