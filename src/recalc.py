"""Recalculate the workbook with the open-source `formulas` engine (no Excel
needed) and expose named outputs. Used by verify.py and the README/memo build.
"""
import pickle
import re
import sys
from pathlib import Path

import formulas
import numpy as np

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / "output" / "PLNT_DCF_Model.xlsx"
CACHE = ROOT / "output" / ".recalc_cache.pkl"


def calculate(path=XLSX):
    xl = formulas.ExcelModel().loads(str(path)).finish()
    sol = xl.calculate()
    out = {}
    for k, v in sol.items():
        m = re.match(r"'\[(.+?)\](.+?)'!([A-Z]+\d+)(?::([A-Z]+\d+))?$", str(k))
        if not m or m.group(4):
            continue
        val = v.value
        if isinstance(val, np.ndarray):
            val = val.ravel()[0]
        out[(m.group(2).upper(), m.group(3))] = val
    return out


def load(path=XLSX, refresh=False):
    if not refresh and CACHE.exists() and CACHE.stat().st_mtime > path.stat().st_mtime:
        return pickle.load(open(CACHE, "rb"))
    out = calculate(path)
    pickle.dump(out, open(CACHE, "wb"))
    return out


def names():
    """Name -> (sheet, col, row) registry saved by build_model.py."""
    return pickle.load(open(ROOT / "output" / ".names.pkl", "rb"))


def get(vals, reg, name, col=None):
    sheet, c, r = reg[name]
    return vals.get((sheet.upper(), f"{col or c}{r}"))


if __name__ == "__main__":
    vals = load(refresh="--refresh" in sys.argv)
    reg = names()
    for n in sys.argv[1:]:
        if n.startswith("--"):
            continue
        if "@" in n:
            nm, c = n.split("@")
            print(n, get(vals, reg, nm, c))
        else:
            print(n, get(vals, reg, n))
