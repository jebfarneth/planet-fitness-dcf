"""Helpers to pull values out of SEC companyfacts JSON with a citation."""
import json
from datetime import date
from pathlib import Path

RAW = Path(__file__).resolve().parents[1] / "data" / "raw"


def load(ticker):
    return json.load(open(RAW / f"companyfacts_{ticker}.json"))


def _days(f):
    if "start" not in f:
        return 0
    s, e = date.fromisoformat(f["start"]), date.fromisoformat(f["end"])
    return (e - s).days


def facts(cf, tag, unit="USD", ns="us-gaap"):
    try:
        return cf["facts"][ns][tag]["units"][unit]
    except KeyError:
        return []


def value(cf, tags, end, kind="duration", months=12, unit="USD", prefer_form=None, ns="us-gaap"):
    """Return (val, tag, form, accn, filed) for the first tag with a matching fact.

    kind='duration' needs a period of ~`months` ending on `end`; kind='instant'
    needs a point-in-time fact on `end`. Among duplicates (same fact repeated in
    later filings) the most recent filing wins, i.e. the latest as-reported or
    recast figure.
    """
    if isinstance(tags, str):
        tags = [tags]
    for tag in tags:
        cands = []
        for f in facts(cf, tag, unit, ns):
            if f["end"] != end:
                continue
            if kind == "instant" and "start" in f:
                continue
            if kind == "duration":
                d = _days(f)
                if not (months * 30.4 - 20 <= d <= months * 30.4 + 20):
                    continue
            if prefer_form and f.get("form") != prefer_form:
                continue
            cands.append(f)
        if cands:
            f = max(cands, key=lambda x: x["filed"])
            return f["val"], tag, f.get("form"), f.get("accn"), f.get("filed")
    return None
