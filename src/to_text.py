"""Convert downloaded PLNT filings (HTML) to plain text for searching / source checks."""
import html
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
out = ROOT / "data" / "text"
out.mkdir(parents=True, exist_ok=True)
for p in sorted((ROOT / "data" / "raw").glob("PLNT_*.htm")):
    t = p.read_text(encoding="utf-8", errors="ignore")
    t = re.sub(r"(?is)<(script|style|ix:header).*?</\1>", "", t)
    t = re.sub(r"(?i)</(p|div|tr|li|h\d)>", "\n", t)
    t = re.sub(r"(?i)</t[dh]>", " | ", t)
    t = re.sub(r"<[^>]+>", "", t)
    t = html.unescape(t).replace("\xa0", " ")
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n\s*\n+", "\n", t)
    name = p.stem.replace("_ex991", "_ex991")
    (out / f"{name}.txt").write_text(t)
    print("wrote", out / f"{name}.txt")
