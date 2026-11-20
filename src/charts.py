"""Football-field chart (light and dark PNG) from the recalculated workbook."""
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import FancyBboxPatch  # noqa: E402

import recalc  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
THEMES = {
    "light": dict(surface="#fcfcfb", ink="#0b0b0b", ink2="#52514e", grid="#e1e0d9", bar="#2a78d6"),
    "dark": dict(surface="#1a1a19", ink="#f0efec", ink2="#c3c2b7", grid="#2c2c2a", bar="#3987e5"),
}


def football_rows():
    vals, reg = recalc.load(), recalc.names()
    sheet, c, r0 = reg["ff_first"]
    rows = []
    for i in range(6):
        r = r0 + i
        lab = {0: "52-week trading range", 1: "Comps: EV / LTM EBITDA (25th-75th pct.)",
               2: "Comps: EV / LTM revenue (25th-75th pct.)", 3: "Comps: P / E (25th-75th pct.)",
               4: "DCF: perpetuity growth (WACC ±0.5%, g ±0.5%)", 5: "DCF: exit multiple (WACC ±0.5%, ±1.0x)"}[i]
        rows.append((lab, vals[("COVER", f"G{r}")], vals[("COVER", f"H{r}")]))
    price = recalc.get(vals, reg, "price")
    vps = recalc.get(vals, reg, "vps")
    return rows, price, vps


def draw(theme, path):
    t = THEMES[theme]
    rows, price, vps = football_rows()
    rows = rows[::-1]
    fig, ax = plt.subplots(figsize=(10, 4.6), dpi=200)
    fig.patch.set_facecolor(t["surface"])
    ax.set_facecolor(t["surface"])
    h = 0.42
    for i, (lab, lo, hi) in enumerate(rows):
        ax.add_patch(FancyBboxPatch((lo, i - h / 2), hi - lo, h, boxstyle="round,pad=0,rounding_size=0.08",
                                    mutation_aspect=0.04, linewidth=0, facecolor=t["bar"]))
        lb = dict(facecolor=t["surface"], edgecolor="none", pad=1.0)
        ax.text(lo - 1.2, i, f"${lo:,.0f}", va="center", ha="right", fontsize=9, color=t["ink2"], bbox=lb, zorder=5)
        ax.text(hi + 1.2, i, f"${hi:,.0f}", va="center", ha="left", fontsize=9, color=t["ink2"], bbox=lb, zorder=5)
    ax.set_yticks(range(len(rows)))
    ax.set_yticklabels([r[0] for r in rows], fontsize=9.5, color=t["ink"])
    ax.axvline(price, color=t["ink2"], linewidth=1.5, linestyle=(0, (4, 3)), zorder=3)
    ax.axvline(vps, color=t["ink"], linewidth=2, zorder=3)
    top = len(rows) - 0.35
    ax.text(price, top, f"Price ${price:,.2f}", ha="right", va="bottom", fontsize=9, color=t["ink2"],
            bbox=dict(facecolor=t["surface"], edgecolor="none", pad=1.5))
    ax.text(vps, top, f"  DCF ${vps:,.2f}", ha="left", va="bottom", fontsize=9, color=t["ink"], fontweight="bold",
            bbox=dict(facecolor=t["surface"], edgecolor="none", pad=1.5))
    lo_all = min(r[1] for r in rows)
    hi_all = max(r[2] for r in rows)
    ax.set_xlim(max(0, lo_all - 15), hi_all + 15)
    ax.set_ylim(-0.6, len(rows) + 0.2)
    ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
    ax.tick_params(axis="x", colors=t["ink2"], labelsize=9, length=0)
    ax.tick_params(axis="y", length=0)
    ax.grid(axis="x", color=t["grid"], linewidth=0.8)
    ax.set_axisbelow(True)
    for s in ax.spines.values():
        s.set_visible(False)
    ax.set_title("Planet Fitness (PLNT): implied value per share by method", loc="left", fontsize=12,
                 color=t["ink"], fontweight="bold", pad=14)
    fig.text(0.01, 0.01, "Prices as of 2026-11-19. Base case. Bars show low-high ranges; source: output/PLNT_DCF_Model.xlsx (Cover tab).",
             fontsize=7.5, color=t["ink2"])
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(path, facecolor=t["surface"])
    plt.close(fig)


if __name__ == "__main__":
    out = ROOT / "docs"
    out.mkdir(exist_ok=True)
    for th in THEMES:
        draw(th, out / f"football_field_{th}.png")
        print("saved", out / f"football_field_{th}.png")
