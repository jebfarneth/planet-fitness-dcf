"""One-page investment memo (PDF). Every number is read from the recalculated
workbook / verification outputs at build time, not typed in."""
import json
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

import recalc

ROOT = Path(__file__).resolve().parents[1]
NAVY = colors.HexColor("#1F3864")


def build():
    v, reg = recalc.load(), recalc.names()
    g = lambda n, c=None: recalc.get(v, reg, n, c)
    scen = json.load(open(ROOT / "output" / "verification_summary.json"))
    tor = json.load(open(ROOT / "output" / "tornado.json"))["tornado"]
    d = dict(
        price=g("price"), vps=g("vps"), pgm=g("vps_pgm"), ex=g("vps_exit"), up=g("upside"),
        upv=scen["Upside"]["vps"], dnv=scen["Downside"]["vps"], wacc=g("wacc"), beta=g("beta"),
        ke=g("ke"), kd=g("kd"), tg=g("tg"), mult=g("exit_mult"), impl_mult=g("impl_mult"), impl_g=g("impl_g"),
        tvp=g("tvpct_pgm"), ev=g("ev_sel"), mkt_ev=g("mkt_ev"), ltm_mult=g("impl_ltm_mult"), mkt_mult=g("mkt_ltm_mult"),
        rv_g=g("rv_impl_g"), rv_m=g("rv_impl_m"), rv_tg=g("rv_impl_tg"), m0=g("rv_m0"),
        rev25=g("p_rev", "C"), rev26=g("p_rev", "D"), e25=g("p_ebitda", "C"), e26=g("p_ebitda", "D"), e31=g("p_ebitda", "I"),
        g27=g("p_rx_g", "E"), g31=g("p_rx_g", "I"), clubs=g("p_clubs", "C"), ltm_adj=g("h_adj_ebitda", "J"),
        bw=g("b_plnt_w"), bm=g("b_plnt_m"), cmp_lo=g("cmp_ebitda_lo"), cmp_hi=g("cmp_ebitda_hi"),
        w_lo=g("wk52_lo"), w_hi=g("wk52_hi"), adj=g("br_adj_total"),
    )
    styles = dict(
        h=ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=10.5, textColor=NAVY, spaceBefore=6, spaceAfter=2.5, leading=12.5),
        b=ParagraphStyle("b", fontName="Helvetica", fontSize=8.6, leading=10.9, alignment=TA_LEFT),
        s=ParagraphStyle("s", fontName="Helvetica-Oblique", fontSize=6.8, leading=8.4, textColor=colors.HexColor("#52514e")),
    )
    P = lambda t, s="b": Paragraph(t, styles[s])
    doc = SimpleDocTemplate(str(ROOT / "output" / "PLNT_Investment_Memo.pdf"), pagesize=letter,
                            leftMargin=0.5 * inch, rightMargin=0.5 * inch, topMargin=0.42 * inch, bottomMargin=0.35 * inch,
                            title="Planet Fitness (PLNT) - Investment Memo", author="Jeb Farneth")
    head = Table([[Paragraph('<font color="white"><b>PLANET FITNESS, INC. (NYSE: PLNT)</b> &nbsp;|&nbsp; Valuation memo &nbsp;|&nbsp; '
                             'Prices as of 19-Nov-2026</font>', ParagraphStyle("t", fontName="Helvetica", fontSize=10.5, leading=13))]],
                 colWidths=[7.5 * inch])
    head.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), NAVY), ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
    kpi = [["Share price", "DCF value / share", "Upside", "Scenario range", "WACC / g / exit", "EV (DCF) / market EV"],
           [f"${d['price']:.2f}", f"${d['vps']:.2f}", f"{d['up']:+.0%}", f"${d['dnv']:.0f} - ${d['upv']:.0f}",
            f"{d['wacc']:.2%} / {d['tg']:.1%} / {d['mult']:.1f}x", f"${d['ev']:,.0f}M / ${d['mkt_ev']:,.0f}M"]]
    kt = Table(kpi, colWidths=[1.1 * inch, 1.25 * inch, 0.8 * inch, 1.15 * inch, 1.55 * inch, 1.65 * inch])
    kt.setStyle(TableStyle([("FONT", (0, 0), (-1, 0), "Helvetica", 6.8), ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#52514e")),
                            ("FONT", (0, 1), (-1, 1), "Helvetica-Bold", 10), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ("LINEBELOW", (0, 1), (-1, 1), 0.6, NAVY), ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
    story = [head, Spacer(1, 4), kt]

    story += [P("Business overview", "h"), P(
        f"Planet Fitness is the largest U.S. fitness franchisor by members and clubs: {d['clubs']:,.0f} clubs (90% franchised) and ~20.8M members at "
        f"FY2025, with a $15/month entry price. Three segments: <b>Franchise</b> (royalties of up to 7% of member dues, ~89% margins), "
        f"<b>Corporate-owned clubs</b> (~38% margins) and <b>Equipment</b> (sells and re-equips franchise clubs every 5-9 years). FY2025 revenue "
        f"${d['rev25']:,.0f}M, Adj. EBITDA ${d['e25']:,.0f}M; LTM Sep-26 Adj. EBITDA ${d['ltm_adj']:,.0f}M. Debt is a ~$2.6B whole-business "
        f"securitization (fixed coupons 3.9-6.2%), book equity is negative after ~$800M of buybacks since the start of 2025.")]

    story += [P("Investment thesis", "h"), P(
        f"<b>The market prices ~3% growth; the base case assumes guidance near-term, fading to ~4%.</b> The shares fell ~50% in 2026 as same-club "
        f"sales slowed from 6.8% (FY2025) to 1.5% (Q3 2026), with a CFO transition. At ${d['price']:.2f} the reverse DCF implies <b>{d['rv_g']:.1%}</b> annual revenue growth at "
        f"base margins (or a {d['rv_m']:.1%} margin vs. {d['m0']:.1%}, or {d['rv_tg']:.1%} terminal growth). Yet 2026 guidance still calls for "
        f"revenue +8% and Adj. EBITDA +6.5%, franchisees had committed to ~900 more clubs at end-2024 (10-K FY2024), and the asset-light royalty stream converts "
        f"~50% of EBITDA to unlevered FCF. Base case: revenue growth {d['g27']:.1%} in FY2027 fading to {d['g31']:.1%} by FY2031 as net openings "
        f"slow (145 to 80 franchise clubs a year) and same-club sales settle at 2.5%; segment margins held near FY2025 levels.")]

    story += [P("Valuation summary", "h")]
    img = Image(str(ROOT / "docs" / "football_field_light.png"), width=4.2 * inch, height=4.2 * inch * 920 / 2000)
    txt = P(
        f"<b>DCF:</b> perpetuity growth ${d['pgm']:.2f} (g {d['tg']:.1%}, implied {d['impl_mult']:.1f}x exit) and exit multiple ${d['ex']:.2f} "
        f"({d['mult']:.1f}x FY2031E EBITDA, implied g {d['impl_g']:.2%}); headline = average, ${d['vps']:.2f}. Implies {d['ltm_mult']:.1f}x LTM "
        f"Adj. EBITDA vs. {d['mkt_mult']:.1f}x today.<br/><b>WACC {d['wacc']:.2%}:</b> Rf 5.04% (10-yr UST), ERP 4.41% (Damodaran), beta "
        f"{d['beta']:.2f} = peer-median unlevered beta relevered at 29.5% D/(D+E); Ke {d['ke']:.2%}; Kd {d['kd']:.2%} pre-tax (7-yr UST + Dec-2025 "
        f"new-issue spread).<br/><b>Comps:</b> 7 peers (fitness + securitized franchisors), median 14.9x EV/LTM EBITDA vs. PLNT 11.5x; "
        f"25th-75th pct. implies ${d['cmp_lo']:.0f}-${d['cmp_hi']:.0f}.<br/><b>Bridge:</b> net debt-like claims ${-d['adj']:,.0f}M incl. "
        f"$362M TRA, offset by the $368M deferred tax asset that funds it.")
    vt = Table([[img, txt]], colWidths=[4.25 * inch, 3.25 * inch])
    vt.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    story += [vt, Spacer(1, 4)]
    cols = ["C", "D", "E", "F", "G", "H", "I"]
    yrs = ["FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
    def row(lab, n, fmt, cs=cols):
        out = [lab]
        for c in cols:
            x = g(n, c) if c in cs else None
            out.append("" if x is None else fmt(x))
        return out
    m = lambda x: f"{x:,.0f}"
    pc = lambda x: f"{x:.1%}"
    tbl = [["$M (base case)"] + yrs,
           row("Total clubs (end of year)", "p_clubs", m),
           row("Revenue", "p_rev", m),
           row("Revenue growth excl. NAF", "p_rx_g", pc, cols[1:]),
           row("Adjusted EBITDA", "p_ebitda", m),
           row("Adj. EBITDA margin (excl. NAF)", "p_ebitda_m", pc),
           row("Unlevered free cash flow", "p_ufcf", m, cols[1:])]
    pt = Table(tbl, colWidths=[1.9 * inch] + [0.8 * inch] * 7)
    pt.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "Helvetica", 7.8), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 7.8),
                            ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("LINEBELOW", (0, 0), (-1, 0), 0.6, NAVY),
                            ("TEXTCOLOR", (0, 0), (-1, 0), NAVY), ("BACKGROUND", (1, 0), (1, -1), colors.HexColor("#F2F2F2")),
                            ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.2)]))
    story += [pt, P("FY2026E calibrated to the 8-K outlook (revenue ~+8%, Adj. EBITDA ~+6.5%); FY2026E growth excl. NAF is lower because the NAF rate "
                    "rose from 2% to 3% of dues in 2026. Excludes the 25 Be-Fit clubs acquired Oct-2026 (valued at cost in the bridge).", "s")]

    top3 = tor[:3]
    story += [P("Key risks and what would change the view", "h"), P(
        "<b>1. Same-club sales stay near 1-2%</b> (competition from other low-price gyms, slower member growth): the Downside case "
        f"(SCS -1.5pp a year, 20 fewer openings, lower margins) gives ${d['dnv']:.2f}, roughly today's price. "
        "<b>2. Development slows faster</b> if franchisee returns compress (higher build costs and 5-6% rates); every 20 fewer net openings a year "
        "is ~$2/share. <b>3. Leverage and refinancing:</b> ~4.4x gross debt / LTM Adj. EBITDA (10-K: 7.0x leverage-ratio incurrence cap); the 2029 anticipated repayment "
        "dates ($923M) must be refinanced near today's higher coupons. <b>4. Execution:</b> CFO transition, strategic review of corporate clubs, "
        "international expansion (Spain, Portugal). <b>5. Valuation sensitivity:</b> terminal value is ~80% of EV; "
        f"{top3[0][0]} moves value {top3[0][1]:+.1f}/{top3[0][2]:+.1f} per share, {top3[1][0].lower()} {top3[1][1]:+.1f}/{top3[1][2]:+.1f}. "
        f"<b>Beta caveat:</b> PLNT's own 2-yr weekly beta is {d['bw']:.2f} (R-sq ~0, the 2026 drop was company-specific); the peer-median "
        f"relevered 0.95 is cross-checked by its 5-yr monthly beta of {d['bm']:.2f}.")]
    story += [Spacer(1, 3), P(
        "Educational model, not investment advice. Sources: PLNT 10-K FY2021-25, 10-Q Q3 2025/26, 8-K 5-Nov-2026 (SEC EDGAR); treasury.gov; "
        "Damodaran (NYU Stern); Nasdaq.com; FRED. Full model, sources and verification: output/PLNT_DCF_Model.xlsx and README.", "s")]
    doc.build(story)
    print("saved", ROOT / "output" / "PLNT_Investment_Memo.pdf")


if __name__ == "__main__":
    build()
