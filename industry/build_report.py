"""Indian auto components: industry deep dive (PDF).

    python industry/build_report.py  ->  industry/Indian_Auto_Components_Deep_Dive.pdf

All figures are from the sources listed at the end of the report; charts are drawn
from the same numbers that appear in the text.
"""

from __future__ import annotations

import io
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (CondPageBreak, Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

HERE = Path(__file__).parent
OUT = HERE / "Indian_Auto_Components_Deep_Dive.pdf"

FONT_DIR = "/usr/share/fonts/truetype/crosextra/"
pdfmetrics.registerFont(TTFont("Carlito", FONT_DIR + "Carlito-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Carlito-Bold", FONT_DIR + "Carlito-Bold.ttf"))
pdfmetrics.registerFont(TTFont("Carlito-Italic", FONT_DIR + "Carlito-Italic.ttf"))
pdfmetrics.registerFontFamily("Carlito", normal="Carlito", bold="Carlito-Bold", italic="Carlito-Italic")
for f in ("Carlito-Regular.ttf", "Carlito-Bold.ttf"):
    font_manager.fontManager.addfont(FONT_DIR + f)
plt.rcParams.update({"font.family": "Carlito", "font.size": 10, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.edgecolor": "#A9B6C2", "axes.labelcolor": "#2C3E50",
                     "xtick.color": "#5D7183", "ytick.color": "#5D7183"})

INK, DARK, AMBER, STEEL, LIGHT, PALE, RED, GREEN = ("#1B2530", "#2C3E50", "#E39B2D", "#5D7183", "#A9B6C2",
                                                     "#F2F4F6", "#B23A3A", "#2E7D5B")

ss = {
    "title": ParagraphStyle("title", fontName="Carlito-Bold", fontSize=26, leading=31, textColor=colors.HexColor(INK)),
    "sub": ParagraphStyle("sub", fontName="Carlito", fontSize=13, leading=17, textColor=colors.HexColor(STEEL)),
    "h1": ParagraphStyle("h1", fontName="Carlito-Bold", fontSize=17, leading=21, textColor=colors.HexColor(INK), keepWithNext=1,
                         spaceBefore=6, spaceAfter=6),
    "h2": ParagraphStyle("h2", fontName="Carlito-Bold", fontSize=12.5, leading=16, textColor=colors.HexColor(DARK), keepWithNext=1,
                         spaceBefore=8, spaceAfter=3),
    "body": ParagraphStyle("body", fontName="Carlito", fontSize=10.5, leading=14.5, textColor=colors.HexColor(INK),
                           alignment=TA_LEFT, spaceAfter=6),
    "bullet": ParagraphStyle("bullet", fontName="Carlito", fontSize=10.5, leading=14.5,
                             textColor=colors.HexColor(INK), leftIndent=12, bulletIndent=2, spaceAfter=3),
    "small": ParagraphStyle("small", fontName="Carlito", fontSize=8.5, leading=11, textColor=colors.HexColor(STEEL)),
    "cap": ParagraphStyle("cap", fontName="Carlito-Italic", fontSize=8.5, leading=11, textColor=colors.HexColor(STEEL),
                          spaceAfter=8),
    "cell": ParagraphStyle("cell", fontName="Carlito", fontSize=9.5, leading=12, textColor=colors.HexColor(INK)),
    "cellb": ParagraphStyle("cellb", fontName="Carlito-Bold", fontSize=9.5, leading=12,
                            textColor=colors.HexColor(INK)),
    "box": ParagraphStyle("box", fontName="Carlito", fontSize=10.5, leading=14.5, textColor=colors.HexColor(INK)),
}
P = lambda t, s="body": Paragraph(t, ss[s])  # noqa: E731
B = lambda t: Paragraph(t, ss["bullet"], bulletText="•")  # noqa: E731
W = A4[0] - 4 * cm  # text width


def fig_image(fig, width=W):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    img = Image(buf)
    img._restrictSize(width, 9 * cm)
    return img


def table(rows, col_widths, header=True, zebra=True, bold_first_col=False):
    data = [[Paragraph(str(c), ss["cellb" if (header and i == 0) or (bold_first_col and j == 0) else "cell"])
             for j, c in enumerate(r)] for i, r in enumerate(rows)]
    t = Table(data, colWidths=col_widths, hAlign="LEFT")
    style = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor(LIGHT)),
             ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PALE)))
    t.setStyle(TableStyle(style))
    return KeepTogether([t])  # never split a table across pages


def callout(text):
    t = Table([[Paragraph(text, ss["box"])]], colWidths=[W])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(PALE)),
                           ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                           ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    return t


# ---------------------------------------------------------------- charts --
def chart_turnover():
    yrs = ["FY19", "FY20", "FY21", "FY22", "FY23", "FY24", "FY25", "FY26"]
    vals = [3.96, 3.50, 3.41, 4.21, 5.60, 6.15, 6.74, 7.59]
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    bars = ax.bar(yrs, vals, color=[LIGHT] * 7 + [AMBER], width=0.62)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.1, f"{v:.2f}", ha="center", fontsize=9, color=DARK)
    ax.set_ylabel("₹ lakh crore")
    ax.set_ylim(0, 8.8)
    ax.set_title("Auto component industry turnover", fontsize=11, color=DARK, loc="left")
    return fig_image(fig)


def chart_channels():
    labels = ["Sales to Indian OEMs", "Exports", "Aftermarket", "Less: imports", "Industry turnover"]
    vals = [6.63, 2.12, 1.08, -2.24, 7.59]
    fig, ax = plt.subplots(figsize=(7.2, 2.7))
    base = 0
    for i, (lab, v) in enumerate(zip(labels, vals)):
        if i == len(vals) - 1:
            ax.barh(i, v, color=AMBER, height=0.55)
            ax.text(v + 0.1, i, f"{v:.2f}", va="center", fontsize=9, color=DARK)
        else:
            start = base if v > 0 else base + v
            ax.barh(i, abs(v), left=start, color=DARK if v > 0 else RED, height=0.55)
            ax.text(start + abs(v) + 0.1, i, f"{v:+.2f}", va="center", fontsize=9, color=DARK)
            base += v
    ax.set_yticks(range(len(labels)), labels)
    ax.invert_yaxis()
    ax.set_xlim(0, 11.5)
    ax.set_xlabel("₹ lakh crore, FY26")
    ax.set_title("Where the ₹7.6 lakh crore comes from", fontsize=11, color=DARK, loc="left")
    return fig_image(fig)


def chart_volumes():
    seg = ["Two-wheelers", "Passenger vehicles", "Three-wheelers", "Commercial vehicles"]
    fy26 = [10.7, 7.9, 12.8, 12.6]
    h1 = [17.6, 28.3, 26.3, 23.8]
    import numpy as np
    x = np.arange(len(seg))
    fig, ax = plt.subplots(figsize=(7.2, 2.9))
    b1 = ax.bar(x - 0.18, fy26, 0.36, color=LIGHT, label="FY26")
    b2 = ax.bar(x + 0.18, h1, 0.36, color=AMBER, label="H1 FY27 (Apr-Sep 2026)")
    for bars in (b1, b2):
        for b in bars:
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.5, f"{b.get_height():.1f}%", ha="center",
                    fontsize=8.5, color=DARK)
    ax.set_xticks(x, seg)
    ax.set_ylabel("YoY growth, domestic wholesales")
    ax.set_ylim(0, 33)
    ax.legend(frameon=False, fontsize=9, loc="upper left")
    ax.set_title("Volume growth stepped up after the GST cut", fontsize=11, color=DARK, loc="left")
    return fig_image(fig)


def chart_e2w():
    oems = ["TVS Motor", "Bajaj Auto", "Ather Energy", "Ola Electric", "Hero (Vida)", "Others"]
    units = [341471, 289323, 239124, 164294, 144313, 1401663 - 1178525]
    fig, ax = plt.subplots(figsize=(7.2, 2.7))
    cols = [DARK, DARK, DARK, LIGHT, DARK, LIGHT]
    bars = ax.barh(oems, [u / 1000 for u in units], color=cols, height=0.6)
    ax.invert_yaxis()
    for b, u in zip(bars, units):
        ax.text(b.get_width() + 4, b.get_y() + b.get_height() / 2, f"{u / 1000:.0f}k  ({u / 1401663:.0%})",
                va="center", fontsize=9, color=DARK)
    ax.set_xlim(0, 430)
    ax.set_xlabel("Electric two-wheelers sold in FY26 ('000 units, Vahan retail)")
    ax.set_title("Incumbent OEMs now lead e-2W sales", fontsize=11, color=DARK, loc="left")
    return fig_image(fig)


def chart_peers():
    names = ["Endurance", "Uno Minda", "Sona BLW", "Gabriel India", "Bharat Forge"]
    pe = [34.9, 53.3, 71.6, 84.2, 126.2]
    roe = [14.9, 20.5, 11.1, 19.8, 11.6]
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    for n, x, y in zip(names, roe, pe):
        ax.scatter(x, y, s=90, color=AMBER if n == "Endurance" else DARK, zorder=3)
        ax.annotate(n, (x, y), xytext=(7, -3), textcoords="offset points", fontsize=9, color=DARK)
    ax.set_xlabel("Return on equity (%)")
    ax.set_ylabel("Trailing P/E (x)")
    ax.set_xlim(9, 24)
    ax.set_ylim(20, 140)
    ax.grid(color=PALE, zorder=0)
    ax.set_title("Valuation vs returns: the sector pays for growth, not ROE", fontsize=11, color=DARK, loc="left")
    return fig_image(fig)


# ---------------------------------------------------------------- report --
def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont("Carlito", 8)
    canvas.setFillColor(colors.HexColor(STEEL))
    if doc.page > 1:
        canvas.drawString(2 * cm, 1.2 * cm, "Indian auto components: industry deep dive · October 2026 · "
                                             "Student research, not investment advice")
        canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, str(doc.page))
    canvas.restoreState()


story = []
S = story.append

# Cover ----------------------------------------------------------------------
S(Spacer(1, 2.5 * cm))
S(P("Indian auto components", "title"))
S(P("Industry deep dive · October 2026", "sub"))
S(Spacer(1, 0.8 * cm))
S(callout("<b>In one paragraph.</b> India's auto component industry reached ₹7.6 lakh crore (about $86 billion) of "
          "turnover in FY26, growing 17% a year over five years, four times faster than vehicle volumes. Three forces "
          "drove it: more content per vehicle, localisation of parts that used to be imported, and exports. The next "
          "leg looks different. Domestic demand is strong after the September 2025 GST cut, but electrification is "
          "reshaping what gets bought, US tariffs have stalled the largest export market, and listed suppliers "
          "already trade at 30-80x earnings. Companies with powertrain-agnostic products and diversified customers "
          "should hold up best."))
S(Spacer(1, 0.6 * cm))
S(P("Key numbers", "h2"))
S(table([
    ["Metric", "Value", "Source"],
    ["Industry turnover, FY26", "₹7.6 lakh crore (+12.7%)", "ACMA"],
    ["5-year / 10-year turnover CAGR", "17% / 12% a year", "BCG-ACMA"],
    ["Exports / imports, FY26", "$24.0bn (+5%) / $25.4bn (+13%)", "ACMA"],
    ["Aftermarket, FY26", "₹1.08 lakh crore (+9%)", "ACMA"],
    ["Domestic vehicle sales, FY26", "28.3 million (+10.4%)", "SIAM"],
    ["Domestic vehicle sales, H1 FY27", "15.8 million (+19.8%)", "SIAM"],
    ["Electric 2W share of 2W sales, FY26", "6.5%", "Vahan via Autocar Pro"],
    ["Exports to the US, FY26", "$6.24bn, ~27% of component exports", "Business Today / ACMA"],
], [6.2 * cm, 6.0 * cm, 4.8 * cm]))
S(Spacer(1, 0.5 * cm))
S(P("Key takeaways", "h2"))
for t in [
    "<b>Components outgrow vehicles.</b> Over the last decade component turnover grew about 12% a year in value "
    "while vehicle volumes grew about 3% a year. The gap is content per vehicle: ABS, alloy wheels, electronics, "
    "airbags and premium trims.",
    "<b>Demand is unusually strong right now.</b> The GST cut from 28% to 18% on small cars and sub-350cc "
    "motorcycles (from 22 Sep 2025) pushed H1 FY27 passenger-vehicle volumes up 28%.",
    "<b>Electrification changes the bill of materials.</b> EV parts were only 4.6% of supplies to Indian OEMs in "
    "FY26, but engine and transmission parts face a long decline while castings, brakes, suspension and "
    "electronics carry over.",
    "<b>Exports have hit a wall in the US.</b> Exports to the US were flat in FY26; about half face a 25% Section 232 "
    "tariff and two Section 301 probes are open. Europe was the fastest-growing export market.",
    "<b>The market prices in a lot.</b> Peer median trailing P/E is around 60x. Endurance, at 35x, is the cheapest "
    "quality name, which is the subject of the accompanying stock pitch.",
]:
    S(B(t))
S(PageBreak())

# 1. Size and structure --------------------------------------------------------
S(P("1. Size and structure", "h1"))
S(P("India's auto component industry makes the parts that go into vehicles: engine and drivetrain parts, "
    "suspension and braking, body and chassis, electricals and electronics, interiors. It sells through three "
    "channels: to vehicle makers (OEMs) for new vehicles, to the replacement market (aftermarket), and abroad. "
    "Turnover, as ACMA measures it, is OEM supplies plus exports plus aftermarket, less imported parts."))
S(chart_turnover())
S(P("Source: BCG-ACMA 'Beyond Resilience' report via Autocar Professional (Sep 2026). FY16 turnover was ₹2.56 lakh "
    "crore. FY20-FY21 dip reflects the BS-VI transition and Covid.", "cap"))
S(P("After two flat years around the pandemic, turnover more than doubled from ₹3.41 lakh crore in FY21 to ₹7.59 "
    "lakh crore in FY26. Growth is slowing from the post-Covid rebound (FY23: +33%) to a more normal 10-13% a year, "
    "but that is still well ahead of vehicle volumes."))
S(chart_channels())
S(P("Source: ACMA FY26 industry performance review; BCG-ACMA. Imports are implied by ACMA's definition of turnover.",
    "cap"))
S(P("Supplies to Indian OEMs are the core of the industry, at ₹6.6 lakh crore and growing 19% a year over five "
    "years. Exports, at ₹2.1 lakh crore ($24bn), are roughly a quarter of turnover. The aftermarket is smaller and "
    "slower (9-11% a year) but steadier, because it follows the number of vehicles on the road rather than new "
    "sales. India still imports slightly more components than it exports ($25.4bn vs $24.0bn in FY26), mostly from "
    "China, Japan and Germany, though localisation has risen from about 60% to above 70% over the decade. Imports of "
    "wheels and rims fell about 74% between FY19 and FY26."))

S(P("Why components grow faster than vehicles", "h2"))
for t in [
    "<b>Content per vehicle.</b> Regulation and buyers are adding parts: ABS on two-wheelers, airbags and electronic "
    "stability control on cars, LED lighting, connected features, alloy wheels.",
    "<b>Premiumisation.</b> Scooters and higher-capacity motorcycles are growing faster than entry commuter bikes, "
    "and SUVs dominate car sales. Each carries more and more expensive parts.",
    "<b>Localisation.</b> Parts that used to be imported (wheels, engine parts, steering) are increasingly made in "
    "India, helped by PLI incentives and import substitution.",
    "<b>Exports and 'China+1'.</b> Global OEMs and Tier-1s are adding Indian suppliers to diversify away from China.",
]:
    S(B(t))
S(Spacer(1, 0.4 * cm))

# 2. Demand -------------------------------------------------------------------
S(CondPageBreak(8 * cm))
S(P("2. Domestic demand: strong cycle after GST 2.0", "h1"))
S(P("FY26 was a record year: domestic wholesales of 28.3 million vehicles, up 10.4%, with every segment at an "
    "all-time high. Two-wheelers, at 21.7 million units, are three-quarters of volume; scooters (+18.5%) grew much "
    "faster than motorcycles (+6.6%). Passenger vehicles grew 7.9% to 4.6 million."))
S(P("The step change came with the GST rationalisation effective 22 September 2025. GST on small cars and "
    "motorcycles up to 350cc fell from 28% (plus cess) to 18%; larger cars moved to a flat 40% with cess removed, "
    "and motorcycles above 350cc rose to 40%. EVs stayed at 5%. The result shows in the first half of FY27:"))
S(chart_volumes())
S(P("Source: SIAM via Outlook Business (FY26) and Autocar Professional (Q2 FY27, 15 Oct 2026). H1 FY27 also benefits "
    "from a low base.", "cap"))
S(P("Part of this is a low base and pent-up demand, so growth of 18-28% should not be extrapolated. SIAM expects "
    "steady growth in FY27 but flags the West Asia conflict (fuel, freight, commodity prices) as the main risk. The "
    "RBI's 25 bp rate increase to 5.5% in October 2026 will also make vehicle loans more expensive at the margin."))
S(P("What this means for suppliers", "h2"))
S(P("Suppliers to scooter makers and small-car makers benefit most from the GST cut. Two-wheeler suppliers have an "
    "additional tailwind from ABS: dual-channel ABS for larger motorcycles is a high-value product that few "
    "Indian suppliers, Endurance among them, make locally. Suppliers exposed to motorcycles above "
    "350cc face a small headwind from the higher tax rate."))
S(Spacer(1, 0.4 * cm))

# 3. Electrification -----------------------------------------------------------
S(CondPageBreak(8 * cm))
S(P("3. Electrification: a change in the bill of materials", "h1"))
S(P("India sold about 2.45 million EVs in FY26. Electric two-wheelers were the largest group at 1.40 million units "
    "(+22%), or 6.5% of all two-wheelers, up from 6.1% in FY25. The market has consolidated around incumbents: TVS, "
    "Bajaj and Hero together sold more than half of all e-2Ws, while Ola Electric's sales halved."))
S(chart_e2w())
S(P("Source: Vahan registrations via Autocar Professional, Apr 2026.", "cap"))
S(P("For component makers, the question is which parts survive. ACMA estimates EV-specific components were only "
    "4.6% of supplies to Indian OEMs in FY26 (excluding battery cells), so the transition is gradual, but the "
    "direction is clear:"))
S(table([
    ["Exposure", "Examples", "Implication"],
    ["At risk (ICE-only)", "Pistons, fuel systems, exhaust, multi-speed transmissions, clutches",
     "Volumes peak and decline; replacement demand lasts longer"],
    ["Powertrain-agnostic", "Suspension, brakes, wheels, aluminium castings, steering, seating, lighting",
     "Carry over to EVs; often heavier-duty versions (EVs weigh more)"],
    ["New EV content", "Battery packs and BMS, motors and controllers, DC-DC converters, chargers",
     "New revenue pools, but competition from Chinese suppliers and OEM in-sourcing"],
], [3.4 * cm, 6.8 * cm, 6.8 * cm], bold_first_col=True))
S(Spacer(1, 0.3 * cm))
S(P("Aluminium die-casting is a good example of a part that gains from electrification: motor housings, battery "
    "enclosures and lightweight structures replace steel. Endurance reports ₹1,806 crore of cumulative EV orders, "
    "spread across castings, brakes, suspension and battery packs, and in Europe nearly half of its five-year order "
    "book is for EVs and another 39% for hybrids."))
S(Spacer(1, 0.4 * cm))

# 4. Exports -------------------------------------------------------------------
S(CondPageBreak(8 * cm))
S(P("4. Exports and trade: the US stalls, Europe opens", "h1"))
S(P("Component exports grew only 5% in FY26 to $24.0 billion, after a 17% five-year CAGR, while imports grew 13% to "
    "$25.4 billion. Engine components, drive transmission and steering made up more than half of exports, and "
    "Europe was the fastest-growing destination."))
S(table([
    ["", "FY25", "FY26", "Change"],
    ["Exports ($bn)", "~22.9*", "24.0", "+5%"],
    ["Imports ($bn)", "~22.5*", "25.4", "+13%"],
    ["Exports to the US ($bn)", "6.18", "6.24", "+1%"],
    ["US share of exports", "", "~27%", ""],
], [6.2 * cm, 3.4 * cm, 3.4 * cm, 4.0 * cm], bold_first_col=True))
S(P("* FY25 export and import figures back-calculated from the FY26 growth rates. Sources: ACMA; Business Today "
    "(7 Jul 2026).", "cap"))
S(P("The US problem", "h2"))
S(P("The US is India's largest component market, but exports there were flat in FY26. Nearly half of those exports "
    "fall under Section 232 and pay a 25% tariff. Two Section 301 investigations are running in parallel: one on "
    "labour practices has proposed a 12.5% tariff (not stacked on Section 232 goods), and one on subsidy-driven "
    "overcapacity has no outcome yet. ACMA disputes both allegations and is pinning its hopes on the India-US trade "
    "negotiations. ACMA has also reported that some US automakers are holding back new sourcing decisions until "
    "tariffs are clear."))
S(P("The European opportunity", "h2"))
S(P("Europe grew fastest in FY26, and trade agreements with the EU and UK should help further. The opportunity is "
    "structural: European OEMs are cutting costs while shifting to EVs and hybrids, and Chinese brands are taking "
    "share in Europe (EU registrations ex-Chinese OEMs grew only 3% in Endurance's Q1 FY27 data). Indian suppliers "
    "that already have European plants, such as Endurance (Italy and Germany), Sona BLW and Motherson, are best "
    "placed. The cost of that presence is energy: Italian electricity prices were up 22% year on year in Q1 FY27."))
S(Spacer(1, 0.4 * cm))

# 5. Policy --------------------------------------------------------------------
S(CondPageBreak(8 * cm))
S(P("5. Policy support", "h1"))
S(table([
    ["Policy", "What it does", "Status"],
    ["PLI for Auto & Auto Components", "Incentives on incremental sales of approved advanced-technology "
     "vehicles and components with certified domestic value addition",
     "₹25,938 cr outlay, FY24-FY28. ₹35,657 cr invested and ₹2,322 cr disbursed by end-2025"],
    ["GST 2.0 (Sep 2025)", "Cut tax on small cars and sub-350cc 2Ws from 28% to 18%; EVs at 5%",
     "In force; main driver of FY27 volume growth"],
    ["Free-trade agreements", "Lower tariffs into the EU and UK; US deal under negotiation",
     "Supportive for exports, US outcome uncertain"],
    ["State incentives", "Capital subsidies for new plants (e.g. Maharashtra package for Endurance raised from "
     "₹600 cr to ₹858 cr)", "Lowers effective capex for greenfield plants"],
], [3.8 * cm, 7.2 * cm, 6.0 * cm], bold_first_col=True))
S(P("Sources: Angel One summary of PLI-Auto progress (Jan 2026); Autocar India (GST, Sep 2025); Endurance Q1 FY27 "
    "results.", "cap"))
S(P("PLI matters less for the established component makers than for companies building new EV and electronics "
    "capacity, because incentives apply to incremental sales of approved advanced-technology products. Disbursements "
    "only began in FY25, so most of the scheme's cash is still ahead."))

# 6. Competitive landscape ------------------------------------------------------
S(CondPageBreak(8 * cm))
S(P("6. Competitive landscape and valuation", "h1"))
S(P("The industry is fragmented, and the top listed companies each specialise in a few product families. Scale "
    "advantages come from tooling and engineering capability, long OEM approval cycles for new parts, and "
    "proximity to OEM plants."))
S(table([
    ["Company", "Core products", "Mkt cap (₹ cr)", "P/E (TTM)", "ROE"],
    ["Endurance Technologies", "Castings, suspension, brakes, alloy wheels; 2W focus + Europe", "33,829", "34.9x",
     "14.9%"],
    ["Uno Minda", "Lighting, switches, alloy wheels, EV components", "64,110", "53.3x", "20.5%"],
    ["Sona BLW Precision", "EV driveline, differential gears, motors", "49,859", "71.6x", "11.1%"],
    ["Gabriel India", "Shock absorbers and suspension", "25,075", "84.2x", "19.8%"],
    ["Minda Corporation", "Wiring harness, locks, sensors", "15,537", "31.0x", "n/a"],
    ["Bharat Forge", "Forgings, industrial and defence", "89,129", "126.2x", "11.6%"],
], [3.9 * cm, 6.3 * cm, 2.4 * cm, 2.1 * cm, 2.3 * cm], bold_first_col=True))
S(P("Sources: Value Research (prices 5-16 Oct 2026), 5paisa (Minda Corp), Screener (Endurance ROE).", "cap"))
S(chart_peers())
S(P("There is little relationship between returns on equity and valuation. Investors are paying for expected growth "
    "and EV exposure: Sona BLW, with an 11% ROE, trades at 72x earnings, while Endurance, with a 15% ROE, trades at "
    "35x. A sector-wide P/E around 60x implies many years of very high earnings growth, which few industrial "
    "businesses have delivered. That is why the accompanying DCF of Endurance comes out well below its "
    "market price, even though it looks cheap against peers."))
S(Spacer(1, 0.4 * cm))

# 7. Risks and outlook ----------------------------------------------------------
S(CondPageBreak(8 * cm))
S(P("7. Risks", "h1"))
S(table([
    ["Risk", "Why it matters", "Who is most exposed"],
    ["Commodity and energy costs", "Aluminium, steel and energy are passed on to OEMs with a lag of one or two "
     "quarters; margins dip in between", "Casting and forging makers; European plants"],
    ["US tariffs", "25% on about half of US-bound exports, more possible from Section 301 probes",
     "Export-heavy suppliers (engine parts, forgings)"],
    ["EV disruption of ICE parts", "Transmission and engine parts lose volume as EV share rises",
     "Single-product ICE suppliers"],
    ["Customer concentration", "Many suppliers depend on one or two OEMs (e.g. Bajaj was ~50% of Endurance's "
     "standalone revenue in FY25)", "Captive-style suppliers"],
    ["Cyclicality", "Vehicle demand follows rates, fuel prices and rural incomes; FY27 growth is flattered by a "
     "low base", "All, especially those that added capacity for the boom"],
    ["Valuation", "Sector multiples assume long, uninterrupted growth", "Highest-multiple names"],
], [3.8 * cm, 8.0 * cm, 5.2 * cm], bold_first_col=True))
S(Spacer(1, 0.3 * cm))
S(CondPageBreak(8 * cm))
S(P("8. Outlook: what to watch", "h1"))
for t in [
    "<b>Margins in H2 FY27.</b> Whether suppliers pass through input-cost increases. Endurance's standalone margin fell "
    "from 13.0% to 11.2% in Q1 FY27.",
    "<b>Volume growth after the festive season</b>, once the GST-driven pull-forward and low base wash out.",
    "<b>The India-US trade deal</b> and the outcome of the Section 301 overcapacity probe.",
    "<b>E-2W penetration</b> moving from 6.5% toward double digits, and how much of the EV bill of materials Indian "
    "suppliers capture versus imports.",
    "<b>Capex discipline.</b> Many suppliers are building new plants; returns depend on filling them.",
]:
    S(B(t))
S(Spacer(1, 0.3 * cm))
S(callout("<b>Bottom line.</b> The industry's long-term growth case is intact and well supported by content per "
          "vehicle, localisation and exports. In the near term, demand is strong but flattered by tax cuts and a "
          "low base, margins are under input-cost pressure, and the US export channel is stuck. With listed peers "
          "on 30-80x earnings, stock selection should favour powertrain-agnostic products, diversified customers "
          "and a reasonable price, and be wary of paying today for growth that is a decade away."))
S(PageBreak())

# Sources ------------------------------------------------------------------------
S(P("Sources", "h1"))
for t in [
    "ACMA FY26 industry performance review, via Autocar Professional, 'Indian auto component industry grows 12.7% to "
    "INR 7.6 lakh crore in FY26'. autocarpro.in/news/acma-indian-auto-component-industry-grows-127-percent-to-inr-76-"
    "lakh-crore-in-fy26-133454",
    "BCG-ACMA 'Beyond Resilience' report, via Autocar Professional, 'Auto component turnover hits Rs 7.6 lakh crore, "
    "grows 17 percent a year over five years' (2 Sep 2026).",
    "SIAM FY26 data, via Outlook Business, 'Automobile wholesales in India clock record 2.83 crore units in FY26'.",
    "SIAM Q2 FY27 data, via Autocar Professional, 'Auto sales up 18.2% in July-Sept, PV sales rise 30.7%' "
    "(15 Oct 2026).",
    "Autocar India, 'Small car and motorcycle GST rates slashed from 28 to 18 percent' (3 Sep 2025).",
    "Autocar Professional, 'Record 1.4 million electric 2Ws sold in FY2026'.",
    "Business Today, 'US targets India's $6 billion auto component exports with twin probes' (7 Jul 2026).",
    "Business Standard, 'US automakers hesitate on Indian components amid tariffs, says ACMA' (Jan 2026).",
    "Angel One, 'PLI-Auto scheme: progress and impact' (1 Jan 2026).",
    "Endurance Technologies FY26 results press release (NSE, 14 May 2026) and Q1 FY27 investor presentation "
    "(NSE, 13 Aug 2026).",
    "ICRA rating rationale, Endurance Technologies (30 Sep 2025).",
    "Value Research, 5paisa and Screener stock pages for peer market data (5-18 Oct 2026).",
    "Trading Economics, India 10-year bond yield and RBI rate news (Sep-Oct 2026).",
]:
    S(B(t))
S(Spacer(1, 0.5 * cm))
S(P("This report is student research written for learning purposes. It is not investment advice. Figures are as "
    "reported by the sources above on the dates shown; some peer data is two weeks older than the rest.", "small"))

doc = SimpleDocTemplate(str(OUT), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm,
                        bottomMargin=2 * cm, title="Indian Auto Components: Industry Deep Dive",
                        subject="Industry research, October 2026")
doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print("Wrote", OUT)
