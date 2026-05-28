"""
build_presentation.py
The Case for Nuclear Power in Zambia (8-slide .pptx).
Editorial / supervisor-grade rebuild.

Palette: forest green (Zambian flag), burnt copper, gold accent,
parchment background, charcoal type. Argument leans hard on
Zambia-specific evidence: Kariba 2024 collapse, drought-cost ledger,
African nuclear leadership gap.

Run:
  pip install -r requirements.txt
  python build_presentation.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (XL_CHART_TYPE, XL_LEGEND_POSITION,
                             XL_LABEL_POSITION)

OUTPUT = "The_Case_for_Nuclear_Power_in_Zambia.pptx"

# Zambian-inspired palette
PARCH = RGBColor(0xF4, 0xEF, 0xE4)
PARCH_DEEP = RGBColor(0xEA, 0xE2, 0xCF)
INK = RGBColor(0x14, 0x18, 0x1F)
INK_SOFT = RGBColor(0x3A, 0x40, 0x4A)
MUTED = RGBColor(0x6B, 0x70, 0x78)
RULE = RGBColor(0xC2, 0xB8, 0xA0)

GREEN = RGBColor(0x14, 0x40, 0x2C)       # Zambian flag green
GREEN_DEEP = RGBColor(0x08, 0x26, 0x18)
GREEN_LIGHT = RGBColor(0x2E, 0x66, 0x49)

COPPER = RGBColor(0xB5, 0x4C, 0x1A)       # Zambian copper
COPPER_LIGHT = RGBColor(0xD9, 0x79, 0x33)
GOLD = RGBColor(0xCF, 0xA3, 0x3C)

CRISIS = RGBColor(0x8C, 0x1F, 0x1F)
CRISIS_DEEP = RGBColor(0x5C, 0x12, 0x12)

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHITE = RGBColor(0xFA, 0xF7, 0xF0)

TITLE_FONT = "Georgia"
BODY_FONT = "Calibri"

CITATIONS = {
    "doe": "U.S. DOE REMS 2024 Report",
    "ndp": "Zambia 7NDP & Vision 2030",
    "iaea": "IAEA Milestone Approach",
    "mines": "Zambia Ministry of Mines — Annual Mining Statistics",
    "zesco": "ZESCO Annual Reports 2023-2024",
    "siavonga": "Siavonga Uranium Baseline Survey (2024)",
    "imf": "IMF Article IV Consultation — Zambia 2024",
    "wna": "World Nuclear Association Country Profiles (2024)",
}

BIG_IDEA = (
    "Zambia is one of the few nations on earth that sits on its own "
    "uranium, runs a 50-year radiological regulator, and holds a "
    "Cabinet-approved nuclear policy — yet still loses billions to "
    "drought every cycle; only enacting the Nuclear Bill turns that "
    "paradox into the drought-proof power our industrialisation demands."
)

THREE_MIN = (
    "Zambia is one of a handful of countries on earth that exports its "
    "own uranium, runs a fifty-year-old Radiation Protection Authority, "
    "holds a Cabinet-approved 2020 Nuclear Policy, and has signed "
    "cooperation agreements with ROSATOM, the U.S. IP3 Allied Nuclear "
    "Partners, and South Korea's KAERI — and yet, in 2026, still "
    "depends on a single river system for eighty-four percent of its "
    "electricity. In late 2024 that single dependency cost us up to "
    "twenty-one hours of load-shedding per day. Copper mines throttled. "
    "The kwacha wobbled. The IMF estimated roughly two percent of GDP "
    "evaporated in drought-driven blackouts. And it is not a one-off: "
    "2015/16 cost us a thousand megawatts and four hundred and forty "
    "million dollars in emergency imports; 2018/19 took eight hundred "
    "and seventy-two megawatts; 2023/24 was the worst drought in forty "
    "years. Meanwhile, in Siavonga, the 2024 Uranium Baseline Survey "
    "detected unmonitored uranium in drinking water and household dust "
    "— communities living beside the very resource that should be "
    "fuelling our future, with no instruments and no baseline. That is "
    "radiological fear. Contrast that with the United States, where my "
    "analysis of the Department of Energy Radiation Exposure Monitoring "
    "System covers twenty-two thousand workers averaging just "
    "fifty-nine point one-two millirem of exposure per year — about "
    "one point two percent of the federal occupational limit, with "
    "workforce growth producing zero rise in individual dose. That is "
    "industrial mastery. Same element, two opposite outcomes; the "
    "variable is the system. While we debate, Egypt is building four "
    "reactors at El Dabaa, Kenya is targeting first criticality by "
    "2034, and South Africa has been operating Koeberg for over forty "
    "years. We have the uranium. We have the regulator. We have the "
    "policy. We have the partners. We have a three-million-tonne "
    "copper ambition that no hydropower fleet on this continent can "
    "power. What we do not have is the legislative signature. Pass the "
    "Nuclear Bill, fund radiological monitoring nationwide starting in "
    "Siavonga, and greenlight Small Modular Reactor pre-feasibility. "
    "The fuel is under our feet. The expertise is on our payroll. The "
    "bill is on the table. Sign it."
)


# ---------- helpers ----------
def set_bg(slide, color):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = color


def rect(slide, left, top, width, height, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def line_rule(slide, left, top, width, color, weight=1.0):
    s = slide.shapes.add_connector(1, left, top, left + width, top)
    s.line.color.rgb = color
    s.line.width = Pt(weight)
    return s


def vline(slide, left, top, height, color, weight=1.0):
    s = slide.shapes.add_connector(1, left, top, left, top + height)
    s.line.color.rgb = color
    s.line.width = Pt(weight)
    return s


def textbox(slide, left, top, width, height, runs, *,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.15):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    for r in runs:
        text = r["text"]
        if text.startswith("\n"):
            p = tf.add_paragraph()
            p.alignment = r.get("align", align)
            p.line_spacing = r.get("line_spacing", line_spacing)
            if "space_after" in r:
                p.space_after = Pt(r["space_after"])
            text = text[1:]
        run = p.add_run()
        run.text = text
        run.font.name = r.get("font", BODY_FONT)
        run.font.size = Pt(r["size"])
        run.font.bold = r.get("bold", False)
        run.font.italic = r.get("italic", False)
        run.font.color.rgb = r.get("color", INK)
    return tb


def t(slide, left, top, width, height, text, *,
      size=14, bold=False, italic=False, color=INK,
      align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
      font=BODY_FONT, line_spacing=1.25):
    return textbox(slide, left, top, width, height,
                   [{"text": text, "size": size, "bold": bold,
                     "italic": italic, "color": color, "font": font}],
                   align=align, anchor=anchor, line_spacing=line_spacing)


def eyebrow(slide, left, top, text):
    return t(slide, left, top, Inches(10), Inches(0.3),
             text, size=10, bold=True, color=COPPER, font=BODY_FONT)


def header_block(slide, eyebrow_text, title, kicker=None):
    eyebrow(slide, Inches(0.6), Inches(0.50), eyebrow_text)
    t(slide, Inches(0.6), Inches(0.80), Inches(12.1), Inches(0.85),
      title, size=30, bold=True, color=GREEN_DEEP, font=TITLE_FONT,
      line_spacing=1.05)
    line_rule(slide, Inches(0.6), Inches(1.65), Inches(0.8), COPPER, 2.0)
    if kicker:
        t(slide, Inches(0.6), Inches(1.78), Inches(12.1), Inches(0.4),
          kicker, size=13, italic=True, color=INK_SOFT)


def footer(slide, n, total, cites):
    line_rule(slide, Inches(0.6), Inches(7.18), Inches(12.1),
              RULE, 0.5)
    t(slide, Inches(0.6), Inches(7.24), Inches(10.5), Inches(0.20),
      "SOURCES: " + " · ".join(cites),
      size=8, color=MUTED, italic=True)
    t(slide, Inches(11.3), Inches(7.24), Inches(1.8), Inches(0.20),
      f"{n} / {total}   E. MWABA",
      size=8, bold=True, color=INK_SOFT, align=PP_ALIGN.RIGHT)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def color_points(series, colors):
    for i, pt in enumerate(series.points):
        if i >= len(colors):
            break
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = colors[i]
        pt.format.line.fill.background()


def style_chart(chart, label_color=INK, axis_color=INK_SOFT, size=10):
    try:
        for plot in chart.plots:
            if plot.has_data_labels:
                dl = plot.data_labels
                dl.font.size = Pt(size + 1)
                dl.font.color.rgb = label_color
                dl.font.bold = True
                dl.font.name = BODY_FONT
    except Exception:
        pass
    try:
        for axis in (chart.category_axis, chart.value_axis):
            tl = axis.tick_labels
            tl.font.size = Pt(size)
            tl.font.color.rgb = axis_color
            tl.font.name = BODY_FONT
    except Exception:
        pass


# -----------------------------------------------------------------------
# SLIDE 1 — COVER + BIG IDEA
# -----------------------------------------------------------------------
def slide_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, PARCH)

    # Left vertical band — Zambian green
    rect(s, Inches(0), Inches(0), Inches(0.45), Inches(7.5), GREEN_DEEP)
    rect(s, Inches(0.45), Inches(0), Inches(0.08), Inches(7.5), COPPER)
    rect(s, Inches(0.53), Inches(0), Inches(0.04), Inches(7.5), GOLD)

    # Eyebrow
    t(s, Inches(1.0), Inches(0.65), Inches(11), Inches(0.3),
      "A POLICY BRIEF  ·  REPUBLIC OF ZAMBIA  ·  2026",
      size=11, bold=True, color=COPPER, font=BODY_FONT)
    line_rule(s, Inches(1.0), Inches(0.95), Inches(0.6), COPPER, 1.5)

    # Title — large editorial serif
    textbox(
        s, Inches(1.0), Inches(1.30), Inches(11.5), Inches(2.6),
        [
            {"text": "The Case for", "size": 44, "bold": False,
             "color": INK_SOFT, "font": TITLE_FONT, "italic": True},
            {"text": "\nNuclear Power", "size": 64, "bold": True,
             "color": GREEN_DEEP, "font": TITLE_FONT, "space_after": 0},
            {"text": "\nin Zambia.", "size": 64, "bold": True,
             "color": COPPER, "font": TITLE_FONT, "space_after": 0},
        ],
        line_spacing=1.0,
    )

    # Big Idea panel
    rect(s, Inches(1.0), Inches(4.85), Inches(11.5), Inches(1.85),
         GREEN_DEEP)
    rect(s, Inches(1.0), Inches(4.85), Inches(0.10), Inches(1.85),
         COPPER)
    t(s, Inches(1.30), Inches(4.95), Inches(11.0), Inches(0.3),
      "THE  BIG  IDEA",
      size=10, bold=True, color=GOLD)
    t(s, Inches(1.30), Inches(5.25), Inches(11.0), Inches(1.5),
      BIG_IDEA,
      size=14, italic=True, color=WHITE, font=TITLE_FONT,
      line_spacing=1.30)

    # Author block
    t(s, Inches(1.0), Inches(6.95), Inches(8), Inches(0.25),
      "PREPARED BY  ELIZABETH MWABA",
      size=10, bold=True, color=INK)
    t(s, Inches(1.0), Inches(7.18), Inches(8), Inches(0.25),
      "Radiological Analyst  ·  Supervisor Review Brief",
      size=9, italic=True, color=MUTED)

    notes(s,
        "[Opening] Three minutes. One argument. Captured in one "
        "sentence on this slide: Zambia is among the very few nations "
        "on earth that already owns the uranium, the regulator, and "
        "the nuclear policy — yet still loses billions every drought "
        "cycle. Only enacting the Nuclear Bill converts that paradox "
        "into drought-proof energy. Let me walk you through why this "
        "is no longer optional.")


# -----------------------------------------------------------------------
# SLIDE 2 — 3-MINUTE STORY
# -----------------------------------------------------------------------
def slide_story(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, PARCH)
    header_block(s,
                 "PROLOGUE  ·  THE 3-MINUTE STORY",
                 "If I had no slides, this is what I would say.",
                 "One paragraph. Three minutes. The full argument.")

    # Pull-quote frame
    rect(s, Inches(0.6), Inches(2.30), Inches(12.1), Inches(4.65),
         OFFWHITE)
    rect(s, Inches(0.6), Inches(2.30), Inches(0.12), Inches(4.65),
         COPPER)
    rect(s, Inches(12.58), Inches(2.30), Inches(0.12), Inches(4.65),
         GREEN_DEEP)

    t(s, Inches(0.95), Inches(2.45), Inches(11.5), Inches(0.3),
      "EXECUTIVE  CADENCE  ·  ~390 WORDS  ·  ~3:00",
      size=9, bold=True, color=COPPER)

    t(s, Inches(0.95), Inches(2.80), Inches(11.5), Inches(4.10),
      THREE_MIN,
      size=11, color=INK, font=BODY_FONT, line_spacing=1.34)

    footer(s, 2, 8,
           [CITATIONS["doe"], CITATIONS["zesco"],
            CITATIONS["siavonga"], CITATIONS["imf"]])
    notes(s,
        "[Read or paraphrase the paragraph at executive cadence. Build "
        "tension: paradox → drought cost ledger → Siavonga fear → DOE "
        "mastery proof → African competition → ask. End on the "
        "imperative: Sign it.]")


# -----------------------------------------------------------------------
# SLIDE 3 — THE HOOK (SIAVONGA vs DOE)
# -----------------------------------------------------------------------
def slide_hook(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, PARCH)
    header_block(s,
                 "01  ·  THE HOOK",
                 "Same element. Two realities.",
                 "Zambians and Americans both live with uranium. "
                 "Only one nation measures it.")

    panel_y = Inches(2.40)
    panel_h = Inches(4.25)

    # LEFT — Siavonga
    rect(s, Inches(0.6), panel_y, Inches(6.05), panel_h, CRISIS_DEEP)
    rect(s, Inches(0.6), panel_y, Inches(0.12), panel_h, GOLD)
    t(s, Inches(0.95), Inches(2.55), Inches(5.5), Inches(0.3),
      "ZAMBIA  ·  SIAVONGA  ·  2024",
      size=10, bold=True, color=GOLD)
    t(s, Inches(0.95), Inches(2.90), Inches(5.5), Inches(1.7),
      "Uranium in the water.\nUranium in the dust.",
      size=30, bold=True, color=WHITE, font=TITLE_FONT,
      line_spacing=1.05)
    t(s, Inches(0.95), Inches(4.85), Inches(5.5), Inches(1.1),
      "Zero monitoring.  Zero baseline.\nZero accountability.",
      size=13, color=RGBColor(0xF2, 0xCF, 0xCF),
      line_spacing=1.3, italic=True)
    line_rule(s, Inches(0.95), Inches(5.95), Inches(0.7), GOLD, 1.5)
    t(s, Inches(0.95), Inches(6.05), Inches(5.5), Inches(0.4),
      "RADIOLOGICAL  FEAR",
      size=12, bold=True, color=GOLD)

    # RIGHT — DOE
    rect(s, Inches(6.85), panel_y, Inches(5.85), panel_h, GREEN_DEEP)
    rect(s, Inches(6.85), panel_y, Inches(0.12), panel_h, COPPER_LIGHT)
    t(s, Inches(7.2), Inches(2.55), Inches(5.4), Inches(0.3),
      "USA  ·  DOE NUCLEAR ENTERPRISE  ·  2014-2024",
      size=10, bold=True, color=COPPER_LIGHT)
    t(s, Inches(7.2), Inches(2.90), Inches(5.4), Inches(1.7),
      "22,000 workers.\n59.12 mrem average.",
      size=30, bold=True, color=WHITE, font=TITLE_FONT,
      line_spacing=1.05)
    t(s, Inches(7.2), Inches(4.85), Inches(5.4), Inches(1.1),
      "Every dose tracked.  Every site audited.\n"
      "Every year. For a decade.",
      size=13, color=RGBColor(0xCF, 0xE0, 0xD2),
      line_spacing=1.3, italic=True)
    line_rule(s, Inches(7.2), Inches(5.95), Inches(0.7), COPPER_LIGHT, 1.5)
    t(s, Inches(7.2), Inches(6.05), Inches(5.4), Inches(0.4),
      "INDUSTRIAL  MASTERY",
      size=12, bold=True, color=COPPER_LIGHT)

    # Kicker
    t(s, Inches(0.6), Inches(6.85), Inches(12.1), Inches(0.3),
      "The gap is not the science. It is the system.",
      size=13, bold=True, color=GREEN_DEEP, font=TITLE_FONT,
      italic=True, align=PP_ALIGN.CENTER)

    footer(s, 3, 8, [CITATIONS["siavonga"], CITATIONS["doe"]])
    notes(s,
        "[The Hook] Picture two scenes. Siavonga: uranium in the "
        "water, uranium in the dust, no instruments — radiological "
        "fear. The U.S. nuclear enterprise: twenty-two thousand workers, "
        "an average of 59.12 mrem a year, every dose tracked, every "
        "year, for a decade — industrial mastery. Same element. Same "
        "physics. Two completely different outcomes. The variable is "
        "the system around it. The Radiation Protection Authority "
        "already builds part of that system in Zambia. We just have "
        "not finished it.")


# -----------------------------------------------------------------------
# SLIDE 4 — THE HOSTAGE MATH (drought cost ledger + energy mix)
# -----------------------------------------------------------------------
def slide_hostage(prs):
    s = prs.slides.add_slide(prs.slide_invariants if False else prs.slide_layouts[6])
    set_bg(s, PARCH)
    header_block(s,
                 "02  ·  THE HOSTAGE MATH",
                 "Zambia's economy is hostage to the rain.",
                 "Eight in every ten kilowatts depend on rainfall we "
                 "no longer get.")

    # CHART — Drought cost ledger (3 dry-year deficits in MW)
    chart_data = CategoryChartData()
    chart_data.categories = ["2015/16", "2018/19", "2023/24"]
    chart_data.add_series("MW lost to drought", (1000, 872, 1430))
    cx, cy, cw, ch = Inches(0.6), Inches(2.30), Inches(7.4), Inches(4.5)
    chart_shape = s.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, cx, cy, cw, ch, chart_data
    )
    chart = chart_shape.chart
    chart.has_title = False
    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(13)
    dl.font.bold = True
    dl.font.color.rgb = CRISIS_DEEP
    dl.font.name = BODY_FONT
    color_points(chart.series[0],
                 [CRISIS, CRISIS, CRISIS_DEEP])
    style_chart(chart, label_color=CRISIS_DEEP, axis_color=INK_SOFT)

    t(s, Inches(0.6), Inches(6.85), Inches(7.4), Inches(0.3),
      "Generation deficit, megawatts (rounded). "
      "Sources: ZESCO; MoE.",
      size=9, italic=True, color=MUTED, align=PP_ALIGN.CENTER)

    # RIGHT — three brutal numbers
    rect(s, Inches(8.25), Inches(2.30), Inches(4.55), Inches(4.5),
         OFFWHITE)
    rect(s, Inches(8.25), Inches(2.30), Inches(0.12), Inches(4.5),
         GREEN_DEEP)

    t(s, Inches(8.50), Inches(2.45), Inches(4.1), Inches(0.3),
      "WHAT DROUGHT COSTS US",
      size=10, bold=True, color=COPPER)

    rows = [
        ("84%",   "of generation",     "from hydropower"),
        ("21 hrs", "of load-shedding", "per day, late 2024"),
        ("$440M", "in emergency",       "imports — 2015/16 alone"),
        ("~2%",   "of GDP lost",        "to 2024 blackouts (IMF)"),
    ]
    y = 2.85
    for big, mid, sm in rows:
        t(s, Inches(8.50), Inches(y), Inches(4.1), Inches(0.55),
          big, size=26, bold=True, color=CRISIS_DEEP,
          font=TITLE_FONT)
        t(s, Inches(8.50), Inches(y + 0.55), Inches(4.1), Inches(0.32),
          f"{mid}  {sm}", size=10, color=INK,
          line_spacing=1.25)
        line_rule(s, Inches(8.50), Inches(y + 0.92), Inches(3.8),
                  RULE, 0.5)
        y += 1.00

    # Kicker
    t(s, Inches(0.6), Inches(7.18), Inches(7.4), Inches(0.3),
      "Three droughts in a decade. The trend, not the exception.",
      size=11, italic=True, bold=True, color=CRISIS_DEEP,
      align=PP_ALIGN.CENTER)

    footer(s, 4, 8,
           [CITATIONS["zesco"], CITATIONS["imf"], CITATIONS["ndp"]])
    notes(s,
        "[Hostage Math] Eighty-four percent of our electricity comes "
        "from hydropower. That single point of failure has now cost us "
        "three major drought events in a decade — each bigger than the "
        "last. The 2023-24 Kariba crisis triggered up to twenty-one "
        "hours of load-shedding a day in late 2024, throttling our "
        "copper mines and, by IMF estimates, evaporating about two "
        "percent of GDP. We cannot industrialise a country on rainfall. "
        "Nuclear is not an option here; it is a necessity.")


# -----------------------------------------------------------------------
# SLIDE 5 — SAFETY MASTERY (DOE proof)
# -----------------------------------------------------------------------
def slide_safety(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, PARCH)
    header_block(s,
                 "03  ·  SAFETY MASTERY",
                 "Nuclear is the safest energy we don't yet use.",
                 "A decade of U.S. DOE data — the playbook Zambia "
                 "would inherit.")

    # CHART — horizontal bar exposure comparison
    chart_data = CategoryChartData()
    chart_data.categories = [
        "Avg. Technician (DOE 2024)",
        "Natural Background",
        "Federal Occupational Limit",
    ]
    chart_data.add_series("mrem / year", (59.12, 300, 5000))
    cx, cy, cw, ch = Inches(0.6), Inches(2.30), Inches(7.4), Inches(4.45)
    cs = s.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, cx, cy, cw, ch, chart_data
    )
    chart = cs.chart
    chart.has_title = False
    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(12)
    dl.font.bold = True
    dl.font.color.rgb = GREEN_DEEP
    dl.font.name = BODY_FONT
    color_points(chart.series[0],
                 [GREEN, GOLD, CRISIS])
    style_chart(chart, label_color=GREEN_DEEP, axis_color=INK_SOFT)

    t(s, Inches(0.6), Inches(6.85), Inches(7.4), Inches(0.3),
      "Annual radiation exposure (mrem). Source: U.S. DOE REMS 2024.",
      size=9, italic=True, color=MUTED, align=PP_ALIGN.CENTER)

    # RIGHT — Hero stat + proofs
    rect(s, Inches(8.25), Inches(2.30), Inches(4.55), Inches(1.95),
         GREEN_DEEP)
    rect(s, Inches(8.25), Inches(2.30), Inches(0.12), Inches(1.95),
         COPPER)
    t(s, Inches(8.50), Inches(2.42), Inches(4.2), Inches(0.3),
      "OF THE FEDERAL LIMIT",
      size=10, bold=True, color=GOLD)
    t(s, Inches(8.50), Inches(2.72), Inches(4.2), Inches(1.0),
      "1.2%", size=72, bold=True, color=WHITE, font=TITLE_FONT)
    t(s, Inches(8.50), Inches(3.78), Inches(4.2), Inches(0.3),
      "59.12 ÷ 5,000 mrem",
      size=11, color=RGBColor(0xCF, 0xE0, 0xD2))

    proofs = [
        ("SCALABLE",
         "Workforce grew 1,133 → 1,167. "
         "Average dose did not move."),
        ("CONTAINED",
         "98% external (shielding). "
         "2% internal (protocol)."),
        ("REPLICABLE",
         "IAEA Milestones — the regulatory path "
         "Zambia (RPA → NSPA) is already on."),
    ]
    y = 4.40
    for head, body in proofs:
        rect(s, Inches(8.25), Inches(y), Inches(4.55), Inches(0.74),
             OFFWHITE)
        rect(s, Inches(8.25), Inches(y), Inches(0.08), Inches(0.74),
             COPPER)
        t(s, Inches(8.45), Inches(y + 0.06), Inches(4.3), Inches(0.25),
          head, size=10, bold=True, color=GREEN_DEEP)
        t(s, Inches(8.45), Inches(y + 0.32), Inches(4.3), Inches(0.42),
          body, size=10, color=INK, line_spacing=1.20)
        y += 0.80

    footer(s, 5, 8, [CITATIONS["doe"], CITATIONS["iaea"]])
    notes(s,
        "[Safety Mastery] Zambians fear nuclear the way most people "
        "fear flying — disproportionate to the data. My analysis of "
        "the U.S. Department of Energy REMS dataset shows the average "
        "nuclear technician absorbs 59.12 mrem per year. That is one "
        "point two percent of the federal occupational limit and "
        "roughly one fifth of the natural background dose any of us "
        "absorb just walking around. The biggest fear in Zambia is "
        "safety; the data says modern monitoring keeps workers safer "
        "than most office radiation environments. And the system "
        "scales — workforce grew while individual dose stayed flat. "
        "This is the analytical capacity we have already built locally; "
        "we are ready to oversee a national programme.")


# -----------------------------------------------------------------------
# SLIDE 6 — STRATEGIC RESOURCE
# -----------------------------------------------------------------------
def slide_resource(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, PARCH)
    header_block(s,
                 "04  ·  STRATEGIC RESOURCE",
                 "We are exporting our future.",
                 "Zambian uranium fuels other countries' reactors. "
                 "We import the electricity it makes.")

    # LEFT — Deposit map proxy
    rect(s, Inches(0.6), Inches(2.30), Inches(6.5), Inches(4.5),
         GREEN_DEEP)
    rect(s, Inches(0.6), Inches(2.30), Inches(6.5), Inches(0.10),
         COPPER)
    t(s, Inches(0.85), Inches(2.50), Inches(6.0), Inches(0.3),
      "KNOWN URANIUM-BEARING ZONES  ·  ZAMBIA",
      size=10, bold=True, color=GOLD)

    deposits = [
        ("LUMWANA  /  KANYEMBA",
         "North-Western Province",
         "Uranium recovered as by-product of copper mining; "
         "stockpiled historically by Barrick."),
        ("GWEMBE  /  SIAVONGA BELT",
         "Southern Province",
         "Karoo-type sandstone uranium — site of the 2024 "
         "Siavonga Baseline Survey."),
        ("MUTANGA  /  DIBWE",
         "Southern Province",
         "Defined uranium resource under historical "
         "GoviEx / African Energy permits."),
    ]
    y = 2.92
    for name, region, body in deposits:
        rect(s, Inches(0.85), Inches(y), Inches(6.0), Inches(1.10),
             RGBColor(0x10, 0x36, 0x24))
        rect(s, Inches(0.85), Inches(y), Inches(0.08), Inches(1.10),
             COPPER)
        t(s, Inches(1.05), Inches(y + 0.08), Inches(5.8), Inches(0.3),
          name, size=13, bold=True, color=WHITE, font=TITLE_FONT)
        t(s, Inches(1.05), Inches(y + 0.36), Inches(5.8), Inches(0.3),
          region, size=10, italic=True, color=GOLD)
        t(s, Inches(1.05), Inches(y + 0.60), Inches(5.8), Inches(0.5),
          body, size=11, color=RGBColor(0xD5, 0xE0, 0xD5),
          line_spacing=1.3)
        y += 1.22

    # RIGHT — value-chain narrative, editorial style
    t(s, Inches(7.4), Inches(2.45), Inches(5.4), Inches(0.4),
      "THE  VALUE-CHAIN  PIVOT",
      size=10, bold=True, color=COPPER)
    line_rule(s, Inches(7.4), Inches(2.78), Inches(0.6), COPPER, 1.5)

    blocks = [
        ("TODAY", CRISIS,
         "Raw-mineral exporter. Uranium leaves Zambia for "
         "enrichment abroad; we capture wellhead value only."),
        ("WITH NUCLEAR POWER", GREEN_DEEP,
         "High-tech energy producer. Domestic fuel feeds domestic "
         "baseload — mines, smelters, hospitals, factories."),
        ("STRATEGIC UPSIDE", COPPER,
         "Import independence. Price-stable electricity. High-skill "
         "jobs anchored to a regulated, exportable competency."),
    ]
    y2 = 3.00
    for head, col, body in blocks:
        t(s, Inches(7.4), Inches(y2), Inches(5.4), Inches(0.3),
          head, size=11, bold=True, color=col)
        t(s, Inches(7.4), Inches(y2 + 0.30), Inches(5.4), Inches(1.1),
          body, size=11.5, color=INK, line_spacing=1.32)
        y2 += 1.30

    footer(s, 6, 8,
           [CITATIONS["mines"], CITATIONS["wna"], CITATIONS["ndp"]])
    notes(s,
        "[Strategic Resource] Zambia exports uranium today. We are a "
        "raw-mineral economy. Why should we ship the very element that "
        "could fuel our own copper mines abroad — only to import the "
        "electricity it produces? Lumwana, Gwembe, Mutanga: known "
        "uranium-bearing zones. By developing nuclear power we move "
        "from raw-material exporter to high-tech energy producer; we "
        "anchor high-skill jobs; we capture the value chain. Why ship "
        "our uranium when we could use it to power our own factories?")


# -----------------------------------------------------------------------
# SLIDE 7 — ECONOMIC IMPACT (Copper goal + African gap)
# -----------------------------------------------------------------------
def slide_economic(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, PARCH)
    header_block(s,
                 "05  ·  ECONOMIC IMPACT",
                 "3 million tonnes of copper need a power source. "
                 "Rain isn't it.",
                 "Zambia's flagship industrial target needs "
                 "high-density, weather-proof power — now.")

    # CHART — column: MW continuous at copper output levels
    chart_data = CategoryChartData()
    chart_data.categories = [
        "Today (~0.8 Mt)",
        "Mid-Path (~1.5 Mt)",
        "Target (3.0 Mt)",
    ]
    chart_data.add_series("MW continuous (mining only)",
                          (320, 600, 1200))
    cx, cy, cw, ch = Inches(0.6), Inches(2.30), Inches(7.4), Inches(3.30)
    cs = s.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, cx, cy, cw, ch, chart_data
    )
    chart = cs.chart
    chart.has_title = False
    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(13)
    dl.font.bold = True
    dl.font.color.rgb = GREEN_DEEP
    dl.font.name = BODY_FONT
    color_points(chart.series[0],
                 [GOLD, COPPER, CRISIS])
    style_chart(chart, label_color=GREEN_DEEP, axis_color=INK_SOFT)

    t(s, Inches(0.6), Inches(5.65), Inches(7.4), Inches(0.3),
      "Continuous power for copper mining at ~3,500 kWh/t. "
      "Sources: MoMMD; ZESCO.",
      size=9, italic=True, color=MUTED, align=PP_ALIGN.CENTER)

    # AFRICAN COMPETITION STRIP (full width below chart)
    rect(s, Inches(0.6), Inches(6.05), Inches(7.4), Inches(1.05),
         GREEN_DEEP)
    rect(s, Inches(0.6), Inches(6.05), Inches(0.10), Inches(1.05),
         COPPER)
    t(s, Inches(0.85), Inches(6.15), Inches(7.0), Inches(0.3),
      "MEANWHILE,  ACROSS  AFRICA",
      size=10, bold=True, color=GOLD)
    textbox(
        s, Inches(0.85), Inches(6.45), Inches(7.0), Inches(0.6),
        [
            {"text": "Egypt: ", "size": 11, "bold": True,
             "color": COPPER_LIGHT},
            {"text": "4 reactors under construction at El Dabaa.  ",
             "size": 11, "color": WHITE},
            {"text": "Kenya: ", "size": 11, "bold": True,
             "color": COPPER_LIGHT},
            {"text": "first criticality target 2034.  ",
             "size": 11, "color": WHITE},
            {"text": "South Africa: ", "size": 11, "bold": True,
             "color": COPPER_LIGHT},
            {"text": "Koeberg, 40+ years.",
             "size": 11, "color": WHITE},
        ],
        line_spacing=1.3,
    )

    # RIGHT — implications
    rect(s, Inches(8.25), Inches(2.30), Inches(4.55), Inches(4.8),
         OFFWHITE)
    rect(s, Inches(8.25), Inches(2.30), Inches(0.12), Inches(4.8),
         COPPER)
    t(s, Inches(8.50), Inches(2.45), Inches(4.1), Inches(0.3),
      "WHAT THIS MEANS",
      size=10, bold=True, color=COPPER)

    pts = [
        ("MINING INTENSITY",
         "Deep-level copper needs ~1,200 MW continuous at full "
         "3 Mt output — hydropower cannot deliver this reliably."),
        ("COST STABILITY",
         "Nuclear fuel costs are decade-stable. Drought imports cost "
         "us $0.50+/kWh in emergency buys."),
        ("STRATEGIC FIT",
         "Zambia's 2,000 MW nuclear target maps almost exactly onto "
         "the 3 Mt copper power gap."),
        ("THE WINDOW",
         "Egypt's first reactor enters service ~2026-28. If we wait, "
         "we buy that power from them."),
    ]
    y = 2.85
    for head, body in pts:
        t(s, Inches(8.50), Inches(y), Inches(4.1), Inches(0.3),
          head, size=10, bold=True, color=GREEN_DEEP)
        t(s, Inches(8.50), Inches(y + 0.28), Inches(4.1), Inches(0.85),
          body, size=10.5, color=INK, line_spacing=1.30)
        y += 1.05

    footer(s, 7, 8,
           [CITATIONS["mines"], CITATIONS["wna"], CITATIONS["zesco"]])
    notes(s,
        "[Economic Impact] Zambia has set a public target of three "
        "million tonnes of copper annually. At industry-standard "
        "energy intensity that is roughly twelve hundred megawatts of "
        "continuous power — for mining alone. Hydropower cannot "
        "deliver that drought after drought. Nuclear can. And there is "
        "a window closing: Egypt is building four reactors at El "
        "Dabaa; Kenya targets first criticality by 2034; South Africa "
        "has been nuclear for over forty years. If we wait, we end up "
        "buying drought-proof electricity from neighbours. This is "
        "not ideology. It is arithmetic.")


# -----------------------------------------------------------------------
# SLIDE 8 — ROADMAP / ASK
# -----------------------------------------------------------------------
def slide_ask(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, GREEN_DEEP)

    rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.10), COPPER)
    rect(s, Inches(0), Inches(0.10), Inches(13.333), Inches(0.04), GOLD)

    t(s, Inches(0.6), Inches(0.50), Inches(10), Inches(0.3),
      "06  ·  THE ASK",
      size=11, bold=True, color=GOLD)
    t(s, Inches(0.6), Inches(0.80), Inches(12.1), Inches(0.9),
      "Three decisions. One generation.",
      size=36, bold=True, color=WHITE, font=TITLE_FONT)
    line_rule(s, Inches(0.6), Inches(1.70), Inches(0.8), COPPER, 2.0)
    t(s, Inches(0.6), Inches(1.85), Inches(12.1), Inches(0.4),
      "From radiological fear to industrial mastery — "
      "the choice is legislative, not technical.",
      size=13, italic=True,
      color=RGBColor(0xCF, 0xE0, 0xD2))

    actions = [
        ("01", "ENACT",
         "Pass the Nuclear Bill — empower NSPA as an IAEA-aligned "
         "independent regulator. Without this signature, every other "
         "step is symbolic."),
        ("02", "EQUIP",
         "Fund radiological monitoring nationwide — starting in "
         "Siavonga. No Zambian community should live next to "
         "uranium without instruments."),
        ("03", "EXECUTE",
         "Greenlight CNST and SMR pre-feasibility — partner with "
         "ROSATOM, NuScale, or Westinghouse using the agreements "
         "Cabinet already approved in 2020."),
    ]
    ay = 2.65
    ah = 1.20
    gap = 0.10
    for i, (num, label, body) in enumerate(actions):
        top = ay + i * (ah + gap)
        # Number block
        rect(s, Inches(0.6), Inches(top), Inches(1.15), Inches(ah),
             COPPER)
        t(s, Inches(0.6), Inches(top + 0.30), Inches(1.15),
          Inches(0.7),
          num, size=36, bold=True, color=GREEN_DEEP,
          font=TITLE_FONT, align=PP_ALIGN.CENTER)
        # Content block
        rect(s, Inches(1.75), Inches(top), Inches(11.0), Inches(ah),
             RGBColor(0x0C, 0x30, 0x1F))
        rect(s, Inches(1.75), Inches(top), Inches(11.0), Inches(0.04),
             GOLD)
        t(s, Inches(2.0), Inches(top + 0.12), Inches(10.5),
          Inches(0.35),
          label, size=18, bold=True, color=GOLD,
          font=TITLE_FONT)
        t(s, Inches(2.0), Inches(top + 0.50), Inches(10.5),
          Inches(0.65),
          body, size=12, color=WHITE, line_spacing=1.30)

    # Closing line
    line_rule(s, Inches(0.6), Inches(6.65), Inches(12.1),
              RGBColor(0x33, 0x55, 0x42), 0.5)
    textbox(
        s, Inches(0.6), Inches(6.75), Inches(12.1), Inches(0.45),
        [
            {"text": "The fuel is under our feet. ",
             "size": 14, "color": WHITE, "font": TITLE_FONT,
             "italic": True},
            {"text": "The expertise is on our payroll. ",
             "size": 14, "color": WHITE, "font": TITLE_FONT,
             "italic": True},
            {"text": "Sign the bill.",
             "size": 16, "bold": True, "color": GOLD,
             "font": TITLE_FONT},
        ],
        align=PP_ALIGN.CENTER,
    )
    t(s, Inches(0.6), Inches(7.20), Inches(12.1), Inches(0.25),
      "8 / 8   ·   E. MWABA   ·   SOURCES: "
      "DOE REMS 2024 · IAEA Milestones · 7NDP & Vision 2030 · "
      "ZESCO · MoMMD · Siavonga 2024 · IMF · WNA",
      size=8, italic=True,
      color=RGBColor(0x8B, 0x97, 0x90),
      align=PP_ALIGN.CENTER)

    notes(s,
        "[The Ask] Three decisions. ENACT: pass the Nuclear Bill, "
        "formalise NSPA as IAEA-aligned regulator. EQUIP: fund "
        "radiological monitoring nationwide, starting with the very "
        "communities in Siavonga that prompted the 2024 Baseline "
        "Survey. EXECUTE: greenlight the Centre for Nuclear Science "
        "and Technology and SMR pre-feasibility using ROSATOM, IP3, "
        "KAERI agreements already on file. The fuel is under our "
        "feet. The expertise is on our payroll. The bill is on the "
        "table. Sign it. Thank you.")


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide_cover(prs)
    slide_story(prs)
    slide_hook(prs)
    slide_hostage(prs)
    slide_safety(prs)
    slide_resource(prs)
    slide_economic(prs)
    slide_ask(prs)

    prs.save(OUTPUT)
    print(f"Generated: {OUTPUT}")


if __name__ == "__main__":
    build()
