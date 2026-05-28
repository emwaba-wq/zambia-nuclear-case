"""
build_presentation.py
Generates: The Case for Nuclear Power in Zambia (8-slide .pptx).
Built for executive / supervisor-grade review.

Structure:
  1. Title + Big Idea
  2. The 3-Minute Story (single paragraph)
  3. The Hook — Siavonga 'Invisible Danger' vs. U.S. 'Precision Monitoring'
  4. The Vision — Nuclear as Zambia's Baseload Future
  5. Safety Mastery — 59.12 mrem proof (DOE REMS 2024)
  6. Strategic Resource — Zambia's Uranium Potential
  7. Economic Impact — Powering the 3-Million-Tonne Copper Goal
  8. Roadmap — Next Steps

Run:
  pip install -r requirements.txt
  python build_presentation.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

OUTPUT = "The_Case_for_Nuclear_Power_in_Zambia.pptx"

NAVY = RGBColor(0x0A, 0x1F, 0x3D)
NAVY_DEEP = RGBColor(0x06, 0x14, 0x2A)
COPPER = RGBColor(0xCC, 0x7A, 0x2E)
COPPER_LIGHT = RGBColor(0xE8, 0xA8, 0x60)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFFWHITE = RGBColor(0xF7, 0xF5, 0xF0)
INK = RGBColor(0x12, 0x18, 0x24)
MUTED = RGBColor(0x6B, 0x72, 0x80)
CRISIS = RGBColor(0xB3, 0x2B, 0x2B)
CRISIS_DEEP = RGBColor(0x7C, 0x18, 0x18)
SLATE = RGBColor(0x2A, 0x33, 0x44)
RULE = RGBColor(0xC8, 0xCE, 0xD6)
SUCCESS = RGBColor(0x1B, 0x5E, 0x3A)
SOLAR = RGBColor(0xE8, 0xA8, 0x60)
COAL = RGBColor(0x55, 0x5B, 0x66)
HFO = RGBColor(0x99, 0x6B, 0x3A)

CITATIONS = {
    "doe": "U.S. DOE REMS 2024 Report",
    "ndp": "Zambia 7NDP & Vision 2030",
    "iaea": "IAEA Milestone Approach",
    "mines": "Zambia Ministry of Mines — Annual Mining Statistics",
    "zesco": "ZESCO Annual Reports 2023-2024",
    "siavonga": "Siavonga Uranium Baseline Survey (2024)",
}

BIG_IDEA = (
    "Zambia must convert its domestic uranium reserves and existing "
    "radiological safety expertise into nuclear baseload power — because "
    "every drought year proves that hydropower alone is no longer an "
    "energy strategy, it is a gamble against a climate we can no longer "
    "afford to lose."
)

THREE_MIN_STORY = (
    "Imagine a country that owns its uranium, runs a Radiation Protection "
    "Authority, has a Cabinet-approved Nuclear Policy, and has signed "
    "cooperation agreements with ROSATOM, the U.S. IP3 Allied Nuclear "
    "Partners, and South Korea's KAERI — and yet, in 2026, still depends "
    "on a single river system for eighty-four percent of its electricity. "
    "That country is Zambia. Every drought repeats the same story: the "
    "2015/16 dry year cost us roughly one thousand megawatts of "
    "generation and four hundred and forty million U.S. dollars in "
    "unbudgeted imports; the 2023–2024 Kariba crisis triggered months of "
    "load-shedding that stalled the very copper mines that fund the "
    "national budget. Meanwhile, in Siavonga, the 2024 Uranium Baseline "
    "Survey detected unmonitored uranium in water and household dust — "
    "communities living beside the resource itself, without instruments, "
    "without a baseline, without accountability. That is radiological "
    "fear. Contrast that with the United States, where my analysis of "
    "the Department of Energy Radiation Exposure Monitoring System "
    "tracks twenty-two thousand workers and shows an average individual "
    "dose of just fifty-nine point one two millirem per year — about one "
    "point two percent of the federal occupational limit, with workforce "
    "growth producing zero increase in individual exposure. That is "
    "industrial mastery. Same element, two completely different "
    "outcomes; the variable is the system. Zambia already participates "
    "in that system through the IAEA Milestone Approach. We have the "
    "uranium. We have the safety institution. We have the policy. We "
    "have the partners. And we have a three-million-tonne copper "
    "ambition that no hydropower fleet on this continent can power "
    "reliably. The missing piece is legislative: pass the Nuclear Bill, "
    "fund radiological monitoring nationwide starting in Siavonga, and "
    "greenlight Small Modular Reactor pre-feasibility. The fuel is "
    "under our feet. The expertise is on our payroll. The bill is on "
    "the table. Sign it."
)


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
        run.font.name = r.get("font", "Calibri")
        run.font.size = Pt(r["size"])
        run.font.bold = r.get("bold", False)
        run.font.italic = r.get("italic", False)
        run.font.color.rgb = r.get("color", INK)
    return tb


def simple_text(slide, left, top, width, height, text, *,
                size=16, bold=False, italic=False, color=INK,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                font="Calibri", line_spacing=1.2):
    return textbox(slide, left, top, width, height,
                   [{"text": text, "size": size, "bold": bold,
                     "italic": italic, "color": color, "font": font}],
                   align=align, anchor=anchor, line_spacing=line_spacing)


def eyebrow(slide, left, top, text, color=COPPER):
    return simple_text(slide, left, top, Inches(10), Inches(0.3),
                       text, size=11, bold=True, color=color)


def slide_header(slide, eyebrow_text, title, kicker=None):
    eyebrow(slide, Inches(0.6), Inches(0.45), eyebrow_text)
    simple_text(slide, Inches(0.6), Inches(0.75), Inches(12.1),
                Inches(0.75),
                title, size=30, bold=True, color=NAVY)
    line_rule(slide, Inches(0.6), Inches(1.45), Inches(1.2), COPPER, 1.5)
    if kicker:
        simple_text(slide, Inches(0.6), Inches(1.55), Inches(12.1),
                    Inches(0.4),
                    kicker, size=13, italic=True, color=MUTED)


def slide_footer(slide, n, total, citations):
    line_rule(slide, Inches(0), Inches(7.25), Inches(13.333),
              RULE, 0.5)
    rect(slide, Inches(0), Inches(7.25), Inches(0.7), Inches(0.05),
         COPPER)
    simple_text(slide, Inches(0.5), Inches(7.30), Inches(10.5),
                Inches(0.20),
                "Sources: " + " · ".join(citations),
                size=8, italic=True, color=MUTED)
    simple_text(slide, Inches(11.4), Inches(7.30), Inches(1.7),
                Inches(0.20),
                f"{n} / {total}  ·  E. Mwaba",
                size=8, bold=True, color=SLATE, align=PP_ALIGN.RIGHT)


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def style_chart_text(chart, color=INK, size=9):
    """Apply consistent text formatting across all chart text elements."""
    try:
        for plot in chart.plots:
            if plot.has_data_labels:
                dl = plot.data_labels
                dl.font.size = Pt(size)
                dl.font.color.rgb = color
                dl.font.name = "Calibri"
    except Exception:
        pass
    try:
        for axis in (chart.category_axis, chart.value_axis):
            tf = axis.tick_labels
            tf.font.size = Pt(size)
            tf.font.color.rgb = color
            tf.font.name = "Calibri"
    except Exception:
        pass


def color_series_points(series, colors):
    for i, pt in enumerate(series.points):
        if i >= len(colors):
            break
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = colors[i]
        pt.format.line.fill.background()


# -----------------------------------------------------------------------
# SLIDE 1 — TITLE + BIG IDEA
# -----------------------------------------------------------------------
def slide_title(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, NAVY_DEEP)

    rect(s, Inches(0.6), Inches(0.6), Inches(0.08), Inches(0.45), COPPER)
    simple_text(s, Inches(0.85), Inches(0.58), Inches(10), Inches(0.4),
                "ENERGY SECURITY  ·  REPUBLIC OF ZAMBIA  ·  2026",
                size=12, bold=True, color=COPPER_LIGHT)

    textbox(
        s, Inches(0.6), Inches(1.55), Inches(12.1), Inches(2.5),
        [
            {"text": "We already own", "size": 60, "bold": True,
             "color": WHITE},
            {"text": "\nthe solution.", "size": 60, "bold": True,
             "color": COPPER_LIGHT, "space_after": 0},
        ],
        line_spacing=1.0,
    )

    line_rule(s, Inches(0.6), Inches(4.15), Inches(2.5), COPPER, 1.5)

    simple_text(s, Inches(0.6), Inches(4.30), Inches(12.1), Inches(0.5),
                "The Case for Nuclear Power in Zambia",
                size=24, color=WHITE)

    # BIG IDEA PANEL
    rect(s, Inches(0.6), Inches(5.10), Inches(12.1), Inches(1.50),
         RGBColor(0x12, 0x28, 0x4A))
    rect(s, Inches(0.6), Inches(5.10), Inches(0.10), Inches(1.50),
         COPPER)
    simple_text(s, Inches(0.85), Inches(5.20), Inches(11.7),
                Inches(0.3),
                "THE  BIG  IDEA",
                size=11, bold=True, color=COPPER_LIGHT)
    simple_text(s, Inches(0.85), Inches(5.50), Inches(11.7),
                Inches(1.05),
                BIG_IDEA,
                size=14, italic=True, color=WHITE, line_spacing=1.3)

    simple_text(s, Inches(0.6), Inches(6.85), Inches(8), Inches(0.25),
                "Prepared by  ELIZABETH MWABA",
                size=10, bold=True, color=WHITE)
    simple_text(s, Inches(0.6), Inches(7.10), Inches(8), Inches(0.20),
                "Radiological Analyst  ·  Supervisor Review Brief",
                size=8, color=RGBColor(0x99, 0xA3, 0xB3))

    set_notes(s,
        "[Opening] Three minutes. One argument. The case I'm making is "
        "captured in one sentence — Zambia must convert its domestic "
        "uranium and its existing radiological safety expertise into "
        "nuclear baseload power, because hydropower alone has stopped "
        "being a strategy and become a climate gamble we cannot keep "
        "paying for.")


# -----------------------------------------------------------------------
# SLIDE 2 — THE 3-MINUTE STORY
# -----------------------------------------------------------------------
def slide_story(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, OFFWHITE)
    slide_header(s,
                 "PROLOGUE  ·  THE 3-MINUTE STORY",
                 "If I had no slides, this is what I would say.",
                 "One paragraph. Three minutes. The full argument.")

    rect(s, Inches(0.6), Inches(2.15), Inches(12.1), Inches(4.85),
         WHITE)
    rect(s, Inches(0.6), Inches(2.15), Inches(0.12), Inches(4.85),
         COPPER)

    simple_text(s, Inches(0.95), Inches(2.30), Inches(11.5),
                Inches(0.3),
                "DELIVERED AT EXECUTIVE CADENCE  ·  ~360 WORDS",
                size=10, bold=True, color=COPPER)

    simple_text(s, Inches(0.95), Inches(2.65), Inches(11.5),
                Inches(4.30),
                THREE_MIN_STORY,
                size=11, color=INK, line_spacing=1.32)

    slide_footer(s, 2, 8,
                 [CITATIONS["doe"], CITATIONS["zesco"],
                  CITATIONS["siavonga"], CITATIONS["ndp"]])
    set_notes(s,
        "[Use this slide if asked 'give me the pitch in one breath.' "
        "Read or paraphrase the paragraph as drafted; it lands on the "
        "final imperative: Sign it.]")


# -----------------------------------------------------------------------
# SLIDE 3 — THE HOOK
# -----------------------------------------------------------------------
def slide_hook(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, OFFWHITE)
    slide_header(s,
                 "01  ·  THE HOOK",
                 "Same element. Two realities.",
                 "Zambians and Americans both live with uranium. "
                 "Only one nation measures it.")

    panel_y = Inches(2.30)
    panel_h = Inches(4.30)

    rect(s, Inches(0.6), panel_y, Inches(6.05), panel_h, CRISIS_DEEP)
    rect(s, Inches(0.6), panel_y, Inches(0.12), panel_h, CRISIS)
    simple_text(s, Inches(0.85), Inches(2.45), Inches(5.6), Inches(0.3),
                "ZAMBIA  ·  SIAVONGA  ·  2024",
                size=10, bold=True, color=COPPER_LIGHT)
    simple_text(s, Inches(0.85), Inches(2.80), Inches(5.6), Inches(1.6),
                "Uranium in the water.\nUranium in the dust.",
                size=28, bold=True, color=WHITE, line_spacing=1.05)
    simple_text(s, Inches(0.85), Inches(4.65), Inches(5.6), Inches(1.1),
                "Zero monitoring.  Zero baseline.\n"
                "Zero accountability.",
                size=13, color=RGBColor(0xF2, 0xCF, 0xCF),
                line_spacing=1.3)
    line_rule(s, Inches(0.85), Inches(5.85), Inches(1.0), COPPER, 1.0)
    simple_text(s, Inches(0.85), Inches(5.95), Inches(5.6), Inches(0.5),
                "RADIOLOGICAL  FEAR",
                size=14, bold=True, color=COPPER_LIGHT)

    rect(s, Inches(6.85), panel_y, Inches(5.85), panel_h, NAVY)
    rect(s, Inches(6.85), panel_y, Inches(0.12), panel_h, COPPER)
    simple_text(s, Inches(7.1), Inches(2.45), Inches(5.5), Inches(0.3),
                "USA  ·  DOE NUCLEAR ENTERPRISE  ·  2014-2024",
                size=10, bold=True, color=COPPER_LIGHT)
    simple_text(s, Inches(7.1), Inches(2.80), Inches(5.5), Inches(1.6),
                "22,000 workers.\n59.12 mrem average.",
                size=28, bold=True, color=WHITE, line_spacing=1.05)
    simple_text(s, Inches(7.1), Inches(4.65), Inches(5.5), Inches(1.1),
                "Every dose tracked.  Every site audited.\n"
                "Every year. For a decade.",
                size=13, color=RGBColor(0xCF, 0xD8, 0xE3),
                line_spacing=1.3)
    line_rule(s, Inches(7.1), Inches(5.85), Inches(1.0), COPPER, 1.0)
    simple_text(s, Inches(7.1), Inches(5.95), Inches(5.5), Inches(0.5),
                "INDUSTRIAL  MASTERY",
                size=14, bold=True, color=COPPER_LIGHT)

    simple_text(s, Inches(0.6), Inches(6.80), Inches(12.1), Inches(0.4),
                "The gap is not the science. It is the system.",
                size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    set_notes(s,
        "[The Hook] Picture two scenes. Siavonga: uranium in water, "
        "uranium in dust, no instruments — radiological fear. The U.S. "
        "nuclear enterprise: twenty-two thousand workers, 59.12 mrem "
        "average, every dose tracked for a decade — industrial mastery. "
        "Same element, two different outcomes. The variable is the "
        "system, not the uranium. Zambia must close that gap, and the "
        "expertise to do it already lives inside our Radiation "
        "Protection Authority.")
    slide_footer(s, 3, 8,
                 [CITATIONS["siavonga"], CITATIONS["doe"]])


# -----------------------------------------------------------------------
# SLIDE 4 — VISION: NUCLEAR AS BASELOAD (CHART: ENERGY MIX DONUT)
# -----------------------------------------------------------------------
def slide_vision(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, OFFWHITE)
    slide_header(s,
                 "02  ·  THE VISION",
                 "Nuclear as Zambia's baseload future.",
                 "We cannot industrialise a country on rainfall. "
                 "Nuclear is weather-proof power.")

    # CHART: ENERGY MIX DONUT
    chart_data = CategoryChartData()
    chart_data.categories = ["Hydro", "Coal", "Solar", "HFO"]
    chart_data.add_series("2019 Mix", (84, 10, 3, 3))
    cx, cy, cw, ch = Inches(0.6), Inches(2.20), Inches(5.8), Inches(4.6)
    chart_shape = s.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT, cx, cy, cw, ch, chart_data
    )
    chart = chart_shape.chart
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(10)
    chart.legend.font.color.rgb = INK
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.show_percentage = False
    dl.show_value = True
    dl.font.size = Pt(11)
    dl.font.bold = True
    dl.font.color.rgb = WHITE
    color_series_points(chart.series[0],
                        [CRISIS, COAL, SOLAR, HFO])

    # Caption under chart
    simple_text(s, Inches(0.6), Inches(6.75), Inches(5.8), Inches(0.4),
                "Zambia's 2019 generation mix (%). Source: MoE / "
                "ZESCO 2019.",
                size=9, italic=True, color=MUTED, align=PP_ALIGN.CENTER)

    # RIGHT — KEY POINTS
    rect(s, Inches(6.85), Inches(2.20), Inches(5.85), Inches(4.6),
         WHITE)
    rect(s, Inches(6.85), Inches(2.20), Inches(0.12), Inches(4.6),
         COPPER)

    simple_text(s, Inches(7.1), Inches(2.35), Inches(5.4), Inches(0.3),
                "WHY THIS MATTERS", size=11, bold=True, color=COPPER)

    points = [
        ("THE PROBLEM",
         "84% hydropower dependence makes Zambia's economy a "
         "hostage to the rain."),
        ("THE SOLUTION",
         "Nuclear delivers 24/7, weather-proof baseload "
         "independent of Kariba reservoir levels."),
        ("THE READINESS",
         "Zambia is an IAEA Member State and the Radiation "
         "Protection Authority (RPA) is already in place."),
    ]
    py = 2.75
    for head, body in points:
        simple_text(s, Inches(7.1), Inches(py), Inches(5.4),
                    Inches(0.3),
                    head, size=11, bold=True, color=NAVY)
        simple_text(s, Inches(7.1), Inches(py + 0.32), Inches(5.4),
                    Inches(1.05),
                    body, size=12, color=INK, line_spacing=1.3)
        py += 1.35

    set_notes(s,
        "[The Vision] Eighty-four percent of Zambia's electricity comes "
        "from hydropower. That is not diversification — that is a single "
        "point of failure with a weather dependency. Nuclear, by "
        "contrast, runs at roughly ninety percent capacity factor "
        "regardless of rainfall. And Zambia is not starting cold: we "
        "are an IAEA Member State, and our Radiation Protection "
        "Authority has been managing radiological risk for decades. We "
        "cannot industrialise a country on intermittent power. Nuclear "
        "isn't just an option; it is a necessity for a drought-resilient "
        "Zambia.")
    slide_footer(s, 4, 8,
                 [CITATIONS["zesco"], CITATIONS["ndp"],
                  CITATIONS["iaea"]])


# -----------------------------------------------------------------------
# SLIDE 5 — SAFETY MASTERY (CHART: 59.12 vs 5,000 mrem)
# -----------------------------------------------------------------------
def slide_safety(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, OFFWHITE)
    slide_header(s,
                 "03  ·  SAFETY MASTERY",
                 "The proof is in the data.",
                 "10-year DOE record across 22,000 workers — "
                 "the safety model Zambia would inherit.")

    # CHART: HORIZONTAL BAR — exposure comparison
    chart_data = CategoryChartData()
    chart_data.categories = [
        "Avg. Technician (DOE 2024)",
        "Natural Background",
        "Federal Occupational Limit",
    ]
    chart_data.add_series("mrem / year", (59.12, 300, 5000))
    cx, cy, cw, ch = Inches(0.6), Inches(2.20), Inches(7.4), Inches(4.4)
    chart_shape = s.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED, cx, cy, cw, ch, chart_data
    )
    chart = chart_shape.chart
    chart.has_title = False
    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(11)
    dl.font.bold = True
    dl.font.color.rgb = NAVY
    color_series_points(chart.series[0],
                        [SUCCESS, COPPER, CRISIS])
    style_chart_text(chart, color=INK, size=10)

    simple_text(s, Inches(0.6), Inches(6.65), Inches(7.4), Inches(0.4),
                "Radiation exposure (mrem/yr). "
                "Source: U.S. DOE REMS 2024.",
                size=9, italic=True, color=MUTED, align=PP_ALIGN.CENTER)

    # RIGHT — HERO STAT + PROOFS
    rect(s, Inches(8.25), Inches(2.20), Inches(4.55), Inches(2.0),
         NAVY)
    rect(s, Inches(8.25), Inches(2.20), Inches(0.12), Inches(2.0),
         COPPER)
    simple_text(s, Inches(8.50), Inches(2.30), Inches(4.2), Inches(0.3),
                "OF THE FEDERAL LIMIT",
                size=11, bold=True, color=COPPER_LIGHT)
    simple_text(s, Inches(8.50), Inches(2.65), Inches(4.2), Inches(1.1),
                "1.2%", size=72, bold=True, color=WHITE)
    simple_text(s, Inches(8.50), Inches(3.75), Inches(4.2), Inches(0.4),
                "59.12 ÷ 5,000 mrem",
                size=12, color=RGBColor(0xCF, 0xD8, 0xE3))

    proofs = [
        ("SCALABLE",
         "Workforce grew 1,133 → 1,167. "
         "Dose did not move."),
        ("CONTAINED",
         "98% external (shielding). "
         "2% internal (protocol)."),
        ("REPLICABLE",
         "IAEA Milestones — the same path Zambia (RPA → NSPA) is on."),
    ]
    py = 4.35
    for head, body in proofs:
        rect(s, Inches(8.25), Inches(py), Inches(4.55), Inches(0.74),
             WHITE)
        rect(s, Inches(8.25), Inches(py), Inches(0.08), Inches(0.74),
             COPPER)
        simple_text(s, Inches(8.45), Inches(py + 0.05), Inches(4.3),
                    Inches(0.25), head,
                    size=10, bold=True, color=NAVY)
        simple_text(s, Inches(8.45), Inches(py + 0.30), Inches(4.3),
                    Inches(0.42), body,
                    size=10, color=INK, line_spacing=1.20)
        py += 0.80

    set_notes(s,
        "[Safety Mastery] My analysis of the U.S. DOE Radiation "
        "Exposure Monitoring System shows the average nuclear technician "
        "absorbs 59.12 mrem per year — that is one point two percent of "
        "the five thousand mrem federal limit. Less than the dose from "
        "routine medical imaging. The biggest fear in Zambia about "
        "nuclear is safety; this data says modern monitoring keeps "
        "workers safer than most office radiation environments. And "
        "between 2021 and 2024 the workforce grew while individual dose "
        "stayed flat. Safety scales. We have the analytical tools — "
        "this dashboard — to oversee a national programme today.")
    slide_footer(s, 5, 8,
                 [CITATIONS["doe"], CITATIONS["iaea"]])


# -----------------------------------------------------------------------
# SLIDE 6 — STRATEGIC RESOURCE: ZAMBIA'S URANIUM
# -----------------------------------------------------------------------
def slide_resource(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, OFFWHITE)
    slide_header(s,
                 "04  ·  STRATEGIC RESOURCE",
                 "We are sitting on our own fuel.",
                 "Zambia exports raw uranium today. "
                 "It can fuel domestic reactors tomorrow.")

    # LEFT — Deposit cards (visual proxy for a map)
    rect(s, Inches(0.6), Inches(2.20), Inches(6.5), Inches(4.6),
         NAVY)
    rect(s, Inches(0.6), Inches(2.20), Inches(6.5), Inches(0.10),
         COPPER)
    simple_text(s, Inches(0.85), Inches(2.40), Inches(6.0), Inches(0.35),
                "KNOWN URANIUM-BEARING ZONES — ZAMBIA",
                size=11, bold=True, color=COPPER_LIGHT)

    deposits = [
        ("LUMWANA / KANYEMBA",
         "North-Western Province",
         "Uranium recovered as by-product of copper mining; "
         "previously stockpiled by Barrick."),
        ("GWEMBE / SIAVONGA BELT",
         "Southern Province",
         "Karoo-type sandstone uranium; site of the 2024 "
         "Siavonga Baseline Survey."),
        ("MUTANGA / DIBWE",
         "Southern Province",
         "Defined uranium resource under historical "
         "GoviEx / African Energy permits."),
    ]
    py = 2.85
    for name, region, body in deposits:
        rect(s, Inches(0.85), Inches(py), Inches(6.0), Inches(1.15),
             RGBColor(0x12, 0x28, 0x4A))
        rect(s, Inches(0.85), Inches(py), Inches(0.08), Inches(1.15),
             COPPER)
        simple_text(s, Inches(1.05), Inches(py + 0.08), Inches(5.8),
                    Inches(0.3),
                    name, size=13, bold=True, color=WHITE)
        simple_text(s, Inches(1.05), Inches(py + 0.36), Inches(5.8),
                    Inches(0.3),
                    region, size=10, italic=True, color=COPPER_LIGHT)
        simple_text(s, Inches(1.05), Inches(py + 0.62), Inches(5.8),
                    Inches(0.5),
                    body, size=11, color=RGBColor(0xCF, 0xD8, 0xE3),
                    line_spacing=1.3)
        py += 1.27

    # RIGHT — VALUE-CHAIN PIVOT
    rect(s, Inches(7.3), Inches(2.20), Inches(5.4), Inches(4.6),
         WHITE)
    rect(s, Inches(7.3), Inches(2.20), Inches(0.12), Inches(4.6),
         COPPER)
    simple_text(s, Inches(7.55), Inches(2.35), Inches(5.0), Inches(0.35),
                "THE VALUE-CHAIN PIVOT",
                size=11, bold=True, color=COPPER)

    pivot = [
        ("TODAY",
         "Raw-mineral exporter — uranium leaves Zambia for "
         "enrichment elsewhere; we capture wellhead value only."),
        ("WITH NUCLEAR POWER",
         "High-tech energy producer — domestic fuel flows into "
         "domestic baseload supporting mines, smelters, hospitals."),
        ("STRATEGIC UPSIDE",
         "Import independence; price-stable electricity; "
         "high-skill jobs anchored to a regulated, exportable "
         "competency."),
    ]
    py = 2.85
    for head, body in pivot:
        simple_text(s, Inches(7.55), Inches(py), Inches(5.0),
                    Inches(0.3),
                    head, size=11, bold=True, color=NAVY)
        simple_text(s, Inches(7.55), Inches(py + 0.30), Inches(5.0),
                    Inches(1.05),
                    body, size=11, color=INK, line_spacing=1.3)
        py += 1.30

    set_notes(s,
        "[Strategic Resource] Zambia has known uranium-bearing zones "
        "across the North-Western and Southern provinces — Lumwana, "
        "Gwembe Valley, Mutanga. We currently export uranium as a raw "
        "mineral. Why? By developing nuclear power we move from "
        "raw-material exporter to high-tech energy producer. We capture "
        "the value-chain, we anchor high-skill jobs, and we stop paying "
        "to import fuel that we are already sending abroad. Why should "
        "we export our uranium as a raw mineral when we could be using "
        "it to power our own mines and factories?")
    slide_footer(s, 6, 8,
                 [CITATIONS["mines"], CITATIONS["ndp"]])


# -----------------------------------------------------------------------
# SLIDE 7 — ECONOMIC IMPACT (CHART: COPPER GOAL MW REQUIREMENT)
# -----------------------------------------------------------------------
def slide_economic(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, OFFWHITE)
    slide_header(s,
                 "05  ·  ECONOMIC IMPACT",
                 "Powering the 3-million-tonne copper ambition.",
                 "Zambia's flagship industrial target needs "
                 "high-density, drought-proof power.")

    # CHART: COLUMN — MW continuous needed at copper output levels
    chart_data = CategoryChartData()
    chart_data.categories = [
        "Current (~0.8 Mt)",
        "Mid-Path (~1.5 Mt)",
        "Target (3.0 Mt)",
    ]
    chart_data.add_series("MW continuous (mining only)",
                          (320, 600, 1200))
    cx, cy, cw, ch = Inches(0.6), Inches(2.20), Inches(7.4), Inches(4.4)
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
    dl.font.size = Pt(12)
    dl.font.bold = True
    dl.font.color.rgb = NAVY
    color_series_points(chart.series[0],
                        [COPPER, COPPER, CRISIS])
    style_chart_text(chart, color=INK, size=10)

    simple_text(s, Inches(0.6), Inches(6.65), Inches(7.4), Inches(0.4),
                "Estimated continuous power for copper mining at "
                "~3,500 kWh/t. Sources: MoMMD; ZESCO.",
                size=9, italic=True, color=MUTED, align=PP_ALIGN.CENTER)

    # RIGHT — IMPLICATIONS
    rect(s, Inches(8.25), Inches(2.20), Inches(4.55), Inches(4.6),
         WHITE)
    rect(s, Inches(8.25), Inches(2.20), Inches(0.12), Inches(4.6),
         COPPER)
    simple_text(s, Inches(8.50), Inches(2.35), Inches(4.1), Inches(0.3),
                "WHAT THIS MEANS",
                size=11, bold=True, color=COPPER)

    pts = [
        ("MINING INTENSITY",
         "Deep-level copper requires massive, uninterrupted power "
         "— hydropower is structurally mismatched."),
        ("COST STABILITY",
         "Nuclear fuel costs are decade-stable vs. volatile "
         "diesel imports and emergency electricity buys."),
        ("STRATEGIC FIT",
         "The 2,000 MW nuclear target maps directly onto the "
         "3 Mt copper power gap."),
    ]
    py = 2.75
    for head, body in pts:
        simple_text(s, Inches(8.50), Inches(py), Inches(4.1),
                    Inches(0.3),
                    head, size=11, bold=True, color=NAVY)
        simple_text(s, Inches(8.50), Inches(py + 0.30), Inches(4.1),
                    Inches(1.0),
                    body, size=11, color=INK, line_spacing=1.3)
        py += 1.30

    set_notes(s,
        "[Economic Impact] Zambia has set a public target of three "
        "million tonnes of copper annually. At industry-standard energy "
        "intensity that requires roughly twelve hundred megawatts of "
        "continuous power — for mining alone, before smelting, before "
        "households, before industrial diversification. Hydropower "
        "cannot deliver that reliably and drought after drought has "
        "proved it. Nuclear can. The 2,000 MW nuclear target maps "
        "almost exactly onto the copper-driven power gap. This is not "
        "ideology — it is arithmetic.")
    slide_footer(s, 7, 8,
                 [CITATIONS["mines"], CITATIONS["zesco"]])


# -----------------------------------------------------------------------
# SLIDE 8 — ROADMAP
# -----------------------------------------------------------------------
def slide_roadmap(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, NAVY_DEEP)

    rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.10), COPPER)

    simple_text(s, Inches(0.6), Inches(0.50), Inches(10), Inches(0.3),
                "06  ·  THE ROADMAP",
                size=11, bold=True, color=COPPER_LIGHT)
    simple_text(s, Inches(0.6), Inches(0.80), Inches(12.1), Inches(0.7),
                "From policy approval to first SMR — one decade.",
                size=28, bold=True, color=WHITE)
    line_rule(s, Inches(0.6), Inches(1.55), Inches(1.2), COPPER, 1.5)
    simple_text(s, Inches(0.6), Inches(1.65), Inches(12.1), Inches(0.4),
                "Three workstreams. Visible milestones. "
                "Accountability windows.",
                size=13, italic=True,
                color=RGBColor(0xCF, 0xD8, 0xE3))

    # TIMELINE GRID
    grid_left = 1.4
    grid_right = 12.7
    grid_y = 2.4
    grid_w = grid_right - grid_left

    # year labels
    years = ["2026", "2028", "2030", "2032", "2034", "2036+"]
    step = grid_w / (len(years) - 1)
    for i, y in enumerate(years):
        x = grid_left + i * step
        simple_text(s, Inches(x - 0.3), Inches(grid_y - 0.05),
                    Inches(0.6), Inches(0.3),
                    y, size=10, bold=True, color=COPPER_LIGHT,
                    align=PP_ALIGN.CENTER)

    line_rule(s, Inches(grid_left), Inches(grid_y + 0.35),
              Inches(grid_w), RGBColor(0x33, 0x44, 0x66), 0.75)

    # Three workstream rows: (label, start_year_index, span_years_in_steps, color)
    # year span: 2026 = idx 0, each step = 2 years
    rows = [
        ("POLICY  &  REGULATION",
         "Enact Nuclear Bill · Formalise NSPA · Site licensing",
         0.0, 2.0, COPPER),
        ("HUMAN  CAPITAL",
         "Medical Physics · Nuclear Eng. · Data Analytics cohorts",
         0.5, 4.0, COPPER_LIGHT),
        ("SMR  IMPLEMENTATION",
         "Partner selection (ROSATOM / NuScale / Westinghouse) · "
         "Pre-feas · EPC · First criticality",
         2.0, 3.0, CRISIS),
    ]
    row_y = grid_y + 0.65
    row_h = 1.10
    for i, (label, body, start, span, col) in enumerate(rows):
        top = row_y + i * row_h

        # bar
        bar_left = grid_left + start * step
        bar_w = span * step
        rect(s, Inches(bar_left), Inches(top + 0.30),
             Inches(bar_w), Inches(0.45), col)
        rect(s, Inches(bar_left), Inches(top + 0.30),
             Inches(0.08), Inches(0.45),
             WHITE)

        # label inside bar (only if bar wide enough — left aligned text overlay)
        simple_text(s, Inches(bar_left + 0.12),
                    Inches(top + 0.32), Inches(bar_w - 0.15),
                    Inches(0.4),
                    label, size=11, bold=True, color=NAVY_DEEP)

        # description below bar
        simple_text(s, Inches(grid_left), Inches(top + 0.80),
                    Inches(grid_w), Inches(0.30),
                    body, size=10,
                    color=RGBColor(0xCF, 0xD8, 0xE3))

    # CLOSING ASK
    rect(s, Inches(0.6), Inches(6.55), Inches(12.1), Inches(0.55),
         RGBColor(0x12, 0x28, 0x4A))
    rect(s, Inches(0.6), Inches(6.55), Inches(0.10), Inches(0.55),
         COPPER)
    textbox(
        s, Inches(0.85), Inches(6.62), Inches(11.7), Inches(0.45),
        [
            {"text": "We have the minerals. We have the regulator. "
             "We have the data capacity. ",
             "size": 13, "color": WHITE},
            {"text": "Zambia is ready to start the nuclear conversation.",
             "size": 13, "bold": True, "color": COPPER_LIGHT},
        ],
        align=PP_ALIGN.LEFT,
    )

    simple_text(s, Inches(0.6), Inches(7.20), Inches(12.1), Inches(0.25),
                "8 / 8   ·   E. Mwaba   ·   Sources: "
                "DOE REMS 2024 · IAEA Milestones · 7NDP & Vision 2030 · "
                "ZESCO 2023-24 · MoMMD · Siavonga Baseline 2024",
                size=8, italic=True,
                color=RGBColor(0x8B, 0x97, 0xA8),
                align=PP_ALIGN.CENTER)

    set_notes(s,
        "[Roadmap] Three workstreams run in parallel. Policy and "
        "regulation lands first: pass the Nuclear Bill in 2026, "
        "formalise NSPA, begin site licensing. Human capital builds "
        "alongside: Medical Physics and Nuclear Engineering cohorts "
        "leveraging local universities, plus data-analytics talent — "
        "the same skill set that built this dashboard. SMR "
        "implementation begins around 2028 with partner selection "
        "drawing on the ROSATOM, IP3 / NuScale, and Westinghouse "
        "relationships already in motion, targeting first criticality "
        "before 2036. We have the minerals, we have the regulatory "
        "framework, and as I've shown today, we have the data capacity. "
        "Zambia is ready to start the nuclear conversation.")


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide_title(prs)
    slide_story(prs)
    slide_hook(prs)
    slide_vision(prs)
    slide_safety(prs)
    slide_resource(prs)
    slide_economic(prs)
    slide_roadmap(prs)

    prs.save(OUTPUT)
    print(f"Generated: {OUTPUT}")


if __name__ == "__main__":
    build()
