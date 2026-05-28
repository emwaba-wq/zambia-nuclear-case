"""
build_presentation.py
Generates: The Case for Nuclear Power in Zambia (5-slide .pptx).
Built for executive / supervisor-grade review.

Story arc (3-minute pitch, Storytelling with Data):
  1. Title         — thesis up front
  2. Hook          — same element, two realities (Siavonga vs. US)
  3. Evidence      — 59.12 mrem as proof of mastery
  4. Stakes        — drought is the new normal; the bill is unpaid
  5. Resolution    — three legislative actions

Run:
  pip install -r requirements.txt
  python build_presentation.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

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

CITATIONS = {
    "siavonga": "Siavonga Uranium Baseline Survey (2024)",
    "zesco": "ZESCO Annual Reports 2023-2024",
    "ndp": "Zambia 7NDP & Vision 2030",
    "iaea": "IAEA Milestone Approach for New Nuclear Programs",
    "doe": "U.S. DOE REMS 2024 Report",
}


def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def rect(slide, left, top, width, height, color, *, line=False):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    if not line:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def line_rule(slide, left, top, width, color, weight_pt=1.0):
    s = slide.shapes.add_connector(1, left, top, left + width, top)
    s.line.color.rgb = color
    s.line.width = Pt(weight_pt)
    return s


def textbox(slide, left, top, width, height, runs, *,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
            line_spacing=1.15):
    """runs: list of dicts: {text, size, bold, color, font, italic, space_after}.
    A run with text starting with '\n' starts a new paragraph."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.0)
    tf.margin_right = Inches(0.0)
    tf.margin_top = Inches(0.0)
    tf.margin_bottom = Inches(0.0)

    first = True
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
        first = False
    return tb


def simple_text(slide, left, top, width, height, text, *,
                size=18, bold=False, italic=False, color=INK,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                font="Calibri", line_spacing=1.15):
    return textbox(slide, left, top, width, height,
                   [{"text": text, "size": size, "bold": bold,
                     "italic": italic, "color": color, "font": font}],
                   align=align, anchor=anchor, line_spacing=line_spacing)


def eyebrow(slide, left, top, width, text, color=COPPER):
    return simple_text(slide, left, top, width, Inches(0.3),
                       text, size=11, bold=True, color=color,
                       font="Calibri")


def slide_footer(slide, n, citations):
    rect(slide, Inches(0), Inches(7.30), Inches(13.333), Inches(0.20),
         RULE)
    rect(slide, Inches(0), Inches(7.30), Inches(0.7), Inches(0.20),
         COPPER)
    simple_text(slide, Inches(0.5), Inches(7.05), Inches(10.5),
                Inches(0.25),
                "Sources: " + " · ".join(citations),
                size=9, color=MUTED, italic=True)
    simple_text(slide, Inches(11.5), Inches(7.05), Inches(1.6),
                Inches(0.25),
                f"{n} / 5  ·  E. Mwaba",
                size=9, bold=True, color=SLATE, align=PP_ALIGN.RIGHT)


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# -----------------------------------------------------------------------
# SLIDE 1 — TITLE
# -----------------------------------------------------------------------
def slide_title(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, NAVY_DEEP)

    rect(s, Inches(0.6), Inches(0.6), Inches(0.08), Inches(0.45), COPPER)
    simple_text(s, Inches(0.85), Inches(0.58), Inches(10), Inches(0.4),
                "ENERGY SECURITY  ·  REPUBLIC OF ZAMBIA  ·  2026",
                size=12, bold=True, color=COPPER_LIGHT)

    textbox(
        s, Inches(0.6), Inches(1.7), Inches(12.1), Inches(2.5),
        [
            {"text": "We already own", "size": 64, "bold": True,
             "color": WHITE},
            {"text": "\nthe solution.", "size": 64, "bold": True,
             "color": COPPER_LIGHT, "space_after": 0},
        ],
        line_spacing=1.0,
    )

    line_rule(s, Inches(0.6), Inches(4.35), Inches(2.5), COPPER, 1.5)

    simple_text(s, Inches(0.6), Inches(4.5), Inches(12.1), Inches(0.5),
                "The Case for Nuclear Power in Zambia",
                size=26, bold=False, color=WHITE)

    rect(s, Inches(0.6), Inches(5.3), Inches(12.1), Inches(0.04),
         RGBColor(0x33, 0x44, 0x66))

    simple_text(
        s, Inches(0.6), Inches(5.5), Inches(12.1), Inches(1.1),
        ("Zambia must leverage its existing radiological safety expertise "
         "and uranium resources to transition to nuclear power as a "
         "drought-proof solution for national energy security."),
        size=15, italic=True, color=RGBColor(0xCF, 0xD8, 0xE3),
        line_spacing=1.25,
    )

    simple_text(s, Inches(0.6), Inches(6.85), Inches(8), Inches(0.3),
                "Prepared by  ELIZABETH MWABA",
                size=11, bold=True, color=WHITE)
    simple_text(s, Inches(0.6), Inches(7.10), Inches(8), Inches(0.25),
                "Radiological Analyst  ·  Supervisor Review Brief",
                size=9, color=RGBColor(0x99, 0xA3, 0xB3))

    set_notes(s,
        "[0:00-0:30 — SETUP]  Three minutes. One argument. Zambia does "
        "not have to choose between energy poverty and environmental "
        "risk — because we already own the solution. We have uranium in "
        "the ground. We have a Radiation Protection Authority that has "
        "spent decades managing radiological risk. And we have a "
        "Cabinet-approved 2020 Nuclear Policy. What we do not yet have "
        "is the legislative signature that turns those assets into "
        "megawatts. That signature is what I am here to argue for.")


# -----------------------------------------------------------------------
# SLIDE 2 — HOOK
# -----------------------------------------------------------------------
def slide_hook(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, OFFWHITE)

    eyebrow(s, Inches(0.6), Inches(0.45), Inches(8), "01  ·  THE HOOK")
    simple_text(s, Inches(0.6), Inches(0.75), Inches(12.1), Inches(0.7),
                "Same element. Two realities.",
                size=34, bold=True, color=NAVY)
    line_rule(s, Inches(0.6), Inches(1.45), Inches(1.2), COPPER, 1.5)
    simple_text(s, Inches(0.6), Inches(1.55), Inches(12.1), Inches(0.5),
                "Zambians and Americans both live with uranium. "
                "Only one nation measures it.",
                size=15, italic=True, color=MUTED)

    panel_y = Inches(2.45)
    panel_h = Inches(4.0)

    # LEFT — CRISIS
    rect(s, Inches(0.6), panel_y, Inches(6.05), panel_h, CRISIS_DEEP)
    rect(s, Inches(0.6), panel_y, Inches(0.12), panel_h, CRISIS)
    simple_text(s, Inches(0.85), Inches(2.6), Inches(5.6), Inches(0.3),
                "ZAMBIA  ·  SIAVONGA",
                size=11, bold=True, color=COPPER_LIGHT)
    simple_text(s, Inches(0.85), Inches(2.95), Inches(5.6), Inches(1.4),
                "Uranium in the water.\nUranium in the dust.",
                size=26, bold=True, color=WHITE, line_spacing=1.05)
    simple_text(s, Inches(0.85), Inches(4.55), Inches(5.6), Inches(1.0),
                "Zero monitoring.  Zero baseline.\nZero accountability.",
                size=13, color=RGBColor(0xF2, 0xCF, 0xCF),
                line_spacing=1.3)
    line_rule(s, Inches(0.85), Inches(5.65), Inches(1.0), COPPER, 1.0)
    simple_text(s, Inches(0.85), Inches(5.75), Inches(5.6), Inches(0.5),
                "RADIOLOGICAL  FEAR",
                size=14, bold=True, color=COPPER_LIGHT)

    # RIGHT — MASTERY
    rect(s, Inches(6.85), panel_y, Inches(5.85), panel_h, NAVY)
    rect(s, Inches(6.85), panel_y, Inches(0.12), panel_h, COPPER)
    simple_text(s, Inches(7.1), Inches(2.6), Inches(5.5), Inches(0.3),
                "USA  ·  DOE NUCLEAR ENTERPRISE",
                size=11, bold=True, color=COPPER_LIGHT)
    simple_text(s, Inches(7.1), Inches(2.95), Inches(5.5), Inches(1.4),
                "22,000 workers.\n59.12 mrem average.",
                size=26, bold=True, color=WHITE, line_spacing=1.05)
    simple_text(s, Inches(7.1), Inches(4.55), Inches(5.5), Inches(1.0),
                "Every dose tracked.  Every site audited.\n"
                "Every year. For a decade.",
                size=13, color=RGBColor(0xCF, 0xD8, 0xE3),
                line_spacing=1.3)
    line_rule(s, Inches(7.1), Inches(5.65), Inches(1.0), COPPER, 1.0)
    simple_text(s, Inches(7.1), Inches(5.75), Inches(5.5), Inches(0.5),
                "INDUSTRIAL  MASTERY",
                size=14, bold=True, color=COPPER_LIGHT)

    # KICKER
    simple_text(s, Inches(0.6), Inches(6.65), Inches(12.1), Inches(0.4),
                "The gap is not the science. It is the system.",
                size=15, bold=True, color=NAVY,
                align=PP_ALIGN.CENTER)

    set_notes(s,
        "[0:30-1:15 — CONFLICT]  Picture two scenes. Siavonga, Zambia: "
        "the 2024 Uranium Baseline Survey found uranium in the drinking "
        "water and uranium in the dust. People live on top of the "
        "resource — with no dosimetry, no baseline, no regulator on the "
        "ground. That is radiological fear. Now picture the U.S. nuclear "
        "enterprise: twenty-two thousand workers, the average technician "
        "absorbs fifty-nine point one two mrem a year — less than two "
        "percent of the federal limit. Every dose tracked. Every year. "
        "Same element. Same physics. Two completely different outcomes. "
        "The variable is not the uranium. The variable is the system "
        "wrapped around it.")
    slide_footer(s, 2, [CITATIONS["siavonga"], CITATIONS["doe"]])


# -----------------------------------------------------------------------
# SLIDE 3 — EVIDENCE
# -----------------------------------------------------------------------
def slide_evidence(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, OFFWHITE)

    eyebrow(s, Inches(0.6), Inches(0.45), Inches(8),
            "02  ·  THE EVIDENCE")
    simple_text(s, Inches(0.6), Inches(0.75), Inches(12.1), Inches(0.7),
                "Nuclear safety is not theory. It's a track record.",
                size=30, bold=True, color=NAVY)
    line_rule(s, Inches(0.6), Inches(1.45), Inches(1.2), COPPER, 1.5)
    simple_text(s, Inches(0.6), Inches(1.55), Inches(12.1), Inches(0.5),
                "A decade of U.S. data (2014-2024) — and a regulatory "
                "playbook Zambia is already following.",
                size=14, italic=True, color=MUTED)

    # HERO STAT BLOCK
    rect(s, Inches(0.6), Inches(2.35), Inches(7.4), Inches(2.9), NAVY)
    rect(s, Inches(0.6), Inches(2.35), Inches(0.15), Inches(2.9), COPPER)
    simple_text(s, Inches(0.95), Inches(2.5), Inches(7.0), Inches(0.35),
                "AVERAGE  INDIVIDUAL  DOSE", size=12, bold=True,
                color=COPPER_LIGHT)
    textbox(
        s, Inches(0.95), Inches(2.95), Inches(7.0), Inches(1.9),
        [
            {"text": "59.12", "size": 110, "bold": True, "color": WHITE},
            {"text": "  mrem", "size": 28, "bold": False,
             "color": COPPER_LIGHT},
        ],
        line_spacing=1.0,
    )
    simple_text(s, Inches(0.95), Inches(4.55), Inches(7.0), Inches(0.4),
                "per technician  ·  per year  ·  10-year mean",
                size=12, color=RGBColor(0xCF, 0xD8, 0xE3))
    line_rule(s, Inches(0.95), Inches(4.95), Inches(7.0),
              RGBColor(0x33, 0x44, 0x66), 0.75)
    textbox(
        s, Inches(0.95), Inches(5.0), Inches(7.0), Inches(0.4),
        [
            {"text": "< 2%", "size": 18, "bold": True,
             "color": COPPER_LIGHT},
            {"text": "  of the 5,000 mrem federal occupational limit",
             "size": 13, "color": WHITE},
        ],
    )

    # THREE PROOFS
    proofs = [
        ("SCALABLE",
         "Workforce grew 1,133 → 1,167. "
         "Average dose did not move."),
        ("CONTAINED",
         "98% external radiation — managed by shielding. "
         "2% internal — controlled by protocol."),
        ("REPLICABLE",
         "IAEA Milestone Approach. The same roadmap Zambia is on, "
         "via RPA → NSPA."),
    ]
    px = 8.25
    py = 2.35
    pw = 4.55
    ph = 0.92
    gap = 0.07
    for i, (label, body) in enumerate(proofs):
        top = py + i * (ph + gap)
        rect(s, Inches(px), Inches(top), Inches(pw), Inches(ph), WHITE)
        rect(s, Inches(px), Inches(top), Inches(0.10), Inches(ph),
             COPPER)
        simple_text(s, Inches(px + 0.25), Inches(top + 0.08),
                    Inches(pw - 0.3), Inches(0.3),
                    label, size=11, bold=True, color=NAVY)
        simple_text(s, Inches(px + 0.25), Inches(top + 0.36),
                    Inches(pw - 0.3), Inches(0.55),
                    body, size=12, color=INK, line_spacing=1.2)

    # KICKER STRIP
    rect(s, Inches(8.25), Inches(5.27), Inches(4.55), Inches(0.92),
         CRISIS)
    simple_text(s, Inches(8.45), Inches(5.42), Inches(4.25), Inches(0.7),
                "Zambia is not starting from zero.\n"
                "Zambia is starting from competence.",
                size=13, bold=True, color=WHITE, line_spacing=1.25)

    simple_text(s, Inches(0.6), Inches(6.55), Inches(12.1), Inches(0.4),
                "If the U.S. can manage 22,000 workers at < 2% of limit, "
                "Zambia can manage one regulator and one reactor.",
                size=14, italic=True, color=SLATE,
                align=PP_ALIGN.CENTER)

    set_notes(s,
        "[1:15-2:00 — EVIDENCE]  Can the system actually be trusted? "
        "Fifty-nine point one two mrem. Across a decade. Across twenty-two "
        "thousand workers. That is the U.S. DOE Radiation Exposure "
        "Monitoring System record — under two percent of the federal "
        "limit. And the line that matters most for us: between 2021 "
        "and 2024 the monitored workforce grew, and individual dose did "
        "not move. Safety scales. Ninety-eight percent of exposure is "
        "external — managed by shielding we already use in our copper "
        "mines. The IAEA Milestone Approach is the exact roadmap "
        "Zambia is already on through the Radiation Protection "
        "Authority. We are not starting from zero. We are starting from "
        "competence.")
    slide_footer(s, 3, [CITATIONS["doe"], CITATIONS["iaea"]])


# -----------------------------------------------------------------------
# SLIDE 4 — STAKES
# -----------------------------------------------------------------------
def slide_stakes(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, OFFWHITE)

    eyebrow(s, Inches(0.6), Inches(0.45), Inches(8),
            "03  ·  THE STAKES")
    simple_text(s, Inches(0.6), Inches(0.75), Inches(12.1), Inches(0.7),
                "Drought is the new normal. Hydropower is the old bet.",
                size=28, bold=True, color=NAVY)
    line_rule(s, Inches(0.6), Inches(1.45), Inches(1.2), COPPER, 1.5)
    simple_text(s, Inches(0.6), Inches(1.55), Inches(12.1), Inches(0.5),
                "Every dry year, Zambia pays in megawatts and in dollars. "
                "We keep paying.",
                size=14, italic=True, color=MUTED)

    # CRISIS STATS ROW
    stats = [
        ("84%", "of generation\nfrom hydropower"),
        ("1,000 MW", "deficit during the\n2015/16 drought"),
        ("$440M", "in unbudgeted\nelectricity imports"),
        ("2023-24", "Kariba crisis.\nIt happened again."),
    ]
    sx = 0.6
    sw = 2.95
    sgap = 0.10
    sy = 2.35
    sh = 1.6
    for i, (big, body) in enumerate(stats):
        left = sx + i * (sw + sgap)
        rect(s, Inches(left), Inches(sy), Inches(sw), Inches(sh),
             CRISIS_DEEP)
        rect(s, Inches(left), Inches(sy), Inches(sw), Inches(0.10),
             CRISIS)
        simple_text(s, Inches(left + 0.2), Inches(sy + 0.2),
                    Inches(sw - 0.3), Inches(0.7),
                    big, size=30, bold=True, color=WHITE)
        simple_text(s, Inches(left + 0.2), Inches(sy + 0.92),
                    Inches(sw - 0.3), Inches(0.65),
                    body, size=11, color=RGBColor(0xF2, 0xCF, 0xCF),
                    line_spacing=1.2)

    # PIVOT LINE
    line_rule(s, Inches(0.6), Inches(4.20), Inches(12.1), RULE, 0.5)
    simple_text(s, Inches(0.6), Inches(4.25), Inches(12.1), Inches(0.4),
                "THE ANSWER  —  ALREADY  ON  PAPER",
                size=11, bold=True, color=COPPER,
                align=PP_ALIGN.CENTER)

    # SOLUTION PANEL
    rect(s, Inches(0.6), Inches(4.75), Inches(12.1), Inches(2.0), NAVY)
    rect(s, Inches(0.6), Inches(4.75), Inches(12.1), Inches(0.08), COPPER)

    textbox(
        s, Inches(0.9), Inches(4.95), Inches(5.0), Inches(1.7),
        [
            {"text": "2,000 MW", "size": 52, "bold": True,
             "color": COPPER_LIGHT},
            {"text": "\nnuclear baseload target",
             "size": 14, "color": WHITE, "space_after": 0},
            {"text": "\nVision 2030  ·  7NDP",
             "size": 11, "color": RGBColor(0xCF, 0xD8, 0xE3),
             "space_after": 0},
        ],
        line_spacing=1.05,
    )

    rect(s, Inches(6.1), Inches(5.0), Inches(0.02), Inches(1.5),
         RGBColor(0x33, 0x44, 0x66))

    points = [
        ("DROUGHT-PROOF",
         "Baseload independent of Kariba levels."),
        ("FUEL-SOVEREIGN",
         "Domestic uranium reserves. No fuel imports."),
        ("ALREADY MOBILISED",
         "2020 Nuclear Policy approved. ROSATOM, IP3, KAERI signed."),
    ]
    py2 = 5.0
    for i, (head, body) in enumerate(points):
        top = py2 + i * 0.50
        simple_text(s, Inches(6.4), Inches(top), Inches(2.3),
                    Inches(0.3),
                    head, size=11, bold=True, color=COPPER_LIGHT)
        simple_text(s, Inches(8.7), Inches(top), Inches(4.1),
                    Inches(0.4),
                    body, size=12, color=WHITE)

    simple_text(s, Inches(0.6), Inches(6.90), Inches(12.1), Inches(0.4),
                "We wrote the plan in 2020. The grid is still waiting "
                "in 2026.",
                size=13, italic=True, bold=True, color=CRISIS,
                align=PP_ALIGN.CENTER)

    set_notes(s,
        "[2:00-2:30 — STAKES]  Now bring it home. Eighty-four percent "
        "of our electricity comes from hydropower. In the 2015/16 "
        "drought we lost a thousand megawatts and spent roughly four "
        "hundred and forty million U.S. dollars on unbudgeted imports. "
        "The 2023-24 ZESCO reports show the Kariba crisis is not a "
        "one-off — it is the new climate baseline. Meanwhile our 2020 "
        "Nuclear Policy commits to a two thousand megawatt target. "
        "ROSATOM, IP3, and KAERI partnerships are signed. The uranium "
        "is in the ground. We wrote the plan in 2020. Six years later, "
        "the grid is still waiting.")
    slide_footer(s, 4,
                 [CITATIONS["zesco"], CITATIONS["ndp"]])


# -----------------------------------------------------------------------
# SLIDE 5 — CALL TO ACTION
# -----------------------------------------------------------------------
def slide_cta(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, NAVY_DEEP)

    rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.10), COPPER)

    simple_text(s, Inches(0.6), Inches(0.55), Inches(10), Inches(0.4),
                "04  ·  THE ASK",
                size=11, bold=True, color=COPPER_LIGHT)
    simple_text(s, Inches(0.6), Inches(0.85), Inches(12.1), Inches(0.9),
                "Three decisions. One generation.",
                size=38, bold=True, color=WHITE)
    line_rule(s, Inches(0.6), Inches(1.7), Inches(1.2), COPPER, 1.5)
    simple_text(s, Inches(0.6), Inches(1.8), Inches(12.1), Inches(0.5),
                "From radiological fear to industrial mastery — "
                "the choice is legislative, not technical.",
                size=14, italic=True,
                color=RGBColor(0xCF, 0xD8, 0xE3))

    actions = [
        ("01", "ENACT",
         "Pass the Nuclear Bill.",
         "Empower NSPA as an IAEA-aligned independent regulator. "
         "Without it, every other step is symbolic."),
        ("02", "EQUIP",
         "Fund radiological monitoring — starting in Siavonga.",
         "Convert invisible risk into measured data. No Zambian "
         "community should live next to uranium without instruments."),
        ("03", "EXECUTE",
         "Greenlight CNST + pre-feasibility for 2,000 MW.",
         "Use the partnerships we already signed. Stop re-debating "
         "what Cabinet already approved in 2020."),
    ]
    ay = 2.65
    ah = 1.20
    gap = 0.10
    for i, (num, label, head, body) in enumerate(actions):
        top = ay + i * (ah + gap)
        rect(s, Inches(0.6), Inches(top), Inches(1.15), Inches(ah),
             COPPER)
        simple_text(s, Inches(0.6), Inches(top + 0.32), Inches(1.15),
                    Inches(0.7),
                    num, size=34, bold=True, color=NAVY_DEEP,
                    align=PP_ALIGN.CENTER)

        rect(s, Inches(1.75), Inches(top), Inches(11.0), Inches(ah),
             RGBColor(0x12, 0x28, 0x4A))
        rect(s, Inches(1.75), Inches(top), Inches(11.0), Inches(0.04),
             RGBColor(0x33, 0x44, 0x66))

        simple_text(s, Inches(2.0), Inches(top + 0.13), Inches(2.0),
                    Inches(0.3),
                    label, size=11, bold=True, color=COPPER_LIGHT)
        simple_text(s, Inches(2.0), Inches(top + 0.40), Inches(10.5),
                    Inches(0.4),
                    head, size=18, bold=True, color=WHITE)
        simple_text(s, Inches(2.0), Inches(top + 0.75), Inches(10.5),
                    Inches(0.5),
                    body, size=12,
                    color=RGBColor(0xCF, 0xD8, 0xE3),
                    line_spacing=1.25)

    line_rule(s, Inches(0.6), Inches(6.65), Inches(12.1),
              RGBColor(0x33, 0x44, 0x66), 0.5)
    textbox(
        s, Inches(0.6), Inches(6.75), Inches(12.1), Inches(0.5),
        [
            {"text": "The fuel is under our feet.  ",
             "size": 14, "color": WHITE},
            {"text": "The expertise is on our payroll.  ",
             "size": 14, "color": WHITE},
            {"text": "The bill is on the table.",
             "size": 14, "bold": True, "color": COPPER_LIGHT},
        ],
        align=PP_ALIGN.CENTER,
    )
    simple_text(s, Inches(0.6), Inches(7.10), Inches(12.1), Inches(0.3),
                "5 / 5   ·   E. Mwaba   ·   "
                "Sources: 7NDP & Vision 2030 · IAEA Milestones · "
                "Siavonga Uranium Baseline Survey 2024",
                size=9, italic=True, color=RGBColor(0x8B, 0x97, 0xA8),
                align=PP_ALIGN.CENTER)

    set_notes(s,
        "[2:30-3:00 — RESOLUTION]  So what am I asking for? Three "
        "decisions. First, ENACT — pass the Nuclear Bill and formalise "
        "the Nuclear Safety and Protection Authority as an independent "
        "regulator, aligned with the IAEA Milestone Approach. Without "
        "that signature, every other step is symbolic. Second, EQUIP "
        "— fund radiological monitoring, starting in Siavonga, so no "
        "community lives beside uranium without instruments. Third, "
        "EXECUTE — greenlight the Centre for Nuclear Science and "
        "Technology and the pre-feasibility studies for the two thousand "
        "megawatt programme. The fuel is under our feet. The expertise "
        "is on our payroll. The bill is on the table. Sign it. "
        "Thank you.")


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide_title(prs)
    slide_hook(prs)
    slide_evidence(prs)
    slide_stakes(prs)
    slide_cta(prs)

    prs.save(OUTPUT)
    print(f"Generated: {OUTPUT}")


if __name__ == "__main__":
    build()
