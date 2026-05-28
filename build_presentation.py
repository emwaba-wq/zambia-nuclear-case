"""
build_presentation.py
Generates: The Case for Nuclear Power in Zambia (5-slide .pptx).
Structure follows the 3-minute Story (Setup -> Conflict -> Resolution)
from Cole Nussbaumer Knaflic's "Storytelling with Data".

Big Idea:
  Zambia must leverage its existing radiological safety expertise and
  uranium resources to transition to nuclear power as a drought-proof
  solution for national energy security.

Run:
  pip install -r requirements.txt
  python build_presentation.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

OUTPUT = "The_Case_for_Nuclear_Power_in_Zambia.pptx"

NAVY = RGBColor(0x0B, 0x2E, 0x4F)
COPPER = RGBColor(0xB8, 0x6B, 0x2A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xF2, 0xF4, 0xF7)
ACCENT = RGBColor(0xC0, 0x39, 0x2B)
TEXT = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x55, 0x5B, 0x66)

CITATIONS = {
    "siavonga": "Siavonga Uranium Baseline Survey (2024)",
    "zesco": "ZESCO Annual Reports 2023-2024",
    "ndp": "Zambia 7th National Development Plan (7NDP) & Vision 2030",
    "iaea": "IAEA Milestone Approach for New Nuclear Programs",
    "doe": "U.S. DOE REMS 2024 Report",
}


def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, left, top, width, height, fill_color, line=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if not line:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def add_text(slide, left, top, width, height, text, *,
             size=18, bold=False, color=TEXT, align=PP_ALIGN.LEFT,
             font="Calibri"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def add_bullets(slide, left, top, width, height, bullets, *,
                size=16, color=TEXT, bullet_char="•"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(6)
        r = p.add_run()
        r.text = f"{bullet_char}  {b}"
        r.font.name = "Calibri"
        r.font.size = Pt(size)
        r.font.color.rgb = color
    return tb


def add_header_band(slide, title, subtitle=None):
    add_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(1.05), NAVY)
    add_rect(slide, Inches(0), Inches(1.05), Inches(13.333), Inches(0.08), COPPER)
    add_text(slide, Inches(0.5), Inches(0.18), Inches(12.5), Inches(0.6),
             title, size=28, bold=True, color=WHITE)
    if subtitle:
        add_text(slide, Inches(0.5), Inches(0.65), Inches(12.5), Inches(0.35),
                 subtitle, size=14, color=RGBColor(0xCF, 0xD8, 0xE3))


def add_footer(slide, slide_no, citations):
    add_rect(slide, Inches(0), Inches(7.05), Inches(13.333), Inches(0.45), NAVY)
    cite_text = "Sources: " + " | ".join(citations)
    add_text(slide, Inches(0.4), Inches(7.10), Inches(11.5), Inches(0.4),
             cite_text, size=9, color=WHITE)
    add_text(slide, Inches(12.3), Inches(7.10), Inches(0.9), Inches(0.4),
             f"{slide_no} / 5", size=10, bold=True, color=COPPER,
             align=PP_ALIGN.RIGHT)


def set_notes(slide, notes_text):
    slide.notes_slide.notes_text_frame.text = notes_text


def slide_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, NAVY)

    add_rect(slide, Inches(0), Inches(3.6), Inches(13.333), Inches(0.08), COPPER)

    add_text(slide, Inches(0.8), Inches(1.6), Inches(11.7), Inches(1.2),
             "The Case for Nuclear Power in Zambia",
             size=48, bold=True, color=WHITE)
    add_text(slide, Inches(0.8), Inches(2.7), Inches(11.7), Inches(0.8),
             "From Radiological Fear to Industrial Mastery",
             size=24, color=COPPER)

    add_text(slide, Inches(0.8), Inches(3.95), Inches(11.7), Inches(2.0),
             ("Zambia must leverage its existing radiological safety "
              "expertise and uranium resources to transition to nuclear "
              "power as a drought-proof solution for national energy "
              "security."),
             size=18, color=RGBColor(0xE6, 0xEC, 0xF2))

    add_text(slide, Inches(0.8), Inches(6.2), Inches(11.7), Inches(0.4),
             "Prepared by: Elizabeth Mwaba",
             size=14, bold=True, color=WHITE)
    add_text(slide, Inches(0.8), Inches(6.55), Inches(11.7), Inches(0.4),
             "A data-driven advocacy brief", size=12,
             color=RGBColor(0xCF, 0xD8, 0xE3))

    notes = (
        "[SETUP — 0:00 to 0:30] Good afternoon. The story I want to tell "
        "you today is not about reactors. It is about a choice: between a "
        "future powered by drought, and a future powered by uranium "
        "Zambia already owns. My Big Idea is simple: Zambia must leverage "
        "its existing radiological safety expertise and uranium resources "
        "to transition to nuclear power as a drought-proof solution for "
        "national energy security. Over the next three minutes I'll walk "
        "you from an invisible danger in Siavonga, to a proven safety "
        "record in the United States, to an actionable policy ask here at "
        "home. The destination is industrial mastery."
    )
    set_notes(slide, notes)
    add_footer(slide, 1, [CITATIONS["ndp"], CITATIONS["iaea"]])


def slide_hook(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_header_band(slide,
                    "The Hook: Siavonga's Invisible Ghost",
                    "The paradox we already live with")

    left_x = Inches(0.5)
    right_x = Inches(6.95)
    panel_w = Inches(5.9)
    panel_y = Inches(1.5)
    panel_h = Inches(5.2)

    add_rect(slide, left_x, panel_y, panel_w, Inches(0.55), ACCENT)
    add_text(slide, left_x, Inches(1.55), panel_w, Inches(0.5),
             "  INVISIBLE DANGER", size=18, bold=True, color=WHITE)
    add_rect(slide, left_x, Inches(2.05), panel_w, panel_h - Inches(0.55),
             LIGHT)
    add_bullets(
        slide, Inches(0.75), Inches(2.2), Inches(5.6), Inches(4.8),
        [
            "Unmonitored environmental uranium detected in Siavonga "
            "water and dust samples.",
            "Communities exposed daily — without dosimetry, without "
            "baseline data, without recourse.",
            "Risk is real, but invisible: no instrumentation, no "
            "regulator presence, no accountability.",
            "This is the status quo of radiological FEAR — "
            "exposure without measurement.",
        ],
        size=15, color=TEXT,
    )

    add_rect(slide, right_x, panel_y, panel_w, Inches(0.55), NAVY)
    add_text(slide, right_x, Inches(1.55), panel_w, Inches(0.5),
             "  PRECISION MONITORING", size=18, bold=True, color=WHITE)
    add_rect(slide, right_x, Inches(2.05), panel_w, panel_h - Inches(0.55),
             LIGHT)
    add_bullets(
        slide, Inches(7.2), Inches(2.2), Inches(5.6), Inches(4.8),
        [
            "U.S. nuclear enterprise: ~22,000 workers, 11,000 with "
            "measurable dose, tracked annually.",
            "Average individual dose: 59.12 mrem — under 2% of the "
            "5,000 mrem federal limit.",
            "Workforce grew (1,133 → 1,167) with NO rise in individual "
            "exposure. Safety is scalable.",
            "This is industrial MASTERY — exposure rigorously measured, "
            "managed, and minimised.",
        ],
        size=15, color=TEXT,
    )

    add_text(slide, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.4),
             "The transition Zambia must make: from FEAR to MASTERY.",
             size=16, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)

    notes = (
        "[CONFLICT — 0:30 to 1:15] Picture two scenes. Scene one: "
        "Siavonga, Zambia. The 2024 Uranium Baseline Survey found "
        "uranium in the water and in the dust — unmonitored, "
        "untracked, in a community that lives on top of the resource. "
        "An invisible ghost. Scene two: the U.S. nuclear security "
        "enterprise. Twenty-two thousand workers. Every measurable "
        "dose recorded. The average technician absorbs 59.12 mrem a "
        "year — less than two percent of the federal limit, and "
        "comparable to routine medical imaging. Same element. Same "
        "physics. Two completely different societal outcomes. The "
        "difference is not the uranium. The difference is the system "
        "around it. That is the gap Zambia must close — and the "
        "expertise to close it already lives inside our Radiation "
        "Protection Authority."
    )
    set_notes(slide, notes)
    add_footer(slide, 2,
               [CITATIONS["siavonga"], CITATIONS["doe"]])


def slide_safety(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_header_band(slide,
                    "Safety Evidence: 59.12 mrem Proves Mastery",
                    "A decade of data (2014-2024) vs. the IAEA benchmark")

    add_rect(slide, Inches(0.5), Inches(1.5), Inches(4.2), Inches(2.5), NAVY)
    add_text(slide, Inches(0.5), Inches(1.55), Inches(4.2), Inches(0.5),
             "AVERAGE DOSE", size=12, bold=True, color=COPPER,
             align=PP_ALIGN.CENTER)
    add_text(slide, Inches(0.5), Inches(2.0), Inches(4.2), Inches(1.4),
             "59.12", size=72, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER)
    add_text(slide, Inches(0.5), Inches(3.25), Inches(4.2), Inches(0.6),
             "mrem / technician / year",
             size=14, color=WHITE, align=PP_ALIGN.CENTER)

    add_rect(slide, Inches(5.0), Inches(1.5), Inches(4.2), Inches(2.5), COPPER)
    add_text(slide, Inches(5.0), Inches(1.55), Inches(4.2), Inches(0.5),
             "FEDERAL LIMIT", size=12, bold=True, color=NAVY,
             align=PP_ALIGN.CENTER)
    add_text(slide, Inches(5.0), Inches(2.0), Inches(4.2), Inches(1.4),
             "5,000", size=72, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER)
    add_text(slide, Inches(5.0), Inches(3.25), Inches(4.2), Inches(0.6),
             "mrem occupational ceiling",
             size=14, color=WHITE, align=PP_ALIGN.CENTER)

    add_rect(slide, Inches(9.5), Inches(1.5), Inches(3.3), Inches(2.5), ACCENT)
    add_text(slide, Inches(9.5), Inches(1.55), Inches(3.3), Inches(0.5),
             "OF THE LIMIT", size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER)
    add_text(slide, Inches(9.5), Inches(2.0), Inches(3.3), Inches(1.4),
             "< 2%", size=72, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER)
    add_text(slide, Inches(9.5), Inches(3.25), Inches(3.3), Inches(0.6),
             "industrial safety margin",
             size=14, color=WHITE, align=PP_ALIGN.CENTER)

    add_text(slide, Inches(0.5), Inches(4.2), Inches(12.3), Inches(0.45),
             "What the data proves — and why Zambia can replicate it:",
             size=18, bold=True, color=NAVY)
    add_bullets(
        slide, Inches(0.7), Inches(4.7), Inches(12.0), Inches(2.2),
        [
            "Scalability: workforce grew from 1,133 to 1,167 (2021-2024); "
            "average dose stayed flat. Safety infrastructure scales.",
            "Containment of risk: 98% of exposure is external (managed by "
            "shielding); only 2% internal — validates respiratory and "
            "clean-room protocols.",
            "Targeted oversight: hot zones (Savannah River, Albuquerque) "
            "are identified and mitigated, not hidden.",
            "Alignment: IAEA Milestone Approach provides the same "
            "regulatory roadmap Zambia is already following through the "
            "Radiation Protection Authority (RPA → NSPA).",
        ],
        size=14, color=TEXT,
    )

    notes = (
        "[EVIDENCE — 1:15 to 2:00] So can a radiological system actually "
        "be trusted? The data says yes. Across a decade — 2014 to 2024 "
        "— the U.S. DOE Radiation Exposure Monitoring System recorded "
        "an average individual dose of 59.12 mrem per technician per "
        "year. The federal limit is 5,000. We are operating at less "
        "than two percent of that ceiling. And here is the line that "
        "matters most for Zambia: between 2021 and 2024 the monitored "
        "workforce grew, and the individual dose did not. Safety scales. "
        "Ninety-eight percent of exposure is external photon radiation "
        "— managed by shielding we already understand. The IAEA "
        "Milestone Approach gives us the exact same playbook, and our "
        "Radiation Protection Authority — soon the Nuclear Safety and "
        "Protection Authority — already speaks this language. We are "
        "not starting from zero. We are starting from competence."
    )
    set_notes(slide, notes)
    add_footer(slide, 3, [CITATIONS["doe"], CITATIONS["iaea"]])


def slide_need(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_header_band(slide,
                    "The Zambian Need: Drought-Proofing 2,000 MW",
                    "Why the status quo is not survivable")

    add_rect(slide, Inches(0.5), Inches(1.5), Inches(6.1), Inches(5.3),
             LIGHT)
    add_text(slide, Inches(0.7), Inches(1.65), Inches(5.7), Inches(0.5),
             "THE HYDRO TRAP", size=16, bold=True, color=ACCENT)

    add_text(slide, Inches(0.7), Inches(2.15), Inches(5.7), Inches(0.5),
             "84%", size=44, bold=True, color=NAVY)
    add_text(slide, Inches(0.7), Inches(3.05), Inches(5.7), Inches(0.5),
             "of Zambia's electricity from hydropower",
             size=13, color=TEXT)

    add_bullets(
        slide, Inches(0.7), Inches(3.6), Inches(5.7), Inches(3.0),
        [
            "2015/16 drought: 1,000 MW power deficit; "
            "~USD 440M unbudgeted electricity imports.",
            "2023-2024 Kariba crisis: prolonged load-shedding; "
            "ZESCO reports continued generation shortfalls.",
            "Drought is no longer the exception — it is the trend.",
        ],
        size=13, color=TEXT,
    )

    add_rect(slide, Inches(6.85), Inches(1.5), Inches(6.0), Inches(5.3),
             NAVY)
    add_text(slide, Inches(7.05), Inches(1.65), Inches(5.7), Inches(0.5),
             "THE NUCLEAR ANSWER", size=16, bold=True, color=COPPER)

    add_text(slide, Inches(7.05), Inches(2.15), Inches(5.7), Inches(0.5),
             "2,000 MW", size=44, bold=True, color=WHITE)
    add_text(slide, Inches(7.05), Inches(3.05), Inches(5.7), Inches(0.5),
             "national nuclear target (Vision 2030 / 7NDP)",
             size=13, color=RGBColor(0xCF, 0xD8, 0xE3))

    add_bullets(
        slide, Inches(7.05), Inches(3.6), Inches(5.7), Inches(3.0),
        [
            "Drought-proof baseload — independent of rainfall and "
            "Kariba reservoir levels.",
            "Domestic uranium reserves provide fuel-supply sovereignty.",
            "National Nuclear Policy (2020) and partnerships with "
            "ROSATOM, IP3/USA, and KAERI already in place.",
            "Closes the projected 2030 demand gap "
            "(forecast: ~27,000 MW) credibly.",
        ],
        size=13, color=WHITE,
    )

    notes = (
        "[CONFLICT DEEPENS — 2:00 to 2:30] Now bring it home. Eighty-four "
        "percent of Zambia's electricity comes from hydropower. In the "
        "2015/16 drought we lost a thousand megawatts of generation and "
        "spent roughly four hundred and forty million U.S. dollars on "
        "unbudgeted imports. The 2023-2024 ZESCO reports show the "
        "Kariba crisis is not a one-off — it is the new climate baseline. "
        "Meanwhile Vision 2030 and the Seventh National Development "
        "Plan commit us to a two thousand megawatt nuclear target, the "
        "National Nuclear Policy was approved in 2020, and partnerships "
        "with ROSATOM, IP3, and KAERI are already signed. The fuel is "
        "literally under our feet. The strategy is on paper. The "
        "missing piece is enactment."
    )
    set_notes(slide, notes)
    add_footer(slide, 4,
               [CITATIONS["zesco"], CITATIONS["ndp"]])


def slide_cta(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.08), COPPER)

    add_text(slide, Inches(0.6), Inches(0.4), Inches(12.1), Inches(0.7),
             "Call to Action: Enact the Nuclear Bill",
             size=34, bold=True, color=WHITE)
    add_text(slide, Inches(0.6), Inches(1.1), Inches(12.1), Inches(0.5),
             "Turning policy into power",
             size=18, color=COPPER)

    add_text(slide, Inches(0.6), Inches(2.0), Inches(12.1), Inches(0.6),
             "Three actions. One generation.",
             size=22, bold=True, color=WHITE)

    actions = [
        ("1. ENACT",
         "Pass the Nuclear Bill — formalise NSPA as independent regulator "
         "in line with the IAEA Milestone Approach."),
        ("2. EQUIP",
         "Fund expansion of Zambia's radiological monitoring — starting "
         "with Siavonga — to convert invisible risk into measured data."),
        ("3. EXECUTE",
         "Advance CNST and pre-feasibility for a 2,000 MW programme "
         "leveraging domestic uranium and existing partnerships."),
    ]
    y = 2.75
    for label, body in actions:
        add_rect(slide, Inches(0.6), Inches(y), Inches(2.1), Inches(0.95),
                 COPPER)
        add_text(slide, Inches(0.6), Inches(y + 0.22), Inches(2.1),
                 Inches(0.5), label, size=18, bold=True, color=NAVY,
                 align=PP_ALIGN.CENTER)
        add_rect(slide, Inches(2.8), Inches(y), Inches(10.0), Inches(0.95),
                 RGBColor(0x14, 0x3A, 0x60))
        add_text(slide, Inches(3.0), Inches(y + 0.18), Inches(9.7),
                 Inches(0.7), body, size=14, color=WHITE)
        y += 1.15

    add_rect(slide, Inches(0.6), Inches(6.2), Inches(12.1), Inches(0.65),
             COPPER)
    add_text(slide, Inches(0.6), Inches(6.30), Inches(12.1), Inches(0.5),
             "From radiological FEAR to industrial MASTERY — "
             "the choice is legislative, not technical.",
             size=15, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    notes = (
        "[RESOLUTION — 2:30 to 3:00] So what do I want you to take away? "
        "Three actions. First, enact: pass the Nuclear Bill and "
        "formalise the Nuclear Safety and Protection Authority as an "
        "independent regulator under the IAEA Milestone Approach. "
        "Second, equip: fund the radiological monitoring expansion, "
        "starting in Siavonga, so that no community lives next to "
        "uranium without instruments. Third, execute: advance the "
        "Centre for Nuclear Science and Technology and the pre-feasibility "
        "studies for the two thousand megawatt programme. The "
        "fifty-nine point one two mrem U.S. record proves the safety "
        "model is real. Our 2020 Nuclear Policy proves the political "
        "will is real. Our uranium proves the fuel is real. The only "
        "thing standing between Zambia and a drought-proof grid is the "
        "signature on a bill. Let us move from fear to mastery. "
        "Thank you."
    )
    set_notes(slide, notes)
    add_footer(slide, 5,
               [CITATIONS["ndp"], CITATIONS["iaea"], CITATIONS["siavonga"]])


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide_title(prs)
    slide_hook(prs)
    slide_safety(prs)
    slide_need(prs)
    slide_cta(prs)

    prs.save(OUTPUT)
    print(f"Generated: {OUTPUT}")


if __name__ == "__main__":
    build()
