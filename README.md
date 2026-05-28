# The Case for Nuclear Power in Zambia

A data-driven advocacy deck arguing that Zambia must leverage its existing
radiological safety expertise and uranium resources to transition to nuclear
power as a drought-proof solution for national energy security.

**Prepared by:** Elizabeth Mwaba

---

## Big Idea

> Zambia must leverage its existing radiological safety expertise and uranium
> resources to transition to nuclear power as a drought-proof solution for
> national energy security.

## Narrative arc (3-minute story)

Built on Cole Nussbaumer Knaflic's *Storytelling with Data* plot structure
(Setup → Conflict → Resolution):

| # | Slide | Beat |
|---|----------------------------------------------|-----------|
| 1 | Title — From Radiological Fear to Industrial Mastery | Setup |
| 2 | The Hook — Siavonga's Invisible Ghost vs. Data Mastery | Conflict |
| 3 | Safety Evidence — 59.12 mrem vs. IAEA standards | Evidence |
| 4 | The Zambian Need — Drought-proofing the 2,000 MW goal | Stakes |
| 5 | Call to Action — Enact the Nuclear Bill | Resolution |

Each slide's speaker notes contain a ~30-second portion of the 3-minute script.

## Repository layout

```
zambia-nuclear-case/
├── build_presentation.py        # python-pptx generator (run this)
├── requirements.txt
├── README.md
├── The_Case_for_Nuclear_Power_in_Zambia.pptx   # generated deliverable
├── data_sources/
│   ├── Summary.pdf              # US radiological safety summary (59.12 mrem proof)
│   └── Analysis.pbix            # Power BI analysis backing Summary.pdf
└── reference_materials/
    ├── 2b.9-Chewe.pdf
    ├── 4.6.Kangwa-ZAMBIA NUCLEAR PROGRAM.pdf
    ├── MINISTERIAL STATEMENT BY THE  MINISTER OF HIGHER EDUCATION.pdf
    └── Stroytelling with data.pdf
```

> **Rule observed:** all research PDFs, the PBIX file, and helper scripts are
> preserved in the repo. Nothing has been deleted.

## How to build the deck

```powershell
python -m pip install -r requirements.txt
python build_presentation.py
```

This produces `The_Case_for_Nuclear_Power_in_Zambia.pptx` at the repo root.

## Citations used in slide footers and speaker notes

- **Siavonga Uranium Baseline Survey (2024)** — local unmonitored exposure
- **ZESCO Annual Reports 2023-2024** — Kariba crisis and ~1,000 MW deficit
- **Zambia 7th National Development Plan (7NDP) & Vision 2030** — national strategy
- **IAEA Milestone Approach for New Nuclear Programs** — regulatory roadmap
- **U.S. DOE REMS 2024 Report** — 59.12 mrem occupational safety benchmark

## Key data points

| Metric | Value | Source |
|---|---|---|
| Avg. individual radiation dose, US nuclear enterprise | 59.12 mrem/yr | DOE REMS 2024 |
| Federal occupational limit | 5,000 mrem/yr | DOE REMS 2024 |
| Margin vs. limit | < 2% | derived |
| Zambia hydropower share of generation | 84% | Kangwa, INPRO 2025 |
| 2015/16 drought deficit | ~1,000 MW | ZESCO / Kangwa |
| 2015/16 unbudgeted import cost | ~USD 440M | ZESCO / Kangwa |
| Zambia nuclear target | 2,000 MW | Ministerial Statement; 7NDP |
| National Nuclear Policy approved | November 2020 | Kangwa, INPRO 2025 |
