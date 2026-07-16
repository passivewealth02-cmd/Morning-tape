# Tot-School Command Center™ — The Complete At-Home Tot-School System

> Not a printable pack — a **complete at-home tot-school operating system** for
> toddlers (roughly 18 months to 3 years). One premium **Google Sheets +
> printable PDF** command center for toddler milestones, first words & concepts,
> gentle weekly themes, a tot-tray & sensory activity planner, board-book
> read-alouds, messy play, outings, a portfolio of firsts and a
> ready-for-preschool checklist. No pressure, all play.

| | |
| - | - |
| **Product** | Tot-School Command Center™ |
| **Target** | At-home tot-school & toddler families · homeschool & preschool-prep parents · first-time & experienced moms · play-based / Montessori-inspired / gentle · nannies & in-home care · anyone loving the toddler years |
| **Angle** | No pressure, all play — gentle rhythms, tot trays & a keepsake of these tiny years. |
| **Formats** | Google Sheets (17-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $14 single · **$22 bundle** (Sheets + PDF) · $28 with preschool add-on · $69 provider / commercial license |

---

## Contents

```
products/tot-school-command-center/
├── README.md
├── Tot_School_Command_Center.xlsx   ← Google Sheets / Excel master (17 tabs)
├── Tot_School_Printables.pdf        ← 12-page print-ready pack (US Letter)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    ├── build_pdf.py
    ├── build_marketing.py
    └── build_marketing_detail.py
```

---

## The 17-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 10 | Sensory Play |
| 2 | Dashboard | 11 | Outings |
| 3 | Tot Profiles | 12 | Attendance |
| 4 | Milestones | 13 | Portfolio |
| 5 | First Words & Concepts | 14 | Supplies |
| 6 | Weekly Themes | 15 | Ready for Preschool |
| 7 | Daily Rhythm | 16 | Goals |
| 8 | Tot Trays | 17 | Settings |
| 9 | Board-Book Log | | |

## The 12 printable PDF pages

Tot-at-a-Glance · Weekly Theme Plan · Daily Rhythm · Milestones Checklist ·
First Words & Concepts · Tot-Tray Planner · Board-Book Log · Outings & Nature ·
Sensory Play Bank · Ready-for-Preschool Checklist · Tot Goals · Portfolio of
Firsts.

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Tots | `=Tots` |
| Tot-School Days | `=DaysDone` (of `DaysGoal`) |
| Words Spoken | `=Words` |
| Milestones Met | `=COUNTIF(MileStatus,"Met")` |
| Themes Done | `=COUNTIF(ThemeStatus,"Done")` |
| Trays Done | `=COUNTIF(TrayStatus,"Done")/COUNTA(TrayName)` |
| Board Books | `=BoardBooks` |
| Outings | `=COUNTA(TripName)` |
| Supplies Spent | `=SuppliesSpent` |
| Preschool-Ready | `=COUNTIF(RdyDone,"Yes")/COUNTA(RdyName)` |
| Goals | `=AVERAGE(GoalProgress)` |
| Growing Score | `=AVERAGE(HealthRange)` |

Marking a milestone Met, a first word/concept got, a tray Done, or a readiness
item Yes all update the dashboard live; the **Growing Score** blends milestones,
words & concepts, themes, trays, readiness and goals into one 0–100% "how they're
growing" number. A **Milestones-Met-by-Domain** bar chart shows the whole-child
balance at a glance.

**Verified sample tot-school** (the Bennetts — Theo, 2, and Nora, 14 months):
Tots **2** · Tot-school days **72** of 120 · Words spoken **45** · Milestones met
**26** of 35 · Themes done **5** of 12 · Trays **62%** done · Board books **24** ·
Outings **6** · Supplies **$259** of $300 · Preschool-ready **88%** · Goals
**71%** · **Growing Score 69%**.

---

## Premium tot-school design

- A true dashboard: 12 KPIs, a "Growing Well?" bars panel, a themes donut, a
  "coming up" list, a Growing gauge & a Milestones-by-Domain bar chart
- A whole-child **Milestones** tracker across 6 domains with Met / Emerging /
  Not Yet color-coding
- A **First Words & Concepts** tracker (colors, shapes, body parts, animal
  sounds, first counting) with a live "concepts got" count and a words counter
- Status color-coding, data-bars & conditional flags throughout; play-based and
  pressure-free — never a test
- **Print-ready PDF pack** on white with a forest-green header band & gold rules
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **Every toddler grows on their own timeline.** The milestones and readiness
> lists are gentle guides, never a test or a diagnosis. This is a planning &
> keepsake tool — if you ever have concerns, talk to your pediatrician.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Tot_School_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Tot_School_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
