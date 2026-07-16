# Preschool Command Center™ — The Complete At-Home Preschool System

> Not a printable pack — a **complete at-home preschool operating system**. One
> premium **Google Sheets + printable PDF** command center for weekly themes, a
> play-based activity planner, a whole-child skills & milestones tracker, ABC &
> 123 mastery, read-alouds, arts & sensory, nature walks, a portfolio and a
> kindergarten-readiness checklist. Play-based and gentle, ages 3–5, one child or
> a houseful.

| | |
| - | - |
| **Product** | Preschool Command Center™ |
| **Target** | At-home preschool & pre-K families · homeschool & tot-school parents · first-time & experienced moms · play-based / Montessori-inspired / eclectic · nannies & in-home care · anyone prepping for kindergarten |
| **Angle** | Play with a plan — themes, skills & sweet memories, all in one calm system. |
| **Formats** | Google Sheets (17-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $14 single · **$22 bundle** (Sheets + PDF) · $28 with tot-school add-on · $69 provider / commercial license |

---

## Contents

```
products/preschool-command-center/
├── README.md
├── Preschool_Command_Center.xlsx   ← Google Sheets / Excel master (17 tabs)
├── Preschool_Printables.pdf        ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 10 | Arts & Sensory |
| 2 | Dashboard | 11 | Field Trips |
| 3 | Child Profiles | 12 | Attendance |
| 4 | Skills & Milestones | 13 | Portfolio |
| 5 | ABC & 123 | 14 | Supplies |
| 6 | Weekly Themes | 15 | Kindergarten Readiness |
| 7 | Daily Rhythm | 16 | Goals |
| 8 | Activity Planner | 17 | Settings |
| 9 | Read-Aloud Log | | |

## The 12 printable PDF pages

Child-at-a-Glance · Weekly Theme Plan · Daily Rhythm · Skills & Milestones ·
ABC & 123 Chart · Activity Planner · Read-Aloud Log · Nature & Field Trips ·
Arts & Sensory Idea Bank · Kindergarten Readiness Checklist · Preschool Goals ·
Portfolio & Keepsakes.

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Children | `=Children` |
| Preschool Days | `=DaysDone` (of `DaysGoal`) |
| Letters Known | `=COUNTIF(AbcRecognize,"Yes")` (of 26) |
| Skills Mastered | `=COUNTIF(SkillStatus,"Mastered")` |
| Themes Done | `=COUNTIF(ThemeStatus,"Done")` |
| Activities Done | `=COUNTIF(ActStatus,"Done")/COUNTA(ActName)` |
| Books Read | `=BooksRead` |
| Field Trips | `=COUNTA(TripName)` |
| Supplies Spent | `=SuppliesSpent` |
| Ready for K | `=COUNTIF(RdyDone,"Yes")/COUNTA(RdyName)` |
| Goals | `=AVERAGE(GoalProgress)` |
| Ready Score | `=AVERAGE(HealthRange)` |

Marking a letter known, a skill Mastered, an activity Done, or a readiness item
Yes all update the dashboard live; the **Ready-for-Kindergarten Score** blends
letters, skills, themes, activities, readiness and goals into one 0–100% "how
they're growing" number. A **Skills-Mastered-by-Domain** bar chart shows the
whole-child balance at a glance.

**Verified sample preschool** (the Bennetts — Millie, 4, and Owen, 3):
Children **2** · Preschool days **96** of 150 · Letters known **24** of 26 ·
Skills mastered **25** of 35 · Themes done **7** of 14 · Activities **62%** done ·
Books **38** · Field trips **7** · Supplies **$350** of $400 · Ready for K
**75%** · Goals **71%** · **Ready Score 70%**.

---

## Premium preschool-software design

- A true dashboard: 12 KPIs, a "Growing Well?" bars panel, a themes donut, a
  "coming up" list, a Ready-for-K gauge & a Skills-by-Domain bar chart
- A whole-child **Skills & Milestones** tracker across 7 domains with
  Mastered / Emerging / Not Yet color-coding
- An **ABC & 123** mastery tab (sees / says / writes per letter) with a live
  "letters known of 26" count
- Status color-coding, data-bars & conditional flags throughout; play-based and
  gentle — never a test
- **Print-ready PDF pack** on white with a forest-green header band & gold rules
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **Every child grows at their own pace.** The readiness checklist is a gentle
> guide, not a test or a diagnosis. This is a planning & keepsake tool.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Preschool_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Preschool_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
