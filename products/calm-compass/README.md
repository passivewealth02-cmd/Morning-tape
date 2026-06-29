# Calm Compass™ — Anxiety & Social Confidence Planner

> A calm, premium wellness operating system in Excel & Google Sheets — build routines, notice patterns, prepare for stressful moments, and celebrate progress.

| | |
| - | - |
| **Product** | Calm Compass™ |
| **Target** | Adults with everyday/social anxiety · college students · young professionals · remote workers · anyone building healthier routines |
| **Angle** | Not just a journal — a cohesive wellness dashboard combining habits, routines, mood, reflection, goals & progress in one system. |
| **Formats** | Excel `.xlsx` (15 sheets) + Google Sheets edition |
| **Pricing** | $17 single · **$24 bundle** · $39 with onboarding mini-course |

> ⚠️ **Wellness & organization tool only.** Calm Compass™ is not a medical
> device and is not a substitute for professional mental health care. The
> disclaimer ships on the **Welcome** tab.

---

## Contents

```
products/calm-compass/
├── README.md                 ← this file
├── Calm_Compass.xlsx         ← Excel master (15 sheets)
├── GOOGLE_SHEETS.md          ← Google Sheets formulas
├── BUILD_INSTRUCTIONS.md     ← reproduction + ship guide
├── ETSY_LISTING.md           ← SEO-optimized listing
└── build/
    ├── build_xlsx.py         ← deterministic workbook generator
    └── build_thumbnail.py    ← hero thumbnail generator
```

---

## The 15-sheet system

| # | Sheet | Role |
| - | ----- | ---- |
| 1 | **Welcome** | Warm onboarding + **mental-health disclaimer** |
| 2 | **Dashboard** | 8 KPI cards + mood/energy/sleep trends + habit consistency |
| 3 | **Daily Check-In** | mood, energy, stress, sleep, water, exercise, mindful min, journaled |
| 4 | **Habits** | 9-habit tracker with auto streaks, daily % & consistency summary |
| 5 | **Daily Planner** | priorities, time blocks, self-care, breaks, notes (printable) |
| 6 | **Social Prep** | plan + reflect on social situations |
| 7 | **Reflection** | 5-prompt daily journal |
| 8 | **Goals** | progress bars + status, by category |
| 9 | **Self-Care** | 6 categories, weekly frequency tracking |
| 10 | **Sleep** | bedtime/wake/quality + weekly sleep-trend chart |
| 11 | **Exercise** | activity + mood-before/after |
| 12 | **Gratitude** | daily gratitude entries |
| 13 | **Resources** | books, podcasts, affirmations, support contacts |
| 14 | **Progress** | wellness score + snapshot + trends |
| 15 | **Settings** | preferences + dropdown lists |

---

## Signature automation

| Metric | How it's computed |
| ------ | ----------------- |
| Avg Mood / Energy / Sleep | `AVERAGE` over the Daily Check-In |
| Habits Completed | `SUM` of daily completed counts |
| Routine Score | `AVERAGE` of daily habit % |
| Journaling Streak | trailing-consecutive helper + `LOOKUP` |
| Best Habit Streak | helper streak column + `MAX` |
| Goals Completed | `COUNTIF(status,"Complete")` |
| **Wellness Score** | gentle blend of mood, routine, sleep & journaling (a guide, not a grade) |

33 named ranges keep the dashboards wired and make the Google Sheets
edition a 1-to-1 mirror. All totals use blank-safe functions (no broken
`#VALUE!` cells), and every pie/bar/line chart is cleanly placed.

---

## Calm, premium design

- Spacious two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards**
- Soft mood color-scales (low = gentle red → high = mint), reversed for stress
- Calm palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`
- Supportive, non-clinical copy throughout

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Calm_Compass.xlsx
python3 build_thumbnail.py               # -> ../marketing/01_hero.png
```

See `BUILD_INSTRUCTIONS.md` for the Google Sheets path and Etsy delivery
package, and `ETSY_LISTING.md` for the SEO-optimized listing.
