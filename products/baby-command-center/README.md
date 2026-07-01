# Baby Command Center™ — The Ultimate Baby Care & Family Organization System

> Not a baby tracker — a **family baby operating system**. One warm, premium Excel & Google Sheets dashboard for feeding, sleep, diapers, growth, health, budget, routines & memories.

| | |
| - | - |
| **Product** | Baby Command Center™ |
| **Target** | First-time & growing families · stay-at-home / working parents · adoptive & foster parents · grandparents & nannies |
| **Angle** | Reduce mental load in the busiest season — one organized home for everything baby. |
| **Formats** | Excel `.xlsx` (21-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $24 single · **$34 bundle** · $49 with onboarding · $99 caregiver/creator license |

> ⚠️ **Organization & record-keeping tool only.** Baby Command Center™ is
> **not** a medical device or a substitute for professional healthcare. The
> disclaimer ships on the **Welcome** tab.

---

## Contents

```
products/baby-command-center/
├── README.md
├── Baby_Command_Center.xlsx        ← Excel master (21-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 21-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Executive Baby Dashboard | 12 | Shopping & Baby Supplies |
| 2 | Baby Profile | 13 | Baby Inventory |
| 3 | Daily Baby Log | 14 | Childcare Organizer |
| 4 | Feeding Command Center | 15 | Routine Planner |
| 5 | Sleep Tracker | 16 | Development Activities |
| 6 | Diaper Tracker | 17 | Memory Book (photos) |
| 7 | Growth Tracker | 18 | Travel with Baby |
| 8 | Milestone Tracker | 19 | Family Goals |
| 9 | Medical Center | 20 | Analytics Dashboard |
| 10 | Appointment Calendar | 21 | Settings |
| 11 | Baby Budget | | |

*(+ a Welcome / Start-Here tab — 22 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Baby's Age (days) | `=TODAY()-BirthDate` |
| Feedings Today | `=COUNTIF(FeedDate,TODAY())` |
| Sleep Today (hrs) | `=SUMIF(SleepDate,TODAY(),SleepHrs)` |
| Diapers Today | `=COUNTIF(DiaperDate,TODAY())` |
| Next Appointment | `MINIFS(ApptDate, …) − TODAY()` |
| Milestones logged | achieved ÷ total (by date) |
| Vaccines done | done ÷ scheduled (Medical) |
| Budget Remaining | `Planned − Actual` |
| Goals / Routine | `AVERAGE(GoalProgress)` · routine checklist % |

Each tracker rolls its own **7-day trend** (feedings/day, sleep hours/day,
diaper counts) into the dashboard charts, plus a **Family Organization Score**
on the Analytics tab. **48 named ranges**, blank-safe formulas, cleanly-placed
charts.

---

## Warm, premium, baby-friendly design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + a gentle executive dashboard
- Milestone "achieved" rows glow mint; reorder-needed & due items flag gold
- Image-placeholder **Memory Book** (birth, first smile, first steps…)
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Baby_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
