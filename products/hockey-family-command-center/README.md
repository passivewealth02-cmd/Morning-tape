# Hockey Family Command Center™ — The Ultimate Hockey Season Operating System

> Not a planner — a **hockey-season operating system**. One premium Excel & Google Sheets dashboard for scheduling, budgeting, equipment, travel, stats, development, nutrition & team organization.

| | |
| - | - |
| **Product** | Hockey Family Command Center™ |
| **Target** | Hockey moms & dads · travel / AAA families · youth hockey parents · team managers · coaches |
| **Angle** | Eliminate stress, prevent missed events, control costs — run the whole season from one place. |
| **Formats** | Excel `.xlsx` (25-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $29 single · **$39 bundle** · $59 with onboarding · $149 team/coach license |

---

## Contents

```
products/hockey-family-command-center/
├── README.md
├── Hockey_Family_Command_Center.xlsx   ← Excel master (25-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 25-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Executive Hockey Dashboard | 14 | Team Roster |
| 2 | Player Profile | 15 | Player Development |
| 3 | Master Season Calendar | 16 | Stats Tracker |
| 4 | Game Day Command Center | 17 | Nutrition Planner |
| 5 | Practice Tracker | 18 | Medical Center |
| 6 | Hockey Budget (26 categories) | 19 | Packing Checklist |
| 7 | Equipment Command Center | 20 | Fundraising Tracker |
| 8 | Skate Sharpening Log | 21 | Team Communication |
| 9 | Stick Inventory | 22 | Photo & Memory Gallery |
| 10 | Tournament Command Center | 23 | Hockey Goals |
| 11 | Travel Planner | 24 | Analytics Dashboard |
| 12 | Hotel Planner | 25 | Settings |
| 13 | Carpool Manager | | |

*(+ a Welcome / Start-Here tab — 26 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Days to Next Game | `MINIFS(CalDate, CalType,"Game", CalDate,">="&TODAY()) − TODAY()` |
| Practices This Week | `COUNTIFS(CalType,"Practice", CalDate, this-week)` |
| Games This Month | `COUNTIFS(CalType,"Game", CalDate, next-30-days)` |
| Next Tournament (days) | `MINIFS` over tournaments |
| Budget Remaining | `BudgetTotalPlanned − BudgetTotalActual` |
| Equipment Alerts | gear with `Replace By ≤ TODAY()+45` |
| Ice Time (30d) | `SUMIFS(PracIce, …)` |
| Schedule Conflicts | distinct calendar days with 2+ events |
| Attendance % | `COUNTIF(PracAttend,"Present") / practices` |
| Travel Miles | `SUM(TravelMiles)` |
| Season Progress % | games played ÷ total games |

Plus a blended **Hockey Readiness Score** on Analytics. **57 named ranges**
keep everything wired; blank-safe totals avoid broken cells; all charts are
cleanly placed.

---

## Premium, cohesive design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + an executive dashboard
- Status color-coding (Present/New = mint, replace-soon = gold, overdue = red)
- Auto flags: conflicting events, equipment due, sharpening reminders, countdowns
- Image-placeholder gallery (team photos, awards, championships)
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Hockey_Family_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
