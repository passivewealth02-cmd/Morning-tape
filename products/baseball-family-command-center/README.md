# Baseball Family Command Center™ — The Ultimate Baseball Season Operating System

> Not a planner — a **baseball-season operating system**. One premium Excel & Google Sheets workbook for scheduling, budgeting, statistics, pitching, equipment, travel, development & team organization.

| | |
| - | - |
| **Product** | Baseball Family Command Center™ |
| **Target** | Baseball moms & dads · Little League / travel / select / club families · team managers · coaches · multi-player families |
| **Angle** | Organize every part of the season, control costs, and track player growth — from one place. |
| **Formats** | Excel `.xlsx` (24-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $29 single · **$39 bundle** · $59 with onboarding · $149 team/coach license |

---

## Contents

```
products/baseball-family-command-center/
├── README.md
├── Baseball_Family_Command_Center.xlsx   ← Excel master (24-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 24-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Executive Baseball Dashboard | 13 | Travel Planner |
| 2 | Player Profile | 14 | Team Roster |
| 3 | Master Season Calendar | 15 | Snack & Volunteer Schedule |
| 4 | Game Day Command Center | 16 | Player Development |
| 5 | Practice Planner | 17 | Nutrition & Hydration |
| 6 | Baseball Budget (20 categories) | 18 | Medical Center |
| 7 | Equipment Command Center | 19 | Packing Checklist |
| 8 | Bat Inventory | 20 | Fundraising Tracker |
| 9 | Glove Care Log | 21 | Photo & Memory Gallery |
| 10 | Player Statistics (AVG/OBP/SLG/OPS) | 22 | Baseball Goals |
| 11 | Pitching Tracker (pitch count + ERA) | 23 | Analytics Dashboard |
| 12 | Tournament Command Center | 24 | Settings |

*(+ a Welcome / Start-Here tab — 25 tabs total.)*

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Days to Next Game | `MINIFS(CalDate, CalType,"Game", …) − TODAY()` |
| Games / Practices This Week | `COUNTIFS` on the calendar |
| Next Tournament (days) | `MINIFS` over tournaments |
| Budget Remaining | `BudgetTotalPlanned − BudgetTotalActual` |
| Equipment Alerts | gear with `Replace By ≤ TODAY()+45` |
| Games Played / Season % | games before today ÷ total games |
| Travel Miles | `SUM(TravelMiles)` |
| Volunteer Hours | `SUM(VolHours)` |
| Dev Progress | `AVERAGE(DevProgress)` |

**Batting stats auto-calculate** — AVG, OBP, SLG & OPS from your game log;
**ERA & pitch counts** from your pitching log (with your own pitch-count
limit alerts). Plus a blended **Baseball Readiness Score** on Analytics.
**65 named ranges**, blank-safe totals, cleanly-placed charts.

---

## Premium, cohesive design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + an executive dashboard
- Status color-coding (Present/New = mint, replace-soon = gold, over pitch limit = red)
- Auto flags: conflicting events, equipment due, pitch-count warnings, countdowns
- Image-placeholder gallery (team photos, awards, championships)
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Baseball_Family_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
