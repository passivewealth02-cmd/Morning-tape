# Home Command Center™ — The Ultimate Household Management System

> Not a planner — a **household operating system**. One premium Excel & Google Sheets dashboard for finances, schedules, meals, cleaning, maintenance, inventories, routines and long-term goals.

| | |
| - | - |
| **Product** | Home Command Center™ |
| **Target** | Stay-at-home & work-from-home parents · busy families · homeschoolers · caregivers · newly married couples · first-time homeowners |
| **Angle** | Reduce mental load — one organized place to run the entire home. |
| **Formats** | Excel `.xlsx` (28-tab system + Welcome) + Google Sheets edition |
| **Pricing** | $29 single · **$39 bundle** · $59 with onboarding · $149 pro/creator license |

---

## Contents

```
products/home-command-center/
├── README.md
├── Home_Command_Center.xlsx     ← Excel master (28-tab system + Welcome)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    └── build_marketing.py
```

---

## The 28-sheet system

| # | Sheet | # | Sheet |
| - | ----- | - | ----- |
| 1 | Executive Home Dashboard | 15 | Home Inventory |
| 2 | Family Directory | 16 | Subscription Tracker |
| 3 | Master Family Calendar | 17 | Family Goals |
| 4 | Daily Command Center | 18 | Savings Planner |
| 5 | Household Budget | 19 | Holiday & Event Planner |
| 6 | Bill Command Center | 20 | Travel Planner |
| 7 | Grocery Planner | 21 | Children's Hub |
| 8 | Pantry Inventory | 22 | Pet Care Center |
| 9 | Fridge & Freezer Inventory | 23 | Wellness & Self-Care |
| 10 | Meal Planner | 24 | Home Project Planner |
| 11 | Recipe Library | 25 | Document Vault Index |
| 12 | Cleaning Command Center | 26 | Analytics Dashboard |
| 13 | Chore Manager | 27 | Home Inspiration Board |
| 14 | Home Maintenance | 28 | Settings |

*(+ a Welcome / Start-Here tab — 29 tabs total.)*

---

## Signature automation (10 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Household Budget Remaining | `=BudgetTotalPlanned-BudgetTotalActual` |
| Monthly Bills Paid | `=COUNTIF(BillPaid,"Yes")` |
| Cleaning Tasks Completed | `=COUNTIF(CleanStatus,"Done")` |
| Pantry Stock Level | `=SUMPRODUCT((PantryName<>"")*(PantryQty>PantryReorder))/COUNTA(PantryName)` |
| Calendar Events This Week | `=SUMPRODUCT((CalDate>=TODAY())*(CalDate<=TODAY()+7))` |
| Meal Plan Completion | filled meal slots ÷ 35 |
| Savings Goal Progress | `=SUM(SavCurrent)/SUM(SavTarget)` |
| Household Projects Active | `=COUNTIF(ProjStatus,"In Progress")` |
| Maintenance Tasks Due | `=SUMPRODUCT((MaintNext<=TODAY()+30)*…)` |
| Family Routine Score | `=AVERAGE(ChorePct)` |

Plus a blended **Household Health Score** on the Analytics tab. **67 named
ranges** keep everything wired; blank-safe totals avoid broken cells; all
charts are cleanly placed (no overlaps, no jumbled labels).

---

## Premium, cohesive design

- Two-row **gold-divider headers** on every tab
- Gold-topped white **KPI cards** + an executive dashboard
- Status color-coding (Done = mint, overdue/low-stock = red, due-soon = gold)
- Automatic flags: overdue bills, low-stock pantry, expiring food, maintenance due
- Image-placeholder boards (Inspiration + before/after project photos)
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

---

## Build & ship

```bash
cd build && python3 build_xlsx.py        # -> ../Home_Command_Center.xlsx
python3 build_marketing.py               # -> ../marketing/01..06.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
