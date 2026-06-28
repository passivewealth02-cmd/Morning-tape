# Back-to-School Command Center (BTSCC)

> The all-in-one dashboard to plan, budget, organize, and prepare for a stress-free school year.

| | |
| - | - |
| **Product** | Back-to-School Command Center |
| **Target** | Parents · homeschool families · college & high school students · teachers · guardians |
| **Angle** | Premium Excel + Google Sheets planner combining budgeting, shopping, scheduling, meal planning, and academic prep in one organized system. |
| **Formats** | Excel `.xlsx` + Google Sheets edition |
| **Pricing** | $19 (single) · **$29 bundle** · $79 with onboarding |

---

## Contents

```
products/back-to-school-command-center/
├── README.md                              ← this file
├── Back_To_School_Command_Center.xlsx     ← Excel master (ship this)
├── GOOGLE_SHEETS.md                       ← Google Sheets formulas
├── BUILD_INSTRUCTIONS.md                  ← reproduction + ship guide
└── build/
    └── build_xlsx.py                      ← deterministic .xlsx generator
```

---

## Workbook architecture — 12 sheets

| # | Sheet | Role | Tab color |
| - | ----- | ---- | --------- |
| 1 | **Dashboard** | 8 KPI cards, 2 charts, quick-nav row | Primary `#1B4F48` |
| 2 | **Students** | Profile cards per child (grade, allergies, contacts) | Accent `#937356` |
| 3 | **Supplies** | 60-row tracker — qty, store, price, priority, purchased flag | Highlight `#75E6C1` |
| 4 | **Budget** | 9 categories with `SUMIFS` actuals pulled from Supplies, donut + bar | Primary `#1B4F48` |
| 5 | **Clothing** | Inventory: size, qty needed, owned, cost, status formula | Accent `#937356` |
| 6 | **Schedule** | Mon–Fri × 8-period weekly grid, color-coded by subject | Highlight `#75E6C1` |
| 7 | **Assignments** | 50-row tracker with Days-Until + overdue/complete formatting | Primary `#1B4F48` |
| 8 | **Extracurricular** | Activity, coach, practice times, fees, equipment | Accent `#937356` |
| 9 | **Calendar** | Event log with upcoming-14-day highlight | Highlight `#75E6C1` |
| 10 | **Meals** | Mon–Fri × breakfast/lunch/snack grid + grocery list | Accent `#937356` |
| 11 | **Emergency** | Family · Medical · School · Insurance — printable | Danger `#C94C4C` |
| 12 | **Settings** | Inputs (budget, first-day, year), all dropdown lists | Surface `#E5D3BA` |

---

## KPI cards driven by formulas

| KPI | Formula |
| --- | ------- |
| Total Budget | `=TotalBudget` |
| Total Spent | `=BudgetTotalActual` |
| Remaining | `=TotalBudget-BudgetTotalActual` |
| Days Until School | `=MAX(FirstDayOfSchool-TODAY(),0)` |
| Items Purchased | `=SUMPRODUCT((SupPurchased="Yes")*1)` |
| Items Remaining | `=SUMPRODUCT((SupItem<>"")*(SupPurchased<>"Yes"))` |
| Supply Completion % | `=IFERROR(.../SUMPRODUCT((SupItem<>"")*1),0)` |
| Assignments Done % | `=IFERROR(.../SUMPRODUCT((AsgStatus<>"")*1),0)` |

---

## Formula engine highlights

- `SUMIFS` budget-actuals from Supplies, per category
- `SUMPRODUCT` for purchase / completion ratios (handles blanks)
- `TODAY()`-driven countdown + days-until-due (recalculates every open)
- 60-row Supplies table with auto `Remaining = Required − Bought`
- 50-row Assignment tracker with three live conditional-format states
- 28 named ranges so the Google Sheets companion mirrors 1-for-1
- Doughnut + Column charts on Dashboard and Budget, all chart series
  expand automatically as you add rows

---

## Brand system applied

| Token | Hex | Where it shows |
| ----- | --- | -------------- |
| Primary | `#1B4F48` | Page headers, table headers, KPI values, navigation chips |
| Accent | `#937356` | KPI labels, field labels, warning tones |
| Surface | `#E5D3BA` | Input cells (Settings, Budget Planned column), totals row |
| Highlight | `#75E6C1` | Positive KPIs, purchased/complete row highlight |
| Danger | `#C94C4C` | Overdue assignment rows, Emergency tab |
| Soft BG | `#FAF7F1` | Field-value backgrounds, schedule grid base |

---

## Build & ship

```bash
cd build && python3 build_xlsx.py
```

Output: `../Back_To_School_Command_Center.xlsx`. See
`BUILD_INSTRUCTIONS.md` for the Google Sheets path, Etsy delivery
package, listing photo plan, pricing tiers, and maintenance protocol.

> **Marketing images** (`marketing/`) are generated separately by
> `build/build_marketing.py` after the workbook is confirmed. See the
> DWBS product folder for the established Driver-Budget thumbnail
> format and 5 supporting listing photos.
