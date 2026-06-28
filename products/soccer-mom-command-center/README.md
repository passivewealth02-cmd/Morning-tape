# Soccer Mom Command Center™ (SMCC)

> The ultimate family dashboard for managing games, practices, budgets, travel, meals, equipment, and team life — all in one premium Excel & Google Sheets system.

| | |
| - | - |
| **Product** | Soccer Mom Command Center™ |
| **Target** | Soccer moms · soccer dads · team managers · youth coaches · busy families with multiple children · travel & club soccer families |
| **Angle** | Replaces a dozen apps, paper calendars, and group texts with one polished season-long source of truth. |
| **Formats** | Excel `.xlsx` + Google Sheets edition |
| **Pricing** | $24 (single) · **$34 bundle** · $89 with onboarding · $149 team license |

---

## Contents

```
products/soccer-mom-command-center/
├── README.md                              ← this file
├── Soccer_Mom_Command_Center.xlsx         ← Excel master (ship this)
├── GOOGLE_SHEETS.md                       ← Google Sheets formulas
├── BUILD_INSTRUCTIONS.md                  ← reproduction + ship guide
└── build/
    └── build_xlsx.py                      ← deterministic .xlsx generator
```

---

## Workbook architecture — 14 sheets

| # | Sheet | Role | Tab color |
| - | ----- | ---- | --------- |
| 1  | **Dashboard** | 8 KPI cards + 2 charts + 8 quick-nav chips | Primary `#1B4F48` |
| 2  | **Players** | Profile cards per child (13 fields each) | Accent `#937356` |
| 3  | **Schedule** | 40-row season match tracker · Days-Out · Win/Loss tint · upcoming-7-day highlight | Highlight `#75E6C1` |
| 4  | **Practices** | 50-row planner · attendance dropdown (Present/Late/Absent/Excused) · 3 conditional states | Highlight `#75E6C1` |
| 5  | **Budget** | 12 categories · variance + % spent · KPI sidebar (cost/month, cost/player) · donut + bar | Primary `#1B4F48` |
| 6  | **Equipment** | 30-row tracker · `Replace By` 60-day amber alert · condition + purchased flags | Accent `#937356` |
| 7  | **Tournaments** | 12-row planner · upcoming-30-day mint highlight · hotel confirmation + entry fee | Primary `#1B4F48` |
| 8  | **Carpool** | 30-row rotation · driver · players · pickup/return · fuel share | Accent `#937356` |
| 9  | **Roster** | 20-row team list · parent, phone, email, position, birthday | Highlight `#75E6C1` |
| 10 | **Meals** | Pre-game · post-game · tournament · snack roster · hydration tracker · shopping list | Accent `#937356` |
| 11 | **Packing** | Six pre-built checklists side-by-side (Practice → Winter Training) | Accent `#937356` |
| 12 | **Mileage** | 40-row travel log · auto Fuel = miles × FuelPerMile · auto Total | Highlight `#75E6C1` |
| 13 | **Communication** | Coach messages · team events · volunteer roster · fundraisers · tasks | Primary `#1B4F48` |
| 14 | **Settings** | Season, budget, fuel rate + 12 named dropdown lists | Surface `#E5D3BA` |

---

## KPI cards driven by formulas

| KPI | Formula |
| --- | ------- |
| Games This Month | `=SUMPRODUCT((GameDates<>"")*(MONTH(...)=MONTH(TODAY()))*(YEAR(...)=YEAR(TODAY())))` |
| Practices This Week | `=SUMPRODUCT((PracDates<>"")*(PracDates-TODAY()>=0)*(PracDates-TODAY()<=7))` |
| Monthly Budget | `=MonthlyBudget` |
| Budget Remaining | `=MonthlyBudget-BudgetTotalActual` |
| Upcoming Tournaments | `=SUMPRODUCT((TourStart<>"")*(TourStart-TODAY()>=0))` |
| Equipment Needed | `=SUMPRODUCT((EquipItem<>"")*(EquipPurchased<>"Yes"))` |
| Family Conflicts | `=SUMPRODUCT((GameDates<>"")*(COUNTIF(PracDates,GameDates)>0))` |
| Practice Attendance % | `=IFERROR(.../SUMPRODUCT((PracAttendance<>"")*(PracAttendance<>"—")*1),0)` |

---

## Formula engine highlights

- `SUMPRODUCT` portfolio aggregations (handles blanks safely)
- `COUNTIF` cross-sheet conflict detection (game ↔ practice clash)
- `TODAY()`-driven Days-Out, upcoming-7, upcoming-30, upcoming-60 windows
- 39 named ranges so the Google Sheets companion is a 1-to-1 mirror
- Mileage Fuel + Total auto-compute via `FuelPerMile` from Settings
- Doughnut + Column charts on Dashboard and Budget, all chart series
  expand automatically as you add rows
- Three live conditional-format states on Practices (Present/Late/Absent)
  and Schedule (Win/Loss/Upcoming/Past)

---

## Brand system applied

| Token | Hex | Where it shows |
| ----- | --- | -------------- |
| Primary | `#1B4F48` | Page headers, table headers, KPI values, nav chips |
| Accent | `#937356` | KPI labels, field labels, warning tones |
| Surface | `#E5D3BA` | Input cells (Settings, Budget planned), totals row |
| Highlight | `#75E6C1` | Positive KPIs, upcoming game / Win highlight |
| Danger | `#C94C4C` | Loss row, Replace condition, Absent attendance |
| Soft BG | `#FAF7F1` | Field-value backgrounds |

---

## Build & ship

```bash
cd build && python3 build_xlsx.py
```

Output: `../Soccer_Mom_Command_Center.xlsx`. See `BUILD_INSTRUCTIONS.md`
for the Google Sheets path, Etsy delivery package, listing photo plan,
4-tier pricing, and maintenance protocol.

> **Marketing images** (`marketing/`) are generated separately by
> `build/build_marketing.py` after the workbook is confirmed. Hero
> follows the locked Driver-Budget Etsy thumbnail format.
