# Restaurant Labor & Scheduling Command Center™ — The Labor-Cost System

> Not a blank schedule — a **complete schedule-to-your-sales system**.
> One premium **Google Sheets + printable PDF** command center for restaurant
> labor: a live weekly schedule grid, a labor-cost engine (scheduled hours × wage
> ÷ sales), an employee roster, a sales forecast with labor targets,
> sales-per-labor-hour, overtime, roles & rates, availability, prime cost, tips and
> labor-by-day.

| | |
| - | - |
| **Product** | Restaurant Labor & Scheduling Command Center™ |
| **Target** | Restaurants & cafés · bars & pubs · quick-service & fast-casual · coffee shops & bakeries · managers & shift leads · anyone who builds a schedule |
| **Angle** | Schedule to your sales, and never over-spend on labor again. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single · **$29 bundle** (Sheets + PDF) · $39 with the scheduling add-on · $99 commercial / multi-location license |

---

## Contents

```
products/labor-scheduling-command-center/
├── README.md
├── Labor_Scheduling_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
├── Labor_Scheduling_Printables.pdf         ← 12-page print-ready pack (US Letter)
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

## The 14-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 8 | Overtime |
| 2 | Dashboard | 9 | Roles & Rates |
| 3 | Labor Cost Calc | 10 | Availability |
| 4 | Weekly Schedule | 11 | Prime Cost |
| 5 | Employees | 12 | Tips |
| 6 | Sales Forecast | 13 | Labor by Day |
| 7 | Sales per Labor Hr | 14 | Settings |

## The 12 printable PDF pages

Weekly Schedule · Labor Cost Worksheet · Employee Roster · Sales Forecast · Sales
per Labor Hour · Overtime Log · Roles & Rates · Time-Off Request · Prime Cost
Worksheet · Tip Sheet · Labor by Day · Shift Swaps & Notes.

---

## Signature automation — schedule to your sales

Build the weekly schedule and every employee's hours and cost calculate live; the
labor-cost engine divides labor by sales for a true labor %; and food cost + labor
roll into a prime-cost number:

```
Employee cost  = Weekly hours × Wage
Labor cost     = Σ (employee cost)
Labor %        = Labor cost ÷ Sales
Sales / labor hr = Sales ÷ Total hours
Prime cost     = Food cost % + Labor %
```

### The 12 dashboard KPIs
Labor Cost · Labor % · Weekly Sales · Total Hours · Sales/Labor Hr · Employees ·
Labor Target · Vs Target · Avg Wage · Prime Cost · OT Hours · Labor Score. The
**Labor Score** blends labor-on-target, SPLH-healthy, overtime-low,
prime-cost-in-check, schedule-covered and wage-in-budget into one 0–100% number.

**Verified sample restaurant** (Maple & Ash, manager Jordan): labor cost **$4,234**
· labor **28.2%** · weekly sales **$15,000** · total hours **280** · sales/labor hr
**$53.57** · **8** employees · labor target **$4,500** · **$266** under budget · avg
wage **$15.12** · prime cost **58.2%** · OT **6** hrs · **Labor Score 90%**.

---

## Premium labor design

- A **live weekly schedule** grid (hours × wage → cost, per employee)
- A **labor-cost** engine (labor ÷ sales = true labor %)
- A **sales forecast** that sets a labor target and target hours
- **Sales per labor hour**, an **overtime** watch list & **prime cost**
- **Roles & rates**, **availability**, **tips** & **labor-by-day**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business tool, not financial, legal or HR advice.** Confirm wages, overtime
> rules and labor law with your own advisors.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Labor_Scheduling_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Labor_Scheduling_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
