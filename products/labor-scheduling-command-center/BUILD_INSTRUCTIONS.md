# Restaurant Labor & Scheduling Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/labor-scheduling-command-center/build
python3 build_xlsx.py      # -> ../Labor_Scheduling_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/legal/HR
   advice" note.
2. **Weekly Schedule** grids 8 employees × Mon–Sun; each row's hours = SUM and cost
   = wage × hours. Week total **280 hours**, **$4,234** labor cost (`LaborCost`).
3. **Sales Forecast** totals **$15,000** (`ForecastSales`) and a labor target of
   **$4,500** at 30%.
4. **Labor Cost Calc** divides labor by sales for **28.2%** labor %, **$53.57**
   sales per labor hour and **$15.12** avg wage.
5. **Prime Cost** = food 30% + labor 28.2% = **58.2%**. **Dashboard** fills 12 KPI
   cards + a Labor Health table + a labor-by-day chart. **Labor Score 90%**
   (overtime-low is the honest weak dimension). OT **6** hrs; **$266** under target.
6. No broken cells; custom tables (Weekly Schedule, Labor Cost Calc, etc.) start in
   column B.

> Note: uses `SUM`, `SUMPRODUCT`, `INDEX`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MIN`,
> `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Labor_Scheduling_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: weekly schedule, labor cost worksheet, employee roster, sales
forecast, sales per labor hour, overtime log, roles & rates, time-off request,
prime cost worksheet, tip sheet, labor by day, and shift swaps & notes.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with prime-cost & labor-by-day
charts), everything-inside (14 tabs), the weekly schedule, sales per labor hour,
the labor-cost engine, and the **12-page printables showcase**. Images 3–5 each
show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "blank schedule vs Command
Center", 09 control-labor in 4 steps, 10 what's-included / who-it's-for / guarantee.
Ten images — fills all 10 Etsy slots. All headline numbers ($4,234 labor · 28.2%
labor · $15,000 sales · 280 hours · $53.57 SPLH · 58.2% prime · 90% score) are
verified against the workbook.

---

## D. Etsy delivery package

```
Labor_Scheduling_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt         ← "Make a Copy" link
Labor_Scheduling_Printables.pdf         ← 12-page print-ready pack
START_HERE.pdf                          ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| LSC-GS   | Google Sheets only | $19 |
| LSC-PDF  | Printable PDF only | $19 |
| LSC-BUNDLE | Sheets + PDF + Quick-Start | **$29** |
| LSC-PLUS | Bundle + scheduling add-on | $39 |
| LSC-PRO  | Commercial / multi-location license | $99 |

- **Steady all-year demand** — every restaurant schedules every week. Priced above
  consumer planners; labor is a restaurant's biggest controllable cost.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward hero;
  the **live weekly schedule** and **labor-% engine** are your strongest
  differentiators — most listings are a blank grid.
- Cross-sell the **Food Cost & Inventory**, **Restaurant** and **Bar** products —
  same buyer, natural "restaurant operations" bundle.

---

## F. Maintenance

- Edit the `SCHEDULE`, `EMPLOYEES`, `FORECAST`, `OVERTIME`, `RATES`,
  `AVAILABILITY`, `TIPS` constants in `build_xlsx.py`; every KPI + the Labor Score
  recompute. Add an employee → add a `SCHEDULE` row (hours & cost follow).
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
