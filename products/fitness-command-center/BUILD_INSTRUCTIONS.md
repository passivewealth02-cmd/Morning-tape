# Fitness & Meal-Prep Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/fitness-command-center/build
python3 build_xlsx.py      # -> ../Fitness_Command_Center.xlsx  (12 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not medical advice" note.
2. **Goals & Stats** shows start **185** → current **176** → goal **165** with
   **9 lb** lost, **11** to go, **45%** to goal.
3. **Macro Tracker** averages calories **2,062** & protein **149 g** (14 days),
   with a calories-vs-target color scale.
4. **Workout Plan** counts **4 / 5** sessions done; **Workout Log** computes
   sets × reps × weight = volume (total **37,515**); **Body Metrics** draws the
   weight-trend line and totals **−9.0 lb**.
5. **Dashboard** fills 12 KPI cards + a Fitness Score table + a weight-trend line
   chart. **Fitness Score 84%**. Steps avg **9,493**, water avg **7.1** cups.
6. No broken cells; custom tables (Goals & Stats, Meal Plan, Recipe Bank, Macro
   Tracker, Workout Log, Body Metrics, Settings) start in column B.

> Note: uses `AVERAGE`, `COUNTIF`, `SUM`, `MIN`, `ABS`, `ROUND`, `IFERROR` —
> opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Fitness_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: weekly meal planner, grocery list (by aisle), macro tracker, recipe
cards, workout plan, workout log, body measurements, weight-progress chart, habit
tracker, meal-prep day checklist, progress & measurements, and goals & why.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with weight-trend chart),
everything-inside (12 tabs), the weekly meal plan, the macro tracker, the workout
log + body-metrics weight trend, and the **12-page printables showcase**. Images
3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic tracker vs Command
Center", 09 get-on-track in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers (176→165 lb
· 9 lb lost · 2,062 cal · 149 g protein · 4/5 workouts · 9,493 steps · 37,515
volume · 84% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Fitness_Command_Center.xlsx        ← Google Sheets / Excel master (12 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Fitness_Printables.pdf              ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| FMC-GS   | Google Sheets only | $16 |
| FMC-PDF  | Printable PDF only | $16 |
| FMC-BUNDLE | Sheets + PDF + Quick-Start | **$24** |
| FMC-PLUS | Bundle + macro-coaching add-on | $32 |
| FMC-PRO  | Coach / commercial-use license | $79 |

- **Sharp December–February peak** (New-Year fitness resolutions) with a spring
  "summer body" tail. List by November and refresh tags in late December.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward
  hero; the **weekly meal plan** and the **weight-trend chart** are your
  strongest differentiators — most listings are a workout OR a meal tracker,
  not both in one system.
- Cross-sell the **Budget & Money Command Center** — both are "New-Year
  resolution" products with the same premium brand.

---

## F. Maintenance

- Edit the `WEIGHTS`, `MACROS`, `RECIPES`, `MEALPLAN`, `GROCERY`, `WORKOUT_PLAN`,
  `WORKOUT_LOG`, `HABITS` constants and the target values in `build_xlsx.py`;
  every KPI + the Fitness Score recompute. Add a recipe → add a `RECIPES` row.
- Keep `build_marketing.py`'s KPIs in sync with the workbook averages.
