# Meal Planning & Grocery Budget Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/meal-plan-command-center/build
python3 build_xlsx.py      # -> ../Meal_Plan_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/nutrition advice"
   note.
2. **Cost Per Meal** sums the flagship's ingredients to **$10.00** (named `RecipeCost`),
   ÷ **4** servings = **COST PER SERVING $2.50** (`CostPerServing`); eat-out $12.50
   minus that = **YOU SAVE / SERVING $10.00** (`SavedServing`).
3. **Weekly Plan** sums 7 dinners to **$66** (`WeeklyPlanCost`); `PlannedDinners` = 7.
4. **Budget** sums weekly spend to **$540** (`SpentMonth`) against a **$600** budget.
5. **Savings** = 120 servings × $12.50 − $540 = **MONTHLY SAVINGS $960**
   (`MonthlySavings`); **Price Book** holds 12 items (`PricedItems`).
6. **Dashboard** fills 12 KPI cards + a Kitchen Health table + cook-vs-eat-out &
   saved-by-month charts. **Kitchen Score 90%** (eating-out-rare is the honest weak
   dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `COUNTA`, `AVERAGE`, `MIN`, `IF`, `IFERROR` — opens in Google
> Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Meal_Plan_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: weekly meal plan, cost per meal, grocery list, price book, recipe
rotation, pantry inventory, grocery budget, eating-out log, meal ideas, savings
summary, monthly summary and a kitchen checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the cook-vs-eat-out & saved charts),
everything-inside (14 tabs), the cost-per-meal engine, the weekly plan, the kitchen
engine (cost + weekly), and the **12-page printables showcase**. Images 3–5 each show a
different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic planner vs Command Center",
09 plan-your-week in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($10 recipe · 4 servings · $2.50
per plate · $12.50 eat-out · $10 saved · $66 week · $600 budget · $540 spent · 90% used
· 120 cooked · $960 savings · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Meal_Plan_Command_Center.xlsx      ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Meal_Plan_Printables.pdf            ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| MPL-GS   | The Google Sheets / Excel file only | $19 |
| MPL-PDF  | The printable PDF only | $19 |
| MPL-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| MPL-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Strong year-round demand** with a January (new-year budgeting) peak. Meal-planning
  and grocery-budget templates are perennial best-sellers on Etsy.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **cost-per-meal engine** and the **weekly plan** are your
  strongest differentiators — most listings are just a blank meal grid.
- Cross-sell the **Net Worth & FIRE** and **Subscription & Bills Audit** templates —
  same budget-minded buyer.

---

## F. Maintenance

- Edit the `INGREDIENTS`, `RECIPES`, `PRICE_BOOK`, `WEEKLY`, `GROCERY`, `PANTRY`,
  `WEEKS_SPEND`, `EATOUT`, `IDEAS`, `MONTHS` constants and the `SERVINGS_N`,
  `EATOUT_SERVING`, `SERVINGS_COOKED_N`, `GROCERY_BUDGET_N`, `EATOUT_MEALS_N`,
  `SAVE_GOAL`, `SERVING_GOAL`, `PRICE_GOAL`, `EATOUT_GOAL` targets in `build_xlsx.py`;
  every KPI + the Kitchen Score recompute. Everything is cross-linked — change a
  recipe or your budget and the cost per serving, savings and score follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
