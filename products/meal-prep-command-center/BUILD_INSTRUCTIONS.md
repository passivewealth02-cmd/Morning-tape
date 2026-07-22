# Meal Prep Business Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/meal-prep-command-center/build
python3 build_xlsx.py      # -> ../Meal_Prep_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/food-safety
   advice" note.
2. **Meal Cost** totals the flagship Grilled Chicken Bowl to a **COST PER MEAL of
   $5.00** (named `MealCost`), with an **ingredient-only cost of $3.20** (named
   `IngredientCost`) driving food-cost %.
3. **Meal Plans** multiplies meals × `MealCost` for each plan cost and shows the live
   margin (5-meal 55%, 10-meal 52%, 15-meal 50%, 21-meal 48%).
4. **Subscribers** pulls each plan's price & meals by `INDEX/MATCH`, so 46
   subscribers roll up to **$4,475 weekly revenue**, **MRR $17,900**, **433**
   meals/week, avg price **$10.33**/meal, food cost **31%**, margin **52%**.
5. **Dashboard** fills 12 KPI cards + a Prep Health table + an MRR-by-month chart.
   **Prep Score 90%** (waste-low is the honest weak dimension). Monthly profit
   **$6,740**; top plan **10 meals / week**.
6. No broken cells; custom tables (Meal Cost, Meal Plans, Subscribers, etc.) start in
   column B.

> Note: uses `INDEX`, `MATCH`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`,
> `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Meal_Prep_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: meal cost card, meal plan pricing, menu, subscriber list, prep list,
ingredient cost list, packaging & delivery, delivery run sheet, waste log, income &
expenses, monthly summary, and a prep-day checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with margin & MRR charts),
everything-inside (14 tabs), the subscribers→MRR engine, the per-meal cost engine,
the pricing engine (meal cost + plans), and the **12-page printables showcase**.
Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic meal list vs Command
Center", 09 launch-your-prep in 4 steps, 10 what's-included / who-it's-for /
works-with. Ten images — fills all 10 Etsy slots. All headline numbers ($5.00 meal
cost · $10.33 avg price · 31% food cost · 52% margin · 46 subscribers · $4,475
weekly · $17,900 MRR · $6,740 profit · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Meal_Prep_Command_Center.xlsx      ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Meal_Prep_Printables.pdf            ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| MPC-GS   | The Google Sheets / Excel file only | $19 |
| MPC-PDF  | The printable PDF only | $19 |
| MPC-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| MPC-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Steady all-year demand** with a January peak (new-year meal-prep resolutions).
  Priced above consumer planners — one lost subscriber is real recurring money.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **subscribers→MRR engine** and the **per-meal cost
  engine** are your strongest differentiators — most listings are just a meal menu.
- Cross-sell the **Recipe Costing**, **Ghost Kitchen** and **Food Cost & Inventory**
  products — same buyer, natural "food business" bundle.

---

## F. Maintenance

- Edit the `MEAL`, `PLANS`, `MENU`, `SUBSCRIBERS`, `PRODUCTION`, `INGREDIENTS`,
  `PACKAGING`, `ORDERS`, `WASTE`, `LEDGER`, `MONTHS` constants and the `TARGET_FC`,
  `MARGIN_GOAL`, `SUB_GOAL`, `PROFIT_GOAL`, `WASTE_LIMIT` targets in `build_xlsx.py`;
  every KPI + the Prep Score recompute. Everything is cross-linked — change the meal
  cost or a subscriber count and MRR and the score follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
