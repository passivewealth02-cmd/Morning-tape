# Recipe Costing & Menu Engineering Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/recipe-costing-command-center/build
python3 build_xlsx.py      # -> ../Recipe_Costing_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not accounting advice" note.
2. **Ingredients** turns pack price ÷ pack size into cost per unit.
3. **Recipe Costing** costs the Classic Burger line by line to **$2.92** / serving;
   that value links into Menu Items via `=BurgerCost`.
4. **Menu Items** shows food-cost % & margin per plate and auto-classes each item.
   Avg food cost **24%**, avg plate cost **$4.75**, avg price **$18.00**, avg
   margin **$13.25**. Classes: **1 Star / 4 Plowhorses / 2 Puzzles / 1 Dog**.
5. **Menu Engineering** counts each class and charts the mix; **Dashboard** fills
   12 KPI cards + a Menu Health table. **Menu Score 92%**. Top margin = Ribeye.
6. No broken cells; custom tables (Ingredients, Recipe Costing, Menu Items, etc.)
   start in column B.

> Note: uses `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`, `INDEX`, `MATCH`,
> `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Recipe_Costing_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: recipe cost card, prep list, a menu-engineering worksheet (the
star/plowhorse/puzzle/dog quadrant), a food-cost pricing guide, ingredient price
log, portion & yield worksheet, menu item P&L, specials planner, batch scaler,
vendor price tracker, waste log, and a weekly food-cost tracker.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (14 tabs), menu
items with food-cost %, the menu-engineering quadrant, the recipe-costing +
price-calculator engine, and the **12-page printables showcase**. Images 3–5 each
show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic calculator vs Command
Center", 09 profit-your-menu in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers (8 items ·
24% food cost · $4.75 / $18.00 / $13.25 · 1/4/2/1 classes · $30,430 revenue · 92%
score) are verified against the workbook.

---

## D. Etsy delivery package

```
Recipe_Costing_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt         ← "Make a Copy" link
Recipe_Costing_Printables.pdf           ← 12-page print-ready pack
START_HERE.pdf                          ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| RCC-GS   | Google Sheets only | $19 |
| RCC-PDF  | Printable PDF only | $19 |
| RCC-BUNDLE | Sheets + PDF + Quick-Start | **$29** |
| RCC-PLUS | Bundle + food-cost add-on | $39 |
| RCC-PRO  | Multi-location / commercial license | $99 |

- **Steady all-year B2B demand** with bumps in January (new-year planning) and
  menu-refresh season. Priced higher than consumer planners — this saves an owner
  real money, so $29 is an easy yes.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward
  hero; the **menu-engineering quadrant** is your strongest differentiator —
  most listings are a plain cost calculator, not a full engineer-the-menu system.
- Cross-sell the **Restaurant Command Center** and the food-service niche
  products (Food Truck, Café, Bakery, Bar) — same owner, natural bundle.

---

## F. Maintenance

- Edit the `INGREDIENTS`, `BURGER`, `MENU`, `PRICECALC`, `YIELD`, `SPECIALS`,
  `BATCH`, `VENDOR`, `WASTE` constants in `build_xlsx.py`; every KPI + the Menu
  Score recompute. Add a menu item → add a `MENU` row (its class calculates).
- Keep `build_marketing.py`'s KPIs and profit figures in sync with the workbook.
