# Bakery Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/bakery-command-center/build
python3 build_xlsx.py      # -> ../Bakery_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not accounting advice" note.
2. **Recipe Costing** costs a Sourdough batch to **$12.80** → **$1.60**/loaf
   (batch ÷ yield 8), linked into the Product List via `=SourdoughBatch`.
3. **Product List** shows unit cost, retail & wholesale price, margin & food-cost %.
   Avg unit cost **$1.13**, avg retail **$5.19**, avg margin **$4.06**, overall
   food cost **22%**; top seller (by revenue) = Butter Croissant.
4. **Pre-Orders** board (6 orders); **Wholesale** accounts total **$2,182**/wk.
   Weekly revenue = retail **$6,495** + wholesale **$2,182** = **$8,677**.
5. **Dashboard** fills 12 KPI cards + a Bakery Health table + a sales-by-day chart.
   **Bakery Score 88%**. Waste **4.5%**.
6. No broken cells; custom tables (Recipe Costing, Product List, etc.) start in column B.

> Note: uses `SUMPRODUCT`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MATCH`,
> `INDEX`, `MIN`, `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Bakery_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: recipe cost card, product price list, pre-order form, wholesale
order sheet, production plan, inventory & par, waste & day-old log, sales log,
ordering sheet, cash & deposits, market day sheet, and a bake-day checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with revenue-mix & sales charts),
everything-inside (14 tabs), the product list (retail & wholesale), the
recipe-costing engine, the pre-orders + wholesale board, and the **12-page
printables showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic price list vs Command
Center", 09 price-your-bakery in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers (8 products
· $1.13 unit cost · 22% food cost · $8,677 revenue · 1,480 units · $2,182
wholesale · 88% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Bakery_Command_Center.xlsx         ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Bakery_Printables.pdf               ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| BKC-GS   | Google Sheets only | $19 |
| BKC-PDF  | Printable PDF only | $19 |
| BKC-BUNDLE | Sheets + PDF + Quick-Start | **$29** |
| BKC-PLUS | Bundle + recipe-costing add-on | $39 |
| BKC-PRO  | Commercial / multi-location license | $99 |

- **Steady all-year demand** with peaks before the holiday baking season and in
  January (new-year business planning). Priced above consumer planners.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward
  hero; the **recipe-costing engine** and **retail + wholesale** pricing are your
  strongest differentiators — most listings are a plain price list.
- Cross-sell the **Recipe Costing**, **Café** and **Home Bakery / Cottage Food**
  products — same buyer, natural "food business" bundle.

---

## F. Maintenance

- Edit the `SOURDOUGH`, `PRODUCTS`, `PREORDERS`, `WHOLESALE`, `PRODUCTION`,
  `INVENTORY`, `WASTE`, `SALES`, `ORDERING`, `CASHDEP`, `MARKETS` constants in
  `build_xlsx.py`; every KPI + the Bakery Score recompute. Add a product → add a
  `PRODUCTS` row (margins & totals follow).
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
