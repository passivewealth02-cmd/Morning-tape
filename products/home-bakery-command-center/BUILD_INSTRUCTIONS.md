# Home Bakery & Cottage Food Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/home-bakery-command-center/build
python3 build_xlsx.py      # -> ../Home_Bakery_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/legal/cottage-
   food-law advice" note.
2. **Price It Right** costs a dozen Custom Cookies: ingredients $6 + packaging $1.50
   + your labor (45 min × $20 = $15) + overhead $1 = **$23.50** true cost; at a $33
   price your effective hourly wage is **$32.67** (`EffHourly`).
3. **Product List** shows ingredient cost, price, margin & monthly revenue on 8
   products. Avg price **$28.25**, avg margin **82%**, food cost **18%**, monthly
   income **$2,908** (`MonthlyIncome`); top seller = Custom Cookies ($660).
4. **Income & Expenses** nets **$1,658** profit; **Custom Orders** = 5 open, $416.
5. **Dashboard** fills 12 KPI cards + a Bakery Health table + a profit-by-month
   chart. **Bakery Score 90%** (waste-low is the honest weak dimension). 113 units.
6. No broken cells; custom tables (Price It Right, Product List, etc.) start in
   column B.

> Note: uses `INDEX`, `MATCH`, `SUMIF`, `SUMPRODUCT`, `COUNTIF`, `COUNTA`,
> `AVERAGE`, `MAX`, `MIN`, `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Home_Bakery_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: price-it worksheet, recipe cost card, product price list, custom
order form, ingredient cost list, cottage food label, market day checklist,
income & expenses, waste log, customer list, monthly summary, and an order calendar.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with food-cost & profit charts),
everything-inside (14 tabs), the product list, the pay-yourself calc, the
price-it-right engine, and the **12-page printables showcase**. Images 3–5 each
show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic price list vs Command
Center", 09 price-your-bakes in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers (8 products ·
$28.25 price · 18% food cost · $2,908 income · $1,658 profit · $32.67/hr · 90%
score) are verified against the workbook.

---

## D. Etsy delivery package

```
Home_Bakery_Command_Center.xlsx    ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Home_Bakery_Printables.pdf          ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| HBC-GS   | Google Sheets only | $19 |
| HBC-PDF  | Printable PDF only | $19 |
| HBC-BUNDLE | Sheets + PDF + Quick-Start | **$29** |
| HBC-PLUS | Bundle + pricing add-on | $39 |
| HBC-PRO  | Commercial license | $99 |

- **Steady all-year demand** with peaks before the holiday baking season and in
  January (new-year business planning). Priced above consumer planners — this
  saves an underpricing baker real money.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward hero;
  the **pay-yourself pricing engine** is your strongest differentiator — most
  listings are a plain price calculator.
- Cross-sell the **Bakery**, **Farmers Market** and **Recipe Costing** products —
  same buyer, natural "food business" bundle.

---

## F. Maintenance

- Edit the `PIR`, `RECIPE`, `PRODUCTS`, `ORDERS`, `INGREDIENTS`, `LABELS`,
  `MARKETS`, `LEDGER`, `WASTE`, `CUSTOMERS`, `MONTHS` constants in `build_xlsx.py`;
  every KPI + the Bakery Score recompute. Add a product → add a `PRODUCTS` row.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
