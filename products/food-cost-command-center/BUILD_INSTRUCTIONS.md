# Food Cost & Inventory Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/food-cost-command-center/build
python3 build_xlsx.py      # -> ../Food_Cost_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not accounting advice" note.
2. **Inventory Count** values 14 items to an **INVENTORY VALUE of $11,800** (count ×
   unit cost), named `InvValue` — which is the ending inventory.
3. **Food Cost Calc** pulls Purchases (**$18,600**) and Sales (**$64,000**) from
   their logs and Ending from the count: Beginning $12,400 + Purchases − Ending =
   **$19,200** food used ÷ Sales = **30.0%** food cost. Turns **1.6**.
4. **Usage & Variance** compares theoretical ($18,650) to actual ($19,200) usage =
   **+2.9%** variance; top category = Meat & seafood ($7,200).
5. **Dashboard** fills 12 KPI cards + an Inventory Health table + a food-cost-by-
   category chart. **Inventory Score 90%** (variance-low is the honest weak
   dimension). To order **$2,705**; 6 vendors.
6. No broken cells; custom tables (Inventory Count, Food Cost Calc, etc.) start in
   column B.

> Note: uses `INDEX`, `MATCH`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`,
> `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Food_Cost_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: inventory count sheet, food-cost worksheet, purchases log, sales log,
usage & variance, par & order guide, vendor list, price comparison, menu costing,
waste log, category breakdown, and a weekly count checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with food-cost & category charts),
everything-inside (14 tabs), the inventory valuation, the usage variance, the
food-cost engine, and the **12-page printables showcase**. Images 3–5 each show a
different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic count sheet vs Command
Center", 09 control-food-cost in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers (30.0% food
cost · $11,800 inventory · $18,600 purchases · $19,200 used · $64,000 sales · 2.9%
variance · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Food_Cost_Command_Center.xlsx      ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Food_Cost_Printables.pdf            ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| FCC-GS   | The Google Sheets / Excel file only | $19 |
| FCC-PDF  | The printable PDF only | $19 |
| FCC-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| FCC-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, or "done-for-you" — that is what
> gets a listing removed and earns a strike.

- **Steady all-year demand** with a January peak (new-year cost control). Priced
  above consumer planners — a single point of food cost is real money.
- Use all 10 photos + a walkthrough video. Lead photo = the feature-forward hero;
  the **period food-cost engine** and **usage variance** are your strongest
  differentiators — most listings are a plain count sheet.
- Cross-sell the **Restaurant**, **Recipe Costing** and **Ghost Kitchen** products
  — same buyer, natural "food business" bundle.

---

## F. Maintenance

- Edit the `INVENTORY`, `PURCHASES`, `SALES`, `USAGE`, `PARORDER`, `VENDORS`,
  `PRICES`, `MENU`, `WASTE`, `CATEGORIES` constants and `BEGINNING` in
  `build_xlsx.py`; every KPI + the Inventory Score recompute. Everything is
  cross-linked — change a count or invoice and the food-cost % follows.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
