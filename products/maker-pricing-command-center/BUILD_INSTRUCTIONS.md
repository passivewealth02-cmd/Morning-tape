# Handmade & Maker Pricing Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/maker-pricing-command-center/build
python3 build_xlsx.py      # -> ../Maker_Pricing_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/tax advice"
   note.
2. **Pricing Calculator** sums 6 materials to **$6.00**, adds labor (0.25 hr × $20 =
   **$5.00**) and overhead **$1.00** → **BASE COST $12.00** (named `BaseCost`). A ×2.5
   markup gives **RETAIL $30.00** (`RetailPrice`) and a ×0.5 factor gives **WHOLESALE
   $15.00** (`WholesalePrice`). Retail margin **60%**, wholesale margin **20%**, and
   the **TRUE HOURLY WAGE $92.00** (`EffHourly`).
3. **Product Line** lists six products, each with base/wholesale/retail and a margin.
4. **Monthly Summary** shows **$2,700** revenue, **90** units, **$1,350** monthly
   profit (`MonthlyProfit`), materials **20%** of revenue.
5. **Overhead & Fees** totals **$360**; **Expenses** totals **$1,200**; **Channels &
   Markets** rolls up revenue and fees by channel.
6. **Dashboard** fills 12 KPI cards + a Maker Health table + where-the-retail-$-goes &
   revenue-by-month charts. **Maker Score 90%** (wholesale margin is the honest weak
   dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MIN`, `IFERROR` — opens
> in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Maker_Pricing_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: pricing worksheet, product price list, order form, supply list, overhead
& fees, sales log, time & labor, expense log, channel tracker, monthly summary, craft
show checklist and a making checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-retail-dollar-goes &
revenue charts), everything-inside (14 tabs), the pricing engine (base cost →
retail/wholesale), the true-hourly-wage engine, the maker engine (pricing + product
line), and the **12-page printables showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic calculator vs Command
Center", 09 run-your-shop in 4 steps, 10 what's-included / who-it's-for / works-with.
Ten images — fills all 10 Etsy slots. All headline numbers ($6 materials · $12 base ·
$30 retail · $15 wholesale · 60% margin · $92/hr · $2,700 revenue · 90 units · $1,350
profit · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Maker_Pricing_Command_Center.xlsx  ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Maker_Pricing_Printables.pdf        ← 12-page print-ready pack
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

- **Strong year-round demand** with a January (new-shop / pricing-reset) peak and a
  fall craft-fair spike. Priced above consumer planners — makers pay for tools that
  fix their pricing.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **price-for-profit engine** and the **true hourly wage**
  are your strongest differentiators — most listings are just a calculator.
- Cross-sell the **Freelancer** cashflow & tax template — same solo-creative buyer.

---

## F. Maintenance

- Edit the `MATERIALS`, `PRODUCTS`, `SUPPLIES`, `ORDERS`, `OVERHEAD`, `SALES`, `TASKS`,
  `EXPENSES`, `REVIEWS`, `CHANNELS_DATA`, `MONTHS` constants and the `LABOR_RATE`,
  `MARKUP`, `WHOLESALE_FACTOR`, `MARGIN_GOAL`, `MARKUP_GOAL`, `HOUR_GOAL`,
  `PROFIT_GOAL`, `WHOLESALE_MARGIN_GOAL` targets in `build_xlsx.py`; every KPI + the
  Maker Score recompute. Everything is cross-linked — change your labor rate or a
  markup and the base cost, retail, wholesale, wage and score follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
