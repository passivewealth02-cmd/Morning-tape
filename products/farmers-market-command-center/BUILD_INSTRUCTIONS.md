# Farmers Market Vendor Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/farmers-market-command-center/build
python3 build_xlsx.py      # -> ../Farmers_Market_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/food-safety
   advice" note.
2. **Products** prices 8 items with a live margin (Sourdough $1.85→$8 = 77%, etc.).
3. **Market Day** pulls each product's price & cost by `INDEX/MATCH`; the 90 units
   sold make **MARKET SALES $768** (`DaySales`), **COGS $187**, and after **$143**
   booth costs the **BOOTH NET is $438** (`BoothNet`), ÷ 8 hours = **NET/HOUR
   $54.75**. COGS **24%**, net margin **57%**, avg basket **$12.00**, top seller
   **Sourdough**.
4. **Markets** sums 4 market days to **MONTHLY SALES $3,000** (`MonthlySales`).
5. **Income & Expenses** takes `MonthlySales` in, subtracts $1,405 of costs for a
   **MONTHLY PROFIT of $1,595**.
6. **Dashboard** fills 12 KPI cards + a Vendor Health table + a sales-by-month chart.
   **Vendor Score 90%** (waste-low is the honest weak dimension). No broken cells;
   custom tables start in column B.

> Note: uses `INDEX`, `MATCH`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`,
> `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Farmers_Market_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: booth P&L, product price list, bake plan, ingredient list, markets log,
booth costs, packaging, customer list, waste log, income & expenses, monthly summary,
and a market-day checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-day-goes & sales
charts), everything-inside (14 tabs), the booth-P&L engine, the markets→monthly-sales
roster, the market-vendor engine (booth P&L + products), and the **12-page printables
showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic price list vs Command
Center", 09 run-your-stall in 4 steps, 10 what's-included / who-it's-for / works-with.
Ten images — fills all 10 Etsy slots. All headline numbers (90 units · $768 sales ·
24% COGS · $438 booth net · $54.75/hr · $12.00 basket · $3,000 monthly sales · $1,595
profit · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Farmers_Market_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt       ← "Make a Copy" link
Farmers_Market_Printables.pdf         ← 12-page print-ready pack
START_HERE.pdf                        ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| FMC-GS   | The Google Sheets / Excel file only | $19 |
| FMC-PDF  | The printable PDF only | $19 |
| FMC-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| FMC-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Strong spring–fall demand** as market season opens, with a craft-fair peak in
  November–December. Priced above consumer planners — one market day that didn't pay
  is real money.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **booth-P&L engine** and the **markets→monthly-sales
  roster** are your strongest differentiators — most listings are just a price list.
- Cross-sell the **Home Bakery**, **Recipe Costing** and **Food Cost & Inventory**
  products — same buyer, natural "food business" bundle.

---

## F. Maintenance

- Edit the `PRODUCTS`, `DAY_SOLD`, `CUSTOMERS`, `BOOTH`, `MARKET_HOURS`,
  `INGREDIENTS`, `BAKEPLAN`, `MARKETS`, `STALLS`, `PACKAGING`, `CUSTOMERS_LIST`,
  `WASTE`, `LEDGER`, `MONTHS` constants and the `TARGET_COGS`, `MARGIN_GOAL`,
  `SALES_GOAL`, `PROFIT_GOAL`, `WASTE_LIMIT` targets in `build_xlsx.py`; every KPI +
  the Vendor Score recompute. Everything is cross-linked — change a product price or
  what you sold and the booth net, the net-per-hour and the score all follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
