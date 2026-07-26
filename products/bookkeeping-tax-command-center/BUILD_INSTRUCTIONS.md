# Small Business Bookkeeping & Tax Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/bookkeeping-tax-command-center/build
python3 build_xlsx.py      # -> ../Bookkeeping_Tax_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not tax/accounting advice"
   note.
2. **Income** sums to **GROSS REVENUE $96,000** (`GrossRevenue`); **COGS & Inventory**
   sums to **$28,800** (`COGSTotal`); **Expenses** sums to **$19,200** (`ExpTotal`).
3. **Schedule C P&L** computes **GROSS PROFIT $67,200** (`GrossProfit`) and **NET
   PROFIT $48,000** (`NetProfit`, Line 31), **50%** net margin (`NetMargin`).
4. Tax block: SE base = net × 92.35%; **SE TAX $6,782** (`SETax`); **INCOME TAX
   $5,353** (`IncomeTax`, on net − ½ SE); **TOTAL TAX $12,135** (`TotalTax`); **SEND
   EACH QUARTER $3,034** (`QuarterlyTax`); effective rate **25.3%**.
5. **Mileage** totals **4,000** miles × $0.70 = **$2,800** (`MileageDeduction`);
   **Sales Tax** shows **$600** still owed; **Invoices** shows outstanding balance.
6. **Dashboard** fills 12 KPI cards + a Books Health table + where-the-revenue-goes &
   net-profit-by-month charts. **Books Score 90%** (receipts-attached is the honest
   weak dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `COUNTA`, `COUNTIF`, `AVERAGE`, `MIN`, `IF`, `IFERROR` —
> opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Bookkeeping_Tax_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: Schedule C P&L, tax worksheet, income log, expense log, COGS & inventory,
mileage log, sales tax log, invoice tracker, receipt tracker, reconciliation, monthly
summary and a tax-time checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the revenue-split & profit charts),
everything-inside (14 tabs), the Schedule C engine, the quarterly tax engine, the books
engine (Schedule C + expenses), and the **12-page printables showcase**. Images 3–5 each
show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic tracker vs Command Center",
09 do-your-books in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($96k revenue · $28.8k COGS ·
$67.2k gross · $19.2k expenses · $48k net · 50% margin · $6,782 SE · $5,353 income ·
$12,135 total · $3,034 quarterly · $2,800 mileage · 90% score) are verified against the
workbook.

---

## D. Etsy delivery package

```
Bookkeeping_Tax_Command_Center.xlsx    ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt         ← "Make a Copy" link
Bookkeeping_Tax_Printables.pdf          ← 12-page print-ready pack
START_HERE.pdf                          ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| BTC-GS   | The Google Sheets / Excel file only | $22 |
| BTC-PDF  | The printable PDF only | $22 |
| BTC-BUNDLE | The spreadsheet + the printable PDF | **$35** |
| BTC-COMM | The same files + a commercial-use file license | $59 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, bookkeeping services, "done-for-you"
> builds, or "free updates / lifetime access" — offering a service (rather than a
> finished file) is what gets a listing removed and earns a strike. A commercial-use
> license is a permission term attached to the *file* and is allowed; doing the
> bookkeeping *for* the buyer is not.

- **The strongest seasonality in the catalogue:** a January–April tax-season surge on
  top of steady year-round demand. Price at the top of the range — this replaces a
  $30/month subscription.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **Schedule C engine** and the **quarterly tax number** are
  your strongest differentiators — most listings are just an expense list.
- Cross-sell the **Freelancer**, **Maker Pricing** and **Etsy Seller Profit**
  templates — same self-employed buyer.

---

## F. Maintenance

- Edit the `INCOME`, `EXPENSES`, `COGS_ITEMS`, `MILEAGE`, `SALES_TAX`, `QUARTERS`,
  `SCHED_C`, `INVOICES`, `RECONCILE`, `MONTHS` constants and the `MILEAGE_RATE`,
  `SE_TAX_RATE`, `SE_BASE_PCT`, `INCOME_TAX_RATE`, `PROFIT_GOAL`, `MARGIN_GOAL`,
  `RECONCILE_GOAL`, `RECEIPTS_PCT` targets in `build_xlsx.py`; every KPI + the Books
  Score recompute. Everything is cross-linked — change revenue or a rate and the
  profit, tax and quarterly payment follow.
- **Update `MILEAGE_RATE` each January** when the IRS publishes the new standard rate.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
