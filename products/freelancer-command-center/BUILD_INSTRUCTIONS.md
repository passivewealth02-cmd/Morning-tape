# Freelancer Cashflow & Tax Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/freelancer-command-center/build
python3 build_xlsx.py      # -> ../Freelancer_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not tax advice; set-aside is
   an estimate" note.
2. **Invoices** sums paid to **INCOME $8,000** (named `MonthlyIncome`) and
   unpaid/sent to **OUTSTANDING $2,400** by `SUMIF`.
3. **Business Expenses** totals **$1,200** (named `BizExp`).
4. **Cashflow & Tax** nets income − expenses = **$6,800**, sets aside 30% = **$2,040**
   (named `TaxSetAside`), leaving **TAKE-HOME $4,760** (named `TakeHome`).
5. **Time & Rates** shows **100** billable hours, **63%** utilization and an effective
   rate of **$68.00** (net ÷ billable).
6. **Dashboard** fills 12 KPI cards + a Freelance Health table + a take-home-by-month
   chart. **Freelance Score 90%** (runway-built is the honest weak dimension). No
   broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `COUNTA`, `AVERAGE`, `MIN`, `ROUND`, `IFERROR` — opens in
> Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Freelancer_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: invoice, cashflow & tax worksheet, client list, rate card, business
expenses, tax vault, mileage & home office, pipeline, savings & runway, subscriptions,
monthly summary, and a weekly checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-income-goes &
take-home charts), everything-inside (14 tabs), the cashflow-&-tax engine, the invoice
board, the freelance engine (cashflow + invoices), and the **12-page printables
showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic invoice vs Command
Center", 09 run-your-freelance in 4 steps, 10 what's-included / who-it's-for /
works-with. Ten images — fills all 10 Etsy slots. All headline numbers ($8,000 income
· $6,800 net · $2,040 tax set-aside · $4,760 take-home · $68/hr · 63% utilization ·
$96,000 year pace · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Freelancer_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Freelancer_Printables.pdf           ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| FLC-GS   | The Google Sheets / Excel file only | $19 |
| FLC-PDF  | The printable PDF only | $19 |
| FLC-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| FLC-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Strong year-round demand** with a hard peak at tax season (Jan–Apr). Priced above
  consumer planners — the self-employed pay for tools that keep the tax straight.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **cashflow-&-tax engine** (auto tax set-aside) and the
  **invoice board** are your strongest differentiators — most listings are just an
  invoice.
- Cross-sell a personal budget and a small-business bookkeeping template — same buyer.

---

## F. Maintenance

- Edit the `INVOICES`, `CLIENTS`, `EXPENSES`, `BILLABLE_HOURS`, `AVAILABLE_HOURS`,
  `RATECARD`, `QUARTERS`, `TAX_SAVED`, `DEDUCTIONS`, `PIPELINE`, `EMERGENCY_FUND`,
  `MONTHLY_NEED`, `SUBS`, `MONTHS` constants and the `INCOME_GOAL`, `MARGIN_GOAL`,
  `RATE_GOAL`, `TAKE_GOAL`, `RUNWAY_GOAL_MONTHS`, `TAX_RATE` targets in
  `build_xlsx.py`; every KPI + the Freelance Score recompute. Everything is
  cross-linked — mark an invoice paid or change the tax rate and the take-home
  follows.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
