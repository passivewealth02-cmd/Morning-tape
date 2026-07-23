# Rental Property & Landlord Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/rental-property-command-center/build
python3 build_xlsx.py      # -> ../Rental_Property_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/tax/investment
   advice" note.
2. **Deal Analyzer** nets the flagship unit: gross **$1,925** − operating **$830** =
   NOI **$1,095**, − mortgage **$735** = **CASH FLOW $360** (named `CashFlow`). Cap
   rate **6.0%**, cash-on-cash **8.0%**, DSCR **1.49**, 1% rule **0.84%**, annual cash
   flow **$4,320**, reserves funded **40%**.
3. **Rent Roll** sums 4 doors to **PORTFOLIO RENT $7,000** (named `PortfolioRent`).
4. **Rent Ledger / Expenses Log** total collected rent and spending by category.
5. **Reserves & Escrow** shows each fund's balance vs goal.
6. **Dashboard** fills 12 KPI cards + a Landlord Health table + a cash-flow-by-month
   chart. **Landlord Score 90%** (reserves-funded is the honest weak dimension). No
   broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `INDEX`, `MATCH`, `COUNTA`, `AVERAGE`, `MIN`, `IFERROR`
> — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Rental_Property_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: deal analyzer, rent roll, tenant sheet, rent ledger, expense log,
maintenance & CapEx, mortgage & loans, reserves tracker, mileage log, renewals &
docs, monthly summary, and a move-in / move-out condition sheet.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-rent-goes & cash-flow
charts), everything-inside (14 tabs), the deal-analyzer engine, the rent roll, the
landlord engine (deal analyzer + rent roll), and the **12-page printables showcase**.
Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic rent sheet vs Command
Center", 09 run-your-rentals in 4 steps, 10 what's-included / who-it's-for /
works-with. Ten images — fills all 10 Etsy slots. All headline numbers ($1,850 rent ·
$1,095 NOI · $360 cash flow · 6.0% cap · 8.0% cash-on-cash · 1.49 DSCR · $4,320 annual
· 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Rental_Property_Command_Center.xlsx  ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt       ← "Make a Copy" link
Rental_Property_Printables.pdf        ← 12-page print-ready pack
START_HERE.pdf                        ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| RPC-GS   | The Google Sheets / Excel file only | $19 |
| RPC-PDF  | The printable PDF only | $19 |
| RPC-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| RPC-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Strong year-round demand** with peaks in spring buying season and at tax time.
  Priced above consumer planners — real-estate investors pay for tools that price a
  deal correctly.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **deal analyzer** and the **rent roll** are your strongest
  differentiators — most listings are just a rent tracker.
- Cross-sell alongside a house-flip or BRRRR analyzer — same investor buyer.

---

## F. Maintenance

- Edit the `MONTHLY_RENT`, `OTHER_INCOME`, `OPEX`, `MORTGAGE_PI`, `PURCHASE_PRICE`,
  `CASH_INVESTED`, `RESERVE_BALANCE`, `UNITS`, `TENANTS`, `LEDGER_PAY`, `EXP_LOG`,
  `MAINT`, `LOANS`, `RESERVES`, `MONTHS` constants and the `CF_GOAL`, `CAP_GOAL`,
  `COC_GOAL`, `DSCR_GOAL`, `EXP_TARGET`, `RESERVE_GOAL_MONTHS` targets in
  `build_xlsx.py`; every KPI + the Landlord Score recompute. Everything is
  cross-linked — change the rent or a cost and the NOI, cash flow and returns follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
