# Photography Business Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/photography-command-center/build
python3 build_xlsx.py      # -> ../Photography_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/tax advice"
   note.
2. **CODB & Break-Even** adds $24,000 overhead + $48,000 salary = **$72,000**, ÷ 48
   shoots = **CODB / SHOOT $1,500** (named `CODBShoot`) — the minimum price per shoot.
3. **Shoot P&L** nets the flagship wedding: $3,000 − $800 = **NET PER SHOOT $2,200**
   (named `NetShoot`), ÷ 25 hours = **EFFECTIVE RATE $88.00**, margin **73%**.
4. **Monthly Summary** sums to **REVENUE YTD $45,000** and **30** bookings, avg
   **$1,500**.
5. **Expenses** totals $24,000, leaving a **NET PROFIT of $21,000**; **Bookings**
   shows **$9,600** booked ahead.
6. **Dashboard** fills 12 KPI cards + a Studio Health table + a revenue-by-month
   chart. **Studio Score 90%** (booked-ahead is the honest weak dimension). No broken
   cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `COUNTA`, `AVERAGE`, `MIN`, `IFERROR` — opens in Google
> Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Photography_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: CODB worksheet, shoot P&L, package menu, bookings calendar, clients &
leads, editing queue, shot list, gear inventory, expense log, mileage log, monthly
summary, and a session checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-shoot-dollar-goes &
revenue charts), everything-inside (14 tabs), the CODB break-even engine, the shoot
P&L, the photography engine (CODB + packages), and the **12-page printables
showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic price list vs Command
Center", 09 run-your-studio in 4 steps, 10 what's-included / who-it's-for /
works-with. Ten images — fills all 10 Etsy slots. All headline numbers ($3,000
package · $2,200 net/shoot · $1,500 break-even · $88/hr · 30 bookings · $45,000
revenue · $21,000 profit · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Photography_Command_Center.xlsx    ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Photography_Printables.pdf          ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| PHC-GS   | The Google Sheets / Excel file only | $19 |
| PHC-PDF  | The printable PDF only | $19 |
| PHC-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| PHC-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Strong year-round demand** with a January (pricing-reset) peak and a spring
  wedding-season spike. Priced above consumer planners — photographers pay for tools
  that fix their pricing.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **CODB break-even engine** and the **per-shoot P&L** are
  your strongest differentiators — most listings are just a price list.
- Cross-sell the **Freelancer** cashflow & tax template — same solo-creative buyer.

---

## F. Maintenance

- Edit the `OVERHEAD_ANNUAL`, `SALARY_TARGET`, `TARGET_SHOOTS`, `PACKAGE_PRICE`,
  `SHOOT_COSTS`, `SHOOT_HOURS`, `PACKAGES`, `BOOKINGS`, `CLIENTS`, `EDITING`, `GEAR`,
  `EXPENSES`, `MILEAGE`, `REVIEWS`, `MONTHS` constants and the `MARGIN_GOAL`,
  `RATE_GOAL`, `PACE_GOAL`, `PROFIT_GOAL`, `BOOKED_AHEAD_GOAL` targets in
  `build_xlsx.py`; every KPI + the Studio Score recompute. Everything is
  cross-linked — change your overhead or a package price and the break-even, net and
  score follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
