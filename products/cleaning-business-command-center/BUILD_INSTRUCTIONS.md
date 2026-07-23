# Cleaning & Service Business Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/cleaning-business-command-center/build
python3 build_xlsx.py      # -> ../Cleaning_Business_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/tax advice"
   note.
2. **Job P&L** prices the flagship clean: $180 − $80 costs = **JOB PROFIT $100**
   (named `JobProfit`), ÷ 2.5 hours = **PROFIT/HOUR $40.00**, margin **56%**.
3. **Clients** rolls up the recurring book to **MRR $4,500** (named `MRR`) across
   **12** clients, plus $1,500 one-time = **MONTHLY REVENUE $6,000** and **31** jobs.
4. **Expenses** totals $3,100, leaving a **MONTHLY PROFIT of $2,900** (named
   `NetProfit`).
5. **Services & Pricing** shows the rate per hour on every service.
6. **Dashboard** fills 12 KPI cards + a Service Health table + a profit-by-month
   chart. **Service Score 90%** (client-base-growing is the honest weak dimension).
   No broken cells; custom tables start in column B.

> Note: uses `SUM`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MIN`, `IFERROR` — opens in Google
> Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Cleaning_Business_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: job quote & P&L, service price list, client list, weekly schedule, leads
& quotes, supply list, team & labor, mileage log, expense log, reviews & referrals,
monthly summary, and a job checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-job-dollar-goes &
profit charts), everything-inside (14 tabs), the job-P&L engine, the recurring client
book, the service-business engine (job P&L + services), and the **12-page printables
showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic price list vs Command
Center", 09 run-your-business in 4 steps, 10 what's-included / who-it's-for /
works-with. Ten images — fills all 10 Etsy slots. All headline numbers ($180 job ·
$100 profit · $40/hr · 31 jobs · $6,000 revenue · $4,500 recurring · $2,900 profit ·
90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Cleaning_Business_Command_Center.xlsx  ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt         ← "Make a Copy" link
Cleaning_Business_Printables.pdf        ← 12-page print-ready pack
START_HERE.pdf                          ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| CSB-GS   | The Google Sheets / Excel file only | $19 |
| CSB-PDF  | The printable PDF only | $19 |
| CSB-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| CSB-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Strong year-round demand** with a spring-cleaning peak. Priced above consumer
  planners — service owners pay for tools that price a job by the hour of profit.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **job-P&L engine** and the **recurring client book** are
  your strongest differentiators — most listings are just a price list.
- Cross-sell the **Freelancer** and a small-business bookkeeping template — same
  solo-owner buyer.

---

## F. Maintenance

- Edit the `JOB_PRICE`, `JOB_SUPPLIES`, `JOB_LABOR`, `JOB_TRAVEL`, `JOB_HOURS`,
  `SERVICES`, `CLIENTS`, `ONE_TIME_REV`, `ONE_TIME_JOBS`, `SCHEDULE`, `LEADS`,
  `SUPPLIES`, `TEAM`, `MILEAGE`, `EXPENSES`, `REVIEWS`, `MONTHS` constants and the
  `MARGIN_GOAL`, `HOUR_GOAL`, `PROFIT_GOAL`, `MRR_GOAL`, `CLIENT_GOAL` targets in
  `build_xlsx.py`; every KPI + the Service Score recompute. Everything is
  cross-linked — change a cost or a client and the job profit, MRR and net profit
  follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
