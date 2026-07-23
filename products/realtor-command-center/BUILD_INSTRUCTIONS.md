# Real Estate Agent Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/realtor-command-center/build
python3 build_xlsx.py      # -> ../Realtor_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/tax advice"
   note.
2. **Commission Calculator** nets the flagship deal: $450,000 × 3.0% = **GCI
   $13,500**, × 70% split = **AGENT COMMISSION $9,450**, − $1,450 costs = **NET PER
   DEAL $8,000** (named `NetDeal`), net margin **85%**.
3. **Closings** sums 8 deals to **GCI YTD $75,600** (named `GCI_YTD`) and **VOLUME
   $3.6M**, avg sale **$450,000**.
4. **Business Expenses** totals $25,600, leaving a **NET INCOME of $50,000** (named
   `NetIncome`).
5. **Pipeline** shows 8 deals worth **$30,000** est. GCI; **Goals & GCI** shows
   **76%** progress to the $100k goal.
6. **Dashboard** fills 12 KPI cards + an Agent Health table + a GCI-by-month chart.
   **Agent Score 90%** (database-nurtured is the honest weak dimension). No broken
   cells; custom tables start in column B.

> Note: uses `SUM`, `COUNTA`, `AVERAGE`, `MIN`, `MAX`, `ROUNDUP`, `IFERROR` — opens in
> Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Realtor_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: commission worksheet, deal pipeline, closings log, buyers & sellers,
listings, lead-source ROI, database plan, GCI goal tracker, business expenses, mileage
log, monthly summary, and a closing checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-commission-goes & GCI
charts), everything-inside (14 tabs), the commission calculator, the deal pipeline, the
real-estate engine (commission + closings), and the **12-page printables showcase**.
Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic calculator vs Command
Center", 09 run-your-business in 4 steps, 10 what's-included / who-it's-for /
works-with. Ten images — fills all 10 Etsy slots. All headline numbers ($13,500
GCI/deal · $8,000 net · 8 closings · $3.6M volume · $75,600 GCI YTD · $50,000 net ·
76% to goal · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Realtor_Command_Center.xlsx        ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Realtor_Printables.pdf              ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| REA-GS   | The Google Sheets / Excel file only | $19 |
| REA-PDF  | The printable PDF only | $19 |
| REA-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| REA-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Strong year-round demand** with a spring buying-season peak and a January
  goal-setting spike. Priced above consumer planners — agents pay for tools that show
  their real GCI and net.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **commission calculator** and the **deal pipeline** are
  your strongest differentiators — most listings are just a calculator.
- Cross-sell the **Rental Property** analyzer — same buyer often invests too.

---

## F. Maintenance

- Edit the `SALE_PRICE`, `COMM_RATE`, `AGENT_SPLIT`, `DEAL_COSTS`, `CLOSINGS`,
  `PIPELINE`, `CLIENTS`, `LISTINGS`, `LEADSOURCES`, `DATABASE`, `EXPENSES`, `MILEAGE`,
  `GCI_GOAL`, `MONTHS` constants and the `MARGIN_GOAL`, `COMM_GOAL`, `PACE_GOAL`,
  `PIPE_GOAL`, `PROFIT_GOAL`, `TOUCH_GOAL` targets in `build_xlsx.py`; every KPI + the
  Agent Score recompute. Everything is cross-linked — change the split or add a
  closing and the GCI, net income and goal progress follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
