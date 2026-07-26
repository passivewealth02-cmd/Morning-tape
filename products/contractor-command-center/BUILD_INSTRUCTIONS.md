# Contractor Job Costing & Bidding Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/contractor-command-center/build
python3 build_xlsx.py      # -> ../Contractor_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial, tax or legal
   advice" note and a reminder that labor burden and insurance vary by state and trade.
2. **Bid Builder**: materials $14,000 + labor (320 hrs × $45 = $14,400) + subs $8,500 +
   equipment $1,200 = **DIRECT COST $38,100** (`DirectCost`); overhead 12% =
   **$4,572** (`Overhead`); **TOTAL COST $42,672** (`TotalCost`).
3. **BID PRICE $53,340** (`BidPrice`) — `=IFERROR(TotalCost/(1-MarginTarget),0)`, i.e.
   cost divided by what's left of the dollar. **PLANNED PROFIT $10,668**
   (`PlannedProfit`).
4. The **markup-is-not-margin** block directly beneath it computes the wrong way for
   contrast: `TotalCost*(1+MarginTarget)` = **$51,206** at a real margin of **16.7%**, a
   **$2,134** shortfall per job. This block is the product's whole sales argument — check
   it renders.
5. **Job Costing** tracks actuals against the bid: **ACTUAL COST $41,178**
   (`ActualCost`), **ACTUAL PROFIT $12,162** (`ActualProfit`), **ACTUAL MARGIN 22.8%**
   (`ActualMargin`). Overruns flag red per line.
6. **Jobs & Pipeline** counts **9** jobs (`JobCount`) and **$186,000** backlog
   (`Backlog`); **Invoices** totals **$31,000** outstanding (`Receivable`); **Bid Log**
   computes a **40.9%** win rate (`WinRate`, 9 of 22) against `WinRateGoal`.
7. **Dashboard** fills 12 KPI cards + a Builder Health table + where-the-bid-dollar-goes
   & revenue-by-month charts. **Builder Score 90%** (bid win rate is the honest weak
   dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `SUMPRODUCT`, `COUNTA`, `COUNTIF`, `AVERAGE`, `ROUND`, `IF`,
> `IFERROR`, `AND` — opens in Google Sheets or Excel 2019/365. `WinRateGoal` is a scalar
> defined name, not a cell reference.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Contractor_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter — designed
to survive a clipboard on a job site. Twelve pages: bid worksheet, job cost sheet,
materials takeoff, labor & crew log, subcontractor log, change order form, daily job log,
bid log, invoice schedule, punch list, monthly summary and a job closeout sheet.

The **change order form** and the **punch list** are the two pages contractors actually
print in volume — they are the pages to feature in the printables showcase.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the bid-dollar & revenue charts),
everything-inside (14 tabs), the **bid engine**, the jobs pipeline, the builder engine
(both), and the **12-page printables showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "napkin estimate vs Command Center",
09 bid-a-job in 4 steps, 10 what's-included / who-it's-for / works-with. Ten images —
fills all 10 Etsy slots. All headline numbers ($38,100 direct · $4,572 overhead · $42,672
total cost · $53,340 bid · $10,668 planned profit · $41,178 actual cost · $12,162 actual
profit · 22.8% actual margin · 9 jobs · $186,000 backlog · $31,000 receivables · 40.9%
win rate · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Contractor_Command_Center.xlsx      ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Contractor_Printables.pdf           ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| CON-GS   | The Google Sheets / Excel file only | $27 |
| CON-PDF  | The printable PDF only | $24 |
| CON-BUNDLE | The spreadsheet + the printable PDF | **$39** |
| CON-COMM | The same files + a commercial-use file license | $65 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom bid building, consultations, **coaching or mentoring**,
> "done-for-you" estimates, or "free updates / lifetime access" — offering a service
> (rather than a finished file) is what gets a listing removed and earns a strike. The
> contractor space is full of "I'll build your estimate" offers; keep yours a file.

- **The highest-value buyer in the shop.** A contractor mis-bidding one kitchen loses
  $2,000+; $39 is a rounding error. This is the priciest product in the catalogue and it
  should be — do not discount it.
- **Demand peaks Jan–Mar** (bidding season, before spring builds) with a second lift in
  September. List it well before January.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; **the markup-vs-margin comparison is the single most persuasive
  image** — it shows a $2,134 hole in a bid the contractor thought was fine.
- Cross-sell the **Bookkeeping & Tax** and **Renovation Budget** templates — same trade
  buyer, and a natural "Trades Bundle".

---

## F. Maintenance

- Edit the `MATERIALS`, `LABOR_HOURS`, `LABOR_RATE`, `SUBS`, `EQUIP`, `OVERHEAD_RATE`,
  `MARGIN_TARGET`, `ACTUAL_COST` constants and the `MARGIN_GOAL`, `OVERHEAD_GOAL`,
  `CO_GOAL`, `AR_GOAL`, `BACKLOG_GOAL`, `WINRATE_GOAL` targets plus the `JOBS`, `CREW`,
  `MATERIAL_LINES`, `SUBLINES`, `CHANGE_ORDERS`, `EQUIPMENT`, `INVOICES`, `BIDS`,
  `MONTHS` tables in `build_xlsx.py`; every KPI + the Builder Score recompute.
- **Re-check `LABOR_RATE` against your fully burdened rate** — payroll taxes, workers'
  comp, liability insurance and vehicle costs, not the wage. A raw wage rate is the
  second-most-common reason a job comes in short.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
