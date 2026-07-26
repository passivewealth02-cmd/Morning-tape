# Trucking Owner-Operator Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/trucking-command-center/build
python3 build_xlsx.py      # -> ../Trucking_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial, tax, legal or
   DOT-compliance advice" note.
2. **Fixed Costs**: tractor $2,200 + trailer $600 + insurance $1,100 + permits $250 +
   ELD $65 + parking $150 + accounting $125 = **FIXED $4,490** (`FixedTotal`); ÷ 11,200
   miles = **FIXED CPM $0.401** (`FixedCPM`).
3. **Variable Costs**: $3.90 ÷ 6.5 MPG = **FUEL CPM $0.600** (`FuelCPM`); + maintenance
   $0.18 + tires $0.045 + tolls $0.035 = **VARIABLE CPM $0.860** (`VarCPM`).
4. **Cost Per Mile**: 10,000 loaded + 1,200 deadhead = **11,200 TOTAL** (`TotalMiles`),
   **DEADHEAD 10.7%** (`DeadheadPct`); **COST PER MILE RUN $1.261** (`TotalCPM`); ×
   11,200 = **TOTAL COST $14,122** (`TotalCostMonth`); ÷ **10,000 loaded** = **COST PER
   LOADED MILE $1.41** (`CostPerLoaded`) — the rate floor.
5. Rate $2.35 → **PROFIT PER MILE $0.938** (`ProfitPerMile`), **COVER 1.66×**
   (`CoverRatio`). The **deadhead warning block** below shows the $2.35 load is really
   **$2.10** per mile actually driven, costing **$2,518** a month. **This block is the
   product's whole sales argument — check it renders.**
6. **Loads** sums to exactly **10,000** loaded, **1,200** deadhead and **$23,500**
   (`MonthlyRevenue`) at an average **$2.35**/loaded mile. Each row computes an all-in
   rate and any load under `CostPerLoaded` goes red.
7. **Fuel Log** sums **1,723** gallons (`GallonsMonth`) → **ACTUAL MPG 6.50**
   (`ActualMPG`), which ties back to the 6.5 in Settings. **IFTA & Miles** sums to
   **11,200** (`IftaTotal`), matching total miles.
8. **Reserve Fund** totals **$4,000** saved (`Reserve`) against a **$10,000** goal.
9. **Monthly Summary**: $23,500 − $14,122 = **PROFIT $9,378** (`MonthlyProfit`), **NET
   MARGIN 39.9%** (`NetMargin`), run-rate **$112,536**.
10. **Dashboard** fills 12 KPI cards + a Road Health table + where-the-$2.35-goes &
    revenue-by-month charts. **Road Score 90%** (maintenance reserve is the honest weak
    dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `COUNTA`, `MIN`, `AVERAGE`, `IF`, `IFERROR` — opens in
> Google Sheets or Excel 2019/365. `kpi_card` in this build adds a `money3`
> (`$#,##0.000`) format because per-mile costs need three decimals.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Trucking_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter — designed
to live on a clipboard in the cab. Twelve pages: cost-per-mile worksheet, fixed costs,
"should I take this load?", trip sheet, fuel log, PM schedule, settlement reconciliation,
IFTA by state, pre-trip inspection, reserve fund, monthly summary and a **rate floor card
for the dash**.

The **rate floor card** and the **"should I take this load?"** page are the two that get
printed and used daily — feature them in the printables showcase.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-$2.35-goes & revenue
charts), everything-inside (14 tabs), the **cost-per-mile engine**, the load log scored
against the floor, the road engine (both), and the **12-page printables showcase**.
Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "free calculator vs Command Center",
09 find-your-rate-floor in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($4,490 fixed · $0.401 fixed CPM ·
$0.860 variable CPM · $1.261 cost per mile run · $1.41 cost per loaded mile · $2.35 rate ·
$0.938 profit/mile · 10,000 loaded miles · 10.7% deadhead · $23,500 revenue · $9,378
profit · 90% score) are verified against the workbook. The where-the-$2.35-goes donut
splits to exactly $2.350: profit $0.938 (39.9%) · fuel $0.672 (28.6%) · fixed $0.449
(19.1%) · maintenance & tires $0.291 (12.4%).

---

## D. Etsy delivery package

```
Trucking_Command_Center.xlsx        ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Trucking_Printables.pdf             ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| TRK-GS   | The Google Sheets / Excel file only | $27 |
| TRK-PDF  | The printable PDF only | $24 |
| TRK-BUNDLE | The spreadsheet + the printable PDF | **$39** |
| TRK-COMM | The same files + a commercial-use file license | $65 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom cost analysis, consultations, **coaching or mentoring**,
> dispatch services, or "free updates / lifetime access" — offering a service (rather
> than a finished file) is what gets a listing removed and earns a strike. The trucking
> niche is full of "I'll be your dispatcher / mentor" offers; keep yours a file.

- **A business buyer with enormous stakes.** An owner-operator running a hundred loads a
  year at ten cents a mile too cheap loses five figures. $39 is nothing against that, and
  this audience is not price-shopping spreadsheets.
- **Demand is steady all year** with a spike in **January** (new authority, new tax year)
  and again when diesel spikes, which is when everyone suddenly wants a cost-per-mile
  number.
- Use all 10 photos + a walkthrough video that shows the *file* (not you offering
  dispatch). Lead photo = the feature-forward hero; **the deadhead block is your single
  most persuasive image** — showing a $2.35 load is really $2.10 is the exact fact that
  makes a driver buy.
- **Say "not DOT-compliance advice" plainly** and note that the pre-trip page doesn't
  replace a required DVIR. Drivers respect that you know the difference.
- Cross-sell the **Bookkeeping & Tax** template — same self-employed buyer with a
  Schedule C, and a natural "Owner-Operator Bundle" at $59.

---

## F. Maintenance

- Edit the `MPG`, `DIESEL_PRICE`, `MAINT_CPM`, `TIRE_CPM`, `TOLL_CPM`, `LOADED_MILES`,
  `DEADHEAD_MILES`, `RATE_PER_MILE` constants and the `PPM_GOAL`, `MARGIN_GOAL`,
  `DEADHEAD_GOAL`, `COVER_GOAL`, `MPG_GOAL`, `RESERVE_GOAL` targets plus the
  `FIXED_LINES`, `VARIABLE_LINES`, `LOADS`, `SETTLEMENTS`, `FUEL`, `MAINTENANCE`,
  `EQUIPMENT`, `IFTA`, `RESERVE`, `MONTHS` tables in `build_xlsx.py`; every KPI + the
  Road Score recompute.
- **Keep the three sample tables tied together**: `LOADS` must sum to `LOADED_MILES` and
  `DEADHEAD_MILES`, `IFTA` must sum to their total, and `FUEL` gallons must divide into
  total miles at `MPG`. They currently do exactly (10,000 / 1,200 / 11,200 / 1,723 gal).
  If you change one, re-run the tie-out check.
- **Re-check `DIESEL_PRICE`** before each relist — a stale diesel price dates every
  screenshot in the listing.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
