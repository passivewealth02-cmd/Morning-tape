# Notary & Loan Signing Agent Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/notary-command-center/build
python3 build_xlsx.py      # -> ../Notary_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not legal, tax or accounting
   advice" note **and** the two state-specific warnings (maximum fee per notarial act;
   bound sequential journal requirements).
2. **Signing Profit**: $125 ÷ 0.9 hr = **LOOKS LIKE $138.89/hr** (`LooksLike`);
   180 pages × $0.035 = **PRINTING $6.30** (`PrintCost`); 38 miles × $0.22 = **DRIVING
   $8.36** (`DriveCost`); **NET PER SIGNING $110.34** (`NetPerSigning`).
3. Prep 0.5 + drive 1.1 + appointment 0.9 = **2.5 HOURS** (`TotalHours`) → **REAL HOURLY
   $44.14** (`RealHourly`), a **$94.75** gap. **This tab is the product's whole sales
   argument — check it renders.**
4. Mileage: 38 × $0.70 = **$26.60** per signing (`MileageDeduction`) × 52 = **$1,383**
   (`MileageMonth`).
5. **Signings Log** computes cost and net per row from the shared `CostPerPage` and
   `VehicleCPM` names; "Overdue" shades amber and "fell through" shades red. **PAID RATE
   96.2%** (`PaidRate`) from 52 signings less 2 unpaid.
6. **Fee Schedule** holds 13 priced services including the **trip fee** for a door
   cancellation, with a state-maximum-fee warning underneath.
7. **Invoices** sums to exactly **52 signings** and **$6,500** (`Invoiced`), matching
   `SigningsMonth × AvgFee`, with **$2,140** outstanding (`Receivable`).
8. **Notarial Journal** records 12 acts (`ActsRecorded`) with the bound-journal warning
   above the table.
9. **Expenses**: fixed lines sum to **$289** (`FixedTotal`); other running costs sum
   separately (`OtherTotal`).
10. **Tax Set-Aside** sums **$2,100** saved (`TaxReserve`) against **$5,250** estimated
    due — the quarterly due column sums to the same $5,250, which is also
    `TaxReserveGoal`.
11. **Monthly Summary**: revenue **$6,500** (`Revenue`) − variable − fixed = **PROFIT
    $5,449** (`Profit`) at **83.8%** margin; break-even **3 signings** (`BreakEven`),
    covered **17.3×** (`CoverRatio`).
12. **Dashboard** fills 12 KPI cards + a Business Health table + a where-the-2.5-hours-go
    donut and revenue-by-month chart. **Signing Score 90%** (tax reserve is the honest
    weak dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `COUNTA`, `MIN`, `MAX`, `ROUNDUP`, `AVERAGE`, `IF`, `IFERROR` — opens
> in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Notary_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter — designed to
live in the signing bag. Twelve pages: what this signing really pays, fee schedule,
signings log, mileage log, **notarial journal sheet**, **signing day checklist**,
**invoice**, who owes you, signing companies, printing & supplies, tax set-aside and a
monthly review.

The **journal sheet**, the **signing day checklist** and the **invoice** are the three
pages agents print in volume — feature them in the printables showcase. Page 5 prints the
bound-journal warning in red above the table; keep it.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-2.5-hours-go donut and
revenue chart), everything-inside (14 tabs), the **real-hourly profit engine**, the fee
schedule netted, the signing engine (both), and the **12-page printables showcase**.
Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "a signing app vs Command Center", 09
know-your-real-rate in 4 steps, 10 what's-included / who-it's-for / works-with. Ten images
— fills all 10 Etsy slots. All headline numbers ($125 fee · $138.89 looks-like · $6.30
printing · $8.36 driving · $110.34 net · $44.14 real hourly · 52 signings · $6,500 revenue
· $5,449 profit · 3 break-even · $1,383 mileage deduction · 90% score) are verified against
the workbook, and the fee-schedule nets ($154.95 / $133.56 / $110.34 / $86.08 / $53.70)
recompute exactly from pages × $0.035 + miles × $0.22.

The crest is a **notary seal** — embossed ring, serrated rim and a five-pointed star.

---

## D. Etsy delivery package

```
Notary_Command_Center.xlsx          ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Notary_Printables.pdf               ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| NOT-GS   | The Google Sheets / Excel file only | $19 |
| NOT-PDF  | The printable PDF only | $17 |
| NOT-BUNDLE | The spreadsheet + the printable PDF | **$25** |
| NOT-COMM | The same files + a commercial-use file license | $45 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, **coaching or mentoring**, "become a signing agent" training,
> business consultations, or "free updates / lifetime access". The NSA space is *saturated*
> with coaching and course offers — that is exactly the category Etsy removes, so yours
> must stay a plain digital file.
>
> Also do not offer notarial services through the listing. You are selling a spreadsheet.

- **Priced lowest in the business list on purpose.** The NSA audience is large, active and
  buys tools constantly, but many are in their first year and price-sensitive. $25 is a
  volume play — expect this one to sell more units than the contractor or trucking
  products even though it carries a smaller ticket.
- **Demand peaks whenever rates drop** (refi waves put thousands of new agents into the
  market) and again every **January** for tax season and mileage-rate updates.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; **the real-hourly engine is your single most persuasive image** —
  "$138.89 an hour is really $44.14 an hour" is a genuinely uncomfortable fact for this
  audience and it converts.
- **Say "check your state's maximum fee" plainly.** It's honest, it protects you, and
  experienced notaries respect that you know fees are statutory.
- Cross-sell the **Bookkeeping & Tax** template — same Schedule C filer. A "Signing Agent
  Bundle" of both at $49 is an easy upsell.

---

## F. Maintenance

- Edit the `AVG_FEE`, `PAGES_PER_SIGNING`, `COST_PER_PAGE`, `MILES_PER_SIGNING`,
  `VEHICLE_COST_PER_MILE`, `IRS_MILEAGE_RATE`, `DRIVE_HOURS`, `APPT_HOURS`, `PREP_HOURS`,
  `SIGNINGS_MONTH`, `UNPAID`, `TAX_SET_ASIDE` constants and the `HOURLY_GOAL`,
  `MARGIN_GOAL`, `COVER_GOAL`, `PAID_GOAL`, `SIGNING_GOAL`, `TAX_RESERVE_GOAL` targets
  plus the `FIXED_LINES`, `FEES`, `SIGNINGS`, `MILEAGE`, `PRINTING`, `INVOICES`,
  `COMPANIES`, `JOURNAL`, `EXPENSES`, `TAXES`, `MONTHS` tables in `build_xlsx.py`.
- **Keep the tie-outs**: `INVOICES` must sum to `SIGNINGS_MONTH` signings and
  `SIGNINGS_MONTH × AVG_FEE` in value, and `TAXES` estimated-due must sum to
  `TAX_RESERVE_GOAL`. They currently do exactly (52 / $6,500 / $5,250).
- **Update `IRS_MILEAGE_RATE` every January.** The IRS publishes a new standard mileage
  rate annually, and a stale rate dates every screenshot in the listing — this is the
  single most important maintenance item on this product.
- Keep `build_marketing.py`'s KPIs and the fee-schedule net figures in sync with the
  workbook.
