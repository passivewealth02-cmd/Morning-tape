# Daycare & Childcare Provider Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/daycare-command-center/build
python3 build_xlsx.py      # -> ../Daycare_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial, tax, legal or
   licensing advice" note and a reminder that ratios and CACFP tiers vary by state.
2. **Costs & Expenses**: assistant $2,400 + payroll tax $220 + insurance $185 + home share
   $640 + curriculum $95 + cleaning $140 + vehicle $85 + training $60 = **FIXED $3,825**
   (`FixedTotal`); food $135 + supplies $22 = **COST PER CHILD $157** (`CostPerChild`);
   × 11 = **VARIABLE $1,727**; **TOTAL COSTS $5,552** (`TotalCosts`).
3. **Rate & Enrollment**: $245 × 4.33 = **TUITION PER CHILD $1,060.85**
   (`TuitionPerChild`) − $157 = **NET PER CHILD $903.85** (`NetPerChild`).
4. $3,825 ÷ $903.85 = 4.23 → **BREAK-EVEN 5 children** (`BreakEven`, `ROUNDUP`), covered
   **2.20×** (`CoverRatio`), **OCCUPANCY 91.7%** (`Occupancy`). The **empty-spot block**
   shows one open spot costs **$904/month → $10,846/year**. **This block is the product's
   whole sales argument — check it renders.**
5. **Children & Families** sums to **$11,063.15** billed (`TuitionBilled`) across 11
   children at mixed rates (infant $285, full-time $245, before/after $145, part-time
   $165). **Tuition & Payments** ties to the same $11,063.15 due, **$9,641.45** collected,
   **$1,421.70** outstanding (`Outstanding`), balances flagged red.
6. **Food Program**: 210 breakfasts × $1.66 + 231 lunches × $3.13 + 231 snacks × $0.97 =
   **CACFP $1,295.70**; food spend $1,485 → **COVERS 87%** (`FoodCoverage`).
7. **Ratios & Schedule**: 11 children ÷ 2 caregivers = **5.5:1**
   (`ChildrenPerCaregiver`), and the COVERED / SHORT STAFFED flag compares caregivers on
   staff against `CaregiversNeeded` summed from the per-age-group requirement.
8. **Compliance & Files** marks a file COMPLETE only when all four documents are "Yes";
   the sample has 7 complete and 4 missing something, flagged red.
9. **Tax Set-Aside** sums **$6,400** saved (`TaxReserve`) against a **$16,000** goal, with
   the quarterly estimated-due column summing to the same $16,000.
10. **Monthly Summary**: revenue **$12,538.85** (`Revenue`) − costs $5,552 = **YOUR PAY
    $6,986.85** (`YourPay`), **MARGIN 55.7%** (`Margin`), **$32.27/hour** (`PayPerHour`)
    → **$24.20** after the set-aside (`PayPerHourNet`), run-rate **$83,842**.
11. **Dashboard** fills 12 KPI cards + a Program Health table + where-the-month-goes &
    revenue-by-month charts. **Care Score 90%** (tax reserve is the honest weak
    dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `COUNTA`, `COUNTIF`, `ROUNDUP`, `MIN`, `MAX`, `AVERAGE`,
> `IF`, `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Daycare_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter. Twelve
pages: rate & break-even worksheet, monthly costs, enrollment form, emergency & allergy
card, daily sheet, attendance sheet, tuition log, CACFP meal count, ratios & daily
schedule, compliance file checklist, tax set-aside and a monthly review.

The **daily sheet**, **enrollment form** and **emergency & allergy card** are the three
pages providers print constantly — all three sell as standalone printables on Etsy, and
all three are included here. Feature them in the printables showcase.

> This build patches the shared `table()` helper so header text (and filled sample values)
> centre **between** column dividers rather than on top of them. Carry that fix forward
> into future products derived from this script.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-month-goes & revenue
charts), everything-inside (14 tabs), the **rate & break-even engine**, the children &
tuition roster, the provider engine (both), and the **12-page printables showcase**.
Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "a printable pack vs Command Center",
09 know-your-numbers in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($245 weekly rate · $1,061
tuition/child · $157 cost/child · $904 net/child · 11 enrolled · 91.7% occupancy · $12,539
revenue · $5,552 costs · $6,987 pay · 5 break-even · $32.27/hour · 90% score) are verified
against the workbook. The where-the-month-goes donut splits to exactly the $12,538.85
revenue: your pay $6,987 (55.7%) · fixed $3,825 (30.5%) · food $1,485 (11.8%) · supplies
$242 (1.9%).

---

## D. Etsy delivery package

```
Daycare_Command_Center.xlsx         ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Daycare_Printables.pdf              ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| DAY-GS   | The Google Sheets / Excel file only | $24 |
| DAY-PDF  | The printable PDF only | $22 |
| DAY-BUNDLE | The spreadsheet + the printable PDF | **$32** |
| DAY-COMM | The same files + a commercial-use file license | $55 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom rate-setting, consultations, **coaching or mentoring**,
> licensing help, or "free updates / lifetime access" — offering a service (rather than a
> finished file) is what gets a listing removed and earns a strike. The childcare-business
> niche is full of "start your daycare with me" coaching; keep yours a file.

- **A huge Etsy audience already buying the cheap version.** Daycare printable packs sell
  in enormous volume at $5–$12 — enrollment forms, daily sheets, emergency cards. All
  three are *included* here, so you compete on quantity **and** offer the thing none of
  them do: the business maths. That's what carries the $32.
- **Demand peaks in August–September** (new program year, new enrolments) with a second
  lift in **January** (rate increases and tax season).
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; **the rate & break-even image is your single most persuasive one**
  — "five children only pay the bills, and one empty spot costs $10,846 a year" is the
  exact fact that makes a provider buy.
- **Say "ratios and CACFP tiers vary by state" plainly.** It's honest and it stops "the
  ratio was wrong for my state" messages.
- Cross-sell the **Preschool Command Center** (curriculum and themes — the same provider
  needs both, and they don't overlap at all) and the **Bookkeeping & Tax** template. A
  "Provider Bundle" of daycare + preschool at $45 is an easy upsell.

---

## F. Maintenance

- Edit the `LICENSED_CAPACITY`, `ENROLLED`, `WEEKLY_RATE`, `WEEKS_PER_MONTH`,
  `FOOD_PER_CHILD`, `SUPPLIES_PER_CHILD`, `OPEN_HOURS_WEEK`, `CAREGIVERS`, `LATE_FEES`,
  `TAX_SET_ASIDE`, `TAX_RATE` constants and the `OCCUPANCY_GOAL`, `MARGIN_GOAL`,
  `COVER_GOAL`, `RATIO_MAX`, `HOURLY_GOAL`, `TAX_RESERVE_GOAL` targets plus the
  `FIXED_LINES`, `CHILDREN`, `PAYMENTS`, `ATTENDANCE`, `MEALS`, `RATIOS`, `COMPLIANCE`,
  `SUPPLIES`, `TAXES`, `MONTHS` tables in `build_xlsx.py`; every KPI + the Care Score
  recompute.
- **Keep the sample tables tied together**: `CHILDREN` rates must match `PAYMENTS` due,
  `RATIOS` children must sum to `ENROLLED`, and `TAXES` estimated-due must sum to
  `TAX_RESERVE_GOAL`. They currently do exactly ($11,063.15 / 11 / $16,000).
- **Re-check the CACFP rates** in `MEALS` before each relist — USDA publishes new Tier I
  rates every July, and stale rates date the screenshots.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
