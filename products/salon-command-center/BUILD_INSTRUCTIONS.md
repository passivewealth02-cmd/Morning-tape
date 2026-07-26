# Salon, Barber & Booth Renter Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/salon-command-center/build
python3 build_xlsx.py      # -> ../Salon_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial, tax or accounting
   advice" note and a reminder that card processing rates vary.
2. **Chair & Rent**: booth rent $900 + supplies $120 + insurance $45 + software $30 +
   education $55 = **FIXED COSTS $1,150** (`FixedCosts`); 160 open hours →
   **RENT PER CHAIR-HOUR $7.19** (`RentPerHour`); 112 booked → **UTILIZATION 70%**
   (`Utilization`).
3. **Service Pricing**: $65.00 price − $6.50 backbar − $2.19 card fee = **SERVICE NET
   $56.31** (`ServiceNet`); rent load 1.25 hrs × $7.19 = **$8.99** (`RentLoad`) →
   **YOU ACTUALLY KEEP $47.32** (`TrueNet`), **TRUE MARGIN 72.8%** (`TrueMargin`),
   **REAL RATE $37.86/hr** (`TrueHourly`).
4. Back on **Chair & Rent**: **BREAK-EVEN 21 clients** (`BreakEven`,
   `ROUNDUP(FixedCosts/ServiceNet,0)`) and **COVER RATIO 4.4×** (`CoverRatio`).
5. **Services Menu** re-runs the whole engine per row using the shared `CardRate`,
   `CardFixed` and `RentPerHour` names. Best per hour = **$55.18** (men's cut), worst =
   **$37.35** (root touch-up). Negative "you keep" values go red. **This tab is the
   product's whole sales argument — check it renders.**
6. **Retail & Backbar** totals **28** units, **$672** revenue, **$336** profit, and an
   **ATTACH RATE 11.2%** (`AttachRate`) against service revenue.
7. **Income & Tips**: revenue **$6,652** (`MonthlyRevenue`) − variable **$1,162.91**
   (`VariableCosts`) − fixed **$1,150** = **MONTHLY PROFIT $4,339** (`MonthlyProfit`),
   **NET MARGIN 65.2%**; + $1,150 tips = **TAKE-HOME $5,489** (`TakeHome`).
8. **Rebooking & Retention**: 92 served, 72 rebooked → **78%** (`RebookRate`); 3 no-shows
   → **3.3%** (`NoShowRate`); **4** new clients (`NewClients`).
9. **Dashboard** fills 12 KPI cards + a Chair Health table + where-the-$65-goes &
   revenue-by-month charts. **Chair Score 90%** (new clients per month is the honest weak
   dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMPRODUCT`, `COUNTA`, `COUNTIF`, `ROUNDUP`, `MIN`, `MAX`, `AVERAGE`,
> `IF`, `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Salon_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter. Twelve
pages: service pricing sheet, chair cost & break-even, service menu costed by the hour,
new client consultation, colour formula card, client card, day sheet, rebooking &
retention, retail sales log, product inventory & reorder, income/tips/tax set-aside and a
monthly summary & chair review.

The **colour formula card** and the **new client consultation** are the two pages
stylists print in volume — both sell as standalone printables on Etsy, and they are
included here. Feature them in the printables showcase.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the where-the-$65-goes & revenue
charts), everything-inside (14 tabs), the **true-ticket engine**, the menu costed by the
hour, the chair engine (both), and the **12-page printables showcase**. Images 3–5 each
show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "a basic tracker vs Command Center",
09 price-your-chair in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($65.00 ticket · $6.50 backbar ·
$2.19 card fee · $7.19 rent/chair-hour · $47.32 kept · 72.8% true margin · 92 clients ·
$6,652 revenue · $4,339 profit · 21 break-even · 70% utilization · 90% score) are
verified against the workbook.

---

## D. Etsy delivery package

```
Salon_Command_Center.xlsx           ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Salon_Printables.pdf                ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| SAL-GS   | The Google Sheets / Excel file only | $22 |
| SAL-PDF  | The printable PDF only | $19 |
| SAL-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| SAL-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom pricing work, consultations, **coaching or mentoring**,
> "done-for-you" menu builds, or "free updates / lifetime access" — offering a service
> (rather than a finished file) is what gets a listing removed and earns a strike. The
> beauty-business niche is saturated with coaching offers; keep yours a file.

- **A huge, warm Etsy audience.** Stylists and booth renters already buy salon printables
  on Etsy in volume; this is the same buyer with a real business problem, so $29 lands
  easily above the $5 printable crowd.
- **Demand peaks in January** (new year, new rent, new prices) with a second lift in
  August–September when leases and booth rents reset.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; **the menu-costed-by-the-hour image is the single most persuasive
  one** — a $35 cut out-earning a $95 colour is the fact that makes a stylist buy.
- Cross-sell the **Bookkeeping & Tax** and **Maker Pricing** templates — same
  self-employed buyer, and a natural "Solo Business Bundle".

---

## F. Maintenance

- Edit the `SERVICE_PRICE`, `BACKBAR_COST`, `CARD_RATE`, `CARD_FIXED`, `SERVICE_HOURS`,
  `OPEN_HOURS`, `BOOKED_HOURS`, `CLIENTS_MONTH`, `REBOOKED`, `NEW_CLIENTS`, `NO_SHOWS`,
  `TIPS_MONTH` constants and the `MARGIN_GOAL`, `REBOOK_GOAL`, `ATTACH_GOAL`,
  `COVER_GOAL`, `NOSHOW_GOAL`, `NEW_CLIENT_GOAL`, `UTIL_GOAL` targets plus the
  `FIXED_LINES`, `SERVICES`, `RETAIL`, `CLIENTS`, `APPTS`, `EXPENSES`, `RETENTION`,
  `INVENTORY`, `MONTHS` tables in `build_xlsx.py`; every KPI + the Chair Score recompute.
- **Re-check `CARD_RATE` and `CARD_FIXED`** against the buyer's own processor — Square,
  Stripe, Vagaro and GlossGenius all differ, and a stale rate dates the screenshots.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
