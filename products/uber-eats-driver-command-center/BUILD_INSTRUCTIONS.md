# Uber Eats Driver Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/uber-eats-driver-command-center/build
python3 build_xlsx.py      # -> ../Uber_Eats_Driver_Command_Center.xlsx  (18-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Driver Dashboard** fills 12 KPI cards + 4 charts (net-earnings trend,
   earnings mix, best shifts, expense breakdown).
3. **Delivery Log** — 16 shifts totalling Gross **$2,544** (Tips **$1,258**),
   **1,502** miles, **90.0** hours, **321** deliveries, **$293** gas; earnings
   data-bars per shift; apps color-code (Uber Eats / DoorDash / Grubhub / Multi).
4. **Expenses** pulls Fuel from the log (`=FuelTotal`) so Total **$1,018** stays
   honest; **Analytics** computes Net **$1,526**, Net **$16.96/hr**, Net
   **$1.02/mi** and a **Driver Health Score of ~87%**.
5. **Tax Center** shows the mileage deduction **$1,051** (IRS $0.70/mi) and a
   quarterly set-aside; **Savings** tracks the emergency fund at **65%**. No
   broken cells.

> Note: uses `SUM`, `SUMIF`, `COUNT`, `MIN/MAX`, `AVERAGE` and `MileageRate` from
> Settings — opens in Excel 2019/365 or Google Sheets. Change the IRS rate in
> Settings and every tax figure updates.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Delivery Log, Expenses, Savings, Analytics & Tax Center,
then the **Dashboard**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py         # -> ../marketing/01..06.png
python3 build_marketing_detail.py              # -> ../marketing/07..10.png
```

**Six app-screenshots** (sidebar of all 18 tabs + real computed KPI numbers +
full tables/charts): hero, everything-inside (18-tab showcase), delivery log,
earnings → real take-home, hotspots + promos, and mobile. (Images 3–5 each show
a different sheet — no repeat of the hero dashboard.)

**Four detailed / benefit-driven images**: 07 feature spotlights, 08 "basic
mileage app vs Command Center" comparison, 09 up-and-running in 4 steps, 10
what's-included / who-it's-for / guarantee. Ten images total — fills all 10 Etsy
photo slots.

---

## D. Etsy delivery package

```
Uber_Eats_Driver_Command_Center.xlsx  ← Excel master (18-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt        ← "Make a Copy" link
START_HERE.pdf                         ← onboarding quick-start
THANK_YOU.pdf                          ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| UEC-EX     | Excel only | $15 |
| UEC-GS     | Google Sheets only | $15 |
| UEC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$22** |
| UEC-PLUS   | Bundle + tax & promo playbook | $29 |
| UEC-PRO    | Fleet / referral license | $79 |

- **Huge, always-on niche** — millions of active delivery drivers, most tracking
  nothing. Bumps in **January** (new-year budgeting + tax season) and **Q1** as
  1099s land.
- Two angles: **"know your real take-home / $ per hour"** (earnings) and **"the
  mileage deduction most drivers miss"** (taxes). Priced for impulse buy.

---

## F. Maintenance

- Edit the `SHIFTS` list and expense inputs in `build_xlsx.py` and rerun; gross,
  net, $/hour, $/mile, miles, the mileage deduction, savings and the Driver
  Health Score all recompute automatically.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Keep `build_marketing.py` numbers in sync with the workbook's sample data.
