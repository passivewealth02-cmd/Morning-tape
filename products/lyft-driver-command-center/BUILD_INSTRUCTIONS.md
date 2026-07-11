# Lyft Driver Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/lyft-driver-command-center/build
python3 build_xlsx.py      # -> ../Lyft_Driver_Command_Center.xlsx  (18-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Driver Dashboard** fills 12 KPI cards + 4 charts (net-earnings trend,
   earnings mix, best shifts, expense breakdown).
3. **Shift Log** — 16 shifts totalling Gross **$4,226**, **2,871** miles,
   **128.5** hours, **316** trips, **$522** fuel; earnings data-bars per shift.
4. **Expenses** pulls Fuel from the log (`=FuelTotal`) so Total **$1,440** stays
   honest; **Analytics** computes Net **$2,786**, Net **$21.68/hr**, Net
   **$0.97/mi** and a **Driver Health Score of ~90%**.
5. **Tax Center** shows the mileage deduction **$2,010** (IRS $0.70/mi) and a
   quarterly set-aside; **Savings** tracks the emergency fund at **70%**. No
   broken cells.

> Note: uses `SUM`, `SUMIF`, `COUNT`, `MIN/MAX`, `AVERAGE` and `MileageRate` from
> Settings — opens in Excel 2019/365 or Google Sheets. Change the IRS rate in
> Settings and every tax figure updates.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Shift Log, Expenses, Savings, Analytics & Tax Center,
then the **Dashboard**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py         # -> ../marketing/01..06.png
python3 build_marketing_detail.py              # -> ../marketing/07..10.png
```

**Six app-screenshots** (sidebar of all 18 tabs + real computed KPI numbers +
full tables/charts): hero, everything-inside (18-tab showcase), shift log,
earnings → real take-home, hot zones + bonuses, and mobile. (Images 3–5 each
show a different sheet — no repeat of the hero dashboard.)

**Four detailed / benefit-driven images**: 07 feature spotlights, 08 "basic
mileage app vs Command Center" comparison, 09 up-and-running in 4 steps, 10
what's-included / who-it's-for / guarantee. Ten images total — fills all 10 Etsy
photo slots.

---

## D. Etsy delivery package

```
Lyft_Driver_Command_Center.xlsx   ← Excel master (18-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| LDC-EX     | Excel only | $15 |
| LDC-GS     | Google Sheets only | $15 |
| LDC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$22** |
| LDC-PLUS   | Bundle + tax & bonus playbook | $29 |
| LDC-PRO    | Fleet / referral license | $79 |

- **Huge, always-on niche** — millions of active rideshare & delivery drivers,
  most of whom track nothing. Bumps in **January** (new-year budgeting + tax
  season) and **Q1** as 1099s land.
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
