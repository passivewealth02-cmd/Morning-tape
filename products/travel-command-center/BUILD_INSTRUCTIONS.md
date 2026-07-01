# Travel Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/travel-command-center/build
python3 build_xlsx.py      # -> ../Travel_Command_Center.xlsx  (21-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Executive Travel Dashboard** fills 12 KPI cards + 4 charts (spending by
   category, budget vs actual, readiness, savings progress).
3. **Expenses** entries roll into **Budget** actuals per category via `SUMIF`;
   Total Budget, Spent, Remaining and % Used all reconcile.
4. **Flights / Hotels / Transport / Activities** drive the "booked/confirmed"
   KPIs; **Documents** and **Packing** drive the readiness percentages.
5. **Itinerary** auto-numbers days and shows countdowns. **Savings** feeds the
   savings-goal ring. **Analytics** blends a **Trip Readiness score**. No broken
   cells (blank-safe `IFERROR` totals).

> Note: uses `MAX`, `SUMIF`, `COUNTIF`, `COUNTA` — opens in Excel 2019/365 or
> Google Sheets.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the trackers, then the **Dashboard** + **Analytics**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense app screenshots (sidebar of all 21 tabs
+ the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (21-tab showcase), executive dashboard, budget command
center, itinerary + flights, and mobile.

---

## D. Etsy delivery package

```
Travel_Command_Center.xlsx        ← Excel master (21-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| TCC-EX     | Excel only | $24 |
| TCC-GS     | Google Sheets only | $24 |
| TCC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$34** |
| TCC-PLUS   | Bundle + trip-planning mini-guide | $49 |
| TCC-PRO    | Travel agency / creator license | $99 |

- **Evergreen with seasonal spikes** — bumps around **Q1 (New-Year trip
  planning)**, **spring break**, **summer** and **honeymoon/holiday** season.
- Lean into the **"know the true cost before you go"** and **group-trip
  split-the-bill** angles — both are strong, distinct search intents.

---

## F. Maintenance

- Edit sample trip data in `build_xlsx.py` and rerun; the budget actuals and
  dashboard KPIs recompute automatically.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Marketing numbers in `build_marketing.py` (`KPIS` list + content functions)
  should be kept in sync with the workbook's sample data.
