# Road Trip Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/road-trip-command-center/build
python3 build_xlsx.py      # -> ../Road_Trip_Command_Center.xlsx  (19-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Road Trip Dashboard** fills 12 KPI cards + 4 charts (budget by category,
   daily mileage, trip readiness, planned vs actual).
3. **Route Planner** totals **1,555 miles / 28.6 hours**; **Fuel Tracker**
   estimates **$350** at 16 MPG and computes cost per mile.
4. **Budget** — 12 categories total **$4,800** planned / **$1,650** actual /
   **$3,150** remaining; cost/traveler **$2,400**, cost/day **$400**.
5. **Vehicle** shows **80%** readiness (8 of 10 OK); **Packing** **77%** packed;
   **3** hotels + **4** campgrounds booked; **Analytics** blends a **Trip
   Readiness Score of ~81%**. No broken cells.

> Note: uses `COUNTIF(S)`, `COUNTA`, `AVERAGE`, `SUM`, date math with `TODAY()`
> — opens in Excel 2019/365 or Google Sheets. The countdown recalculates daily.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Route, Budget, Fuel, Vehicle, Stays & Camping, then the
**Dashboard** + **Analytics**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense app screenshots (sidebar of all 19 tabs
+ the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (19-tab showcase), route planner, budget command center,
vehicle + packing, and mobile. (Images 3–5 each show a different sheet — no
repeat of the hero dashboard.)

---

## D. Etsy delivery package

```
Road_Trip_Command_Center.xlsx     ← Excel master (19-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| RTC-EX     | Excel only | $19 |
| RTC-GS     | Google Sheets only | $19 |
| RTC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$29** |
| RTC-PLUS   | Bundle + printable trip kit (route cards, checklists) | $39 |
| RTC-PRO    | Creator / travel-agency license | $99 |

- **Evergreen adventure niche** — RV, van-life and national-park travel are
  booming. Bumps in **spring** (summer-trip planning) and **Q4** (holiday
  road trips).
- Two angles: **"never break down or blow the budget"** (peace of mind) and
  **"one file for the whole adventure"** (organization). Strong van-life and
  national-park-tour reach.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun; the countdown, total mileage,
  fuel economy, budget remaining, packing %, vehicle readiness and the Trip
  Readiness Score recompute automatically.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Marketing numbers in `build_marketing.py` (`KPIS` list + content functions)
  should be kept in sync with the workbook's sample data.
