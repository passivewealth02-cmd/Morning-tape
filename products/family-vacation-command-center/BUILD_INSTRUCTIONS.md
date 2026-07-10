# Family Vacation Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/family-vacation-command-center/build
python3 build_xlsx.py      # -> ../Family_Vacation_Command_Center.xlsx  (24-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Family Travel Dashboard** fills 12 KPI cards + 4 charts (budget by
   category, vacation-fund growth, trip readiness, planned vs actual).
3. **Budget Command Center** — 10 categories total **$8,500** planned /
   **$5,200** actual / **$3,300** remaining; cost/person **$1,700**, cost/day
   **$1,214**. Savings fund totals **$6,800**.
4. **Packing** shows **72%** complete (23 of 32); **Reservations** **92%**
   confirmed; **Documents** **83%** ready; **Activities** lists **14**.
5. **Analytics** blends a **Family Trip Readiness Score of ~82%**. No broken
   cells.

> Note: uses `COUNTIF`, `COUNTA`, `AVERAGE`, date math with `TODAY()` — opens
> in Excel 2019/365 or Google Sheets. The countdown recalculates every day.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Budget, Packing, Reservations & Documents, then the
**Dashboard** + **Analytics**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense app screenshots (sidebar of all 24 tabs
+ the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (24-tab showcase), budget command center, packing command
center, itinerary + reservations, and mobile. (Images 3–5 each show a different
sheet — no repeat of the hero dashboard.)

---

## D. Etsy delivery package

```
Family_Vacation_Command_Center.xlsx   ← Excel master (24-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt       ← "Make a Copy" link
START_HERE.pdf                        ← onboarding quick-start
THANK_YOU.pdf                         ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| FVC-EX     | Excel only | $19 |
| FVC-GS     | Google Sheets only | $19 |
| FVC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$29** |
| FVC-PLUS   | Bundle + printable trip kit (packing cards, itinerary) | $39 |
| FVC-PRO    | Travel-planner / creator / agency license | $99 |

- **Evergreen, high-emotion niche** — families invest heavily in trips. Bumps
  in **January** (summer planning), **spring break** and **Q3/Q4** (holiday
  travel).
- Two angles: **"remove the travel stress"** (organization) and **"never blow
  the vacation budget again"** (money). Strong **gift** angle for grandparents.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun; the countdown, budget vs
  actual, savings progress, completion %s and the Readiness Score recompute
  automatically.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Marketing numbers in `build_marketing.py` (`KPIS` list + content functions)
  should be kept in sync with the workbook's sample data.
