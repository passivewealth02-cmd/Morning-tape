# Airbnb Host Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/airbnb-host-command-center/build
python3 build_xlsx.py      # -> ../Airbnb_Host_Command_Center.xlsx  (19-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Executive Host Dashboard** fills 12 KPI cards + 4 charts (expense
   breakdown, revenue by property, occupancy by property, booking sources).
3. **Financial Command Center** rolls income & 13 expense categories into a
   live P&L — Revenue $11,600, Expenses $6,750, Net Profit $4,850, Margin 42%.
4. **Calendar** drives occupancy (59/90 = 66%), ADR ($170) & length of stay
   (3.9), plus upcoming check-in/out counts. **Reservations** nets each booking.
5. **Cleaning / Maintenance / Inventory** drive pending turnovers (4), open
   tasks (4) & low-stock alerts (4). **Analytics** blends a **Business Health
   Score** (80%). No broken cells (blank-safe `IFERROR` totals).

> Note: uses `SUMPRODUCT`, `COUNTIFS`, `SUMIF`, `COUNTA` — opens in Excel
> 2019/365 or Google Sheets.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Calendar/Financial/Reservations engines, then the
**Dashboard** + **Analytics**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense app screenshots (sidebar of all 19 tabs
+ the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (19-tab showcase), executive dashboard, financial command
center (P&L), booking calendar + reservation manager, and mobile.

---

## D. Etsy delivery package

```
Airbnb_Host_Command_Center.xlsx   ← Excel master (19-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| AHC-EX     | Excel only | $29 |
| AHC-GS     | Google Sheets only | $29 |
| AHC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$44** |
| AHC-PLUS   | Bundle + host profitability playbook | $59 |
| AHC-PRO    | Agency / property-manager / creator license | $129 |

- **Premium business tool** — hosts happily pay more for something that saves
  hours and grows profit. Price above hobby planners.
- Bumps around **Q1 (new-host tax-season setup)**, **spring** (summer-season
  prep) and whenever platforms change their fee structures.
- Lean into two angles: **"replace your 5 messy spreadsheets"** and **"know
  your real profit every month."** Both are high-intent for serious hosts.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun; the P&L, occupancy and
  dashboard KPIs recompute automatically.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Marketing numbers in `build_marketing.py` (`KPIS` list + content functions)
  should be kept in sync with the workbook's sample data.
