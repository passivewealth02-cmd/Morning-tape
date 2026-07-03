# Pickleball Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/pickleball-command-center/build
python3 build_xlsx.py      # -> ../Pickleball_Command_Center.xlsx  (17-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Executive Dashboard** fills 12 KPI cards + 4 charts (match results,
   court time by month, skill progress, spending).
3. **Match Tracker** — 24 matches, **67% win rate** (16–8); W/L glow mint/red.
4. **Partners** compute records from matches (Chris P. 10–4 is top);
   **Tournaments** show 2 gold medals ($350 prize) + 3 upcoming registrations.
5. **Skills** chart start-vs-now (avg 6.8/10); **Equipment** flags 2 items
   "due soon"; **Budget** shows $1,665 of $2,500 (67%).
6. **Analytics** blends a **Pickleball Performance Score** (65%). No broken
   cells (blank-safe `IFERROR` totals).

> Note: uses `SUMPRODUCT`, `COUNTIFS`, `AVERAGEIF`, `INDEX/MATCH` — opens in
> Excel 2019/365 or Google Sheets.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Match Tracker & trackers, then the **Dashboard** +
**Analytics**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense app screenshots (sidebar of all 17 tabs
+ the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (17-tab showcase), executive dashboard, match tracker,
skills + tournaments, and mobile.

---

## D. Etsy delivery package

```
Pickleball_Command_Center.xlsx    ← Excel master (17-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| PCC-EX     | Excel only | $19 |
| PCC-GS     | Google Sheets only | $19 |
| PCC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$29** |
| PCC-PLUS   | Bundle + drills & strategy playbook | $39 |
| PCC-PRO    | Coach / club / creator license | $79 |

- **Fastest-growing sport in the US** — huge, still-expanding market. Bumps in
  **spring/summer** (outdoor season) and **January** (new-year goals).
- Lean into two angles: **"improve your game"** (competitive players) and
  **"finally organized"** (rec players & seniors). Level-agnostic = broadest reach.
- Great **club/coach** upsell (PCC-PRO) for managing multiple players/leagues.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun; win %, court time, partner
  records and the performance score recompute automatically.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Marketing numbers in `build_marketing.py` (`KPIS` list + content functions)
  should be kept in sync with the workbook's sample data.
