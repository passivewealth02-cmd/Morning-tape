# Baseball Family Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/baseball-family-command-center/build
python3 build_xlsx.py      # -> ../Baseball_Family_Command_Center.xlsx  (24-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + start-here steps.
2. **Executive Baseball Dashboard** fills 12 KPI cards + 4 charts (spending,
   games vs practices, skill progress, batting totals).
3. **Calendar** drives the countdown / conflict KPIs — add two events on the
   same day and watch the conflict flag appear.
4. **Budget** (20 categories) drives the money KPIs and cost-per-game.
   **Equipment** flags gear due within 45 days.
5. **Stats** auto-calculates AVG / OBP / SLG / OPS from the game log;
   **Pitching** computes ERA and flags any outing over your Pitch Count Limit
   (Settings). **Analytics** computes a blended **Baseball Readiness Score**.
   No broken cells (blank-safe totals).

> Note: `MINIFS`/`COUNTIFS` are used for countdowns — open in Excel 2019/365
> or Google Sheets (both supported).

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the data sheets, then the **Dashboard** + **Analytics**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense "app screenshots" (sidebar of all 24
tabs + the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (24-tab showcase), executive dashboard, the full 20-row
budget, player statistics (auto slash line) + season calendar, and mobile.

---

## D. Etsy delivery package

```
Baseball_Family_Command_Center.xlsx   ← Excel master (24-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt       ← "Make a Copy" link
START_HERE.pdf                        ← onboarding quick-start
THANK_YOU.pdf                         ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| BFC-EX     | Excel only | $29 |
| BFC-GS     | Google Sheets only | $29 |
| BFC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$39** |
| BFC-PLUS   | Bundle + onboarding mini-guide | $59 |
| BFC-PRO    | Team / coach client-resource license | $149 |

- **Seasonality:** strong bumps at **registration / tryouts (Jan–Mar)** and
  **spring/summer (Apr–Jul)** during peak season; travel-ball families buy
  year-round.
- The **Pro license** is the high-ticket play — team managers and coaches buy
  it to hand to every family on the roster.
- Run a launch sale (~30–40% off) to build review velocity.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun.
- Brand styles, dropdowns, KPI cards & chart-summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
