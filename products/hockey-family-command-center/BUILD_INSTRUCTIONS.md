# Hockey Family Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/hockey-family-command-center/build
python3 build_xlsx.py      # -> ../Hockey_Family_Command_Center.xlsx  (25-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + start-here steps.
2. **Executive Hockey Dashboard** fills 12 KPI cards + 4 charts (spending,
   practices vs games, skill progress, season totals).
3. **Calendar** drives the countdown / conflict KPIs — add two events on the
   same day and watch the conflict count rise.
4. **Budget** (26 categories) drives the money KPIs and cost-per-game.
   **Equipment** flags gear due within 45 days; **Sharpening** reminds at 14 days.
5. **Practice** feeds ice-time & attendance; **Stats** totals G/A/PTS;
   **Development** & **Goals** feed progress bars. **Analytics** computes a
   blended **Hockey Readiness Score**. No broken cells (blank-safe totals).

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

Six 2000×2000 PNGs, rendered as dense "app screenshots" (sidebar of all 25
tabs + the real computed KPI numbers + fully populated tables/charts):
hero, everything-inside (25-tab showcase), executive dashboard, the full
26-row budget, season calendar + equipment, and mobile.

---

## D. Etsy delivery package

```
Hockey_Family_Command_Center.xlsx   ← Excel master (25-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
START_HERE.pdf                      ← onboarding quick-start
THANK_YOU.pdf                       ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| HFC-EX     | Excel only | $29 |
| HFC-GS     | Google Sheets only | $29 |
| HFC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$39** |
| HFC-PLUS   | Bundle + onboarding mini-guide | $59 |
| HFC-PRO    | Team / coach client-resource license | $149 |

- **Seasonality:** strong bumps at **tryout/registration time (Aug–Sep)** and
  **New Year**; steady demand through the season (Oct–Mar). Spring/AAA bump.
- The **Pro license** is the high-ticket play — team managers and coaches buy
  it to hand to every family on the roster.
- Run a launch sale (~30–40% off) to build review velocity.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun.
- Brand styles, dropdowns, KPI cards & chart-summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
