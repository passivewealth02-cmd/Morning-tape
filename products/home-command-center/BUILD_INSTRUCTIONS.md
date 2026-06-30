# Home Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/home-command-center/build
python3 build_xlsx.py      # -> ../Home_Command_Center.xlsx  (28-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + start-here steps.
2. **Executive Home Dashboard** fills 10 KPI cards + 6 charts (spending,
   bill status, cleaning, chores, pantry, goals).
3. **Budget** & **Bills** drive money KPIs; mark a bill Paid and watch the
   dashboard move. Overdue bills + low-stock pantry rows flag automatically.
4. **Cleaning / Chores / Meal Planner / Savings / Goals** feed the routine,
   meal-plan, savings and goal KPIs.
5. **Analytics** computes 8 health scores + a blended **Household Health
   Score**. No broken cells (blank-safe totals).

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the data sheets, then the **Dashboard** + **Analytics**.
Keep feature parity (ARRAYFORMULA / QUERY / FILTER where useful).

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs: hero, dashboard, budget+bills, cleaning/chores/meals,
household health score, mobile.

---

## D. Etsy delivery package

```
Home_Command_Center.xlsx          ← Excel master (28-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| HCC-EX     | Excel only | $29 |
| HCC-GS     | Google Sheets only | $29 |
| HCC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$39** |
| HCC-PLUS   | Bundle + onboarding mini-guide | $59 |
| HCC-PRO    | Creator / coach client-resource license | $149 |

- **Seasonality:** strong bumps in **January** (new-year organization),
  **August/September** (back-to-school), and **late December** (fresh-start
  resolutions). Steady demand year-round.
- The **Pro license** is the high-ticket play — organizing coaches, mom
  bloggers, and home-management creators resell it to their audience.
- Run a launch sale (~30–40% off) to build review velocity.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun.
- Brand styles, dropdowns, KPI cards & chart-summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
