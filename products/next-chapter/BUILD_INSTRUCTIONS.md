# Next Chapter™ — Build Instructions

> ⚠️ Organizational & planning tool only — not legal, financial, or
> mental-health advice. Keep the Welcome-tab disclaimer in every version.

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/next-chapter/build
python3 build_xlsx.py      # -> ../Next_Chapter.xlsx  (20 sheets)
```

### Verifying
1. **Welcome** shows the intro + disclaimer box.
2. **Life Dashboard** fills 8 KPI cards + budget/asset donuts + progress/net-worth bars.
3. **Tasks** & **Documents** drive completion KPIs; mark items Complete/Collected and watch the dashboard move.
4. **Expenses → Budget** actuals flow via `SUMIF`; **Finances** computes net worth.
5. No broken cells (blank-safe totals).

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: create all 20 tabs, build **Settings** first,
define the cross-sheet named ranges, then the data sheets (Tasks,
Documents, Finances, Budget, Expenses, Goals, Appointments), then the
**Dashboard** + **Analytics**. Keep the Welcome disclaimer.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs: hero, dashboard, tasks, finances, analytics, mobile.

---

## D. Etsy delivery package

```
Next_Chapter.xlsx                 ← Excel master (20 sheets)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding + disclaimer
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| NC-EX     | Excel only | $24 |
| NC-GS     | Google Sheets only | $24 |
| NC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$34** |
| NC-PLUS   | Bundle + onboarding mini-guide | $59 |
| NC-PRO    | Coach / mediator client-resource license | $129 |

- **Steady, year-round demand** (no strong seasonality). January and
  September see small bumps ("fresh start" periods).
- The **Pro license** is the high-ticket play — divorce coaches and
  mediators hand it to clients.
- **Compliance:** never give legal/financial/mental-health advice in the
  listing or product. Keep all copy to "organization & planning."

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun.
- Brand styles, dropdowns, and conditional formats are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
