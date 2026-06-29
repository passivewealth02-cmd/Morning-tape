# Calm Compass™ — Build Instructions

> ⚠️ Wellness & organization tool only — not medical advice. Keep the
> Welcome-tab disclaimer in every version you ship.

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/calm-compass/build
python3 build_xlsx.py      # -> ../Calm_Compass.xlsx  (15 sheets)
```

Deterministic — same script, same workbook every run.

### Verifying

Open in Excel 365 / 2021+. On open:
1. **Welcome** shows the warm intro + disclaimer box.
2. **Dashboard** fills 8 KPI cards from the sample data + 4 trend charts.
3. **Check-In** + **Habits** drive every KPI; toggle a habit "Yes" and the
   routine score, streak, and habit-consistency bar all update.
4. **Progress** shows the blended Wellness Score and snapshot metrics.
5. No `#VALUE!`/broken cells (totals use blank-safe functions).

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`:
1. Create all 15 tabs.
2. **Settings** first (dropdown lists), then the cross-sheet named ranges.
3. Build **Check-In** & **Habits** (they feed the dashboards), incl. the
   streak helper columns.
4. **Goals, Sleep**, and the simple logs.
5. **Dashboard** + **Progress** KPIs and charts.
6. Keep the **Welcome disclaimer**. Optional Apps Script daily nudge.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py    # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs: hero, dashboard, daily check-in, habit tracker,
progress/wellness score, mobile.

---

## D. Etsy delivery package

```
Calm_Compass.xlsx                 ← Excel master (15 sheets)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding + disclaimer
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| CC-EX     | Excel only | $17 |
| CC-GS     | Google Sheets only | $17 |
| CC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$24** |
| CC-PLUS   | Bundle + a short "calm routines" mini-guide | $39 |

- **Seasonality:** strongest in **January** (new-year wellness) and around
  **back-to-school / fall**; steady year-round.
- Run a launch sale to build velocity. Use all 10 photos + a calm video.
- **Compliance:** never claim to treat/diagnose anxiety. Keep messaging to
  "wellness, organization, routines, reflection."

---

## F. Maintenance

- Edit sample data / habit names in `build_xlsx.py` and rerun.
- Brand styles, dropdowns, and conditional formats are centralized.
- New sheet: add a `build_<name>()` and slot it into the `order` list.
