# Baby Command Center™ — Build Instructions

> ⚠️ Organization & record-keeping tool only — not a medical device or a
> substitute for professional healthcare. Keep the Welcome-tab disclaimer in
> every version.

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/baby-command-center/build
python3 build_xlsx.py      # -> ../Baby_Command_Center.xlsx  (21-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + the healthcare disclaimer.
2. **Executive Baby Dashboard** fills 12 KPI cards + 4 charts (feedings/day,
   sleep hours/day, weight over time, baby spending).
3. **Feeding / Sleep / Diapers** log entries; today's totals + 7-day trends
   update automatically and drive the dashboard.
4. **Growth / Milestones / Medical / Appointments** feed the growth chart,
   milestone %, vaccine % and next-appointment countdown.
5. **Budget** computes plan vs actual & remaining. **Analytics** blends a
   **Family Organization Score**. No broken cells (blank-safe totals).

> Note: `MINIFS` is used for countdowns — open in Excel 2019/365 or Google Sheets.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the trackers, then the **Dashboard** + **Analytics**. Keep
the Welcome disclaimer.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense app screenshots (sidebar of all 21 tabs
+ the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (21-tab showcase), executive dashboard, feeding + diaper
trackers, growth + milestones, and mobile.

---

## D. Etsy delivery package

```
Baby_Command_Center.xlsx          ← Excel master (21-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start + disclaimer
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| BCC-EX     | Excel only | $24 |
| BCC-GS     | Google Sheets only | $24 |
| BCC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$34** |
| BCC-PLUS   | Bundle + newborn onboarding mini-guide | $49 |
| BCC-PRO    | Nanny / caregiver / creator license | $99 |

- **Steady, evergreen demand** — new babies every day. Bumps around **baby
  showers**, **Q4 (gifts)** and **January** (new-parent organization).
- Great **baby-shower gift** angle — market it as "the gift every new parent
  actually needs."
- **Compliance:** keep all copy to organization & record-keeping — never imply
  medical guidance. This protects the listing and the buyer.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun; log generators keep the
  7-day trends populated.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
