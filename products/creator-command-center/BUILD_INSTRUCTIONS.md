# Creator Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/creator-command-center/build
python3 build_xlsx.py      # -> ../Creator_Command_Center.xlsx  (24-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here guide (and the password-manager
   security note).
2. **Executive Business Dashboard** fills 12 KPI cards + 4 charts (revenue by
   source, audience growth, content by platform, expense breakdown).
3. **Revenue Command Center** rolls 10 income streams into a live P&L —
   Revenue $8,160, Expenses $2,390, Net Profit $5,770, Margin 71%,
   run-rate $97,920.
4. **Sponsorship CRM** shows 4 active deals (Signed/Delivered/Negotiating);
   **Performance** ranks top content (148k-view TikTok #1).
5. **Calendar** counts 8 Scheduled; **Analytics** blends a **Business Health
   Score** (75%) and charts audience growth to 72.4k. No broken cells.

> Note: uses `SUMIFS`, `COUNTIFS`, `RANK`, `INDEX/MATCH`-style refs — opens in
> Excel 2019/365 or Google Sheets.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Revenue/Expense engines & trackers, then the
**Dashboard** + **Analytics**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense app screenshots (sidebar of all 24 tabs
+ the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (24-tab showcase), business dashboard, revenue command
center, sponsorship CRM + content calendar, and mobile.

---

## D. Etsy delivery package

```
Creator_Command_Center.xlsx       ← Excel master (24-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| CCC-EX     | Excel only | $29 |
| CCC-GS     | Google Sheets only | $29 |
| CCC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$44** |
| CCC-PLUS   | Bundle + creator-business playbook | $59 |
| CCC-PRO    | Agency / creator-educator license | $149 |

- **Huge, fast-growing market** — the creator economy. Sell to creators who
  buy tools that make them look/feel professional. Price as premium software.
- Bumps in **January** (new-year systems) and around **platform monetization
  news**. Creator-educators are natural affiliates for the PRO license.
- Lean into two angles: **"run your content like a business"** (monetizing
  creators) and **"stop juggling 12 tools"** (overwhelmed creators).

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun; revenue, profit, health and
  all dashboard KPIs recompute automatically.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Marketing numbers in `build_marketing.py` (`KPIS` list + content functions)
  should be kept in sync with the workbook's sample data.
