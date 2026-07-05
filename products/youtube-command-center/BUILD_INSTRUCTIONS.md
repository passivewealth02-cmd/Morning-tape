# YouTube Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/youtube-command-center/build
python3 build_xlsx.py      # -> ../YouTube_Command_Center.xlsx  (30-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Executive YouTube Dashboard** fills 12 KPI cards + 4 charts (subscriber
   growth, revenue by source, top videos, expense breakdown).
3. **Content Master Calendar** — 6 published in the last 28 days; status glows
   mint / gold / soft; **Pipeline** shows live progress %.
4. **Analytics Command Center** ranks 6 recent videos and computes a
   **Channel Health Score of ~79%** across 6 dimensions.
5. **Business Finance Center** — 8 income streams total **$9,840/mo**
   ($118,080 run-rate), expenses **$2,410**, Net **$7,430** (75% margin).
   **Sponsorship CRM** shows 4 active deals (lead → paid). No broken cells.

> Note: uses `SUMPRODUCT`, `COUNTIFS`, `AVERAGEIF`, `LEFT/MID/FIND` — opens in
> Excel 2019/365 or Google Sheets.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Calendar, Analytics, Finance & CRM, then the
**Dashboard**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py     # -> ../marketing/01..06.png
```

Six 2000×2000 PNGs, rendered as dense app screenshots (sidebar of all 30 tabs
+ the real computed KPI numbers + fully populated tables/charts): hero,
everything-inside (30-tab showcase), analytics command center, business finance
center, calendar + sponsorship CRM, and mobile. (Images 3–5 each show a
different sheet — no repeat of the hero dashboard.)

---

## D. Etsy delivery package

```
YouTube_Command_Center.xlsx       ← Excel master (30-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| YTC-EX     | Excel only | $19 |
| YTC-GS     | Google Sheets only | $19 |
| YTC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$29** |
| YTC-PLUS   | Bundle + creator growth & monetization playbook | $39 |
| YTC-PRO    | Agency / creator / MCN license | $99 |

- **Evergreen, spend-happy niche** — creators reinvest in tools that make them
  money. Bumps in **January** (new-year goals) and **Q3/Q4** (planning season).
- Lean into two angles: **"run your channel like a business"** (monetization)
  and **"one file for your whole channel"** (organization). Faceless-channel and
  agency upsells (YTC-PRO) extend reach.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun; revenue, net profit, upload
  consistency, sponsorship pipeline and the Channel Health Score recompute
  automatically. Publishing status derives from each row's date.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Marketing numbers in `build_marketing.py` (`KPIS` list + content functions)
  should be kept in sync with the workbook's sample data.
