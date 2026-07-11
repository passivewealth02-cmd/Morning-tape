# Podcast Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/podcast-command-center/build
python3 build_xlsx.py      # -> ../Podcast_Command_Center.xlsx  (24-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Show Dashboard** fills 12 KPI cards + 4 charts (downloads trend, revenue by
   source, top episodes, expense breakdown).
3. **Episode Calendar** — 4 episodes published in the last 28 days (publishing
   consistency **80%** of the 5 goal); Published/Scheduled/Recording color-code.
4. **Analytics** ranks episodes (top 18.4K downloads) and blends a **Show Health
   Score of ~83%**; **Memberships** total **260 members / $2,630**, which flows
   into the **Finance Center** (Revenue **$9,130**, Net **$7,230**, 79% margin).
5. **Sponsors** show 3 active (Booked/Live) in a lead → paid pipeline. No broken
   cells.

> Note: uses `COUNTIFS`, `COUNTIF`, `AVERAGE`, date math with `TODAY()` — opens
> in Excel 2019/365 or Google Sheets. The published-episode count recalculates
> daily.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Calendar, Analytics, Sponsors, Members & Finance, then
the **Dashboard**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py         # -> ../marketing/01..06.png
python3 build_marketing_detail.py              # -> ../marketing/07..10.png
```

**Six app-screenshots** (sidebar of all 24 tabs + real computed KPI numbers +
full tables/charts): hero, everything-inside (24-tab showcase), analytics,
finance, guest CRM + sponsors, and mobile. (Images 3–5 each show a different
sheet — no repeat of the hero dashboard.)

**Four detailed / benefit-driven images**: 07 feature spotlights, 08 "basic
episode tracker vs Command Center" comparison, 09 up-and-running in 4 steps, 10
what's-included / who-it's-for / guarantee. Ten images total — fills all 10 Etsy
photo slots.

---

## D. Etsy delivery package

```
Podcast_Command_Center.xlsx       ← Excel master (24-tab system + Welcome)
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
| PCC-PLUS   | Bundle + launch & sponsorship playbook | $39 |
| PCC-PRO    | Network / producer license | $99 |

- **Growing creator niche** — podcasting keeps expanding and sponsors pay well.
  Bumps in **January** (new-show launches) and **Q4** (ad season).
- Two angles: **"turn listens into income"** (monetization) and **"one file for
  your whole show"** (production). Network and producer upsells extend reach.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun; subscribers, downloads, revenue,
  net profit, members, publishing consistency and the Show Health Score
  recompute automatically. Calendar dates drive the published count.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Keep `build_marketing.py` numbers in sync with the workbook's sample data.
