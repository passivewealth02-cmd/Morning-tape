# TikTok Command Center™ — Build Instructions

---

## A. Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/tiktok-command-center/build
python3 build_xlsx.py      # -> ../TikTok_Command_Center.xlsx  (24-tab system + Welcome)
```

### Verifying
1. **Welcome** shows the intro + a Start-Here quick guide.
2. **Creator Dashboard** fills 12 KPI cards + 4 charts (follower growth, revenue
   by source, top videos, expense breakdown).
3. **Content Calendar** — 22 posts in the last 28 days (posting consistency
   **92%** of the 24 goal); Viral/Posted/Scheduled color-code.
4. **Analytics** ranks 6 videos (top 3.2M views) and blends a **Creator Health
   Score of ~88%**; **TikTok Shop** totals **$14,233 GMV / $1,883** commission,
   which flows into the **Finance Center** (Revenue **$8,583**, Net **$6,633**,
   77% margin).
5. **Brand Deals** shows 4 active (lead → paid). No broken cells.

> Note: uses `COUNTIFS`, `COUNTIF`, `AVERAGE`, date math with `TODAY()` — opens
> in Excel 2019/365 or Google Sheets. The posting count recalculates daily.

---

## B. Google Sheets build

Follow `GOOGLE_SHEETS.md`: build **Settings** first, define the cross-sheet
named ranges, then the Calendar, Analytics, Shop, Brand Deals & Finance, then
the **Dashboard**.

---

## C. Marketing images

```bash
cd build && python3 build_marketing.py         # -> ../marketing/01..06.png
python3 build_marketing_detail.py              # -> ../marketing/07..10.png
```

**Six app-screenshots** (sidebar of all 24 tabs + real computed KPI numbers +
full tables/charts): hero, everything-inside (24-tab showcase), analytics,
finance, trends + brand deals, and mobile. (Images 3–5 each show a different
sheet — no repeat of the hero dashboard.)

**Four detailed / benefit-driven images** (reuse the branded helpers via
`from build_marketing import *`): 07 feature spotlights, 08 "basic calendar vs
Command Center" comparison, 09 up-and-running in 4 steps, 10 what's-included /
who-it's-for / guarantee. Ten images total — fills all 10 Etsy photo slots.

---

## D. Etsy delivery package

```
TikTok_Command_Center.xlsx        ← Excel master (24-tab system + Welcome)
GOOGLE_SHEETS_TEMPLATE_LINK.txt   ← "Make a Copy" link
START_HERE.pdf                    ← onboarding quick-start
THANK_YOU.pdf                     ← brand thank-you card
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| TTC-EX     | Excel only | $19 |
| TTC-GS     | Google Sheets only | $19 |
| TTC-BUNDLE | Excel + Google Sheets + Quick-Start PDF | **$29** |
| TTC-PLUS   | Bundle + creator growth & monetization playbook | $39 |
| TTC-PRO    | Agency / creator / MCN license | $99 |

- **Fast-growing creator niche** — TikTok Shop & UGC have exploded. Bumps in
  **January** (new-year goals) and around **Q4 / holiday** shopping.
- Two angles: **"turn views into income"** (monetization) and **"one file for
  your whole account"** (organization). Faceless-page and agency upsells extend
  reach.

---

## F. Maintenance

- Edit sample data in `build_xlsx.py` and rerun; followers, revenue, net profit,
  Shop GMV, posting consistency and the Creator Health Score recompute
  automatically. Calendar dates drive the posting count.
- Brand styles, dropdowns, KPI cards & summary tables are centralized.
- New sheet: add a builder and slot it into the `order` list in `main()`.
- Marketing numbers in `build_marketing.py` (`KPIS` list + content functions)
  should be kept in sync with the workbook's sample data.
