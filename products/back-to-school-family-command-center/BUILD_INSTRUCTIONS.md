# Back-to-School Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/back-to-school-family-command-center/build
python3 build_xlsx.py      # -> ../Back_to_School_Command_Center.xlsx  (17 tabs + Settings)
```

### Verifying
1. **Start Here** shows the mom-of-six intro + a 6-step quick guide.
2. **Family Dashboard** fills 12 KPI cards + readiness bars + a Budget-by-Category
   donut + a "due this week" list + a Readiness gauge.
3. **Child Profiles** has a block for each of the 6 Rivera kids (scales to 8).
4. Checking a supply / paying a fee / marking a form updates the **Budget** and
   the **Readiness Score** (blended **82%**). Sample: Supplies **85%**, Fees
   **88%**, Forms **80%**, Uniforms **75%**, Budget **$1,850** of **$2,500**.
5. Events within 30 days count = **9**; open to-dos = **10**. No broken cells.

> Note: uses `COUNTIF/COUNTIFS`, `SUMIF`, `AVERAGE`, `TODAY()` date math — opens
> in Google Sheets or Excel 2019/365. The countdown & event counts recalc daily.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Back_to_School_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band & gold rules. 300 DPI,
US Letter (2550×3300). Twelve pages: child info, school contacts, first-day prep,
backpack checklist, weekly schedule, supply checklist, clothing sizes, lunchbox
planner, field-trip log, meeting notes, year goals, and first/last-day memory pages.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (17 tabs),
a-profile-for-every-child, budget vs actual, supplies + events, and the
**12-page printables showcase** (real PDF thumbnails). Images 3–5 each show a
different tab — no repeat of the hero.

**Four detailed images**: 07 feature spotlights (incl. printable thumbnails),
08 "basic checklist vs Command Center", 09 up-and-running in 4 steps, 10
what's-included / who-it's-for / guarantee. Ten images — fills all 10 Etsy slots.

---

## D. Etsy delivery package

```
Back_to_School_Command_Center.xlsx    ← Google Sheets / Excel master (17 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt        ← "Make a Copy" link
Back_to_School_Printables.pdf          ← 12-page print-ready pack
START_HERE.pdf                         ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| BTS-GS   | Google Sheets only | $12 |
| BTS-PDF  | Printable PDF only | $12 |
| BTS-BUNDLE | Sheets + PDF + Quick-Start | **$18** |
| BTS-PLUS | Bundle + lunch / meal-plan add-on | $24 |
| BTS-PRO  | Shop / commercial-use license | $59 |

- **Seasonal spike** — heavy demand **July–September** (and a smaller January
  bump). Large-family + multi-school angle is far more differentiated than a
  basic checklist, and the **mom-of-six** story is impossible to copy.
- Two angles: **"one calm system for the whole house"** and **"Google Sheets +
  printables in one"**. Bundle is the hero SKU.

---

## F. Maintenance

- Edit the `CHILDREN`, `SUPPLIES`, `CLOTHING`, `FEES`, `DOCS`, `BUDGET` constants
  and `FIRST_DAY` in `build_xlsx.py`; every KPI and the Readiness Score recompute.
- To add kid #7/#8: add rows to `CHILDREN` and a profile block renders automatically.
- Printable pages live in `build_pdf.py` (one function per page) — restyle once,
  re-run. Keep `build_marketing.py` numbers in sync with the workbook.
