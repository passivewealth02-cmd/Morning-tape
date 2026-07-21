# College Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/college-command-center/build
python3 build_xlsx.py      # -> ../College_Command_Center.xlsx  (15 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "confirm details with each
   college" note.
2. **College List** flags each school Reach/Match/Safety; checking off essays,
   recs, form & submission fills the progress bar (`COUNTIF/4`). Avg progress
   **69%**; **4 of 8** submitted (3 reach / 3 match / 2 safety).
3. **Dashboard** fills 12 KPI cards + an "Are We On Track?" table + a
   Reach/Match/Safety donut. **Ready Score 63%**.
4. **Scholarships** sums awarded (**$12,000**); **Net Price** highlights the
   lowest true cost (**$2,000**/yr) with a color scale.
5. Marking a school submitted, an essay Final, or a task done updates the
   dashboard live. **Next Deadline** uses `MINIFS` over unsubmitted schools.
6. No broken cells; custom tables (College List, Net Price, Deadlines,
   Applicant Profile) start in column B.

> Note: uses `COUNTIF`, `AVERAGE`, `SUMIF`, `MIN`, `MINIFS`, `IFERROR` — opens in
> Google Sheets or Excel 2019/365 (MINIFS needs 2019+ / Sheets).

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../College_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: applicant profile, college list, application tracker, essay &
supplement tracker, recommendation tracker, test scores, activities résumé,
scholarship log, net-price comparison, visits & interviews, decisions & compare,
and a master to-do & deadlines.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (15 tabs), the
college list (reach/match/safety + progress), net-price comparison, essays +
scholarships, and the **12-page printables showcase**. Images 3–5 each show a
different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic checklist vs Command
Center", 09 run-the-season in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers (8 colleges
· 4 submitted · 69% · 50% essays · $12K aid · $2K net · 63% score) are verified
against the workbook.

---

## D. Etsy delivery package

```
College_Command_Center.xlsx        ← Google Sheets / Excel master (15 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
College_Printables.pdf              ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| CAC-GS   | Google Sheets only | $16 |
| CAC-PDF  | Printable PDF only | $16 |
| CAC-BUNDLE | Sheets + PDF + Quick-Start | **$24** |
| CAC-PLUS | Bundle + high-school transcript add-on | $32 |
| CAC-PRO  | Counselor / commercial-use license | $79 |

- **Sharp August–January peak** (application season) with an early spring
  scholarship tail. The net-price comparison and reach/match/safety balance are
  more differentiated than a plain deadline checklist.
- Cross-sell with the **High-School Transcript Command Center** (same student,
  Ella) — transcript → college app is a natural bundle. Two angles: **"apply
  smart"** and **"compare offers before you commit."**

---

## F. Maintenance

- Edit the `STUDENT`, `COLLEGES`, `ESSAYS`, `RECS`, `SCHOLARSHIPS`, `NETPRICE`,
  `DECISIONS`, `TASKS` constants in `build_xlsx.py`; every KPI + the Ready Score
  recompute. Add a college → add a `COLLEGES` row (progress renders).
- Printable pages live in `build_pdf.py` (one function per page). Keep
  `build_marketing.py` numbers in sync with the workbook.
