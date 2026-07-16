# High-School Transcript Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/highschool-transcript-command-center/build
python3 build_xlsx.py      # -> ../HS_Transcript_Command_Center.xlsx  (17 tabs)
```

### Verifying
1. **Start Here** shows the intro + quick guide + a "confirm your state & target
   colleges" note.
2. **Dashboard** fills 12 KPI cards + a College-Ready gauge + GPA-by-year and
   credits-by-subject charts + an application-file list.
3. **Course Records** is the engine: type a *letter grade* and grade points,
   quality points and weighted points all appear via `VLOOKUP(GradeScale)` and
   `VLOOKUP(LevelScale)`. Add/remove a course row and every total re-computes.
4. **GPA Calculator** shows Unweighted **3.81**, Weighted **4.06**, Honors/AP
   **9**, Courses **27**; **Credit Tracker** shows **24.0** of 24 credits.
5. **Official Transcript** renders school header, per-year blocks, a summary box
   (UW 3.81 / W 4.06 / 24.0 credits) and a signature line — matching PDF page 1.
6. No broken cells; custom tables start in column B (never the width-2 margin A).

> Note: uses `VLOOKUP`, `COUNTIF/COUNTIFS`, `SUMIF`, `AVERAGE`, `IFERROR` — opens
> in Google Sheets or Excel 2019/365. Change `CreditsReq`, `GPATarget`,
> `ServiceGoal` or `RigorGoal` in Settings and the College-Ready score adjusts.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../HS_Transcript_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Page 1 is the print-ready **Official Transcript** (custom two-column year-block
layout, summary box, signature line). Pages 2–12: 4-year plan, course-records /
grade sheet, GPA worksheet, credit & graduation checklist, test-score log,
awards & honors, service-hours log, course descriptions, reading list, college
application tracker, and activities résumé. 300 DPI, US Letter.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (16 tabs shown),
official transcript, course records with auto-GPA, test scores + activities
(the application file), and the **12-page printables showcase** (real PDF
thumbnails). Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights (College-Ready gauge, GPA table,
application file, transcript thumbnails), 08 "blank template vs Command Center",
09 up-and-running in 4 steps, 10 what's-included / who-it's-for / guarantee. Ten
images — fills all 10 Etsy slots. All headline numbers (3.81 / 4.06 / 24.0 /
1380 / 30 / 120 / 7 / 5 / 98%) are verified against the workbook.

---

## D. Etsy delivery package

```
HS_Transcript_Command_Center.xlsx    ← Google Sheets / Excel master (17 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt       ← "Make a Copy" link
HS_Transcript_Printables.pdf          ← 12-page print-ready pack
START_HERE.pdf                        ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| HST-GS   | Google Sheets only | $16 |
| HST-PDF  | Printable PDF only | $16 |
| HST-BUNDLE | Sheets + PDF + Quick-Start | **$24** |
| HST-PLUS | Bundle + 4-year planning add-on | $32 |
| HST-PRO  | Counselor / commercial-use license | $79 |

- **Evergreen niche with an August–November peak** (application season) and a
  spring bump (course planning). The auto-GPA + official-transcript angle is far
  more differentiated than a blank transcript printable.
- Two angles: **"GPA that calculates itself"** and **"a real, print-ready
  transcript for applications."** Bundle is the hero SKU; the counselor license
  and 4-year-plan add-on lift average order value.

---

## F. Maintenance

- Edit `COURSES`, `TESTS`, `ACTIVITIES`, `AWARDS`, `SERVICE`, `GRADREQ`,
  `DESCRIPTIONS`, `READING`, `PLAN` and the `GRADE_SCALE` / `LEVELS` constants in
  `build_xlsx.py`; every KPI, the GPA and the College-Ready score recompute. Add
  a course → add a `COURSES` row (a transcript line renders in both the workbook
  and the PDF).
- Printable pages live in `build_pdf.py` (one function per page); page 1 imports
  `COURSES`/`GRADE_SCALE` from `build_xlsx.py` so the transcript can never drift
  from the workbook. Keep `build_marketing.py` numbers in sync with the workbook.
