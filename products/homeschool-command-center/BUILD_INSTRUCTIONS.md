# Homeschool Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/homeschool-command-center/build
python3 build_xlsx.py      # -> ../Homeschool_Command_Center.xlsx  (18 tabs + Settings)
```

### Verifying
1. **Start Here** shows the intro + a 6-step quick guide + a "confirm your state"
   note.
2. **Homeschool Dashboard** fills 12 KPI cards + on-track bars + a Budget-by-
   Category donut + a "coming up" list + an On-Track gauge.
3. **Student Profiles** has a block for each of the 4 Bennett students (scales up).
4. Logging Attendance/Hours, checking off Lessons, grading Assignments and
   marking a resource "Have" updates the dashboard + On-Track Score (**75%**).
   Sample: Days **120**/180, Hours **640**/900, Lessons **78%**, Graded **85%**,
   Curriculum **$1,240**, Records **80%**.
5. No broken cells; custom tables (Budget, Goals, Report Card) start in column B.

> Note: uses `COUNTIF/COUNTIFS`, `SUMIF`, `AVERAGE` — opens in Google Sheets or
> Excel 2019/365. Change required days/hours in Settings and the score adjusts.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Homeschool_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band & gold rules. 300 DPI,
US Letter. Twelve pages: student-at-a-glance, weekly lesson plan, daily schedule,
attendance & hours, subject plan, reading log, field-trip log, grade sheet, book
list, goals, records checklist, and report card / transcript.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (18 tabs), a-
profile-for-every-student, curriculum & resources, lesson planner + attendance,
and the **12-page printables showcase** (real PDF thumbnails). Images 3–5 each
show a different tab.

**Four detailed images**: 07 feature spotlights (incl. printable thumbnails),
08 "basic lesson plan vs Command Center", 09 up-and-running in 4 steps, 10
what's-included / who-it's-for / guarantee. Ten images — fills all 10 Etsy slots.

---

## D. Etsy delivery package

```
Homeschool_Command_Center.xlsx    ← Google Sheets / Excel master (18 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt    ← "Make a Copy" link
Homeschool_Printables.pdf          ← 12-page print-ready pack
START_HERE.pdf                     ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| HSC-GS   | Google Sheets only | $14 |
| HSC-PDF  | Printable PDF only | $14 |
| HSC-BUNDLE | Sheets + PDF + Quick-Start | **$22** |
| HSC-PLUS | Bundle + high-school transcript add-on | $29 |
| HSC-PRO  | Co-op / commercial-use license | $69 |

- **Big, evergreen niche with a July–September peak** (planning season) and a
  January bump. The records/compliance + On-Track angle is more differentiated
  than a basic lesson-plan printable, and it fits every method.
- Two angles: **"plan the year, keep the records"** and **"Google Sheets +
  printables in one."** Bundle is the hero SKU; add a state-specific note in the
  listing FAQ.

---

## F. Maintenance

- Edit the `STUDENTS`, `WEEKS`, `CURRICULUM`, `LESSONS`, `ASSIGNMENTS`, `RECORDS`,
  `GOALS`, `FIELDTRIPS` constants in `build_xlsx.py`; every KPI + the On-Track
  Score recompute. Add a student → add a `STUDENTS` row (a profile block renders).
- Printable pages live in `build_pdf.py` (one function per page) — restyle once,
  re-run. Keep `build_marketing.py` numbers in sync with the workbook.
