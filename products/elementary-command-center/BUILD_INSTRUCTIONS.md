# Elementary Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/elementary-command-center/build
python3 build_xlsx.py      # -> ../Elementary_Command_Center.xlsx  (17 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "grows at their own pace"
   note.
2. **Dashboard** fills 12 KPI cards + an "Are We On Track?" bars panel + a
   skills-mastered donut + a "coming up" list + an On-Track gauge + a
   Skills-by-Subject bar chart.
3. **Subject Mastery** covers 8 subjects (Mastered / Practicing / Not Yet); the
   summary block feeds the dashboard bar chart. **16 of 28** mastered.
4. **Math Facts** shows % fluent per operation & child with a live **67%**
   average; **Sight Words** shows known/total progress bars (**210 of 220**).
5. Marking a skill Mastered, editing a math-facts %, logging a book or grading an
   assignment updates the dashboard + On-Track Score (**78%**). Sample: Days
   **120**/180, Graded **79%**, Habits **86%**.
6. No broken cells; custom tables (Weekly Plan, Math Facts, Sight Words, Report
   Card) start in column B.

> Note: uses `COUNTIF/COUNTIFS`, `AVERAGE`, `IFERROR` — opens in Google Sheets or
> Excel 2019/365. Change goals in Settings and the score adjusts.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Elementary_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band & gold rules. 300 DPI,
US Letter. Twelve pages: student-at-a-glance, weekly lesson plan, subject-mastery
checklist, math-facts chart (color-a-box), sight words & spelling, reading log &
levels, assignment & grade sheet, report card, attendance, field-trip log, a
habits & character grid, and awards & milestones.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard), everything-inside (17 tabs), the
subject-mastery tracker, math-facts fluency (climbing bars), reading log + report
card, and the **12-page printables showcase**. Images 3–5 each show a different
tab.

**Four detailed images**: 07 feature spotlights, 08 "basic tracker vs Command
Center", 09 up-and-running in 4 steps, 10 what's-included / who-it's-for /
guarantee. Ten images — fills all 10 Etsy slots. All headline numbers (16/28 ·
67% · 210/220 · 52/60 · 79% · 86% · 78%) are verified against the workbook.

---

## D. Etsy delivery package

```
Elementary_Command_Center.xlsx     ← Google Sheets / Excel master (17 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Elementary_Printables.pdf           ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing

| SKU | Format | Price |
| --- | ------ | ----- |
| ECC-GS   | Google Sheets only | $14 |
| ECC-PDF  | Printable PDF only | $14 |
| ECC-BUNDLE | Sheets + PDF + Quick-Start | **$22** |
| ECC-PLUS | Bundle + K–12 all-grades add-on | $29 |
| ECC-PRO  | Co-op / commercial-use license | $69 |

- **Evergreen niche, July–September planning peak** (+ a January bump). The
  mastery + math-facts + reading-level angle is far more differentiated than a
  plain grade book, and it fits every method.
- Sells as part of an education ladder (tot-school → preschool → elementary →
  high-school transcript) — cross-list & bundle for higher AOV.

---

## F. Maintenance

- Edit the `STUDENTS`, `MASTERY_ROWS`, `MATHFACTS`, `SIGHTWORDS`, `READING`,
  `ASSIGNMENTS`, `HABITS`, `FIELDTRIPS`, `CURRICULUM`, `AWARDS` constants in
  `build_xlsx.py`; every KPI + the On-Track Score recompute. Add a student → add
  a `STUDENTS` row (a profile block renders).
- Printable pages live in `build_pdf.py` (one function per page). Keep
  `build_marketing.py` numbers in sync with the workbook.
