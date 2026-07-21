# Elementary Homeschool Command Center™ — The Complete K–5 Learning System

> Not a worksheet pack — a **complete K–5 planning, mastery & records system**.
> One premium **Google Sheets + printable PDF** command center with a weekly
> lesson plan, a subject-mastery tracker, math-facts fluency, sight words &
> spelling, reading log & levels, assignments & grades, a print-ready report
> card, attendance, curriculum, field trips, habits & character, awards and a
> book list. Built for the elementary years (grades K–5), any method.

| | |
| - | - |
| **Product** | Elementary Homeschool Command Center™ |
| **Target** | K–5 homeschooling families · multi-age & large families · new & veteran homeschoolers · Classical / Charlotte Mason / eclectic · after-schoolers & summer-slide parents · anyone who keeps records |
| **Angle** | See what's mastered, teach what's next — the elementary years, organized. |
| **Formats** | Google Sheets (17-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $14 single · **$22 bundle** (Sheets + PDF) · $29 with the K–12 add-on · $69 co-op / commercial license |

---

## Contents

```
products/elementary-command-center/
├── README.md
├── Elementary_Command_Center.xlsx   ← Google Sheets / Excel master (17 tabs)
├── Elementary_Printables.pdf        ← 12-page print-ready pack (US Letter)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    ├── build_pdf.py
    ├── build_marketing.py
    └── build_marketing_detail.py
```

---

## The 17-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 10 | Report Card |
| 2 | Dashboard | 11 | Attendance |
| 3 | Student Profiles | 12 | Curriculum |
| 4 | Weekly Plan | 13 | Field Trips |
| 5 | Subject Mastery | 14 | Habits |
| 6 | Math Facts | 15 | Awards |
| 7 | Sight Words | 16 | Book List |
| 8 | Reading Log | 17 | Settings |
| 9 | Assignments | | |

## The 12 printable PDF pages

Student-at-a-Glance · Weekly Lesson Plan · Subject Mastery Checklist · Math Facts
Chart · Sight Words & Spelling · Reading Log & Levels · Assignment & Grade Sheet ·
Report Card · Attendance & Days · Field-Trip Log · Habits & Character Chart ·
Awards & Milestones.

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Students | `=Students` |
| School Days | `=DaysDone` (of `DaysRequired`) |
| Subjects | `=Subjects` |
| Skills Mastered | `=COUNTIF(MasteryStatus,"Mastered")` |
| Math Facts | `=MathAvg` (average % fluent) |
| Sight Words | `=SightWordsKnown` (of goal) |
| Books Read | `=BooksRead` (of `ReadingGoal`) |
| Graded | `=COUNTIF(AssignGraded,"Graded")/COUNTA(AssignName)` |
| Field Trips | `=COUNTA(TripName)` |
| Habits | `=COUNTIF(HabitDone,"Yes")/COUNTA(HabitDone)` |
| Mastery | `=COUNTIF(MasteryStatus,"Mastered")/COUNTA(MasteryName)` |
| On-Track | `=AVERAGE(HealthRange)` |

Marking a skill Mastered, updating a math-facts %, logging a book or grading an
assignment all update the dashboard live; the **On-Track Score** blends subject
mastery, math facts, sight words, graded work, reading and habits into one
0–100% number, and a **Skills-Mastered-by-Subject** bar chart shows the balance.

**Verified sample family** (the Bennetts — Lucy, grade 3, and Jack, grade 1):
Students **2** · Days **120** of 180 · Subjects **8** · Skills mastered **16** of
28 · Math facts **67%** fluent · Sight words **210** of 220 · Books **52** of 60 ·
Graded **79%** · Field trips **6** · Habits **86%** · Mastery **57%** ·
**On-Track 78%**.

---

## Premium elementary-software design

- A true dashboard: 12 KPIs, an "Are We On Track?" bars panel, a skills-mastered
  donut, a "coming up" list, an On-Track gauge & a Skills-by-Subject bar chart
- A **subject-mastery** tracker across 8 subjects (Mastered / Practicing / Not Yet)
- **Math-facts fluency** by operation & child, and **sight-word** progress bars
- Reading log with guided-reading levels, a print-ready report card & habit streaks
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **An encouraging guide, not a standardized test.** Every child grows at their
> own pace; requirements vary by state — confirm your own rules. This is a
> planning & record-keeping tool.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Elementary_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Elementary_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
