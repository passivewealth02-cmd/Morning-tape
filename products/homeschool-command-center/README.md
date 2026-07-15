# Homeschool Command Center™ — The Complete Homeschool Planning & Records System

> Not a lesson-plan template — a **complete homeschool operating system**. One
> premium **Google Sheets + printable PDF** command center for curriculum,
> lessons, attendance & hours, grades, reading, field trips, portfolio and the
> records your state may ask for. Any method, any grade, 1 to many kids.

| | |
| - | - |
| **Product** | Homeschool Command Center™ |
| **Target** | Homeschooling families (any method) · large & multi-age families · new & veteran homeschoolers · Classical / Charlotte Mason / eclectic · co-op & hybrid families · anyone who keeps state records |
| **Angle** | Plan the year, log the days, keep the records — homeschool with confidence. |
| **Formats** | Google Sheets (18-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $14 single · **$22 bundle** (Sheets + PDF) · $29 with high-school transcript add-on · $69 co-op / commercial license |

---

## Contents

```
products/homeschool-command-center/
├── README.md
├── Homeschool_Command_Center.xlsx   ← Google Sheets / Excel master (18 tabs + Settings)
├── Homeschool_Printables.pdf        ← 12-page print-ready pack (US Letter)
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

## The 18-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 10 | Reading Log |
| 2 | Homeschool Dashboard | 11 | Field Trips & Activities |
| 3 | Student Profiles | 12 | Co-op & Extracurriculars |
| 4 | Year Planner (2026–2027) | 13 | Portfolio |
| 5 | Curriculum & Resources | 14 | Curriculum Budget |
| 6 | Subject / Course Plan | 15 | Goals & Milestones |
| 7 | Lesson Planner | 16 | Report Card / Transcript |
| 8 | Attendance & Hours | 17 | Records & Compliance |
| 9 | Assignments & Grades | 18 | Book List & Wish List |
| | | | *(+ Settings)* |

## The 12 printable PDF pages

Student-at-a-Glance · Weekly Lesson Plan · Daily Homeschool Schedule ·
Attendance & Hours Log · Subject / Course Plan · Reading Log · Field-Trip Log ·
Assignment & Grade Sheet · Book List & Wish List · Homeschool Goals ·
Records & Compliance Checklist · Report Card / Transcript

---

## Signature automation (12 dashboard KPIs)

| Metric | How it's computed |
| ------ | ----------------- |
| Students | `=Students` |
| School Days | `=DaysDone` (of `DaysRequired`) |
| Hours Logged | `=HoursLogged` (of `HoursGoal`) |
| Days Left | `=DaysRequired-DaysDone` |
| Subjects | `=Subjects` |
| Lessons Done | `=COUNTIF(LessonStatus,"Done")/COUNTA(LessonUnit)` |
| Graded | `=COUNTIF(AssignGraded,"Graded")/COUNTA(AssignName)` |
| Books Read | `=BooksRead` |
| Curriculum Spent | `=CurricSpent` (SUMIF of acquired resources) |
| Field Trips | `=COUNTA(TripName)` |
| Goals | `=AVERAGE(GoalProgress)` |
| On-Track Score | `=AVERAGE(HealthRange)` |

Logging attendance & hours, checking off lessons, grading an assignment, or
marking a resource acquired all update the dashboard live; the **On-Track Score**
blends school days, instructional hours, lessons, grades, goals and
records-&-compliance into one 0–100% "are we on pace" number.

**Verified sample family** (the Bennetts — 4 students, grades K–8):
Students **4** · School days **120** of 180 · Hours **640** of 900 · Days left
**60** · Subjects **8** · Lessons **78%** done · Graded **85%** · Books **46** ·
Curriculum **$1,240** of $1,500 · Field trips **7** · Goals **72%** ·
**On-Track 75%**.

---

## Premium homeschool-software design

- Two-row **gold-divider headers** on every tab; a true dashboard (12 KPIs +
  on-track bars, a budget donut, a "coming up" list & an On-Track gauge)
- Status color-coding (Done / In Progress / Not Started; Have / Ordered / Need;
  Graded / Turned In / Not Yet), data-bars and conditional flags throughout
- Bends to any method — Classical, Charlotte Mason, unit study, eclectic — and
  scales from 1 student to a full house
- **Print-ready PDF pack** on white with a forest-green header band & gold rules
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **Not legal advice.** Homeschool requirements vary by state/country — always
> confirm your own jurisdiction's rules. This is a planning & record-keeping tool.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Homeschool_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Homeschool_Printables.pdf (+ page PNGs)
python3 build_marketing.py                 # -> ../marketing/01..06.png
python3 build_marketing_detail.py          # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
