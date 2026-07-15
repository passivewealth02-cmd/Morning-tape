# Homeschool Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Student Profiles, Year Planner, Curriculum, Subject Plan,
Lesson Planner, Attendance, Assignments, Reading Log, Field Trips, Co-op,
Portfolio, Budget, Goals, Report Card, Records, Book List, Settings**.

> Build **Settings** first (family + required days/hours + dropdown lists), then
> Curriculum, Lesson Planner, Attendance, Assignments, Goals & Records, then the
> Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `FamilyName`, `SchoolYear`, `HomeState`, `Students` (4),
`Subjects` (8), `DaysRequired` (180), `HoursGoal` (900),
`CurriculumBudget` (1500), `BooksRead` (46).

Lists: `GradeList, SubjectList, StatusList, AcquiredList, GradedList,
ApproachList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CurricCost` | `Curriculum!D5:D44` | `CurricSpent` | `Curriculum!D45` |
| `CurricStatus` | `Curriculum!E5:E44` | `DaysDone` | `Attendance!B45` |
| `LessonUnit` | `'Lesson Planner'!C5:C64` | `HoursLogged` | `Attendance!C45` |
| `LessonStatus` | `'Lesson Planner'!D5:D64` | `AttDays` | `Attendance!B5:B44` |
| `AssignName` | `Assignments!D5:D64` | `AttHours` | `Attendance!C5:C44` |
| `AssignGraded` | `Assignments!F5:F64` | `GoalProgress` | `Goals!D5:D10` |
| `RecName` | `Records!A5:A28` | `RecDone` | `Records!B5:B28` |
| `TripName` | `'Field Trips'!B5:B34` | `HealthRange` | `Dashboard!C13:C18` |
| `BudgetSpent` | `Budget!D14` | `BudgetPlanTotal` | `Budget!C14` |

---

## 3. Dashboard — the 12 KPIs

```sheets
Students        =Students
School Days     =DaysDone
Hours Logged    =HoursLogged
Days Left       =DaysRequired-DaysDone
Subjects        =Subjects
Lessons Done    =IFERROR(COUNTIF(LessonStatus,"Done")/COUNTA(LessonUnit),0)
Graded          =IFERROR(COUNTIF(AssignGraded,"Graded")/COUNTA(AssignName),0)
Books Read      =BooksRead
Curriculum      =CurricSpent          (=SUMIF(CurricStatus,"Have",CurricCost))
Field Trips     =COUNTA(TripName)
Goals           =IFERROR(AVERAGE(GoalProgress),0)
On-Track        =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Budget by Category (donut) from the Budget tab (actual column). Turn off
auto data labels.

---

## 4. On-Track Score (6 dimensions)

```sheets
School days (of 180)  =IFERROR(DaysDone/DaysRequired,0)
Instructional hours   =IFERROR(HoursLogged/HoursGoal,0)
Lessons done          =IFERROR(COUNTIF(LessonStatus,"Done")/COUNTA(LessonUnit),0)
Assignments graded    =IFERROR(COUNTIF(AssignGraded,"Graded")/COUNTA(AssignName),0)
Goals & milestones    =IFERROR(AVERAGE(GoalProgress),0)
Records & compliance  =IFERROR(COUNTIF(RecDone,"Yes")/COUNTA(RecName),0)
On-Track Score        =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `COUNTIF`/`COUNTIFS`, `SUMIF`, `AVERAGE`, `IFERROR`, plus `QUERY`
("what's not done", "grades by student") and `FILTER`/`SORT`. A profile per
student, curriculum list and lesson planner are plain tables — duplicate a
student block for kid #5+.

---

## 5. Printables

The 12-page PDF is print-ready as-is (US Letter). Google Sheets users can also
print any tab: File ▸ Print ▸ fit to width.

> Requirements vary by state/country — confirm your own rules. This is a
> record-keeping tool, not legal advice.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
