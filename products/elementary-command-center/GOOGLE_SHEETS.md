# Elementary Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Student Profiles, Weekly Plan, Subject Mastery, Math
Facts, Sight Words, Reading Log, Assignments, Report Card, Attendance, Curriculum,
Field Trips, Habits, Awards, Book List, Settings**.

> Build **Settings** first (family + students + goals + lists), then Subject
> Mastery, Math Facts, Sight Words, Assignments and Habits, then the Dashboard.
> Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `FamilyName`, `SchoolYear`, `Students` (2), `Subjects` (8),
`DaysRequired` (180), `ReadingGoal` (60), `BooksRead` (52),
`SightWordsKnown` (210), `SightWordGoal` (220).

Lists: `SubjectList, MasteryList, GradedList, StatusList, AcquiredList,
GradeList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `MasterySubject` | `'Subject Mastery'!A5:A44` | `MasteryName` | `'Subject Mastery'!B5:B44` |
| `MasteryStatus` | `'Subject Mastery'!D5:D44` | `MathPct` | `'Math Facts'!D5:D10` |
| `MathAvg` | `'Math Facts'!D11` (average) | `SWKnown` | `'Sight Words'!D5:D10` |
| `SWTotal` | `'Sight Words'!E5:E10` | `AssignName` | `Assignments!D5:D64` |
| `AssignGraded` | `Assignments!F5:F64` | `TripName` | `'Field Trips'!B5:B34` |
| `HabitDone` | `Habits!B5:B24` | `DaysDone` | `Attendance!B45` |
| `SubjSumVals` | `'Subject Mastery'!C47:C54` | `HealthRange` | `Dashboard!C13:C18` |

---

## 3. Dashboard — the 12 KPIs

```sheets
Students         =Students
School Days      =DaysDone
Subjects         =Subjects
Skills Mastered  =COUNTIF(MasteryStatus,"Mastered")
Math Facts       =MathAvg
Sight Words      =SightWordsKnown
Books Read       =BooksRead
Graded           =IFERROR(COUNTIF(AssignGraded,"Graded")/COUNTA(AssignName),0)
Field Trips      =COUNTA(TripName)
Habits           =IFERROR(COUNTIF(HabitDone,"Yes")/COUNTA(HabitDone),0)
Mastery          =IFERROR(COUNTIF(MasteryStatus,"Mastered")/COUNTA(MasteryName),0)
On-Track         =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Skills Mastered by Subject (bar) from the summary block on the Subject
Mastery tab (`=COUNTIFS(MasterySubject,<subject>,MasteryStatus,"Mastered")`).

Math Facts average: `MathAvg =IFERROR(AVERAGE(MathPct),0)` at the total row.

---

## 4. On-Track Score (6 dimensions)

```sheets
Skills mastered       =IFERROR(COUNTIF(MasteryStatus,"Mastered")/COUNTA(MasteryName),0)
Math facts fluency    =MathAvg
Sight words           =IFERROR(SightWordsKnown/SightWordGoal,0)
Assignments graded    =IFERROR(COUNTIF(AssignGraded,"Graded")/COUNTA(AssignName),0)
Reading goal          =IFERROR(BooksRead/ReadingGoal,0)
Habits & character    =IFERROR(COUNTIF(HabitDone,"Yes")/COUNTA(HabitDone),0)
On-Track Score        =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `COUNTIF`/`COUNTIFS`, `AVERAGE`, `IFERROR`, plus `QUERY`
("skills still practicing", "grades by subject") and `FILTER`/`SORT`. Mastery,
math facts, sight words & assignments are plain tables — add a row and the
dashboard follows.

---

## 5. Printables

The 12-page PDF is print-ready as-is (US Letter). Google Sheets users can also
print any tab: File ▸ Print ▸ fit to width.

> Every child grows at their own pace — an encouraging guide, not a test.
> Confirm your state's requirements. This is a record-keeping tool.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
