# High-School Transcript Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Student Profile, Official Transcript, Course Records,
GPA Calculator, Credit Tracker, 4-Year Plan, Test Scores, Activities, Awards,
Service, Course Descriptions, Reading List, Grad Requirements, Portfolio,
Settings**.

> Build **Settings** first (student + school + grade/level scales + dropdowns),
> then Course Records (the engine), then GPA Calculator, Credit Tracker, Grad
> Requirements, then the Dashboard and Official Transcript. Add the named ranges
> below (Data ▸ Named ranges).

---

## 1. Settings — controls, scales & lists

Controls: `StudentName` (Ella Bennett), `SchoolName` (Bennett Family Academy),
`ClassOf` (2027), `GradDate`, `CreditsReq` (24), `GPATarget` (3.5),
`ServiceGoal` (100), `RigorGoal` (8, i.e. Honors+AP target).

**GradeScale** (`Settings!E7:F17`) — letter → grade points:

```
A+ 4.0   A 4.0   A- 3.7   B+ 3.3   B 3.0   B- 2.7
C+ 2.3   C 2.0   C- 1.7   D 1.0    F 0.0
```

**LevelScale** (`Settings!E21:F24`) — course level → weight bump:

```
Regular 0.0   Honors 0.5   AP 1.0   Dual Credit 1.0
```

Lists: `AreaList` (subject areas), `LevelList` (Regular/Honors/AP/Dual Credit),
`YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CrsYear` | `'Course Records'!A5:A49` | `CrsArea` | `'Course Records'!B5:B49` |
| `CrsLevel` | `'Course Records'!D5:D49` | `CrsCredits` | `'Course Records'!E5:E49` |
| `CrsQual` | `'Course Records'!H5:H49` | `CrsWtd` | `'Course Records'!I5:I49` |
| `CreditsEarned` | `'Course Records'!E50` | `QualTotal` | `'Course Records'!H50` |
| `WtdTotal` | `'Course Records'!I50` | `CourseCount` | `'GPA Calculator'!C17` |
| `GPAUW` | `'GPA Calculator'!C14` | `GPAW` | `'GPA Calculator'!C15` |
| `HonorsAP` | `'GPA Calculator'!C16` | `ServiceHours` | `Service!C35` |
| `ReqName` | `'Grad Requirements'!A5:A24` | `ReqDone` | `'Grad Requirements'!B5:B24` |
| `TestName` | `'Test Scores'!A5:A28` | `ActName` | `Activities!A5:A34` |
| `AwardName` | `Awards!A5:A28` | `HealthRange` | `Dashboard!C13:C18` |

---

## 3. Course Records — the GPA engine

For each course row *r* (grade in **F**, credits in **E**, level in **D**):

```sheets
G (grade points)     =IFERROR(VLOOKUP(F5,GradeScale,2,FALSE),"")
H (quality points)   =IFERROR(G5*E5,"")
I (weighted points)  =IFERROR((G5+IFERROR(VLOOKUP(D5,LevelScale,2,FALSE),0))*E5,"")
```

Totals row (row 50): `=SUM(E5:E49)`, `=SUM(H5:H49)`, `=SUM(I5:I49)`.

---

## 4. GPA Calculator

```sheets
Unweighted GPA   =IFERROR(QualTotal/CreditsEarned,0)     -> 3.81
Weighted GPA     =IFERROR(WtdTotal/CreditsEarned,0)      -> 4.06
Honors + AP      =COUNTIF(CrsLevel,"Honors")+COUNTIF(CrsLevel,"AP")+COUNTIF(CrsLevel,"Dual Credit")  -> 9
Course count     =COUNTA(CrsYear)                         -> 27
```

---

## 5. Dashboard — the 12 KPIs

```sheets
GPA (Weighted)    =GPAW
GPA (Unweighted)  =GPAUW
Credits Earned    =CreditsEarned
Grad Progress     =IFERROR(MIN(CreditsEarned/CreditsReq,1),0)
Courses           =CourseCount
Honors / AP       =HonorsAP
Class Of          =ClassOf
Best SAT          =1380     (best SAT on the Test Scores tab)
Best ACT          =30       (best ACT on the Test Scores tab)
Service Hrs       =ServiceHours
Activities        =COUNTA(ActName)
College-Ready     =IFERROR(AVERAGE(HealthRange),0)
```

Charts: GPA by year (bar), Credits by subject (bar/donut). Turn off auto labels.

---

## 6. College-Ready Score (6 dimensions)

```sheets
GPA vs target          =IFERROR(MIN(GPAUW/GPATarget,1),0)
Credits vs required    =IFERROR(MIN(CreditsEarned/CreditsReq,1),0)
Grad requirements met  =IFERROR(COUNTIF(ReqDone,"Yes")/COUNTA(ReqName),0)
Test scores on file    =IFERROR(MIN(COUNTA(TestName)/3,1),0)
Course rigor (H+AP)    =IFERROR(MIN(HonorsAP/RigorGoal,1),0)
Service & activities   =IFERROR(MIN(ServiceHours/ServiceGoal,1),0)
College-Ready Score    =IFERROR(AVERAGE(C13:C18),0)       -> 98%
```

Power features: `VLOOKUP` (grade & level scales), `COUNTIF`/`COUNTIFS`, `SUMIF`,
`AVERAGE`, `IFERROR`, plus `QUERY` ("courses by year", "which requirements are
open") and `FILTER`/`SORT`. Course Records is a plain table — add a row and the
GPA re-computes.

---

## 7. Official Transcript & printables

The **Official Transcript** tab renders a print-like layout (school header,
per-year course blocks, a summary box showing **UW 3.81 / W 4.06 / 24.0
credits**, and a signature line) — identical to page 1 of the PDF. The 12-page
PDF is print-ready as-is (US Letter). Google Sheets users can also print any tab:
File ▸ Print ▸ fit to width.

> Not an accredited transcript service — confirm your state's and your target
> colleges' requirements. This is a record-keeping tool, not legal advice.

---

## 8. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
