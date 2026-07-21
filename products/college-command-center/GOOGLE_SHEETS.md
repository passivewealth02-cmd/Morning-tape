# College Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Applicant Profile, College List, Essays,
Recommendations, Test Scores, Activities, Scholarships, Net Price, Visits,
Decisions, To-Do, Deadlines, Settings**.

> Build **College List** first (the engine), then Essays, Recommendations,
> Scholarships and To-Do, then the Dashboard. Add the named ranges below
> (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Applicant`, `ClassOf`, `GPAW`, `BestSAT`, `BestACT`, `Major`,
`SchGoal` (8).

Lists: `TypeList, YesNoList, EssayList, DecisionList, SchList, RecList, AppList`.

---

## 2. College List — the engine

For each school (essays / recs / form / submitted as Yes/No in **E:H**):

```sheets
Progress (I)  =IFERROR(COUNTIF(E5:H5,"Yes")/4,0)
```

Named: `CollegeName` (B), `CollegeType` (C), `CollegeDeadline` (D),
`CollegeSubmitted` (H), `CollegeProg` (I).

---

## 3. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CollegeName` | `'College List'!B5:B12` | `CollegeType` | `'College List'!C5:C12` |
| `CollegeDeadline` | `'College List'!D5:D12` | `CollegeSubmitted` | `'College List'!H5:H12` |
| `CollegeProg` | `'College List'!I5:I12` | `EssayStatus` | `Essays!D5:D44` |
| `RecSubmitted` | `Recommendations!D5:D20` | `SchStatus` | `Scholarships!D5:D34` |
| `SchAmount` | `Scholarships!B5:B34` | `AidAwarded` | `Scholarships!B35` |
| `NetPrice` | `'Net Price'!E5:E10` | `Decision` | `Decisions!B5:B24` |
| `TaskName` | `'To-Do'!A5:A44` | `TaskDone` | `'To-Do'!B5:B44` |
| `HealthRange` | `Dashboard!C13:C18` | | |

---

## 4. Dashboard — the 12 KPIs

```sheets
Colleges         =COUNTA(CollegeName)
Apps Submitted   =COUNTIF(CollegeSubmitted,"Yes")
Avg Progress     =IFERROR(AVERAGE(CollegeProg),0)
Essays Done      =IFERROR(COUNTIF(EssayStatus,"Final")/COUNTA(EssayStatus),0)
Recs Secured     =COUNTIF(RecSubmitted,"Yes")
Scholarships     =COUNTIF(SchStatus,"Applied")+COUNTIF(SchStatus,"Awarded")
Aid Awarded      =AidAwarded          (=SUMIF(SchStatus,"Awarded",SchAmount))
Next Deadline    =IFERROR(MINIFS(CollegeDeadline,CollegeSubmitted,"No"),"")
Acceptances      =COUNTIF(Decision,"Accepted")
Best Net Price   =IFERROR(MIN(NetPrice),0)
Tasks Done       =IFERROR(COUNTIF(TaskDone,"Yes")/COUNTA(TaskName),0)
Ready Score      =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Reach · Match · Safety (donut) from `COUNTIF(CollegeType,…)`.

---

## 5. Ready Score (6 dimensions)

```sheets
Application progress  =IFERROR(AVERAGE(CollegeProg),0)
Apps submitted        =IFERROR(COUNTIF(CollegeSubmitted,"Yes")/COUNTA(CollegeName),0)
Essays final          =IFERROR(COUNTIF(EssayStatus,"Final")/COUNTA(EssayStatus),0)
Recs secured          =IFERROR(COUNTIF(RecSubmitted,"Yes")/COUNTA(RecSubmitted),0)
Scholarships applied  =IFERROR((COUNTIF(SchStatus,"Applied")+COUNTIF(SchStatus,"Awarded"))/SchGoal,0)
Tasks done            =IFERROR(COUNTIF(TaskDone,"Yes")/COUNTA(TaskName),0)
Ready Score           =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `COUNTIF`, `COUNTA`, `AVERAGE`, `SUMIF`, `MIN`, `MINIFS`,
`IFERROR`, plus `QUERY`/`FILTER` ("what's due this week", "schools not submitted")
and color scales (net price: lowest = green).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter). Print any tab: File ▸ Print ▸
fit to width.

> An organizing tool, not admissions advice — confirm every deadline, fee &
> requirement directly with each college.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
