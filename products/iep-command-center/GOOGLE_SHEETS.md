# IEP Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Student Profile, IEP Goals, Progress Monitoring,
Services, Accommodations, Therapy Log, Behavior, Meetings, Communication,
Health & Meds, Strengths, Records, Wins, Settings**.

> Build **IEP Goals** first (the engine), then Services, Accommodations,
> Monitoring and Records, then the Dashboard. Add the named ranges below
> (Data ▸ Named ranges). Keep the file private.

---

## 1. Settings — controls & lists

Controls: `ChildName`, `Grade`, `PlanType`, `CaseManager`, `NextReview`,
`DataTarget` (30).

Lists: `AreaList, YesNoList, SettingList, ServiceList, MeetList, DoneList`.

---

## 2. IEP Goals — the engine

For each goal (baseline in **D**, target in **E**, current in **F**):

```sheets
Progress (G)  =IFERROR(MIN(MAX((F5-D5)/(E5-D5),0),1),0)
```

Named: `GoalArea` (B), `GoalName` (C), `GoalPct` (G).

---

## 3. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `GoalArea` | `'IEP Goals'!B5:B9` | `GoalName` | `'IEP Goals'!C5:C9` |
| `GoalPct` | `'IEP Goals'!G5:G9` | `DataDate` | `'Progress Monitoring'!A5:A64` |
| `ServiceName` | `Services!B5:B8` | `ServiceMin` | `Services!E5:E8` |
| `ServiceDeliv` | `Services!F5:F8` | `ServiceMinTotal` | `Services!E9` |
| `ServiceDelivTotal` | `Services!F9` | `AccomActive` | `Accommodations!C5:C28` |
| `AccomName` | `Accommodations!A5:A28` | `TherapyDate` | `'Therapy Log'!A5:A44` |
| `MeetDate` | `Meetings!A5:A24` | `WinName` | `Wins!A5:A34` |
| `RecDone` | `Records!B5:B24` | `RecName` | `Records!A5:A24` |
| `HealthRange` | `Dashboard!C13:C18` | | |

---

## 4. Dashboard — the 12 KPIs

```sheets
IEP Goals         =COUNTA(GoalName)
Avg Progress      =IFERROR(AVERAGE(GoalPct),0)
Goals On Pace     =COUNTIF(GoalPct,">=0.6")
Services          =COUNTA(ServiceName)
Service Min/Wk    =ServiceMinTotal
Accommodations    =COUNTIF(AccomActive,"Yes")
Therapy Logged    =COUNTA(TherapyDate)
Data Points       =COUNTA(DataDate)
Meetings          =COUNTA(MeetDate)
Wins              =COUNTA(WinName)
Supports in Place =IFERROR(COUNTIF(AccomActive,"Yes")/COUNTA(AccomActive),0)
Progress Score    =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Progress Toward Each Goal (bar) from `GoalPct` with `GoalArea` categories.

---

## 5. Progress Score (6 dimensions)

```sheets
Goal progress            =IFERROR(AVERAGE(GoalPct),0)
Goals on pace            =IFERROR(COUNTIF(GoalPct,">=0.6")/COUNTA(GoalPct),0)
Accommodations in place  =IFERROR(COUNTIF(AccomActive,"Yes")/COUNTA(AccomActive),0)
Services delivered       =IFERROR(ServiceDelivTotal/ServiceMinTotal,0)
Data collection          =IFERROR(MIN(COUNTA(DataDate)/DataTarget,1),0)
Records ready            =IFERROR(COUNTIF(RecDone,"Yes")/COUNTA(RecName),0)
Progress Score           =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `IF`, `IFERROR`, `MIN`/`MAX`, `COUNTIF`, `AVERAGE`, `SUM`, plus
`QUERY`/`FILTER` ("goals below pace", "unmet accommodations", "open follow-ups").

---

## 6. Printables & privacy

The 12-page PDF is print-ready as-is (US Letter). Keep this file private — it may
contain sensitive information; share only with those who need it.

> Not medical, legal or educational advice. An organizing & advocacy tool that
> does not create or replace an IEP/504 — always work with your child's IEP team.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
