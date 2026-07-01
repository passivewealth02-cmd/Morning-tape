# Baby Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Profile, Daily Log, Feeding, Sleep, Diapers,
Growth, Milestones, Medical, Appointments, Budget, Shopping, Inventory,
Childcare, Routine, Development, Memories, Travel, Family Goals, Analytics,
Settings**.

> Build **Settings** first (baby details + dropdown lists), then the trackers,
> then the Dashboard. Add the named ranges below (Data ▸ Named ranges).

> ⚠️ Keep the **Welcome-tab disclaimer** in the Google Sheets version too —
> organization & record-keeping tool, not medical advice.

---

## 1. Settings — controls & lists

Controls: `BabyName` (C6), `BirthDate` (C7), `MonthlyBudget` (C8),
`FeedGoal` (C9), `SleepGoal` (C10).

Lists: `FeedTypeList, SleepCatList, ApptTypeList, ExpenseCatList,
MilestoneCatList, SizeList, SupplyCatList, DiaperTypeList, StatusList,
YesNoList, RoutineTypeList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `FeedDate` | `Feeding!A5:A64` | `MileDate` | `Milestones!C5:C34` |
| `FeedType` | `Feeding!C5:C64` | `MileName` | `Milestones!B5:B34` |
| `SleepDate` | `Sleep!A5:A54` | `MedType` | `Medical!B5:B44` |
| `SleepHrs` | `Sleep!E5:E54` | `MedStatus` | `Medical!E5:E44` |
| `DiaperDate` | `Diapers!A5:A64` | `ApptDate` | `Appointments!C5:C44` |
| `DiaperType` | `Diapers!C5:C64` | `GrowthDate` | `Growth!A5:A28` |
| `BudgetActual` | `Budget!C5:C16` | `GoalProgress` | `'Family Goals'!D5:D34` |
| `BudgetTotalPlanned` | `Budget!B17` | `RoutineDone` | `Routine!J5:J26` |

---

## 3. Trackers — daily totals & 7-day trends

```sheets
Feeding 7-day  =COUNTIF(FeedDate, <day cell>)
Sleep 7-day    =SUMIF(SleepDate, <day cell>, SleepHrs)
Diaper 7-day   =COUNTIF(DiaperDate, <day cell>)
Diaper by type =COUNTIF(DiaperType, "Wet")   (etc.)
```

Use `ARRAYFORMULA` for the summary columns, or `QUERY`:
`=QUERY(Feeding!A5:A, "select A, count(A) group by A order by A", 0)`.

---

## 4. Dashboard — the 12 KPIs

```sheets
Baby's Age (days)  =TODAY()-BirthDate
Feedings Today     =COUNTIF(FeedDate,TODAY())
Sleep Today        =SUMIF(SleepDate,TODAY(),SleepHrs)
Diapers Today      =COUNTIF(DiaperDate,TODAY())
Next Appt          =IF(COUNTIF(ApptDate,">="&TODAY())=0,0,MINIFS(ApptDate,ApptDate,">="&TODAY())-TODAY())
Milestones         =IFERROR(COUNTIF(MileDate,"<>"&"")/MAX(COUNTA(MileName),1),0)
Vaccines Done      =IFERROR(COUNTIFS(MedType,"Vaccination",MedStatus,"Done")/MAX(COUNTIF(MedType,"Vaccination"),1),0)
Budget Left        =BudgetTotalPlanned-BudgetTotalActual
Goals Progress     =IFERROR(AVERAGE(GoalProgress),0)
Routine Done       =IFERROR(COUNTIF(RoutineDone,"Yes")/MAX(COUNTA(RoutineDone),1),0)
```

Charts: Feedings/Day (line), Sleep Hours/Day (line), Weight Over Time (line),
Baby Spending (donut). Turn off auto data labels.

---

## 5. Analytics — Family Organization Score

```sheets
=IFERROR(AVERAGE(
   MIN(FeedingsToday/FeedGoal,1),  MIN(SleepToday/SleepGoal,1),
   1-BudgetUsed,  MilestonesLogged,  VaccinesDone,  AVERAGE(GoalProgress)), 0)
```

Power features: `ARRAYFORMULA`, `QUERY`, `FILTER`/`SORT` ("next 3
appointments"), `UNIQUE`, all wrapped in `IFERROR`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, mint "achieved" rows.
Keep it warm, gentle and consistent — that calm is the whole feel.
