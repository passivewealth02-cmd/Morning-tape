# Hockey Family Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same
tab order: **Welcome, Dashboard, Player Profile, Calendar, Game Day,
Practice, Budget, Equipment, Sharpening, Sticks, Tournaments, Travel, Hotels,
Carpool, Roster, Development, Stats, Nutrition, Medical, Packing, Fundraising,
Team Comms, Gallery, Goals, Analytics, Settings**.

> Build **Settings** first — controls + dropdown lists. Then add the
> cross-sheet named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `PlayerName` (C6), `TeamName` (C7), `MonthlyBudget` (C8),
`SeasonStart` (C9), `SeasonEnd` (C10), `HomeArena` (C11).

Dropdown lists: `EventTypeList, LeagueList, DivisionList, PositionList,
ExpenseCatList, EquipStatusList, SkillLevelList, TourneyTypeList,
TravelStatusList, PracticeTypeList, AttendList, ShootsList, YesNoList,
ExpenseGroupList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CalType` | `Calendar!A5:A74` | `EquipReplace` | `Equipment!G5:G44` |
| `CalDate` | `Calendar!C5:C74` | `EquipName` | `Equipment!A5:A44` |
| `CalConflict` | `Calendar!H5:H74` | `SharpDate` | `Sharpening!A5:A44` |
| `PracDate` | `Practice!A5:A64` | `BudgetCat` | `Budget!A5:A30` |
| `PracAttend` | `Practice!F5:F64` | `BudgetGroup` | `Budget!B5:B30` |
| `PracIce` | `Practice!G5:G64` | `BudgetActual` | `Budget!D5:D30` |
| `TravelMiles` | `Travel!E5:E44` | `BudgetTotalPlanned` | `Budget!C31` |
| `DevProgress` | `Development!C5:C24` | `BudgetTotalActual` | `Budget!D31` |
| `GoalProgress` | `Goals!D5:D34` | `StatPTS` | `Stats!E5:E29` |

---

## 3. Calendar — countdowns & conflicts

```sheets
Days (col G)      =IF(C5="","",C5-TODAY())
Conflict (col H)  =IF(C5="","",IF(AND(COUNTIF($C$5:C5,C5)=1,COUNTIF(CalDate,C5)>1),1,0))
Conflict format   =AND($C5<>"",COUNTIF(CalDate,$C5)>1)  → highlight row
```

`MINIFS`, `COUNTIFS`, `SUMPRODUCT` all work in Google Sheets — no array
entry needed. Use `ARRAYFORMULA` for the G/H helper columns if you prefer one
formula per column.

---

## 4. Dashboard — the 12 KPIs

```sheets
Days to Next Game  =IF(COUNTIFS(CalType,"Game",CalDate,">="&TODAY())=0,0,
                      MINIFS(CalDate,CalType,"Game",CalDate,">="&TODAY())-TODAY())
Practices / Week   =COUNTIFS(CalType,"Practice",CalDate,">="&TODAY(),CalDate,"<="&TODAY()+7)
Games / Month      =COUNTIFS(CalType,"Game",CalDate,">="&TODAY(),CalDate,"<="&TODAY()+30)
Next Tournament    =IF(COUNTIFS(CalType,"Tournament",CalDate,">="&TODAY())=0,0,
                      MINIFS(CalDate,CalType,"Tournament",CalDate,">="&TODAY())-TODAY())
Monthly Budget     =MonthlyBudget
Budget Left        =BudgetTotalPlanned-BudgetTotalActual
Equip. Alerts      =SUMPRODUCT((EquipReplace<>"")*(EquipReplace<=TODAY()+45))
Ice Time (30d)     =SUMIFS(PracIce,PracDate,">="&TODAY()-30,PracDate,"<="&TODAY())
Conflicts          =SUM(CalConflict)
Attendance         =IFERROR(COUNTIF(PracAttend,"Present")/MAX(COUNTA(PracDate),1),0)
Travel Miles       =SUM(TravelMiles)
Season Progress    =IFERROR(COUNTIFS(CalType,"Game",CalDate,"<"&TODAY())/MAX(COUNTIF(CalType,"Game"),1),0)
```

Charts: Hockey Spending (donut by group), Practices vs Games (donut), Skill
Progress (bar), Season Totals (bar). Turn off auto data labels (value only).

---

## 5. Budget — costs & cost-per-game

```sheets
Remaining (E)   =C5-D5
% Used   (F)    =IFERROR(D5/C5,0)
Totals          C31=SUM(C5:C30)   D31=SUM(D5:D30)
Cost / Game     =IFERROR(BudgetTotalActual/MAX(COUNTIF(CalType,"Game"),1),0)
By group (chart)=SUMIF(BudgetGroup,<group>,BudgetActual)
```

---

## 6. Equipment, Sharpening, Stats

```sheets
Equipment due    =AND($G5<>"", $G5<=TODAY()+45)              → gold highlight
Sharpening due   =TODAY()-MAX(SharpDate)>=14 → "⚠ Sharpen soon"
Stats points     =N(C5)+N(D5)        Season totals =SUM(StatG) / SUM(StatA) / SUM(StatPTS)
```

---

## 7. Analytics — Hockey Readiness Score

```sheets
=IFERROR(AVERAGE(
   1-BudgetUsed,  Attendance,  EquipmentReady,
   AVERAGE(DevProgress),  AVERAGE(GoalProgress),  SeasonComplete), 0)
```

Power features: `ARRAYFORMULA` for helper columns, `QUERY` for roll-ups,
`FILTER`/`SORT` for "next 3 games", `UNIQUE` for lists, all wrapped in
`IFERROR`.

---

## 8. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, mint/gold/red status
coding. Keep every tab consistent — that cohesion is the premium feel.
