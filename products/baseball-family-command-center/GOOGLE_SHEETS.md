# Baseball Family Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same
tab order: **Welcome, Dashboard, Player Profile, Calendar, Game Day,
Practice, Budget, Equipment, Bats, Gloves, Stats, Pitching, Tournaments,
Travel, Roster, Volunteers, Development, Nutrition, Medical, Packing,
Fundraising, Gallery, Goals, Analytics, Settings**.

> Build **Settings** first — controls + dropdown lists. Then add the
> cross-sheet named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `PlayerName` (C6), `TeamName` (C7), `MonthlyBudget` (C8),
`SeasonStart` (C9), `SeasonEnd` (C10), `PitchLimit` (C11), `GameInnings`
(C12), `HomePark` (C13).

Dropdown lists: `EventTypeList, LeagueList, DivisionList, PositionList,
ExpenseCatList, EquipTypeList, EquipStatusList, VolRoleList, PracticeTypeList,
TourneyTypeList, SkillLevelList, AttendList, YesNoList, ExpenseGroupList,
BatsList, ThrowsList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CalType` | `Calendar!A5:A74` | `EquipReplace` | `Equipment!G5:G44` |
| `CalDate` | `Calendar!C5:C74` | `EquipName` | `Equipment!A5:A44` |
| `CalConflict` | `Calendar!H5:H74` | `BudgetGroup` | `Budget!B5:B24` |
| `PracDate` | `Practice!A5:A64` | `BudgetActual` | `Budget!D5:D24` |
| `PracAttend` | `Practice!D5:D64` | `BudgetTotalPlanned` | `Budget!C25` |
| `PracHours` | `Practice!F5:F64` | `BudgetTotalActual` | `Budget!D25` |
| `TravelMiles` | `Travel!E5:E44` | `StatAB` | `Stats!C5:C34` |
| `VolHours` | `Volunteers!E5:E44` | `StatH` | `Stats!D5:D34` |
| `DevProgress` | `Development!C5:C24` | `StatBB` | `Stats!J5:J34` |
| `GoalProgress` | `Goals!D5:D34` | `PitchCount` | `Pitching!D5:D29` |

---

## 3. Dashboard — the 12 KPIs

```sheets
Days to Next Game  =IF(COUNTIFS(CalType,"Game",CalDate,">="&TODAY())=0,0,
                      MINIFS(CalDate,CalType,"Game",CalDate,">="&TODAY())-TODAY())
Games / Week       =COUNTIFS(CalType,"Game",CalDate,">="&TODAY(),CalDate,"<="&TODAY()+7)
Practices / Week   =COUNTIFS(CalType,"Practice",CalDate,">="&TODAY(),CalDate,"<="&TODAY()+7)
Next Tournament    =IF(COUNTIFS(CalType,"Tournament",CalDate,">="&TODAY())=0,0,
                      MINIFS(CalDate,CalType,"Tournament",CalDate,">="&TODAY())-TODAY())
Budget Left        =BudgetTotalPlanned-BudgetTotalActual
Equip. Alerts      =SUMPRODUCT((EquipReplace<>"")*(EquipReplace<=TODAY()+45))
Games Played       =COUNTIFS(CalType,"Game",CalDate,"<"&TODAY())
Travel Miles       =SUM(TravelMiles)
Volunteer Hrs      =SUM(VolHours)
Dev Progress       =IFERROR(AVERAGE(DevProgress),0)
Season Complete    =IFERROR(COUNTIFS(CalType,"Game",CalDate,"<"&TODAY())/MAX(COUNTIF(CalType,"Game"),1),0)
```

---

## 4. Player Statistics — the slash line

```sheets
AVG =IFERROR(SUM(H)/SUM(AB),0)
OBP =IFERROR((SUM(H)+SUM(BB))/(SUM(AB)+SUM(BB)),0)
1B  = SUM(H)-SUM(2B)-SUM(3B)-SUM(HR)
TB  = 1B + 2*SUM(2B) + 3*SUM(3B) + 4*SUM(HR)
SLG =IFERROR(TB/SUM(AB),0)
OPS = OBP + SLG
```

Format AVG/OBP/SLG with `.000`. Use `ARRAYFORMULA` if you want per-row rate
stats.

---

## 5. Pitching — ERA & pitch-count alerts

```sheets
Strike %   =IF(Pitches="","",Strikes/Pitches)
ERA        =IFERROR(SUM(ER)/SUM(IP)*GameInnings,0)
Warning    =AND($D5<>"", $D5>PitchLimit)   → red highlight (over your limit)
```

---

## 6. Budget, Calendar, Equipment

```sheets
Budget:   Remaining =C-D   ·  %Used =IFERROR(D/C,0)  ·  Cost/Game =BudgetTotalActual/COUNTIF(CalType,"Game")
Calendar: Conflict  =IF(C5="","",IF(AND(COUNTIF($C$5:C5,C5)=1,COUNTIF(CalDate,C5)>1),1,0))
Equipment: due      =AND($G5<>"",$G5<=TODAY()+45) → gold highlight
By group (chart): =SUMIF(BudgetGroup,<group>,BudgetActual)
```

---

## 7. Analytics — Baseball Readiness Score

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
