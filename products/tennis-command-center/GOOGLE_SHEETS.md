# Tennis Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Player Profile, Matches, Match Analytics,
Tournaments, Practice, Skills, Fitness, Equipment, Budget, Partners, Coaching,
Goals, Travel, Nutrition, Media, Season, Analytics, Settings**.

> Build **Settings** first (player details + dropdown lists), then the Match
> Tracker & trackers, then the Dashboard. Add the named ranges below
> (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `PlayerName` (C6), `AgeDiv` (C7), `PlayerRating` (C8),
`CurrentRank` (C9), `HomeClub` (C10), `MonthlyBudget` (C11), `SeasonTarget` (C12),
`HoursTarget` (C13).

Lists: `SurfaceList, FormatList, TournLevelList, SkillList, ExpenseCatList,
EquipTypeList, GoalCatList, FocusList, ResultList, CondList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `MatchResult` | `Matches!K5:K64` | `EquipName` | `Equipment!A5:A34` |
| `MatchSurface` | `Matches!D5:D64` | `EquipReplace` | `Equipment!G5:G34` |
| `MatchSetsW` | `Matches!G5:G64` | `BudgetActual` | `Budget!C5:C15` |
| `MatchGamesW` | `Matches!I5:I64` | `TournPlace` | `Tournaments!G5:G34` |
| `PracticeHrs` | `Practice!B5:B44` | `GoalCategory` | `Goals!B5:B11` |
| `SkillCurrent` | `Skills!C5:C17` | `GoalProgress` | `Goals!E5:E11` |
| `RateVal` | `Analytics!C17:C22` | `HealthRange` | `Analytics!C7:C12` |

---

## 3. Match Tracker — auto win %, sets, games

```sheets
Result (per row)   =IF(G5="","",IF(G5>H5,"W","L"))     (from set counts)
Win %              =IFERROR(COUNTIF(MatchResult,"W")/COUNTA(MatchResult),0)
Sets Won           =SUM(MatchSetsW)
Games Won          =SUM(MatchGamesW)
Surface record     =COUNTIFS(MatchSurface,"Hard",MatchResult,"W")
```

Enter set-by-set games (Sets W/L, Games W/L); Result derives itself. Glow "W"
mint and "L" red with conditional formatting.

---

## 4. Dashboard — the 12 KPIs

```sheets
Matches Played    =COUNTA(MatchResult)
Win Percentage    =IFERROR(COUNTIF(MatchResult,"W")/COUNTA(MatchResult),0)
Current Ranking   ="#"&CurrentRank
Current Rating    =PlayerRating
Practice Hours    =SUM(PracticeHrs)
Tournament Wins   =COUNTIF(TournPlace,"1st")
Sets Won          =SUM(MatchSetsW)
Games Won         =SUM(MatchGamesW)
Fitness Score     =IFERROR(AVERAGEIF(GoalCategory,"Fitness",GoalProgress),0)
Gear to Replace   =SUMPRODUCT((EquipReplace<=TODAY()+45)*(EquipReplace<>"")*(EquipName<>""))
Monthly Budget    =MonthlyBudget
Performance       =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Match Results (donut), Rating Progress UTR (line), Skill Progress
start-vs-now (bar), Budget Breakdown (donut). Match Analytics adds a serve-stats
column chart. Turn off auto data labels.

---

## 5. Analytics — Player Performance Score

```sheets
Win rate              =IFERROR(COUNTIF(MatchResult,"W")/COUNTA(MatchResult),0)
Skill level           =IFERROR(AVERAGE(SkillCurrent)/10,0)
Fitness               =IFERROR(AVERAGEIF(GoalCategory,"Fitness",GoalProgress),0)
Tournament success    =IFERROR(COUNTIF(TournPlace,"1st")/MAX(COUNTA(TournPlace),1),0)
Goal progress         =IFERROR(AVERAGE(GoalProgress),0)
Practice consistency  =IFERROR(MIN(SUM(PracticeHrs)/HoursTarget,1),0)
Performance Score     =IFERROR(AVERAGE(C7:C12),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("surface splits", "last 5 matches"),
`FILTER`/`SORT`/`UNIQUE`, all wrapped in `IFERROR`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, W/L color flags. Keep it
premium and consistent — that polish is what makes it feel like academy software.
