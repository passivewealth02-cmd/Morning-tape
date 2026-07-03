# Pickleball Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Player Profile, Matches, Tournaments, Practice,
Skills, Equipment, Budget, Fitness, Partners, Club & League, Travel, Nutrition,
Goals, Gallery, Analytics, Settings**.

> Build **Settings** first (player details + dropdown lists), then the Match
> Tracker & trackers, then the Dashboard. Add the named ranges below
> (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `PlayerName` (C6), `PlayLevel` (C7), `PlayerRating` (C8),
`AgeDiv` (C9), `HomeClub` (C10), `MonthlyBudget` (C11), `SeasonTarget` (C12).

Lists: `SkillList, MatchTypeList, TournLevelList, ExpenseCatList,
EquipTypeList, GoalCatList, ClubList, ResultList, CondList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `MatchResult` | `Matches!G5:G64` | `EquipName` | `Equipment!A5:A34` |
| `MatchPartner` | `Matches!C5:C64` | `EquipReplace` | `Equipment!G5:G34` |
| `MatchType` | `Matches!E5:E64` | `BudgetActual` | `Budget!C5:C15` |
| `TournName` | `Tournaments!A5:A34` | `PartnerName` | `Partners!A5:A7` |
| `TournDate` | `Tournaments!C5:C34` | `PartnerWinPct` | `Partners!E5:E7` |
| `TournPlace` | `Tournaments!F5:F34` | `PartnerRecord` | `Partners!F5:F7` |
| `TournPrize` | `Tournaments!G5:G34` | `GoalCategory` | `Goals!B5:B10` |
| `PracticeDate` | `Practice!A5:A44` | `GoalProgress` | `Goals!E5:E10` |
| `PracticeHrs` | `Practice!C5:C44` | `SkillCurrent` | `Skills!C5:C16` |
| `CourtHrs` | `Analytics!C17:C22` | `HealthRange` | `Analytics!C7:C12` |

---

## 3. Match Tracker — win % engine

```sheets
Win %              =IFERROR(COUNTIF(MatchResult,"W")/COUNTA(MatchResult),0)
Partner Wins       =COUNTIFS(MatchPartner, <name>, MatchResult, "W")
Partner Record     =Wins&"-"&Losses
```

Use conditional formatting to glow "W" mint and "L" red.

---

## 4. Dashboard — the 12 KPIs

```sheets
Matches Played    =COUNTA(MatchResult)
Win Percentage    =IFERROR(COUNTIF(MatchResult,"W")/COUNTA(MatchResult),0)
Current Rating    =PlayerRating
Tournament Wins   =COUNTIF(TournPlace,"1st")
Practice (30 d)   =COUNTIFS(PracticeDate,">="&TODAY()-30)
Court Time        =SUM(PracticeHrs)+COUNTA(MatchResult)
Monthly Budget    =MonthlyBudget
Gear to Replace   =SUMPRODUCT((EquipReplace<=TODAY()+60)*(EquipReplace<>"")*(EquipName<>""))
Fitness Progress  =IFERROR(AVERAGEIF(GoalCategory,"Fitness",GoalProgress),0)
Upcoming Events   =COUNTIF(TournDate,">="&TODAY())
Top Partner       =INDEX(PartnerRecord,MATCH(MAX(PartnerWinPct),PartnerWinPct,0))
Season Progress   =IFERROR(COUNTA(MatchResult)/SeasonTarget,0)
```

Charts: Match Results (donut), Court Time by Month (line), Skill Progress
start-vs-now (bar), Spending by Category (donut). Turn off auto data labels.

---

## 5. Analytics — Pickleball Performance Score

```sheets
Win rate               =IFERROR(COUNTIF(MatchResult,"W")/COUNTA(MatchResult),0)
Skill level            =IFERROR(AVERAGE(SkillCurrent)/10,0)
Fitness consistency    =IFERROR(AVERAGEIF(GoalCategory,"Fitness",GoalProgress),0)
Tournament success     =IFERROR(COUNTIF(TournPlace,"1st")/MAX(COUNTA(TournPlace),1),0)
Goal progress          =IFERROR(AVERAGE(GoalProgress),0)
Practice consistency   =IFERROR(MIN(COUNTIFS(PracticeDate,">="&TODAY()-30)/8,1),0)
Performance Score      =IFERROR(AVERAGE(C7:C12),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("last 5 matches"), `FILTER`/`SORT`,
`UNIQUE`, all wrapped in `IFERROR`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, W/L color flags. Keep it
premium and consistent — that polish is what makes it feel like software.
