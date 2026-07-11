# Podcast Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Show Profile, Calendar, Pipeline, Episodes, Guests,
Recording, Show Notes, Analytics, Platforms, Sponsors, Members, Clips,
Repurposing, Finance, Expenses, Equipment, Reviews, Brand Kit, Gallery, Goals,
Audience, Collabs, Settings**.

> Build **Settings** first (show + goals + dropdown lists), then the Calendar,
> Analytics, Sponsors, Members & Finance, then the Dashboard. Add the named
> ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `ShowName`, `HostName`, `Category`, `RevenueGoal` (10000),
`EpGoal` (5), `DownloadGoal` (50000), `ConsumptionTarget` (0.80),
`SponsorTarget` (4), `GrowthGoal` (2000).

Lists: `FormatList, StatusList, RevCatList, ExpCatList, SponStageList,
GuestStatusList, GoalCatList, PriorityList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CalDate` | `Calendar!A5:A64` | `RevenueTotal` | `Finance!B12` |
| `CalStatus` | `Calendar!F5:F64` | `ExpenseTotal` | `Finance!G13` |
| `Subscribers` | `Analytics!C6` | `MemberCount` | `Members!C9` |
| `SubGrowth` | `Analytics!C7` | `MemberRev` | `Members!D9` |
| `Downloads28` | `Analytics!C8` | `SponStage` | `Sponsors!C5:C34` |
| `AvgPerEp` | `Analytics!C9` | `GoalProgress` | `Goals!E5:E11` |
| `Consumption` | `Analytics!C11` | `HealthRange` | `Analytics!F7:F12` |
| `CPM` | `Analytics!C13` | `DlVal` | `Analytics!C27:C32` |

---

## 3. Members, Calendar & publishing

```sheets
Tier monthly          =B5*C5                 (price × members)
Members total         =SUM(C5:C8)
Member revenue        =SUM(D5:D8)
Episodes (28d)        =COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Published")
Publishing consistency=IFERROR(MIN(Episodes28/EpGoal,1),0)
```

The Finance "Memberships" line is `=MemberRev`, so Members and Finance stay in sync.

---

## 4. Dashboard — the 12 KPIs

```sheets
Subscribers          =Subscribers
Downloads (28d)      =Downloads28
Avg / Episode        =AvgPerEp
Episodes (28d)       =COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Published")
Consumption          =Consumption
Monthly Revenue      =RevenueTotal
Net Profit           =RevenueTotal-ExpenseTotal
Ad Rate (CPM)        =CPM
Active Sponsors      =COUNTIF(SponStage,"Booked")+COUNTIF(SponStage,"Live")
Members              =MemberCount
Publishing           =IFERROR(MIN(Episodes28/EpGoal,1),0)
Show Health          =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Downloads (line), Revenue by Source (donut), Top Episodes by Downloads
(bar), Expense Breakdown (donut). Turn off auto data labels.

---

## 5. Analytics — Show Health Score

```sheets
Revenue vs goal        =IFERROR(MIN(RevenueTotal/RevenueGoal,1),0)
Publishing consistency =IFERROR(MIN(Episodes28/EpGoal,1),0)
Downloads vs goal      =IFERROR(MIN(Downloads28/DownloadGoal,1),0)
Consumption            =IFERROR(MIN(Consumption/ConsumptionTarget,1),0)
Sponsor pipeline       =IFERROR(MIN(ActiveSponsors/SponsorTarget,1),0)
Audience growth        =IFERROR(MIN(SubGrowth/GrowthGoal,1),0)
Show Health Score      =IFERROR(AVERAGE(F7:F12),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("top episodes", "downloads by platform",
"sponsors due this week"), `FILTER`/`SORT`/`UNIQUE`, all wrapped in `IFERROR`.

Cover & Clip Gallery: click a cell ▸ Insert ▸ Image ▸ **Image in cell**, or
paste `=IMAGE("your-art-link")`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, status color flags. Keep it
premium and consistent — that polish is what makes it feel like producer
software, not a spreadsheet.
