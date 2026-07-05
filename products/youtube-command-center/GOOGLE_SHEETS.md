# YouTube Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Channel Profile, Calendar, Pipeline, Ideas,
Scripts, Thumbnails, SEO, Analytics, Long Form, Shorts, Community, Live,
Playlists, Sponsors, Affiliate, Products, Finance, Equipment, Repurposing,
Brand Kit, Assets, AI Prompts, Goals, Collabs, Audience, Taxes, Annual Plan,
Gallery, Settings**.

> Build **Settings** first (channel details + dropdown lists), then the
> Calendar, Analytics, Finance & CRM, then the Dashboard. Add the named ranges
> below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `ChannelName`, `CreatorName`, `Niche`, `RevenueGoal` (12000),
`UploadGoal` (7), `DealTarget` (6), `AVDTarget` (6), plus subscriber / watch-hour
goals.

Lists: `PillarList, VideoTypeList, StatusList, PriorityList, StageList,
SponStageList, GoalCatList, PlatformList, ExpenseCatList, RevCatList,
YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CalDate` | `Calendar!A5:A64` | `RevMonthly` | `Finance!B5:B12` |
| `CalStatus` | `Calendar!F5:F64` | `RevenueTotal` | `Finance!B13` |
| `SubNow` | `Analytics!C6` | `ExpMonthly` | `Finance!G6:G13` |
| `Views28` | `Analytics!C7` | `ExpenseTotal` | `Finance!G14` |
| `WatchHrs` | `Analytics!C8` | `SponStage` | `Sponsors!E5:E34` |
| `AvgCTR` | `Analytics!C9` | `SponRate` | `Sponsors!D5:D34` |
| `AvgViewDur` | `Analytics!C10` | `GoalProgress` | `Goals!E5:E11` |
| `RPM` | `Analytics!C12` | `GoalCategory` | `Goals!B5:B11` |
| `HealthRange` | `Analytics!F7:F12` | `VidViews` | `Analytics!C17:C22` |

---

## 3. Calendar — publishing status calculates itself

```sheets
Priority (per row)   =IF(D5="Long-form","High","Medium")
Published (28d)      =COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Published")
Upload consistency   =IFERROR(MIN(Published28/UploadGoal,1),0)
```

Glow "Published" mint, "Scheduled" gold, "Editing/Filming" soft with
conditional formatting.

---

## 4. Dashboard — the 12 KPIs

```sheets
Subscribers        =SubNow
Views (28d)        =Views28
Watch Hours        =WatchHrs
Published (28d)    =COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Published")
Monthly Revenue    =RevenueTotal
Net Profit         =RevenueTotal-ExpenseTotal
RPM                =RPM
Avg CTR            =AvgCTR
Avg View Duration  =AvgViewDur
Brand Deals        =COUNTIF(SponStage,"Signed")+COUNTIF(SponStage,"Delivered")+COUNTIF(SponStage,"Negotiation")
Upload Consistency =IFERROR(MIN(COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Published")/UploadGoal,1),0)
Channel Health     =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Subscriber Growth (line), Revenue by Source (donut), Top Videos by
Views (bar), Expense Breakdown (donut). Turn off auto data labels.

---

## 5. Finance & Analytics

```sheets
Annual (est.)      =B5*12
% of Rev           =IFERROR(B5/RevenueTotal,0)
Total Revenue      =SUM(B5:B12)
Net profit         =RevenueTotal-ExpenseTotal
Profit margin      =IFERROR((RevenueTotal-ExpenseTotal)/RevenueTotal,0)

Channel Health dimensions
Revenue vs goal      =IFERROR(MIN(RevenueTotal/RevenueGoal,1),0)
Upload consistency   =IFERROR(MIN(Published28/UploadGoal,1),0)
Audience / CTR       =IFERROR(MIN(AvgCTR/0.08,1),0)
Retention            =IFERROR(MIN((LEFT(AvgViewDur,FIND(":",AvgViewDur)-1)+MID(AvgViewDur,FIND(":",AvgViewDur)+1,2)/60)/AVDTarget,1),0)
Sponsorship pipeline =IFERROR(MIN(ActiveDeals/DealTarget,1),0)
Goal progress        =IFERROR(AVERAGE(GoalProgress),0)
Channel Health Score =IFERROR(AVERAGE(F7:F12),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("top videos", "revenue by month"),
`FILTER`/`SORT`/`UNIQUE`, all wrapped in `IFERROR`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, status color flags. Keep it
premium and consistent — that polish is what makes it feel like media-company
software, not a spreadsheet.
