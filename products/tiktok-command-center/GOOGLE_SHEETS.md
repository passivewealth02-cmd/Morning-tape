# TikTok Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Profile, Calendar, Pipeline, Ideas, Hooks, Trends,
Hashtags, Analytics, LIVE, Shop, Affiliate, Brand Deals, Series, Finance,
Expenses, Equipment, Repurposing, Brand Kit, Gallery, Goals, Audience, Collabs,
Settings**.

> Build **Settings** first (handle + goals + dropdown lists), then the Calendar,
> Analytics, Shop, Brand Deals & Finance, then the Dashboard. Add the named
> ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Handle`, `CreatorName`, `Niche`, `RevenueGoal` (10000),
`PostGoal` (24), `DealTarget` (5), `AVDTarget` (0.55), `GrowthGoal` (20000).

Lists: `PillarList, VideoTypeList, StatusList, RevCatList, ExpCatList,
HookTypeList, GoalCatList, SponStageList, TrendStatusList, PriorityList,
YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CalDate` | `Calendar!A5:A64` | `RevenueTotal` | `Finance!B12` |
| `CalStatus` | `Calendar!F5:F64` | `ExpenseTotal` | `Finance!G12` |
| `FollowerNow` | `Analytics!C6` | `ShopGMV` | `Shop!F13` |
| `FollowerGrowth` | `Analytics!C7` | `ShopEarn` | `Shop!G13` |
| `Views28` | `Analytics!C8` | `SponStage` | `Brand Deals!E5:E34` |
| `EngRate` | `Analytics!C10` | `GoalProgress` | `Goals!E5:E11` |
| `Completion` | `Analytics!C11` | `HealthRange` | `Analytics!F7:F12` |
| `VidViews` | `Analytics!C17:C22` | `FollowVal` | `Analytics!C27:C32` |

---

## 3. Shop, Calendar & posting

```sheets
GMV (per product)     =C5*E5                (price × units)
Your earnings         =F5*D5                (GMV × commission %)
Shop total GMV        =SUM(F5:F12)
Posted (28d)          =COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Posted")+COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Viral")
Posting consistency   =IFERROR(MIN(Posted28/PostGoal,1),0)
```

The Finance "TikTok Shop" line is `=ShopEarn`, so Shop and Finance stay in sync.

---

## 4. Dashboard — the 12 KPIs

```sheets
Followers            =FollowerNow
Views (28d)          =Views28
Engagement Rate      =EngRate
Avg Completion       =Completion
Posted (28d)         =COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Posted")+COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Viral")
Monthly Revenue      =RevenueTotal
Net Profit           =RevenueTotal-ExpenseTotal
Shop Sales (GMV)     =ShopGMV
Brand Deals          =COUNTIF(SponStage,"Signed")+COUNTIF(SponStage,"Delivered")+COUNTIF(SponStage,"Negotiation")
Follower Growth      =FollowerGrowth
Posting Consistency  =IFERROR(MIN(Posted28/PostGoal,1),0)
Creator Health       =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Follower Growth (line), Revenue by Source (donut), Top Videos by Views
(bar), Expense Breakdown (donut). Turn off auto data labels.

---

## 5. Analytics — Creator Health Score

```sheets
Revenue vs goal      =IFERROR(MIN(RevenueTotal/RevenueGoal,1),0)
Posting consistency  =IFERROR(MIN(Posted28/PostGoal,1),0)
Engagement           =IFERROR(MIN(EngRate/0.10,1),0)
Retention            =IFERROR(MIN(Completion/AVDTarget,1),0)
Brand pipeline       =IFERROR(MIN(ActiveDeals/DealTarget,1),0)
Follower growth      =IFERROR(MIN(FollowerGrowth/GrowthGoal,1),0)
Creator Health Score =IFERROR(AVERAGE(F7:F12),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("top videos", "rising trends", "deals
due this week"), `FILTER`/`SORT`/`UNIQUE`, all wrapped in `IFERROR`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, status color flags. Keep it
premium and consistent — that polish is what makes it feel like creator
software, not a spreadsheet.
