# Instagram Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Profile, Calendar, Grid Planner, Pipeline, Ideas,
Captions, Reels, Stories, Hashtags, Analytics, Community, Shop, Affiliate,
Brand Deals, Finance, Expenses, Equipment, Repurposing, Brand Kit, Gallery,
Goals, Audience, Settings**.

> Build **Settings** first (handle + goals + dropdown lists), then the Calendar,
> Analytics, Shop, Brand Deals & Finance, then the Dashboard. Add the named
> ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Handle`, `CreatorName`, `Niche`, `RevenueGoal` (9000),
`PostGoal` (20), `DealTarget` (5), `RetTarget` (0.55), `GrowthGoal` (12000).

Lists: `PillarList, PostTypeList, StatusList, RevCatList, ExpCatList,
CaptionTypeList, GoalCatList, SponStageList, PriorityList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CalDate` | `Calendar!A5:A64` | `RevenueTotal` | `Finance!B12` |
| `CalStatus` | `Calendar!F5:F64` | `ExpenseTotal` | `Finance!G12` |
| `FollowerNow` | `Analytics!C6` | `ShopGMV` | `Shop!F13` |
| `FollowerGrowth` | `Analytics!C7` | `ShopEarn` | `Shop!G13` |
| `Reach28` | `Analytics!C8` | `SponStage` | `Brand Deals!E5:E34` |
| `EngRate` | `Analytics!C9` | `GoalProgress` | `Goals!E5:E11` |
| `Saves28` | `Analytics!C10` | `HealthRange` | `Analytics!F7:F12` |
| `Completion` | `Analytics!C12` | `FollowVal` | `Analytics!C27:C32` |

---

## 3. Shop, Calendar & posting

```sheets
GMV (per product)     =C5*E5                (price × units)
Your earnings         =F5*D5                (GMV × commission %)
Shop total GMV        =SUM(F5:F12)
Posted (28d)          =COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Posted")+COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Boosted")
Posting consistency   =IFERROR(MIN(Posted28/PostGoal,1),0)
```

The Finance "Instagram Shop" line is `=ShopEarn`, so Shop and Finance stay in sync.

---

## 4. Dashboard — the 12 KPIs

```sheets
Followers            =FollowerNow
Reach (28d)          =Reach28
Engagement Rate      =EngRate
Saves (28d)          =Saves28
Posted (28d)         =COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Posted")+COUNTIFS(CalDate,">="&TODAY()-28,CalStatus,"Boosted")
Monthly Revenue      =RevenueTotal
Net Profit           =RevenueTotal-ExpenseTotal
Shop Sales (GMV)     =ShopGMV
Brand Deals          =COUNTIF(SponStage,"Signed")+COUNTIF(SponStage,"Delivered")+COUNTIF(SponStage,"Negotiation")
Follower Growth      =FollowerGrowth
Posting Consistency  =IFERROR(MIN(Posted28/PostGoal,1),0)
Creator Health       =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Follower Growth (line), Revenue by Source (donut), Top Posts by Reach
(bar), Expense Breakdown (donut). Turn off auto data labels.

---

## 5. Analytics — Creator Health Score

```sheets
Revenue vs goal      =IFERROR(MIN(RevenueTotal/RevenueGoal,1),0)
Posting consistency  =IFERROR(MIN(Posted28/PostGoal,1),0)
Engagement           =IFERROR(MIN(EngRate/0.07,1),0)
Reel retention       =IFERROR(MIN(Completion/RetTarget,1),0)
Brand pipeline       =IFERROR(MIN(ActiveDeals/DealTarget,1),0)
Follower growth      =IFERROR(MIN(FollowerGrowth/GrowthGoal,1),0)
Creator Health Score =IFERROR(AVERAGE(F7:F12),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("top posts", "saves by pillar", "deals
due this week"), `FILTER`/`SORT`/`UNIQUE`, all wrapped in `IFERROR`.

Grid Planner & Gallery: click a cell ▸ Insert ▸ Image ▸ **Image in cell**, or
paste `=IMAGE("your-cover-link")`.

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
