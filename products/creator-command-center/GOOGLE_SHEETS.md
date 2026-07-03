# Creator Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Brand, Calendar, Pipeline, Ideas, Performance,
Revenue, Sponsorships, Affiliate, Products, Expenses, Deal Calendar, Goals,
Assets, Accounts, Email List, Collabs, Launch Planner, Audience, SEO,
Repurposing, Gallery, Analytics, Settings**.

> Build **Settings** first (business details + dropdown lists), then the
> Revenue/Expense engines, then the Dashboard. Add the named ranges below
> (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `CreatorName` (C6), `BrandName` (C7), `PrimaryPlatform` (C8),
`HomeCurr` (C9), `RevenueGoal` (C10), `ContentGoal` (C11), `DealTarget` (C12),
`MarginTarget` (C13).

Lists: `PlatformList, ContentTypeList, ContentStatusList, RevCatList,
ExpCatList, CampaignList, GoalCatList, SponStageList, PriorityList,
InvoiceStatusList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CalStatus` | `Calendar!G5:G64` | `RevenueTotal` | `Revenue!B15` |
| `CalPlatform` | `Calendar!B5:B64` | `RevSponsor/Affiliate/Digital` | `Revenue!B5/B6/B8` |
| `PerfDate` | `Performance!C5:C44` | `ExpenseTotal` | `Expenses!B16` |
| `PerfPlatform` | `Performance!B5:B44` | `SponStage` | `Sponsorships!G5:G34` |
| `PerfViews` | `Performance!D5:D44` | `PipeProgress` | `Pipeline!F5:F44` |
| `GoalProgress` | `Goals!E5:E10` | `GoalCategory` | `Goals!B5:B10` |
| `AudNow` / `AudPrev` | `Analytics!C22 / C21` | `HealthRange` | `Analytics!C7:C12` |

---

## 3. Revenue & Expenses — the live P&L

```sheets
Revenue % of total   =IFERROR(B5/$B$15,0)
Annual (est.)        =B5*12
Net profit           =RevenueTotal-ExpenseTotal
Profit margin        =IFERROR((RevenueTotal-ExpenseTotal)/RevenueTotal,0)
Annual run-rate      =RevenueTotal*12
```

---

## 4. Dashboard — the 12 KPIs

```sheets
Published (30d)     =COUNTIFS(PerfDate,">="&TODAY()-30)
Scheduled           =COUNTIF(CalStatus,"Scheduled")
Revenue / Month     =RevenueTotal
Sponsorships        =RevSponsor
Affiliate           =RevAffiliate
Digital Products    =RevDigital
Expenses / Month    =ExpenseTotal
Net Profit          =RevenueTotal-ExpenseTotal
Brand Deals Active  =COUNTIF(SponStage,"Signed")+COUNTIF(SponStage,"Delivered")+COUNTIF(SponStage,"Negotiating")
Audience Growth     =IFERROR(AudNow/AudPrev-1,0)
Completion Rate     =IFERROR(MIN(COUNTIFS(PerfDate,">="&TODAY()-30)/ContentGoal,1),0)
Business Health     =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Revenue by Source (donut), Audience Growth (line), Content by Platform
(column, via `COUNTIF(PerfPlatform,…)`), Expense Breakdown (donut). Turn off
auto data labels.

---

## 5. Analytics — Business Health Score

```sheets
Revenue vs goal          =IFERROR(MIN(RevenueTotal/RevenueGoal,1),0)
Profit margin            =IFERROR(MIN(((RevenueTotal-ExpenseTotal)/RevenueTotal)/MarginTarget,1),0)
Publishing consistency   =IFERROR(MIN(COUNTIFS(PerfDate,">="&TODAY()-30)/ContentGoal,1),0)
Content completion       =IFERROR(AVERAGE(PipeProgress),0)
Sponsorship pipeline     =IFERROR(MIN((Signed+Delivered+Negotiating)/DealTarget,1),0)
Goal progress            =IFERROR(AVERAGE(GoalProgress),0)
Health Score             =IFERROR(AVERAGE(C7:C12),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("top 5 by views":
`=QUERY(Performance!A5:D,"select A,D order by D desc limit 5",0)`),
`FILTER`/`SORT`/`UNIQUE`, all wrapped in `IFERROR`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, status heat-maps. Keep it
premium and consistent — that polish is what makes it feel like SaaS.
