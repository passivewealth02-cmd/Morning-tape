# Restaurant Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Profile, Menu, Inventory, Par Levels, Suppliers,
Sales, Labor, Staff, P&L, Expenses, Cash & Tips, Reservations, Marketing,
Reviews, Compliance, Checklists, Waste, Payments, Training, Analytics,
Settings**.

> Build **Settings** first (restaurant details + cost targets + dropdown lists),
> then Menu, Inventory, Sales, Labor, Expenses & P&L, then the Dashboard +
> Analytics. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `RestName`, `Concept`, `Seats` (84), `FoodTarget` (0.30),
`LaborTarget` (0.32), `PrimeTarget` (0.60), `MarginTarget` (0.15),
`CheckGoal` (32).

Lists: `MenuCatList, ExpCatList, InvCatList, UnitList, StationList, ShiftList,
DaypartList, ResStatusList, PayStatusList, PromoStatusList, ComplyStatusList,
PriorityList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `RevenueTotal` | `Sales!D35` (month total) | `LaborTotal` | `Labor!F14` |
| `CoversTotal` | `Sales!C35` | `OpexTotal` | `Expenses!B16` |
| `FoodCostTotal` | `P&L!C8` (COGS input) | `InvValue` | `Inventory!G32` |
| `InvStatus` | `Inventory!H5:H31` | `MenuFCPct` | `Menu!E5:E34` |
| `AvgRating` | `Reviews!C12` | `HealthRange` | `Analytics!C7:C12` |
| `WeekSales` | `Analytics!C17:C20` | `WeekLabel` | `Analytics!B17:B20` |

---

## 3. Menu, Inventory & Sales

```sheets
Food Cost % (per item)  =IFERROR(D5/C5,0)          (cost ÷ price)
Margin (per item)       =C5-D5
Inventory value (row)   =D5*F5                       (on-hand × unit cost)
Low-stock flag          =IF(D5<E5,"Low","OK")
Avg check (per day)     =IFERROR(D5/C5,0)            (net sales ÷ covers)
Month revenue           =SUM(D5:D34)
Month covers            =SUM(C5:C34)
```

---

## 4. Dashboard — the 12 KPIs

```sheets
Monthly Revenue   =RevenueTotal
Covers            =CoversTotal
Avg Check         =IFERROR(RevenueTotal/CoversTotal,0)
Food Cost %       =IFERROR(FoodCostTotal/RevenueTotal,0)
Labor Cost %      =IFERROR(LaborTotal/RevenueTotal,0)
Prime Cost %      =IFERROR((FoodCostTotal+LaborTotal)/RevenueTotal,0)
Net Profit        =RevenueTotal-FoodCostTotal-LaborTotal-OpexTotal
Profit Margin     =IFERROR((RevenueTotal-FoodCostTotal-LaborTotal-OpexTotal)/RevenueTotal,0)
Inventory Value   =InvValue
Low-Stock Items   =COUNTIF(InvStatus,"Low")
Avg Rating        =AvgRating
Health Score      =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Cost Structure (donut: food/labor/overhead/profit), Weekly Sales
(column), Operations Health (bar), Prime Cost (donut). Turn off auto data labels.

---

## 5. Analytics — Operations Health Score

```sheets
Food cost control    =IFERROR(MIN(FoodTarget/(FoodCostTotal/RevenueTotal),1),0)
Labor cost control   =IFERROR(MIN(LaborTarget/(LaborTotal/RevenueTotal),1),0)
Prime cost control   =IFERROR(MIN(PrimeTarget/((FoodCostTotal+LaborTotal)/RevenueTotal),1),0)
Profit margin        =IFERROR(MIN(((RevenueTotal-FoodCostTotal-LaborTotal-OpexTotal)/RevenueTotal)/MarginTarget,1),0)
Guest rating         =IFERROR(AvgRating/5,0)
Inventory in-stock   =IFERROR(COUNTIF(InvStatus,"OK")/COUNTA(InvStatus),0)
Health Score         =IFERROR(AVERAGE(C7:C12),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("sales by daypart", "items below par",
"AP due this week"), `FILTER`/`SORT`/`UNIQUE`, all wrapped in `IFERROR`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, status color flags. Keep it
premium and consistent — that polish is what makes it feel like restaurant
software, not a spreadsheet.
