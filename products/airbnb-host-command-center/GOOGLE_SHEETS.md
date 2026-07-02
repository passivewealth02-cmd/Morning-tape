# Airbnb Host Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Property Profile, Calendar, Reservations,
Financial, Pricing, Guests, Cleaning, Maintenance, Inventory, Reviews,
Messages, Suppliers, Taxes, Multi-Property, Goals, Improvements, Analytics,
Settings**.

> Build **Settings** first (business details + dropdown lists), then the
> Calendar, Financial & Reservations engines, then the Dashboard. Add the named
> ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `BizName` (C6), `PrimaryProp` (C7), `NumProps` (C8),
`HomeCurr` (C9), `OccTarget` (C10), `MarginTarget` (C11), `RateTarget` (C12),
`ReportMonth` (C13).

Lists: `PropertyList, PropTypeList, PlatformList, ExpenseCatList, MaintCatList,
SupplierCatList, GoalCatList, CurrencyList, BookStatusList, CleanStatusList,
PayStatusList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `CalCheckIn` | `Calendar!C5:C…` | `FinRevenue` | `Financial!C9` |
| `CalCheckOut` | `Calendar!D5:D…` | `FinNightly` | `Financial!C6` |
| `CalNights` | `Calendar!E5:E…` | `FinExpenses` | `Financial!F19` |
| `CalPlatform` | `Calendar!B5:B…` | `FinNetProfit` | `Financial!I8` |
| `CalProperty` | `Calendar!G5:G…` | `FinMargin` | `Financial!I9` |
| `ResNet` | `Reservations!J5:J…` | `RevOverall` | `Reviews!B5:B…` |
| `CleanStatus` | `Cleaning!D5:D…` | `PropName` | `Multi-Property!B5:B7` |
| `MaintStatus` | `Maintenance!G5:G…` | `PropRevenue` | `Multi-Property!C5:C7` |
| `MaintTask` | `Maintenance!A5:A…` | `GoalProgress` | `Goals!E5:E…` |
| `InvQty` | `Inventory!C5:C…` | `InvReorder` | `Inventory!D5:D…` |
| `InvItem` | `Inventory!B5:B…` | `HealthRange` | `Analytics!C7:C12` |

---

## 3. Financial — the live P&L

```sheets
Total Revenue    =SUM(C6:C8)                 (nightly + cleaning + extras)
Total Expenses   =SUM(F6:F18)                (13 categories)
Net Profit       =FinRevenue-FinExpenses
Profit Margin    =IFERROR(FinNetProfit/FinRevenue,0)
Revenue / Property =IFERROR(FinRevenue/NumProps,0)
```

---

## 4. Dashboard — the 12 KPIs

```sheets
Monthly Revenue     =FinRevenue
Net Profit          =FinNetProfit
Occupancy Rate      =IFERROR(SUM(CalNights)/(COUNTA(PropName)*30),0)
Avg Nightly Rate    =IFERROR(FinNightly/SUM(CalNights),0)
Upcoming Check-Ins  =COUNTIFS(CalCheckIn,">="&TODAY(),CalCheckIn,"<="&TODAY()+7)
Upcoming Check-Outs =COUNTIFS(CalCheckOut,">="&TODAY(),CalCheckOut,"<="&TODAY()+7)
Avg Length of Stay  =IFERROR(SUM(CalNights)/COUNT(CalNights),0)
Guest Rating        =IFERROR(AVERAGE(RevOverall),0)
Cleaning Pending    =COUNTIF(CleanStatus,"Scheduled")
Maintenance Due     =COUNTIF(MaintStatus,"Open")+COUNTIF(MaintStatus,"In Progress")
Low-Stock Items     =SUMPRODUCT((InvQty<=InvReorder)*(InvItem<>""))
Business Health     =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Expense Breakdown (donut), Revenue by Property (column), Occupancy by
Property (bar), Booking Sources (donut), 6-month revenue trend (line). Turn off
auto data labels.

---

## 5. Analytics — Business Health Score

```sheets
Occupancy vs target   =IFERROR(MIN((SUM(CalNights)/(COUNTA(PropName)*30))/OccTarget,1),0)
Profit margin vs tgt  =IFERROR(MIN(FinMargin/MarginTarget,1),0)
Guest satisfaction    =IFERROR(AVERAGE(RevOverall)/5,0)
Maintenance handled   =IFERROR(1-(COUNTIF(MaintStatus,"Open")+COUNTIF(MaintStatus,"In Progress"))/MAX(COUNTA(MaintTask),1),0)
Inventory stocked     =IFERROR(1-SUMPRODUCT((InvQty<=InvReorder)*(InvItem<>""))/MAX(COUNTA(InvItem),1),0)
Goal progress         =IFERROR(AVERAGE(GoalProgress),0)
Health Score          =IFERROR(AVERAGE(C7:C12),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("next 3 check-ins"), `FILTER`/`SORT`,
`UNIQUE`, all wrapped in `IFERROR`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, data-bar rows, mint/gold/
red status flags. Keep it warm, premium and consistent — that polish is what
makes it feel like software, not a spreadsheet.
