# Uber Eats Driver Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Driver Profile, Delivery Log, Earnings, Mileage,
Fuel Log, Vehicle, Expenses, Budget, Tax Center, Savings, Promos, Hotspots,
Ratings, Analytics, Planner, Gallery, Settings**.

> Build **Settings** first (goals + dropdown lists), then the Delivery Log
> (the earnings engine), then Expenses, Savings, Analytics & the Tax Center,
> then the Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `DriverName`, `VehicleName`, `HomeCity`, `NetGoal` (1800),
`HourlyTarget` (16), `DeliveryGoal` (300), `ShiftGoal` (16), `SavingsGoal` (4000),
`TaxReserveGoal` (1000), `MileageRate` (0.70), `TaxRate` (0.25).

Lists: `PlatformList, ExpCatList, BudgetCatList, MaintList, GoalCatList,
PriorityList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `ShiftPlatform` | `'Delivery Log'!B5:B64` | `GrossEarn` | `'Delivery Log'!H65` |
| `ShiftHours` | `'Delivery Log'!C5:C64` | `TotalHours` | `'Delivery Log'!C65` |
| `ShiftOrders` | `'Delivery Log'!D5:D64` | `TotalOrders` | `'Delivery Log'!D65` |
| `ShiftTips` | `'Delivery Log'!F5:F64` | `TotalTips` | `'Delivery Log'!F65` |
| `ShiftEarn` | `'Delivery Log'!H5:H64` | `TotalMiles` | `'Delivery Log'!I65` |
| `ShiftMiles` | `'Delivery Log'!I5:I64` | `FuelTotal` | `'Delivery Log'!J65` |
| `ExpenseTotal` | `Expenses!B14` | `NetEarn` | `Analytics!C8` |
| `PerHour` | `Analytics!C10` | `PerMile` | `Analytics!C11` |
| `SavingsSaved` | `Savings!D7` | `TaxSetAside` | `Savings!D8` |
| `HealthRange` | `Analytics!F7:F12` | `TrendVal` | `Analytics!C27:C32` |

---

## 3. Delivery Log — the earnings engine

```sheets
Shift earnings        =E5+F5+G5              (base + tips + promo)
Gross earnings (mo)   =SUM(H5:H64)
Total miles           =SUM(I5:I64)
Total hours           =SUM(C5:C64)
Gas total             =SUM(J5:J64)
```

The Expenses "Fuel" line is `=FuelTotal`, so the log and expenses stay in sync —
no double entry.

---

## 4. Dashboard — the 12 KPIs

```sheets
Net Earnings     =GrossEarn-ExpenseTotal
Gross Earnings   =GrossEarn
Tips             =SUM(ShiftTips)
Net $ / Hour     =IFERROR((GrossEarn-ExpenseTotal)/TotalHours,0)
Net $ / Mile     =IFERROR((GrossEarn-ExpenseTotal)/TotalMiles,0)
Miles Driven     =TotalMiles
Online Hours     =TotalHours
Deliveries       =TotalOrders
Expenses         =ExpenseTotal
Tax Deduction    =TotalMiles*MileageRate
Savings          =IFERROR(MIN(SavingsSaved/SavingsGoal,1),0)
Driver Health    =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Net Earnings (line), Earnings Mix (donut), Best Shifts (bar), Expense
Breakdown (donut). Turn off auto data labels.

---

## 5. Analytics — Driver Health Score

```sheets
Net earnings vs goal   =IFERROR(MIN(NetEarn/NetGoal,1),0)
Net $/hour vs target   =IFERROR(MIN(PerHour/HourlyTarget,1),0)
Deliveries vs goal     =IFERROR(MIN(TotalOrders/DeliveryGoal,1),0)
Emergency fund         =IFERROR(MIN(SavingsSaved/SavingsGoal,1),0)
Driving consistency    =IFERROR(MIN(COUNT(ShiftMiles)/ShiftGoal,1),0)
Tax reserve            =IFERROR(MIN(TaxSetAside/TaxReserveGoal,1),0)
Driver Health Score    =IFERROR(AVERAGE(F7:F12),0)
```

## 6. Tax Center — mileage vs actual

```sheets
Mileage deduction   =TotalMiles*MileageRate
Better method       =IF(TotalMiles*MileageRate>=ExpenseTotal,"Mileage","Actual")
Deduction (best)    =MAX(TotalMiles*MileageRate,ExpenseTotal)
Taxable (est.)      =MAX(GrossEarn-TotalMiles*MileageRate,0)
Set aside this month=MAX(GrossEarn-TotalMiles*MileageRate,0)*TaxRate
Quarterly estimate  =MAX(GrossEarn-TotalMiles*MileageRate,0)*TaxRate*3
```

Power features: `ARRAYFORMULA`, `QUERY` ("best shifts", "miles by day", "promos
due this week"), `SUMIF` by app, `FILTER`/`SORT`, all wrapped in `IFERROR`.

Receipts & Vehicle Gallery: click a cell ▸ Insert ▸ Image ▸ **Image in cell**,
or paste `=IMAGE("your-receipt-link")`.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, status color flags. Keep it
premium and consistent — that polish is what makes it feel like driver software,
not a spreadsheet.
