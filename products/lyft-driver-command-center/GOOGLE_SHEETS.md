# Lyft Driver Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Driver Profile, Shift Log, Earnings, Mileage,
Fuel Log, Vehicle, Expenses, Budget, Tax Center, Savings, Bonuses, Hot Zones,
Ratings, Analytics, Planner, Gallery, Settings**.

> Build **Settings** first (goals + dropdown lists), then the Shift Log
> (the earnings engine), then Expenses, Savings, Analytics & the Tax Center,
> then the Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `DriverName`, `VehicleName`, `HomeCity`, `NetGoal` (3000),
`HourlyTarget` (20), `TripsGoal` (300), `ShiftGoal` (16), `SavingsGoal` (6000),
`TaxReserveGoal` (1800), `MileageRate` (0.70), `TaxRate` (0.25).

Lists: `PlatformList, ExpCatList, BudgetCatList, MaintList, GoalCatList,
PriorityList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `ShiftPlatform` | `'Shift Log'!B5:B64` | `GrossEarn` | `'Shift Log'!H65` |
| `ShiftHours` | `'Shift Log'!C5:C64` | `TotalHours` | `'Shift Log'!C65` |
| `ShiftTrips` | `'Shift Log'!D5:D64` | `TotalTrips` | `'Shift Log'!D65` |
| `ShiftTips` | `'Shift Log'!F5:F64` | `TotalTips` | `'Shift Log'!F65` |
| `ShiftEarn` | `'Shift Log'!H5:H64` | `TotalMiles` | `'Shift Log'!I65` |
| `ShiftMiles` | `'Shift Log'!I5:I64` | `FuelTotal` | `'Shift Log'!J65` |
| `ExpenseTotal` | `Expenses!B14` | `NetEarn` | `Analytics!C8` |
| `PerHour` | `Analytics!C10` | `PerMile` | `Analytics!C11` |
| `SavingsSaved` | `Savings!D7` | `TaxSetAside` | `Savings!D8` |
| `HealthRange` | `Analytics!F7:F12` | `TrendVal` | `Analytics!C27:C32` |

---

## 3. Shift Log — the earnings engine

```sheets
Shift earnings        =E5+F5+G5              (fares + tips + bonus)
Gross earnings (mo)   =SUM(H5:H64)
Total miles           =SUM(I5:I64)
Total hours           =SUM(C5:C64)
Fuel total            =SUM(J5:J64)
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
Trips            =TotalTrips
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
Trips vs goal          =IFERROR(MIN(TotalTrips/TripsGoal,1),0)
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

Power features: `ARRAYFORMULA`, `QUERY` ("best shifts", "miles by day", "bonuses
due this week"), `SUMIF` by platform, `FILTER`/`SORT`, all wrapped in `IFERROR`.

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
