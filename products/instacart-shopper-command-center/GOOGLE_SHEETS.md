# Instacart Shopper Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Shopper Profile, Batch Log, Earnings, Mileage,
Fuel Log, Vehicle, Expenses, Budget, Tax Center, Savings, Boosts, Stores,
Ratings, Analytics, Planner, Gallery, Settings**.

> Build **Settings** first (goals + dropdown lists), then the Batch Log
> (the earnings engine), then Expenses, Savings, Analytics & the Tax Center,
> then the Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `DriverName`, `VehicleName`, `HomeCity`, `NetGoal` (2000),
`HourlyTarget` (17), `BatchGoal` (90), `ShiftGoal` (16), `SavingsGoal` (4000),
`TaxReserveGoal` (900), `MileageRate` (0.70), `TaxRate` (0.25).

Lists: `PlatformList, ExpCatList, BudgetCatList, MaintList, GoalCatList,
PriorityList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `ShiftPlatform` | `'Batch Log'!B5:B64` | `GrossEarn` | `'Batch Log'!I65` |
| `ShiftHours` | `'Batch Log'!C5:C64` | `TotalHours` | `'Batch Log'!C65` |
| `ShiftBatches` | `'Batch Log'!D5:D64` | `TotalBatches` | `'Batch Log'!D65` |
| `ShiftItems` | `'Batch Log'!E5:E64` | `TotalItems` | `'Batch Log'!E65` |
| `ShiftTips` | `'Batch Log'!G5:G64` | `TotalTips` | `'Batch Log'!G65` |
| `ShiftEarn` | `'Batch Log'!I5:I64` | `TotalMiles` | `'Batch Log'!J65` |
| `ShiftMiles` | `'Batch Log'!J5:J64` | `FuelTotal` | `'Batch Log'!K65` |
| `ExpenseTotal` | `Expenses!B14` | `NetEarn` | `Analytics!C8` |
| `PerHour` | `Analytics!C10` | `PerMile` | `Analytics!C11` |
| `SavingsSaved` | `Savings!D7` | `TaxSetAside` | `Savings!D8` |
| `HealthRange` | `Analytics!F7:F12` | `TrendVal` | `Analytics!C27:C32` |

---

## 3. Batch Log — the earnings engine

```sheets
Shift earnings        =F5+G5+H5              (batch pay + tips + boost)
Gross earnings (mo)   =SUM(I5:I64)
Total miles           =SUM(J5:J64)
Total hours           =SUM(C5:C64)
Total items           =SUM(E5:E64)
Gas total             =SUM(K5:K64)
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
Batches          =TotalBatches
Expenses         =ExpenseTotal
Tax Deduction    =TotalMiles*MileageRate
Savings          =IFERROR(MIN(SavingsSaved/SavingsGoal,1),0)
Shopper Health   =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Net Earnings (line), Earnings Mix (donut), Best Shifts (bar), Expense
Breakdown (donut). Turn off auto data labels.

---

## 5. Analytics — Shopper Health Score

```sheets
Net earnings vs goal   =IFERROR(MIN(NetEarn/NetGoal,1),0)
Net $/hour vs target   =IFERROR(MIN(PerHour/HourlyTarget,1),0)
Batches vs goal        =IFERROR(MIN(TotalBatches/BatchGoal,1),0)
Emergency fund         =IFERROR(MIN(SavingsSaved/SavingsGoal,1),0)
Shopping consistency   =IFERROR(MIN(COUNT(ShiftMiles)/ShiftGoal,1),0)
Tax reserve            =IFERROR(MIN(TaxSetAside/TaxReserveGoal,1),0)
Shopper Health Score   =IFERROR(AVERAGE(F7:F12),0)
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

> Because grocery shoppers drive fewer miles per dollar, the **actual-expense**
> method often beats the standard mileage deduction — the Tax Center picks the
> bigger one for you.

Power features: `ARRAYFORMULA`, `QUERY` ("best shifts", "batches by store", "boosts
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
premium and consistent — that polish is what makes it feel like shopper software,
not a spreadsheet.
