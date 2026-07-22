# Restaurant Labor & Scheduling Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Labor Cost Calc, Weekly Schedule, Employees, Sales
Forecast, Sales per Labor Hr, Overtime, Roles & Rates, Availability, Prime Cost,
Tips, Labor by Day, Settings**.

> Build the **Weekly Schedule** and **Sales Forecast** first, then **Labor Cost
> Calc** (it pulls their totals), then the rest and the Dashboard. Add the named
> ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Restaurant`, `Manager`, `TargetLaborPct` (0.30), `SPLHGoal` (50),
`OTLimit` (10), `PrimeGoal` (0.60), `WageGoal` (16), `Currency`.

Lists: `RoleList, StatusList, YesNoList`.

---

## 2. Weekly Schedule — the grid

```sheets
Hours   =SUM(Mon:Sun)                       → SchedHours
Cost    =Wage*Hours                         → SchedCost
WEEK TOTAL Hours =SUM(SchedHours)           → TotalHours
WEEK TOTAL Cost  =SUM(SchedCost)            → LaborCost
```

Named: `SchedName, SchedWage, SchedHours, SchedCost, TotalHours, LaborCost`.

---

## 3. Sales Forecast & Labor Cost Calc — the engine

```sheets
Forecast   Labor Target =Forecast*TargetLaborPct ; WEEK TOTAL Sales → ForecastSales
Labor Calc Labor %      =LaborCost/ForecastSales           → LaborPct
           Sales/labor hr =ForecastSales/TotalHours         → SPLH
           Avg wage      =LaborCost/TotalHours              → AvgWage
```

Named: `ForecastDay, ForecastAmt, ForecastSales, LaborTarget, LaborPct, SPLH,
AvgWage`.

---

## 4. Dashboard — the 12 KPIs

```sheets
Labor Cost       =LaborCost
Labor %          =LaborPct
Weekly Sales     =ForecastSales
Total Hours      =TotalHours
Sales/Labor Hr   =SPLH
Employees        =COUNTA(SchedName)
Labor Target     =LaborTarget
Vs Target        =LaborTarget-LaborCost
Avg Wage         =AvgWage
Prime Cost       =PrimeCost
OT Hours         =OTHours
Labor Score      =IFERROR(AVERAGE(HealthRange),0)
```

Overtime adds `OTHrs, OTHours`; Prime Cost adds `FoodPct, PrimeCost` (=FoodPct +
LaborPct); Labor by Day adds `LaborDayAmt` (SUMPRODUCT of wage × day-hours, drives
the chart).

---

## 5. Labor Score (6 dimensions)

```sheets
Labor on target          =IFERROR(MIN(TargetLaborPct/LaborPct,1),0)
Sales per labor hr healthy =IFERROR(MIN(SPLH/SPLHGoal,1),0)
Overtime low             =IFERROR(1-MIN(OTHours/OTLimit,1),0)
Prime cost in check       =IFERROR(MIN(PrimeGoal/PrimeCost,1),0)
Schedule covered          =IFERROR(COUNTIF(SchedHours,">0")/COUNTA(SchedName),0)
Wage in budget            =IFERROR(MIN(WageGoal/AvgWage,1),0)
Labor Score               =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `SUM`, `SUMPRODUCT`, `INDEX`, `COUNTIF`, `COUNTA`, `AVERAGE`,
`MIN`, `IFERROR`, color scales (SPLH, labor score) and conditional formatting.

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter) — a blank weekly schedule, roster,
sales forecast, overtime log, time-off request & a tip sheet for the wall. Print
any tab: File ▸ Print ▸ fit to width.

> A business tool, not financial, legal or HR advice — confirm wages, overtime
> rules and labor law with your own advisors.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
