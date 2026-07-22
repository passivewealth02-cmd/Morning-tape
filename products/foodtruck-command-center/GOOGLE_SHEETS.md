# Food Truck Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Menu & Cost, Events, Break-Even, Daily Sales, Commissary
& Overhead, Inventory & Par, Fuel & Mileage, Permits, Supplies, Bookings, Cash &
Tips, Settings**.

> Build **Events** and **Commissary & Overhead** first (they feed break-even &
> the dashboard), then the rest and the Dashboard. Add the named ranges below
> (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Truck`, `Owner`, `TargetFC` (0.30), `EventGoal` (10), `MarginGoal`
(0.35), `Currency`.

Lists: `UnitList, EvTypeList, PermitList, YesNoList`.

---

## 2. Events — the P&L engine

```sheets
Net Profit (I)  =Sales-Food-Fuel-Fee-Staff
Food % (J)      =IFERROR(Food/Sales,0)
Totals          =SUM(...)  → TotalSales, TotalFood, TotalProfit
```

Named: `EvName, EvSales, EvFood, EvNet`, plus totals `TotalSales, TotalFood,
TotalProfit`.

---

## 3. Break-Even

```sheets
Break-even events / month  =IFERROR(OverheadTotal/AVERAGE(EvNet),0)   → BreakEvenEvents
Break-even sales / month   =IFERROR(OverheadTotal/(1-TotalFood/TotalSales-0.32),0)
```

`OverheadTotal` is the sum of your Commissary & Overhead fixed costs.

---

## 4. Dashboard — the 12 KPIs

```sheets
Events            =COUNTA(EvName)
Total Sales       =TotalSales
Total Profit      =TotalProfit
Avg Profit/Event  =AVERAGE(EvNet)
Food Cost         =IFERROR(TotalFood/TotalSales,0)
Avg Sales/Event   =AVERAGE(EvSales)
Break-Even        =BreakEvenEvents
Overhead/Mo       =OverheadTotal
Top Event         =INDEX(EvName,MATCH(MAX(EvNet),EvNet,0))
Best Sales        =INDEX(EvName,MATCH(MAX(EvSales),EvSales,0))
Permits OK        =COUNTIF(PermitStatus,"OK")&"/"&COUNTA(PermitStatus)
Truck Score       =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Net Profit by Event (bar) from the Events net column.

---

## 5. Truck Score (6 dimensions)

```sheets
Profit margin        =IFERROR(MIN((TotalProfit/TotalSales)/MarginGoal,1),0)
Food cost on target  =IFERROR(MIN(TargetFC/(TotalFood/TotalSales),1),0)
Events booked        =IFERROR(MIN(COUNTA(EvName)/EventGoal,1),0)
Permits current      =IFERROR(COUNTIF(PermitStatus,"OK")/COUNTA(PermitStatus),0)
Profitable events    =IFERROR(COUNTIF(EvNet,">0")/COUNTA(EvName),0)
Overhead covered     =IFERROR(MIN(TotalProfit/OverheadTotal,1),0)
Truck Score          =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `SUMIF`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MATCH`, `INDEX`,
`MIN`, `IFERROR`, data bars (net profit), color scales (food cost, truck score)
and conditional formatting (permit status, booking status, low inventory).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter) — an event P&L sheet, a
break-even worksheet, a permit tracker & a monthly P&L for the truck. Print any
tab: File ▸ Print ▸ fit to width.

> A business tool, not financial, tax or legal advice — confirm figures & permit
> rules with the proper authorities.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
