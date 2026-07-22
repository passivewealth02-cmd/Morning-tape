# Cafe & Coffee Shop Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Cup Cost, Menu Board, Daypart Sales, Weekly Sales, Labor
& Prime, Bean & Milk, Inventory & Par, Waste Log, Ordering, Cash & Tips,
Regulars, Settings**.

> Build **Cup Cost** and **Menu Board** first (the engine), then Daypart, Weekly,
> Labor and the rest, then the Dashboard. Add the named ranges below (Data ▸
> Named ranges).

---

## 1. Settings — controls & lists

Controls: `Cafe`, `Owner`, `TargetBev` (0.28), `TargetLabor` (0.30), `TicketGoal`
(8), `Currency`.

Lists: `UnitList, CatList, RoleList, YesNoList`.

---

## 2. Cup Cost & Menu Board — the engine

```sheets
Cup Cost   Ext. (F)  =Qty*CostUnit ;  CUP COST =SUM(...)  → LatteCost
Menu Board Bev % (G) =IFERROR(CupCost/Price,0) ;  Margin (H) =Price-CupCost
           Overall beverage % =SUMPRODUCT(MenuCost,MenuUnits)/SUMPRODUCT(MenuPrice,MenuUnits)  → BevCostPct
```

Named: `LatteCost, MenuItem, MenuCost, MenuPrice, MenuUnits, MenuMargin,
BevCostPct`.

---

## 3. Daypart, Labor & Prime cost

```sheets
Daypart   Avg Ticket =Sales/Transactions ; DayTotal / TxnTotal (named totals)
Labor     Labor % (E) =Labor/Sales ; LaborPct = week total labor ÷ sales
Prime     PrimeCost = BevCostPct + LaborPct   (the make-or-break number)
```

Named: `DaypartName, DaypartSales, DayTotal, TxnTotal, WeekSalesTotal,
LaborTotal, LaborPct, PrimeCost, WasteTotal`.

---

## 4. Dashboard — the 12 KPIs

```sheets
Menu Items     =COUNTA(MenuItem)
Avg Cup Cost   =AVERAGE(MenuCost)
Avg Price      =AVERAGE(MenuPrice)
Avg Margin     =AVERAGE(MenuMargin)
Bev Cost       =BevCostPct
Labor Cost     =LaborPct
Daily Sales    =DayTotal
Transactions   =TxnTotal
Avg Ticket     =IFERROR(DayTotal/TxnTotal,0)
Top Daypart    =INDEX(DaypartName,MATCH(MAX(DaypartSales),DaypartSales,0))
Prime Cost     =PrimeCost
Café Score     =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Sales by Daypart (doughnut) from the Daypart sales column.

---

## 5. Café Score (6 dimensions)

```sheets
Beverage cost on target =IFERROR(MIN(TargetBev/BevCostPct,1),0)
Labor on target         =IFERROR(MIN(TargetLabor/LaborPct,1),0)
Prime cost healthy      =IFERROR(MIN(0.62/PrimeCost,1),0)
Avg ticket vs goal      =IFERROR(MIN((DayTotal/TxnTotal)/TicketGoal,1),0)
Margin per cup          =IFERROR(MIN(AVERAGE(MenuMargin)/4,1),0)
Low waste               =IFERROR(1-MIN((WasteTotal/WeekSalesTotal)/0.06,1),0)
Café Score              =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `SUMPRODUCT`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MATCH`,
`INDEX`, `MIN`, `IFERROR`, color scales (beverage %, café score) and conditional
formatting (low inventory).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter) — a cup-cost card, daypart &
weekly logs, a labor sheet & an open/close checklist for the bar. Print any tab:
File ▸ Print ▸ fit to width.

> A business tool, not financial or accounting advice — confirm figures with your
> own books.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
