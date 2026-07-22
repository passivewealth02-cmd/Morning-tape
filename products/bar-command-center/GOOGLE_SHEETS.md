# Bar & Pub Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Pour Cost, Drink Menu, Keg & Draft, Inventory Variance,
Liquor Inventory, Happy Hour, Weekly Sales, Waste & Spill, Ordering, Cash & Tips,
Events & Tabs, Settings**.

> Build **Pour Cost** and **Drink Menu** first (the engine), then Keg & Draft,
> Inventory Variance, Happy Hour and the rest, then the Dashboard. Add the named
> ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Bar`, `Owner`, `TargetPour` (0.22), `MarginGoal` (6), `VarLimit`
(0.10), `Currency`.

Lists: `UnitList, CatList, StatusList, YesNoList`.

---

## 2. Pour Cost & Drink Menu — the engine

```sheets
Pour Cost   Ext. (F)   =Qty*CostUnit ;  POUR COST =SUM(...)  → OldFashionedCost
Drink Menu  Pour %     =PourCost/Price
            Margin     =Price-PourCost
            Wk Rev (I) =Price*Units
            PourCostPct = SUMPRODUCT(DrinkCost,DrinkUnits)/SUMPRODUCT(DrinkPrice,DrinkUnits)
```

Named: `OldFashionedCost, DrinkItem, DrinkCost, DrinkPrice, DrinkUnits,
DrinkMargin, DrinkRev`, plus `WeeklySales` and `PourCostPct` (weighted).

---

## 3. Keg & Draft, Inventory Variance & Happy Hour

```sheets
Keg & Draft   Cost/pint  =KegCost/Pints
              Profit/keg =(PricePint-CostPint)*Pints  → AvgKegProfit
Variance      Variance % =(Actual-Theoretical)/Theoretical → VariancePct
Happy Hour    HH margin  =HHPrice-PourCost ; HHAvgMargin = AVERAGE(HH margin %)
```

Named: `KegProfit, AvgKegProfit, VariancePct, HHAvgMargin, WasteTotal`.

---

## 4. Dashboard — the 12 KPIs

```sheets
Drinks           =COUNTA(DrinkItem)
Avg Pour Cost    =AVERAGE(DrinkCost)
Avg Price        =AVERAGE(DrinkPrice)
Avg Margin       =AVERAGE(DrinkMargin)
Pour Cost %      =PourCostPct
Top Seller       =INDEX(DrinkItem,MATCH(MAX(DrinkRev),DrinkRev,0))
Weekly Sales     =WeeklySales
Weekly Units     =SUM(DrinkUnits)
Avg Profit/Keg   =AvgKegProfit
Inv Variance     =VariancePct
Happy-Hr Margin  =HHAvgMargin
Bar Score        =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Sales by Day (column) from the Weekly Sales tab.

---

## 5. Bar Score (6 dimensions)

```sheets
Pour cost on target     =IFERROR(MIN(TargetPour/PourCostPct,1),0)
Margin per drink        =IFERROR(MIN(AVERAGE(DrinkMargin)/MarginGoal,1),0)
Menu fully costed       =IFERROR(COUNTIF(DrinkCost,">0")/COUNTA(DrinkItem),0)
Inventory variance low  =IFERROR(1-MIN(VariancePct/VarLimit,1),0)
Happy-hour margin       =IFERROR(MIN(HHAvgMargin/0.5,1),0)
Gross margin            =IFERROR(MIN((1-PourCostPct)/0.75,1),0)
Bar Score               =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `SUMPRODUCT`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MATCH`,
`INDEX`, `MIN`, `IFERROR`, color scales (pour cost %, bar score), data bars
(keg profit) and conditional formatting (variance, low inventory).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter) — a pour-cost card, keg & draft
log, inventory variance sheet, liquor count sheet & an open/close checklist for
behind the bar. Print any tab: File ▸ Print ▸ fit to width.

> A business tool, not financial or accounting advice — confirm figures with your
> own books.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
