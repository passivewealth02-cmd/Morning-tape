# Food Cost & Inventory Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Food Cost Calc, Inventory Count, Purchases Log, Sales Log,
Usage & Variance, Par & Ordering, Vendors, Price Tracker, Menu Costing, Waste Log,
Categories, Settings**.

> Build **Inventory Count, Purchases Log & Sales Log** first, then **Food Cost
> Calc** (the engine pulls their totals), then the rest and the Dashboard. Add the
> named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Business`, `Owner`, `TargetFC` (0.30), `VarLimit` (0.05), `TurnGoal`
(1.5), `WasteLimit` (0.03), `Currency`.

Lists: `UnitList, CatList, StatusList, YesNoList`.

---

## 2. Inventory Count, Purchases & Sales — the inputs

```sheets
Inventory   Ext. Value  =Count*UnitCost ;  INVENTORY VALUE =SUM(...)  → InvValue
Purchases   PURCHASES TOTAL =SUM(PurchAmt)                            → PurchTotal
Sales       SALES TOTAL     =SUM(SalesAmt)                            → SalesTotal
```

Named: `InvItem, InvExt, InvValue, PurchAmt, PurchTotal, SalesAmt, SalesTotal`.

---

## 3. Food Cost Calc — the engine

```sheets
Beginning inventory          (input)                     → Beginning
+ Purchases (from log)        =PurchTotal
− Ending inventory (count)    =InvValue
= Food used (COGS)            =Beginning+PurchTotal−InvValue  → FoodCostDollar
÷ Food sales (from log)       =SalesTotal
= FOOD COST %                 =FoodCostDollar/SalesTotal       → FoodCostPct
Avg inventory                 =(Beginning+InvValue)/2          → AvgInv
Inventory turns               =FoodCostDollar/AvgInv           → Turns
```

---

## 4. Dashboard — the 12 KPIs

```sheets
Food Cost %      =FoodCostPct
Inventory Value  =InvValue
Purchases        =PurchTotal
Food Used        =FoodCostDollar
Sales            =SalesTotal
Top Category     =INDEX(CatName,MATCH(MAX(CatActual),CatActual,0))
Variance         =VariancePct
Items Tracked    =COUNTA(InvItem)
To Order         =ToOrderValue
Inv Turns        =Turns
Vendors          =COUNTA(VendorName)
Inventory Score  =IFERROR(AVERAGE(HealthRange),0)
```

Usage & Variance adds `CatName, CatActual, VariancePct`; Par & Ordering adds
`ParItem, ParLevel, OrderVal, ToOrderValue`; Categories adds `CatLabel, CatSpend`
(drives the by-category chart); Vendors adds `VendorName`; Waste adds `WasteTotal`.

---

## 5. Inventory Score (6 dimensions)

```sheets
Food cost on target      =IFERROR(MIN(TargetFC/FoodCostPct,1),0)
Inventory variance low    =IFERROR(1-MIN(VariancePct/VarLimit,1),0)
Inventory counted         =IFERROR(COUNTIF(InvExt,">0")/COUNTA(InvItem),0)
Pars set                  =IFERROR(COUNTIF(ParLevel,">0")/COUNTA(ParItem),0)
Inventory turns healthy   =IFERROR(MIN(Turns/TurnGoal,1),0)
Gross margin              =IFERROR(MIN((1-FoodCostPct)/0.7,1),0)
Inventory Score           =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `INDEX`, `MATCH`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`,
`IFERROR`, color scales (food %, score), data bars (category spend) and
conditional formatting (high variance, below-par stock).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter) — an inventory count sheet,
food-cost worksheet, order guide, vendor list & a weekly count checklist. Print
any tab: File ▸ Print ▸ fit to width.

> A business tool, not financial or accounting advice — confirm figures with your
> own books.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
