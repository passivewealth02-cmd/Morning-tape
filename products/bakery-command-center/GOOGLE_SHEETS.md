# Bakery Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Recipe Costing, Product List, Pre-Orders, Wholesale,
Production Plan, Inventory & Par, Waste Log, Sales Log, Ordering, Cash & Deposits,
Market Days, Settings**.

> Build **Recipe Costing** and **Product List** first (the engine), then
> Pre-Orders, Wholesale and the rest, then the Dashboard. Add the named ranges
> below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Bakery`, `Owner`, `TargetFC` (0.30), `MarginGoal` (3), `PreOrderGoal`
(7), `Currency`.

Lists: `UnitList, CatList, StatusList, YesNoList`.

---

## 2. Recipe Costing & Product List — the engine

```sheets
Recipe   Ext. (F)   =Qty*CostUnit ;  BATCH COST =SUM(...)  → SourdoughBatch
         Cost/unit  =BatchCost/Yield
Product  Unit Cost  =BatchCost/Yield
         Retail Margin =Retail-UnitCost
         Food %     =UnitCost/Retail
         Wk Rev (K) =Retail*Units
```

Named: `SourdoughBatch, ProdItem, ProdUnitCost, ProdRetail, ProdMargin,
ProdWkUnits, ProdRev`, plus `RetailRev` and `FoodCostPct` (weighted).

---

## 3. Pre-Orders & Wholesale

```sheets
Pre-Orders  PreOrderValue = SUM(PreOrderPrice)
Wholesale   Weekly Rev = Qty*UnitPrice ; WholeRev = SUM(weekly rev)
```

Named: `PreOrderName, PreOrderPrice, PreOrderStatus, PreOrderValue, WholeRev,
WasteTotal`.

---

## 4. Dashboard — the 12 KPIs

```sheets
Products       =COUNTA(ProdItem)
Avg Unit Cost  =AVERAGE(ProdUnitCost)
Avg Retail     =AVERAGE(ProdRetail)
Avg Margin     =AVERAGE(ProdMargin)
Food Cost      =FoodCostPct
Top Seller     =INDEX(ProdItem,MATCH(MAX(ProdRev),ProdRev,0))
Weekly Revenue =RetailRev+WholeRev
Weekly Units   =SUM(ProdWkUnits)
Pre-Orders     =COUNTA(PreOrderName)
Wholesale Rev  =WholeRev
Waste %        =IFERROR(WasteTotal/RetailRev,0)
Bakery Score   =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Sales by Day (column) from the Sales Log.

---

## 5. Bakery Score (6 dimensions)

```sheets
Food cost on target   =IFERROR(MIN(TargetFC/FoodCostPct,1),0)
Margin per unit       =IFERROR(MIN(AVERAGE(ProdMargin)/MarginGoal,1),0)
Products fully costed  =IFERROR(COUNTIF(ProdUnitCost,">0")/COUNTA(ProdItem),0)
Pre-orders vs goal     =IFERROR(MIN(COUNTA(PreOrderName)/PreOrderGoal,1),0)
Low waste              =IFERROR(1-MIN((WasteTotal/RetailRev)/0.08,1),0)
Gross margin           =IFERROR(MIN((1-FoodCostPct)/0.7,1),0)
Bakery Score           =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `SUMPRODUCT`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MATCH`,
`INDEX`, `MIN`, `IFERROR`, color scales (food cost %, bakery score), data bars
(wholesale revenue) and conditional formatting (order status, low inventory).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter) — a recipe cost card, product
price list, pre-order form & a bake-day checklist for the kitchen. Print any tab:
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
