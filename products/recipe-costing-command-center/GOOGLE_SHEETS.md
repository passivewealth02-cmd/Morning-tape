# Recipe Costing & Menu Engineering Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Ingredients, Recipe Costing, Menu Items, Menu
Engineering, Price Calculator, Sales Mix, Portion & Yield, Specials & LTO, Batch
& Prep, Vendor Prices, Waste Log, Settings**.

> Build **Ingredients** and **Recipe Costing** first (the engine), then **Menu
> Items** (which drives everything), then the rest and the Dashboard. Add the
> named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Restaurant`, `Owner`, `TargetFC` (0.30), `MarginGoal` (12),
`Currency`.

Lists: `UnitList, StationList, CatList, YesNoList`.

---

## 2. Ingredients & Recipe Costing — the engine

```sheets
Ingredients   Cost/Unit (F)  =IFERROR(PackPrice/PackSize,0)     → named IngCostUnit
Recipe Cost   Ext. Cost (F)  =Qty*CostUnit
              Total          =SUM(ext costs)
              Cost/Serving   =IFERROR(Total/Yield,0)            → named BurgerCost
```

---

## 3. Menu Items — food cost %, margin, class

```sheets
Food Cost % (G)  =IFERROR(PlateCost/Price,0)
Margin $ (H)     =Price-PlateCost
Class (I)        =IF(AND(Units>=AvgUnits,Margin>=AvgMargin),"Star",
                   IF(Units>=AvgUnits,"Plowhorse",
                   IF(Margin>=AvgMargin,"Puzzle","Dog")))
```

`AvgUnits =AVERAGE(MenuUnits)` and `AvgMargin =AVERAGE(MenuMargin)` are the
menu's own dividing lines. The flagship item's cost links to the engine:
`PlateCost =BurgerCost`.

Named: `MenuItem, MenuCost, MenuPrice, MenuUnits, MenuFC, MenuMargin, MenuClass,
AvgUnits, AvgMargin, AvgFC, AvgCost, AvgPrice`.

---

## 4. Dashboard — the 12 KPIs

```sheets
Menu Items       =COUNTA(MenuItem)
Avg Food Cost    =AvgFC
Target Food Cost =TargetFC
Avg Plate Cost   =AvgCost
Avg Menu Price   =AvgPrice
Avg Margin       =AvgMargin
Stars            =COUNTIF(MenuClass,"Star")
Plowhorses       =COUNTIF(MenuClass,"Plowhorse")
Puzzles          =COUNTIF(MenuClass,"Puzzle")
Dogs             =COUNTIF(MenuClass,"Dog")
Top Margin       =INDEX(MenuItem,MATCH(MAX(MenuMargin),MenuMargin,0))
Menu Score       =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Menu Mix by Class (doughnut) from the Menu Engineering counts.

---

## 5. Menu Score (6 dimensions)

```sheets
Food cost on target      =IFERROR(MIN(TargetFC/AvgFC,1),0)
Margin vs goal           =IFERROR(MIN(AvgMargin/MarginGoal,1),0)
Menu fully costed        =IFERROR(COUNTIF(MenuCost,">0")/COUNTA(MenuItem),0)
Items above margin goal  =IFERROR(COUNTIF(MenuMargin,">="&MarginGoal)/COUNTA(MenuItem),0)
Menu balance (few dogs)  =IFERROR(1-COUNTIF(MenuClass,"Dog")/COUNTA(MenuItem),0)
Contribution margin      =IFERROR(MIN((1-AvgFC)/0.7,1),0)
Menu Score               =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `SUMIF`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`, `INDEX`,
`MATCH`, `IFERROR`, color scales (food cost %, menu score) and conditional
formatting (class colors, vendor price creep).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter) — recipe cost cards, a
menu-engineering worksheet & a weekly food-cost tracker for the line. Print any
tab: File ▸ Print ▸ fit to width.

> A business tool, not financial or accounting advice — confirm figures with your
> own books.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
