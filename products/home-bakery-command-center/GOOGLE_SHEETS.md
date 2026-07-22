# Home Bakery & Cottage Food Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Price It Right, Recipe Costing, Product List, Custom
Orders, Ingredient Costs, Labeling & Allergens, Markets & Events, Income &
Expenses, Waste Log, Customers, Monthly Summary, Settings**.

> Build the **Product List** and **Price It Right** first, then the rest and the
> Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Bakery`, `Owner`, `TargetFC` (0.25), `MarginGoal` (0.70), `WageGoal`
(20), `ProfitGoal` (1500), `WasteLimit` (0.05), `Currency`.

Lists: `UnitList, StatusList, AllergenList, YesNoList`.

---

## 2. Price It Right — the pay-yourself engine

```sheets
Your labor    =Minutes/60 × Your rate
TRUE COST     =Ingredients + Packaging + Labor + Overhead   → TrueCost
Profit        =Price − TrueCost
YOUR HOURLY   =(Price − Ingredients − Packaging − Overhead) ÷ (Minutes/60) → EffHourly
```

---

## 3. Recipe Costing & Product List

```sheets
Recipe    BATCH COST =SUM(ingredients)  → BatchCost ; Cost/unit =BatchCost/Yield
Product   Margin     =(Price − Ing. cost) ÷ Price
          Revenue    =Price × Units/mo
```

Named: `BatchCost, ProdItem, ProdCost, ProdPrice, ProdMargin, ProdUnits, ProdRev`,
plus `MonthlyIncome` (=SUM(ProdRev)) and `FoodCostPct` (weighted).

---

## 4. Orders, Markets & Ledger

```sheets
Custom Orders  Open order value =SUMIF(OrderStatus,"<>Delivered",OrderPrice) → OrderValue
Income & Exp   Net profit =MonthlyIncome − SUM(ExpAmt)  → NetProfit
```

Named: `OrderCust, OrderPrice, OrderStatus, OrderValue, TotalIncome, ExpAmt,
ExpTotal, NetProfit, WasteTotal, MonthProfit`.

---

## 5. Dashboard — the 12 KPIs

```sheets
Products       =COUNTA(ProdItem)
Avg Price      =AVERAGE(ProdPrice)
Food Cost      =FoodCostPct
Avg Margin     =AVERAGE(ProdMargin)
Monthly Income =MonthlyIncome
Top Seller     =INDEX(ProdItem,MATCH(MAX(ProdRev),ProdRev,0))
Monthly Profit =NetProfit
Your Hourly    =EffHourly
Open Orders    =COUNTIF(OrderStatus,"<>Delivered")
Order Value    =OrderValue
Monthly Units  =SUM(ProdUnits)
Bakery Score   =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Profit by Month (column) from the Monthly Summary tab.

---

## 6. Bakery Score (6 dimensions)

```sheets
Food cost on target  =IFERROR(MIN(TargetFC/FoodCostPct,1),0)
Margins healthy      =IFERROR(MIN(AVERAGE(ProdMargin)/MarginGoal,1),0)
Products priced       =IFERROR(COUNTIF(ProdPrice,">0")/COUNTA(ProdItem),0)
Paying yourself       =IFERROR(MIN(EffHourly/WageGoal,1),0)
Profitable            =IFERROR(MIN(NetProfit/ProfitGoal,1),0)
Waste low             =IFERROR(1-MIN((WasteTotal/MonthlyIncome)/WasteLimit,1),0)
Bakery Score          =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `INDEX`, `MATCH`, `SUMIF`, `SUMPRODUCT`, `COUNTIF`, `COUNTA`,
`AVERAGE`, `MAX`, `MIN`, `IFERROR`, color scales and conditional formatting.

---

## 7. Printables & brand palette

The 12-page PDF is print-ready (US Letter) — a price-it worksheet, custom order
form & a cottage-food label for the kitchen.

> A business tool, not financial, legal or cottage-food-law advice — check your
> state's cottage food rules.

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
