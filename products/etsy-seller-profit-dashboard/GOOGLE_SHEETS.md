# Etsy Seller Profit Dashboard™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same
tab order: **Welcome, Dashboard, Orders, Product Calc, Fees, Ads, Expenses,
Library, Monthly, Insights, Goals, Cash Flow, Tax Prep, Strategy, Settings**.

> Build **Settings** first (fee rates + dropdown lists), then Orders — every
> other tab reads from it. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — fee rates & goals

Controls: `ShopName` (C6), `FeeListing` (C7), `FeeTxn` (C8), `FeeProc` (C9),
`FeeProcFixed` (C10), `FeeOffsite` (C11), `FeeSubscription` (C12), `TaxRate`
(C13), `RevGoal` (C14), `ProfitGoal` (C15).

Lists: `CategoryList, ExpenseTypeList, FeeTypeList, AdTypeList, GoalTypeList,
StatusList, YesNoList`.

> 2026 US defaults: listing $0.20, transaction 6.5%, processing 3% + $0.25,
> offsite ads 15%. **Confirm your exact rates in Etsy.**

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `OrdID` | `Orders!A5:A44` | `AdProduct` | `Ads!B5:B29` |
| `OrdProduct` | `Orders!C5:C44` | `AdSpend` | `Ads!C5:C29` |
| `OrdQty` | `Orders!D5:D44` | `AdOrders` | `Ads!F5:F29` |
| `OrdSale` | `Orders!F5:F44` | `AdRevenue` | `Ads!G5:G29` |
| `OrdShip` | `Orders!G5:G44` | `LibName` | `Library!A5:A10` |
| `OrdOffsite` | `Orders!H5:H44` | `LibRevenue` | `Library!F5:F10` |
| `OrdFees` | `Orders!I5:I44` | `LibProfit` | `Library!G5:G10` |
| `OrdRefunded` | `Orders!K5:K44` | `LibMargin` | `Library!H5:H10` |
| `OrdProfit` | `Orders!L5:L44` | `CalcName` | `Product Calc!A5:A10` |
| `FeesTotal` | `Fees!B12` | `ExpensesTotal` | `Expenses!I5` |

---

## 3. Order Tracker — the engine

```sheets
Sale Total (F)  =D5*E5
Etsy Fees (I)   =IF(F5="","", FeeListing*D5 + FeeTxn*(F5+G5)
                  + (FeeProc*(F5+G5)+FeeProcFixed)
                  + IF(H5="Yes", FeeOffsite*(F5+G5), 0))
COGS (J)        =IFERROR(SUMIF(CalcName,C5,CalcCOGS)*D5,0)
Net Profit (L)  =IF(F5="","", IF(K5="Yes", 0, F5+G5-I5-J5))
```

Refund rule (conditional format A5:L44): `=$K5="Yes"` → red.

Use `ARRAYFORMULA` to fill F/I/J/L down the whole column in one formula each.

---

## 4. Dashboard — the 10 KPIs

```sheets
Total Revenue  =SUM(OrdSale)+SUM(OrdShip)
Net Profit     =SUM(OrdProfit)-SUM(AdSpend)-ExpensesTotal
Etsy Fees      =FeesTotal
Ad Spend       =SUM(AdSpend)
Refunds        =SUMIFS(OrdSale,OrdRefunded,"Yes")
Profit Margin  =IFERROR(NetProfit/Revenue,0)
Avg Order      =IFERROR(Revenue/COUNTA(OrdID),0)
Orders         =COUNTA(OrdID)
Best Seller    =INDEX(LibName,MATCH(MAX(LibRevenue),LibRevenue,0))
Needs Work     =INDEX(LibName,MATCH(MIN(LibProfit),LibProfit,0))
```

Charts: Revenue vs Profit (line, from Monthly), Etsy Fees (donut, from Fees),
Profit by Product (bar, from Library), Ad Revenue vs Spend (from Ads).

---

## 5. Fees, Library, Ads

```sheets
Fees:    Transaction =FeeTxn*(SUM(OrdSale)+SUM(OrdShip))
         Processing  =FeeProc*(SUM(OrdSale)+SUM(OrdShip))+FeeProcFixed*COUNTA(OrdID)
         Offsite     =FeeOffsite*SUMIFS(OrdSale,OrdOffsite,"Yes")
Library: Revenue =SUMIF(OrdProduct,A5,OrdSale)   Profit =SUMIF(OrdProduct,A5,OrdProfit)
Ads:     CTR =Clicks/Impressions   ROAS =Revenue/Spend
```

`QUERY` is great for the Library ranking: `=QUERY(Orders!C5:L, "select C,
sum(F), sum(L) group by C order by sum(L) desc label ...", 0)`.

---

## 6. Cash Flow · Tax · Strategy

```sheets
Net Cash Flow  =(SUM(OrdSale)+SUM(OrdShip))-FeesTotal-SUM(AdSpend)-ExpensesTotal
Tax reserve    =MAX(NetProfit*TaxRate,0)
Best product   =INDEX(LibName,MATCH(MAX(LibProfit),LibProfit,0))
Needs work     =INDEX(LibName,MATCH(MIN(LibProfit),LibProfit,0))
```

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Clean KPI cards, gold-topped, soft shadows, red-refund / mint-win coding.
Keep it a calm finance-app feel — that clarity is the whole product.
