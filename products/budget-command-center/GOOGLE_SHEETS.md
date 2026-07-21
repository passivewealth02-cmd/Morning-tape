# Budget & Money Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Income, Monthly Budget, Bills, Expense Log, Savings
Goals, Sinking Funds, Debt Snapshot, Net Worth, Subscriptions, Year View,
No-Spend, Settings**.

> Build **Income** and **Monthly Budget** first (the engine), then Bills, Savings
> Goals, Sinking Funds, Debt Snapshot and Net Worth, then the Dashboard. Add the
> named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Household`, `BudgetMonth`, `SavingsTarget` (0.20),
`EFTarget` ($15,000), `Currency` (5).

Lists: `CategoryList, YesNoList, PaidList, FreqList`.

---

## 2. Income & Monthly Budget — the engine

```sheets
Income (Income!C)            =SUM(C5:C7)                       → named Income
Remaining (Monthly Budget)  =D{r}-E{r}                        planned − actual
Left to budget (header)     =Income-BudgetPlanTotal
Budget totals               =SUM(D…) / =SUM(E…)               → BudgetPlanTotal / BudgetActualTotal
```

Named on Monthly Budget: `BudgetCat` (B), `BudgetGroup` (C), `BudgetPlanned` (D),
`BudgetActual` (E), plus the total cells `BudgetPlanTotal` and `BudgetActualTotal`.

---

## 3. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `Income` | `Income!$C$8` | `BudgetGroup` | `'Monthly Budget'!C7:C20` |
| `BudgetActual` | `'Monthly Budget'!E7:E20` | `BudgetPlanTotal` | `'Monthly Budget'!$D$21` |
| `BudgetActualTotal` | `'Monthly Budget'!$E$21` | `BillStatus` | `Bills!D5:D28` |
| `BillAmount` | `Bills!B5:B28` | `GoalPct` | `'Savings Goals'!E5:E8` |
| `EFSaved` | `'Savings Goals'!$D$5` | `SinkTotal` | `'Sinking Funds'!$C$…` |
| `DebtTotal` | `'Debt Snapshot'!$C$…` | `NetWorth` | `'Net Worth'!$F$…` |
| `SubTotal` | `Subscriptions!$B$…` | `SavingsTarget` | `Settings!$C$8` |
| `EFTarget` | `Settings!$C$9` | `HealthRange` | `Dashboard!C13:C18` |

---

## 4. Dashboard — the 12 KPIs

```sheets
Income          =Income
Spent           =BudgetActualTotal
Left to Budget  =Income-BudgetPlanTotal
Saved           =SUMIF(BudgetGroup,"Savings",BudgetActual)
Savings Rate    =IFERROR(SUMIF(BudgetGroup,"Savings",BudgetActual)/Income,0)
Bills Paid      =IFERROR(COUNTIF(BillStatus,"Paid")/COUNTA(BillStatus),0)
Net Worth       =NetWorth
Total Debt      =DebtTotal
Savings Goals   =IFERROR(AVERAGE(GoalPct),0)
Sinking Funds   =SinkTotal
Subscriptions   =SubTotal
Health Score    =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Spending by Category (donut) from the Monthly Budget actual column.

---

## 5. Budget Health Score (6 dimensions)

```sheets
On budget            =IFERROR(IF(BudgetActualTotal<=BudgetPlanTotal,1,BudgetPlanTotal/BudgetActualTotal),0)
Savings rate         =IFERROR(MIN((SUMIF(BudgetGroup,"Savings",BudgetActual)/Income)/SavingsTarget,1),0)
Bills paid on time   =IFERROR(COUNTIF(BillStatus,"Paid")/COUNTA(BillStatus),0)
Emergency fund       =IFERROR(MIN(EFSaved/EFTarget,1),0)
Savings goals        =IFERROR(AVERAGE(GoalPct),0)
Sinking funds ready  =IFERROR(MIN(SinkTotal/1650,1),0)
Health Score         =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `SUMIF`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MIN`, `IFERROR`, data
bars (budget actual & goal progress), a color scale (debt by rate, health score)
and conditional formatting (over-budget = red, bill status Paid/Due/Overdue).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter). Print any tab: File ▸ Print ▸
fit to width.

> A personal budgeting tool, not financial, tax or investment advice — for big
> decisions, talk to a qualified professional.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
