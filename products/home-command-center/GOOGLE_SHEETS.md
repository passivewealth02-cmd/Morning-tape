# Home Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same
tab order: **Welcome, Dashboard, Family Directory, Calendar, Daily, Budget,
Bills, Grocery, Pantry, Fridge & Freezer, Meal Planner, Recipes, Cleaning,
Chores, Maintenance, Home Inventory, Subscriptions, Family Goals, Savings,
Holidays & Events, Travel, Children's Hub, Pet Care, Wellness, Projects,
Documents, Analytics, Inspiration, Settings**.

> Build **Settings** first — it holds the control cells and all dropdown
> lists. Then add the cross-sheet named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & dropdown lists

Controls: `HouseholdName` (C6), `MonthlyIncome` (C7), `MonthlyBudget` (C8),
`WeekStart` (C9).

Dropdown lists: `FamilyList, ExpenseCatList, GroceryCatList, PantryCatList,
CleanFreqList, ChoreFreqList, MaintCatList, GoalCatList, MealTypeList,
StatusList, PriorityList, CalCatList, StoreList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `BudgetCat` | `Budget!A5:A22` | `CleanStatus` | `Cleaning!E5:E54` |
| `BudgetActual` | `Budget!C5:C22` | `CleanTask` | `Cleaning!A5:A54` |
| `BudgetTotalPlanned` | `Budget!B23` | `CleanFreq` | `Cleaning!C5:C54` |
| `BudgetTotalActual` | `Budget!C23` | `ChoreAssigned` | `Chores!B5:B44` |
| `BillName` | `Bills!A5:A44` | `ChoreStatus` | `Chores!D5:D44` |
| `BillAmount` | `Bills!C5:C44` | `ChorePct` | `Chores!F5:F44` |
| `BillDue` | `Bills!D5:D44` | `MaintNext` | `Maintenance!E5:E54` |
| `BillPaid` | `Bills!F5:F44` | `SavTarget` | `Savings!B5:B10` |
| `GrocEst` | `Grocery!E5:E64` | `SavCurrent` | `Savings!C5:C10` |
| `GrocActual` | `Grocery!G5:G64` | `GoalName` | `'Family Goals'!A5:A44` |
| `PantryName` | `Pantry!A5:A64` | `GoalProgress` | `'Family Goals'!D5:D44` |
| `PantryQty` | `Pantry!C5:C64` | `GoalStatus` | `'Family Goals'!E5:E44` |
| `PantryReorder` | `Pantry!D5:D64` | `ProjStatus` | `Projects!H5:H16` |
| `PantryCat` | `Pantry!B5:B64` | `SubMonthly` | `Subscriptions!D5:D44` |
| `CalDate` | `Calendar!C5:C64` | `MealWeek` | `'Meal Planner'!B6:H10` |

---

## 3. Budget ← cash flow & savings rate

```sheets
D5  =B5-C5                                   ' Remaining (fill down)
E5  =IFERROR(C5/B5,0)                         ' % Used
B23 =SUM(B5:B22)   C23 =SUM(C5:C22)           ' totals
Cash Flow    =MonthlyIncome-BudgetTotalActual
Savings Rate =IFERROR(SUMIF(BudgetCat,"Savings",BudgetActual)/MonthlyIncome,0)
```

Donut: **Monthly Spending** = `Actual` by `Category`. Validate the Category
column to `ExpenseCatList`.

---

## 4. Bills — overdue flagging

```sheets
Overdue rule (conditional format on A5:H44):
=AND($D5<>"", $D5<TODAY(), $F5<>"Yes")        → red
Paid rule:  =$F5="Yes"                          → mint
Paid count =COUNTIF(BillPaid,"Yes")   Unpaid =COUNTIF(BillPaid,"No")
```

---

## 5. Pantry & Fridge — stock & expiry alerts

```sheets
Low stock (Pantry A5:G64):  =AND($A5<>"", $C5<=$D5)          → red
Expiring (Pantry E col):    =AND($E5<>"", $E5<=TODAY()+14)   → gold
Use-soon (Fridge):          =AND($E5<>"", $E5<=TODAY()+3)    → gold
By-category counts:         =COUNTIF(PantryCat, <category>)
```

---

## 6. Cleaning & Chores — completion

```sheets
Cleaning by frequency:  Done =COUNTIFS(CleanFreq,<freq>,CleanStatus,"Done")
                        %    =IFERROR(Done/MAX(COUNTIF(CleanFreq,<freq>),1),0)
Chore by member:        =IFERROR(AVERAGEIFS(ChorePct,ChoreAssigned,<member>),0)
```

---

## 7. Meal Planner & Savings

```sheets
Meal plan complete = (COUNTIF(MealWeek,"<>"&"") - COUNTIF(MealWeek,"—")) / 35
Savings funded     = C5/B5   (per fund)   ·   SUM(SavCurrent)/SUM(SavTarget) (overall)
```

---

## 8. Dashboard — the 10 KPIs

```sheets
Budget Remaining    =BudgetTotalPlanned-BudgetTotalActual
Bills Paid          =COUNTIF(BillPaid,"Yes")
Cleaning Done       =COUNTIF(CleanStatus,"Done")
Pantry Stock        =IFERROR(SUMPRODUCT((PantryName<>"")*(PantryQty>PantryReorder))/MAX(COUNTA(PantryName),1),0)
Events This Week    =SUMPRODUCT((CalDate<>"")*(CalDate>=TODAY())*(CalDate<=TODAY()+7))
Meal Plan           =IFERROR((COUNTIF(MealWeek,"<>"&"")-COUNTIF(MealWeek,"—"))/35,0)
Savings Goal        =IFERROR(SUM(SavCurrent)/SUM(SavTarget),0)
Projects Active     =COUNTIF(ProjStatus,"In Progress")
Maintenance Due     =SUMPRODUCT((MaintNext<>"")*(MaintNext<=TODAY()+30))
Routine Score       =IFERROR(AVERAGE(ChorePct),0)
```

Charts: Monthly Spending (donut), Bill Status (donut), Cleaning Progress
(bar), Chore Completion (bar), Pantry by Category (donut), Household Goals
(bar). Turn off auto data labels (value only) so nothing piles up.

---

## 9. Analytics — Household Health Score

```sheets
=IFERROR(AVERAGE(
   Bills Paid %, Budget Used %, Cleaning Done %, Chore Completion %,
   Pantry Stocked %, Savings Funded %, Goal Progress %, Meal Plan %), 0)
```

Google-Sheets power features to use: `ARRAYFORMULA` for the Remaining/%Used
columns, `QUERY` for the Analytics roll-ups, `FILTER`/`SORT` for "bills due
this week", `UNIQUE` for category lists, all wrapped in `IFERROR`.

---

## 10. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, mint/gold/red status
coding. Keep every tab consistent — that cohesion is the premium feel.
