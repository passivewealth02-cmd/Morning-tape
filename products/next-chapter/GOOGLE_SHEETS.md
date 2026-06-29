# Next Chapter™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the 20-sheet Excel
workbook. Same tab order: **Welcome, Dashboard, Profile, Timeline, Tasks,
Documents, Finances, Budget, Property, Debts, Parenting, Expenses,
Appointments, Contacts, Goals, Life Rebuild, Journal, Doc Checklist,
Analytics, Settings**.

> Build **Settings** first — it defines the control cells and dropdown
> lists. Then add the cross-sheet named ranges below.

> ⚠️ Keep the **Welcome-tab disclaimer** in the Google Sheets version too.
> Next Chapter™ is an organizational & planning tool — **not** legal,
> financial, or mental-health advice. Laws vary by jurisdiction.

---

## 1. Settings — controls & named ranges

Control cells: `ProcessStart` (C6), `MonthlyBudget` (C7), `UserName` (C8).

Dropdown lists (Data → Named ranges):
`TaskCatList, PriorityList, StatusList, DocStatusList, ExpenseCatList,
GoalCatList, OwnerList, AssetTypeList, AssetCatList, PayMethodList,
RoleList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to |
| ----- | --------- |
| `TaskName` | `Tasks!B5:B44` |
| `TaskDue` | `Tasks!D5:D44` |
| `TaskStatus` | `Tasks!E5:E44` |
| `DocName` | `Documents!A5:A44` |
| `DocStatus` | `Documents!E5:E44` |
| `SnapType` | `Finances!B5:B34` |
| `SnapCat` | `Finances!C5:C34` |
| `SnapValue` | `Finances!D5:D34` |
| `TotalAssets` | `Finances!G5` |
| `TotalDebts` | `Finances!G6` |
| `NetWorth` | `Finances!G7` |
| `SavingsBalance` | `Finances!G8` |
| `AllocLabels` | `Finances!F13:F18` |
| `AllocValues` | `Finances!G13:G18` |
| `BudgetCat` | `Budget!A5:A14` |
| `BudgetActual` | `Budget!C5:C14` |
| `BudgetTotalPlanned` | `Budget!B15` |
| `BudgetTotalActual` | `Budget!C15` |
| `ExpDate` | `Expenses!A5:A64` |
| `ExpCat` | `Expenses!B5:B64` |
| `ExpAmount` | `Expenses!C5:C64` |
| `ParentSchedule` | `Parenting!B7:B13` |
| `ApptDate` | `Appointments!C5:C44` |
| `GoalName` | `Goals!A5:A44` |
| `GoalProgress` | `Goals!D5:D44` |
| `GoalStatus` | `Goals!E5:E44` |

---

## 3. Tasks & Documents

**Tasks** columns: `A Category · B Task · C Owner · D Due · E Status ·
F Notes`. Validate `E5:E44` → `StatusList`, `A5:A44` → `TaskCatList`.
Conditional formatting: Status = **Complete** → mint; an overdue test
(`=AND($D5<>"",$D5<TODAY(),$E5<>"Complete")`) → soft red.

**Documents** columns: `A Document · B Who Has It · C Source · D Needed By ·
E Status · F Notes`. Validate `E5:E44` → `DocStatusList`.
**Collected** → mint.

---

## 4. Finances — net worth snapshot

Columns: `A Item · B Type (Asset/Debt) · C Category · D Value · E Owner`.

Summary block (right side):

```sheets
TotalAssets    G5 =SUMIF(SnapType,"Asset",SnapValue)
TotalDebts     G6 =SUMIF(SnapType,"Debt",SnapValue)
NetWorth       G7 =TotalAssets-TotalDebts
SavingsBalance G8 =SUMIFS(SnapValue,SnapType,"Asset",SnapCat,"Cash & Savings")
```

Asset-allocation table (`AllocLabels` / `AllocValues`), one row per asset
category:

```sheets
G13 =SUMIFS(SnapValue,SnapType,"Asset",SnapCat,F13)   ' fill down to G18
```

Charts: **Asset Allocation** doughnut (`AllocValues` by `AllocLabels`) and
**Assets vs Debts** column chart. Validate `B5:B34` → `AssetTypeList`,
`C5:C34` → `AssetCatList`, `E5:E34` → `OwnerList`.

---

## 5. Budget ← Expenses

**Expenses** columns: `A Date · B Category · C Amount · D Method · E Notes`.
Validate `B5:B64` → `ExpenseCatList`, `D5:D64` → `PayMethodList`.

**Budget** pulls actuals automatically:

```sheets
C5  =SUMIF(ExpCat,A5,ExpAmount)          ' actual for this category, fill down
D5  =B5-C5                                ' remaining = planned - actual
B15 =SUM(B5:B14)                          ' BudgetTotalPlanned
C15 =SUM(C5:C14)                          ' BudgetTotalActual
```

Conditional formatting on `D` (remaining): negative → soft red.

---

## 6. Parenting · Appointments · Goals

```sheets
ParentSchedule  Parenting!B7:B13   ' one row per weekday (who has the kids)
ApptDate        Appointments!C5:C44
GoalProgress    Goals!D5:D44       ' percent, data bar
GoalStatus      Goals!E5:E44       ' validate -> StatusList
```

Goals: validate `B` (category) → `GoalCatList`, format `D` as a percent
data bar.

---

## 7. Dashboard KPIs

```sheets
Days in Process       =MAX(TODAY()-ProcessStart,0)
Tasks Completed       =COUNTIF(TaskStatus,"Complete")
Documents Collected   =COUNTIF(DocStatus,"Collected")
Upcoming Appointments =SUMPRODUCT((ApptDate<>"")*(ApptDate>=TODAY()))
Monthly Budget        =MonthlyBudget
Savings Balance       =SavingsBalance
Parenting Set         =IFERROR(COUNTA(ParentSchedule)/7,0)            [0%]
Overall Progress      =IFERROR((
                          COUNTIF(TaskStatus,"Complete")/MAX(COUNTA(TaskName),1)
                        + COUNTIF(DocStatus,"Collected")/MAX(COUNTA(DocName),1)
                        + COUNTIF(GoalStatus,"Complete")/MAX(COUNTA(GoalName),1)
                        )/3,0)                                         [0%]
```

Charts: **Monthly Budget** doughnut (planned vs actual via
`BudgetTotalPlanned`/`BudgetTotalActual`), **Asset Allocation** doughnut,
**Overall Progress** bar, **Assets vs Debts** bar. Anchor them on a clean
grid so nothing overlaps; turn off auto data labels (show value only).

---

## 8. Analytics

Roll-ups from the data tabs:

```sheets
Tasks by status       =COUNTIF(TaskStatus,"Complete") / "In Progress" / "Not Started"
Documents collected   =COUNTIF(DocStatus,"Collected") / total
Spending by category  =SUMIF(ExpCat, <cat>, ExpAmount)   ' one row per category
Net worth             =NetWorth
Goal completion       =COUNTIF(GoalStatus,"Complete")/MAX(COUNTA(GoalName),1)
```

---

## 9. Apps Script (optional) — appointment reminder

```javascript
function nextChapterNudge() {
  // Optional: email yourself the next upcoming appointment / task.
  MailApp.sendEmail(Session.getActiveUser().getEmail(),
    "Next Chapter — upcoming this week",
    "Check your Appointments and Tasks tabs for what's coming up. One step at a time.");
}
function installTrigger() {
  ScriptApp.newTrigger('nextChapterNudge').timeBased().everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
}
```

---

## 10. Brand palette

| Token | Hex |
| ----- | --- |
| Primary | `#1B4F48` |
| Accent (Gold) | `#937356` |
| Gold Light | `#C9A86A` |
| Surface | `#E5D3BA` |
| Mint | `#75E6C1` |
| Ivory | `#FBF8F2` |
| Warning (disclaimer) | `#FBF0E2` |

Two-row gold-divider headers, gold-topped KPI cards, status color-coding
(Complete = mint, overdue = soft red, needed = gold). Keep it calm,
professional, and uncluttered.
