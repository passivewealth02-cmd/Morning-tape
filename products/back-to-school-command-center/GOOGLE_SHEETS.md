# Back-to-School Command Center — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook.
Twelve tabs, identical layout: **Dashboard, Students, Supplies, Budget,
Clothing, Schedule, Assignments, Extracurricular, Calendar, Meals,
Emergency, Settings**.

---

## 1. Settings (build first — defines every named range)

| Cell | Named range | Default | Format |
| ---- | ----------- | ------- | ------ |
| `C5` | `SchoolYear` | `2026-2027` | text |
| `C6` | `FirstDayOfSchool` | `2026-09-02` | `mm/dd/yyyy` |
| `C7` | `TotalBudget` | `1200` | `"$"#,##0.00` |
| `C8` | `FamilySize` | `2` | `0` |
| `C9` | `TaxRate` | `0.075` | `0.00%` |
| `C10` | `TodayDate` | `=TODAY()` | `mm/dd/yyyy` |

Dropdown lists (named in **Data → Named ranges**):

- `CategoryList` → `Settings!E6:E14`
- `PriorityList` → `Settings!F6:F8`
- `StatusList` → `Settings!G6:G9`
- `SubjectList` → `Settings!H6:H15`
- `StoreList` → `Settings!E21:E26`
- `GradeList` → `Settings!F21:F35`
- `YesNoList` → `Settings!G21:G22`
- `TransportList` → `Settings!H21:H25`

---

## 2. Supplies (rows 5–64)

Column layout: `A Category · B Item · C Req · D Bought · E Remaining
· F Store · G Est · H Actual · I Purchased · J Priority · K Notes`.

**E5 (Remaining, whole column):**

```sheets
=ARRAYFORMULA(IF(C5:C64="","",IF(D5:D64="",C5:C64,MAX(C5:C64-D5:D64,0))))
```

Replace `MAX` with `IF(C5:C64-D5:D64<0,0,C5:C64-D5:D64)` if you want
per-row guards without `ARRAYFORMULA` collisions:

```sheets
=ARRAYFORMULA(IF(C5:C64="","",
  IF(IFERROR(D5:D64,0)>=C5:C64,0,C5:C64-IFERROR(D5:D64,0))))
```

Validation:

- `A5:A64` → list from `CategoryList`
- `F5:F64` → list from `StoreList`
- `I5:I64` → list from `YesNoList`
- `J5:J64` → list from `PriorityList`

Conditional formatting:

| Range | Formula | Background |
| ----- | ------- | ---------- |
| `A5:K64` | `=$I5="Yes"` | `#E3F8EF` |
| `J5:J64` | `=$J5="High"` | `#EFE0CC` |
| `E5:E64` | `=$E5>0` | `#FBF0E2` |

Named ranges for cross-sheet aggregation:

- `SupActual` → `Supplies!H5:H64`
- `SupPurchased` → `Supplies!I5:I64`
- `SupItem` → `Supplies!B5:B64`
- `SupReq` → `Supplies!C5:C64`
- `SupBought` → `Supplies!D5:D64`

---

## 3. Budget (rows 5–13 + total on 14)

```sheets
A5:A13   Category names (typed: Supplies, Clothing, ... Miscellaneous)
B5       Planned amount (input)
C5       =IFERROR(SUMIFS(SupActual,Supplies!$A$5:$A$64,A5),0)
D5       =B5-C5
E5       =IFERROR(C5/B5,0)
```

Drag rows 5→13 down. Row 14 totals:

```sheets
B14  =SUM(B5:B13)
C14  =SUM(C5:C13)
D14  =SUM(D5:D13)
E14  =IFERROR(C14/B14,0)
```

Named ranges:

- `BudgetTotalPlanned` → `B14`
- `BudgetTotalActual` → `C14`

Data bar on `E5:E13` via conditional formatting → color scale 0% → 100%
using primary `#1B4F48`.

Charts:

- **Spending Breakdown** (Doughnut) — labels `A5:A13`, data `C5:C13`
- **Budget vs Actual** (Column) — labels `A5:A13`, series `B5:B13` & `C5:C13`

---

## 4. Assignments (rows 5–54)

Column layout: `A Subject · B Assignment · C Due Date · D Priority ·
E Status · F Grade · G Notes · H Days Until`.

```sheets
H5 (Days Until):
=ARRAYFORMULA(IF(C5:C54="","",C5:C54-TODAY()))
```

Validation:

- `A5:A54` → `SubjectList`
- `D5:D54` → `PriorityList`
- `E5:E54` → `StatusList`

Conditional formatting (apply to `A5:H54`):

| Rule | Formula | Background |
| ---- | ------- | ---------- |
| Overdue | `=AND($C5<>"",$C5<TODAY(),$E5<>"Complete")` | `#FBE6E6` |
| Complete | `=$E5="Complete"` | `#E3F8EF` |
| Due this week | `=AND($C5<>"",$C5-TODAY()>=0,$C5-TODAY()<=7,$E5<>"Complete")` | `#FBF0E2` |

Named ranges:

- `AsgStatus` → `Assignments!E5:E54`
- `AsgDue` → `Assignments!C5:C54`
- `AsgSubject` → `Assignments!A5:A54`

---

## 5. Calendar (rows 5–44)

Columns: `A Date · B Day · C Event · D Notes · E Type · F Student`.

Conditional formatting (apply to `A5:F44`):

- Upcoming 14 days: `=AND($A5<>"",$A5-TODAY()>=0,$A5-TODAY()<=14)` → `#DCF5EC`
- Past: `=AND($A5<>"",$A5<TODAY())` → `#F1F1F1`

Auto-fill day-of-week if you'd rather not type it:

```sheets
B5 =ARRAYFORMULA(IF(A5:A44="","",TEXT(A5:A44,"ddd")))
```

---

## 6. Dashboard

### KPI row 1 (cells B5, D5, F5, H5)

```sheets
B5 =TotalBudget                                            [$0.00]
D5 =BudgetTotalActual                                      [$0.00]
F5 =TotalBudget-BudgetTotalActual                          [$0.00]
H5 =MAX(FirstDayOfSchool-TODAY(),0)                        [0]
```

### KPI row 2 (cells B8, D8, F8, H8)

```sheets
B8 =SUMPRODUCT((SupPurchased="Yes")*1)                     [0]
D8 =SUMPRODUCT((SupItem<>"")*(SupPurchased<>"Yes"))        [0]
F8 =IFERROR(
     SUMPRODUCT((SupPurchased="Yes")*1)
     / SUMPRODUCT((SupItem<>"")*1), 0)                     [0%]
H8 =IFERROR(
     SUMPRODUCT((AsgStatus="Complete")*1)
     / SUMPRODUCT((AsgStatus<>"")*1), 0)                   [0%]
```

### Embedded charts

| Chart | Type | Source |
| ----- | ---- | ------ |
| Spending by Category | Doughnut | `Budget!A5:A13`, `Budget!C5:C13` |
| Budget vs Actual | Column | `Budget!A5:A13`, `Budget!B5:B13` + `C5:C13` |
| Assignment Status | Stacked Column | `=QUERY(AsgStatus, "select Col1, count(Col1) where Col1<>'' group by Col1")` |
| Days-to-First-Day | Single big card | `=MAX(FirstDayOfSchool-TODAY(),0)` |

### Quick Navigation chips (row 11)

Insert a 7-cell row with internal links — right-click → **Insert link** →
pick the target sheet for each chip (Shopping List, Budget, Calendar,
Students, Assignments, Meals, Emergency). Style: fill `#1B4F48`, white
bold text, centered.

---

## 7. Apps Script — countdown auto-refresh + monthly snapshot

`Extensions → Apps Script`:

```javascript
function refreshCountdown() {
  SpreadsheetApp.flush();
}

function monthlySnapshot() {
  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName('Reports') || ss.insertSheet('Reports');
  const calc = (name, cell) =>
    ss.getSheetByName(name).getRange(cell).getValue();
  log.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM'),
    calc('Dashboard', 'B5'),   // Total budget
    calc('Dashboard', 'D5'),   // Spent
    calc('Dashboard', 'F8'),   // Supply completion %
    calc('Dashboard', 'H8'),   // Assignment completion %
  ]);
}

function installTriggers() {
  ScriptApp.newTrigger('refreshCountdown')
    .timeBased().everyDays(1).atHour(5).create();
  ScriptApp.newTrigger('monthlySnapshot')
    .timeBased().onMonthDay(1).atHour(6).create();
}
```

Run `installTriggers` once.

---

## 8. Brand palette (Format → Theme + custom)

| Token | Hex |
| ----- | --- |
| Primary | `#1B4F48` |
| Accent | `#937356` |
| Surface | `#E5D3BA` |
| Highlight | `#75E6C1` |
| Soft BG | `#FAF7F1` |
| Background | `#FFFFFF` |
| Text | `#333333` |
| Success | `#75E6C1` |
| Warning | `#937356` |
| Danger | `#C94C4C` |

Header rows: fill `#1B4F48`, white bold text, all-caps.
Alternating rows: white / `#F4ECDE`.
Input cells (Settings + Budget B5:B13): fill `#E5D3BA`, primary bold text.
