# Soccer Mom Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook.
Fourteen tabs in this order: **Dashboard, Players, Schedule, Practices,
Budget, Equipment, Tournaments, Carpool, Roster, Meals, Packing,
Mileage, Communication, Settings**.

---

## 1. Settings (build first — defines every named range)

| Cell | Named range | Default | Format |
| ---- | ----------- | ------- | ------ |
| `C5` | `Season` | `Fall 2026` | text |
| `C6` | `SeasonStart` | `2026-08-24` | `mm/dd/yyyy` |
| `C7` | `SeasonEnd` | `2026-11-22` | `mm/dd/yyyy` |
| `C8` | `League` | `U10 Travel` | text |
| `C9` | `Division` | `Premier` | text |
| `C10` | `MonthlyBudget` | `650` | `"$"#,##0.00` |
| `C11` | `PlayerCount` | `2` | `0` |
| `C12` | `FuelPerMile` | `0.18` | `"$"#,##0.000` |
| `C13` | `TodayDate` | `=TODAY()` | `mm/dd/yyyy` |

Dropdown lists (Data → Named ranges):

- `ExpenseList`, `PositionList`, `PracticeTypeList`, `ConditionList`,
  `VenueList` — rows 6–end in cols E:I
- `ResultList`, `LeagueList`, `DivisionList`, `FieldList`, `YesNoList`
  — rows 21–end in cols E:I
- `AttendanceList`, `UniformList` — col B

---

## 2. Schedule (rows 5–44)

Column layout: `A Date · B Opponent · C Home/Away · D Kickoff · E Arrival ·
F Field · G Address · H Weather · I Uniform · J Result · K Notes · L Days Out`.

**L5 (Days Out — whole column):**

```sheets
=ARRAYFORMULA(IF(A5:A44="","",A5:A44-TODAY()))
```

Validation:

- `C5:C44` → `VenueList`
- `F5:F44` → `FieldList`
- `I5:I44` → `UniformList`
- `J5:J44` → `ResultList`

Conditional formatting (apply to `A5:L44`):

| Rule | Formula | Background |
| ---- | ------- | ---------- |
| Upcoming 7 days | `=AND($A5<>"",$A5-TODAY()>=0,$A5-TODAY()<=7)` | `#DCF5EC` |
| Past games | `=AND($A5<>"",$A5<TODAY())` | `#F1F1F1` |
| Result = Win | `=$J5="Win"` (apply to `J5:J44` only) | `#E3F8EF` |
| Result = Loss | `=$J5="Loss"` (apply to `J5:J44` only) | `#FBE6E6` |

Named ranges:

- `GameDates → Schedule!A5:A44`
- `GameOpponents → Schedule!B5:B44`
- `GameResults → Schedule!J5:J44`

---

## 3. Practices (rows 5–54)

Columns: `A Date · B Time · C Location · D Focus Area · E Attendance ·
F Coach Notes · G Drills`.

Validation:

- `C5:C54` → `FieldList`
- `D5:D54` → `PracticeTypeList`
- `E5:E54` → `AttendanceList`

Conditional formatting (apply to `E5:E54`):

| Value | Background |
| ----- | ---------- |
| `Present` | `#E3F8EF` |
| `Late` | `#FBF0E2` |
| `Absent` | `#FBE6E6` |

Named ranges:

- `PracDates → Practices!A5:A54`
- `PracAttendance → Practices!E5:E54`

---

## 4. Budget (rows 5–16 + total on 17)

```sheets
B5:B16   Planned amount (input)
C5       Actual (typed or =SUMIFS from a transaction log if you want
         to wire it; default leaves manual entry)
D5       =B5-C5
E5       =IFERROR(C5/B5,0)
```

Drag down. Row 17 totals:

```sheets
B17  =SUM(B5:B16)
C17  =SUM(C5:C16)
D17  =SUM(D5:D16)
E17  =IFERROR(C17/B17,0)
```

Named ranges:

- `BudgetTotalPlanned → Budget!B17`
- `BudgetTotalActual → Budget!C17`

KPI sidebar (G5:H9):

```sheets
H5  =BudgetTotalPlanned                  [$0.00]
H6  =BudgetTotalActual                   [$0.00]
H7  =BudgetTotalPlanned-BudgetTotalActual[$0.00;[Red]-$0.00]
H8  =BudgetTotalActual/MAX(1,(SeasonEnd-SeasonStart)/30) [$0.00]
H9  =BudgetTotalActual/MAX(PlayerCount,1)                [$0.00]
```

Charts:

- **Spending Breakdown** (Doughnut) — labels `A5:A16`, data `C5:C16`
- **Budget vs Actual** (Column) — labels `A5:A16`, series `B5:B16` + `C5:C16`

---

## 5. Equipment (rows 5–34)

Columns: `A Category · B Item · C Size · D Qty · E Purchased ·
F Condition · G Replace By · H Cost · I Notes`.

Validation:

- `E5:E34` → `YesNoList`
- `F5:F34` → `ConditionList`

Conditional formatting (apply to `A5:I34`):

| Rule | Formula | Background |
| ---- | ------- | ---------- |
| Replace By within 60 days | `=AND($G5<>"",$G5-TODAY()>=0,$G5-TODAY()<=60)` | `#FBF0E2` |
| Condition = Replace | `=$F5="Replace"` (apply to `F5:F34`) | `#FBE6E6` |
| Purchased = No | `=$E5="No"` (apply to `E5:E34`) | `#FBE6E6` |
| Purchased = Yes | `=$E5="Yes"` (apply to `E5:E34`) | `#E3F8EF` |

Named ranges:

- `EquipItem → Equipment!B5:B34`
- `EquipPurchased → Equipment!E5:E34`
- `EquipReplaceBy → Equipment!G5:G34`

---

## 6. Tournaments (rows 5–16)

Conditional formatting: `=AND($B5<>"",$B5-TODAY()>=0,$B5-TODAY()<=30)`
→ `#DCF5EC`.

Named ranges:

- `TourStart → Tournaments!B5:B16`
- `TourName → Tournaments!A5:A16`

---

## 7. Mileage (rows 5–44)

```sheets
D5 (Fuel $)  =ARRAYFORMULA(IF(C5:C44="","",C5:C44*FuelPerMile))
H5 (Total)   =ARRAYFORMULA(IF(C5:C44="","",
                D5:D44 + IFERROR(E5:E44,0)
                       + IFERROR(F5:F44,0)
                       + IFERROR(G5:G44,0)))
```

Named ranges:

- `MileageMiles → Mileage!C5:C44`
- `MileageTotal → Mileage!H5:H44`
- `MileageDate → Mileage!A5:A44`

Monthly travel cost chart (Insert → Chart, Column):

```sheets
Use QUERY in a side panel:
=QUERY({MileageDate, MileageTotal},
  "select Col1, sum(Col2)
   where Col1 is not null
   group by Col1 pivot month(Col1)+1",0)
```

---

## 8. Dashboard

### KPI row 1 (B5, D5, F5, H5)

```sheets
B5 =SUMPRODUCT((GameDates<>"")
              *(MONTH(IFERROR(GameDates,0))=MONTH(TODAY()))
              *(YEAR(IFERROR(GameDates,0))=YEAR(TODAY())))      [0]
D5 =SUMPRODUCT((PracDates<>"")
              *(PracDates-TODAY()>=0)
              *(PracDates-TODAY()<=7))                          [0]
F5 =MonthlyBudget                                                [$0.00]
H5 =MonthlyBudget - BudgetTotalActual                            [$0.00]
```

### KPI row 2 (B8, D8, F8, H8)

```sheets
B8 =SUMPRODUCT((TourStart<>"")*(TourStart-TODAY()>=0))           [0]
D8 =SUMPRODUCT((EquipItem<>"")*(EquipPurchased<>"Yes"))          [0]
F8 =SUMPRODUCT((GameDates<>"")
              *(COUNTIF(PracDates,GameDates)>0))                 [0]
H8 =IFERROR(
     SUMPRODUCT((PracAttendance="Present")*1)
     / SUMPRODUCT((PracAttendance<>"")*(PracAttendance<>"—")*1),
   0)                                                            [0%]
```

### Quick navigation chips (row 11)

8 cells with internal links to: Schedule, Budget, Equipment, Roster,
Travel/Mileage, Meals, Carpool, Communication. Fill `#1B4F48`, bold
white text, centered.

### Embedded charts

| Chart | Type | Source |
| ----- | ---- | ------ |
| Spending by Category | Doughnut | `Budget!A5:A16`, `Budget!C5:C16` |
| Budget vs Actual | Column | `Budget!A5:A16`, `B5:B16` + `C5:C16` |
| Games vs Practices (monthly) | Stacked Column | QUERY on `GameDates` + `PracDates` |
| Attendance Trend | Line | `=QUERY(PracAttendance,"select count(Col1) where Col1<>'' group by Col1")` |
| Travel by Month | Column | QUERY on `MileageDate`/`MileageTotal` (see §7) |

---

## 9. Apps Script — auto-refresh + monthly snapshot

```javascript
function refreshDaily() {
  SpreadsheetApp.flush();
}

function monthlySnapshot() {
  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName('Reports') || ss.insertSheet('Reports');
  const v = (sheet, cell) => ss.getSheetByName(sheet).getRange(cell).getValue();
  log.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM'),
    v('Dashboard', 'B5'),   // Games this month
    v('Dashboard', 'H5'),   // Budget remaining
    v('Dashboard', 'B8'),   // Upcoming tournaments
    v('Dashboard', 'H8'),   // Practice attendance %
  ]);
}

function installTriggers() {
  ScriptApp.newTrigger('refreshDaily')
    .timeBased().everyDays(1).atHour(5).create();
  ScriptApp.newTrigger('monthlySnapshot')
    .timeBased().onMonthDay(1).atHour(6).create();
}
```

Run `installTriggers` once.

---

## 10. Brand palette

| Token | Hex | Where |
| ----- | --- | ----- |
| Primary | `#1B4F48` | Page headers, table headers, KPI values, nav chips |
| Accent | `#937356` | KPI labels, field labels, warning |
| Surface | `#E5D3BA` | Input cells, totals row |
| Highlight | `#75E6C1` | Positive KPIs, upcoming game highlight |
| Danger | `#C94C4C` | Replace / Absent / Loss rows |
| Soft BG | `#FAF7F1` | Field-value backgrounds |
| Muted Row | `#F4ECDE` | Alternating row stripes |

Header rows: fill `#1B4F48`, white bold all-caps.
Alternating data rows: white / `#F4ECDE`.
Input cells (Settings + Budget Planned): fill `#E5D3BA`, primary bold text.
