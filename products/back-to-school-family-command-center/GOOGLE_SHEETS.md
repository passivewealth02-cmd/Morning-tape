# Back-to-School Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Family Dashboard, Child Profiles, Contacts, Calendar, Events,
Supplies, Clothing, Budget, Fees, Extracurriculars, Lunch & Grocery, Homework,
PT Comms, Absences, Grades, Documents, Settings**.

> Build **Settings** first (family + first day + dropdown lists), then the
> Supplies, Clothing, Fees, Budget, Documents & Events tabs, then the Family
> Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `FamilyName`, `SchoolYear`, `FirstDay` (date), `Children` (6),
`Schools` (4), `BudgetTotal` (=BudgetPlanTotal), `Parent1`, `Parent2`.

Lists: `GradeList, EventTypeList, SupCatList, StatusList, PaidList, ReadyList,
SubjectList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `SupItem` | `Supplies!A5:A64` | `BudgetPlanTotal` | `Budget!B14` |
| `SupBought` | `Supplies!F5:F64` | `BudgetSpent` | `Budget!C14` |
| `ClothItem` | `Clothing!B5:B44` | `FeeAmtTotal` | `Fees!C15` |
| `ClothReady` | `Clothing!D5:D44` | `FeePaidTotal` | `Fees!D15` |
| `DocName` | `Documents!A5:A34` | `EventDate` | `Events!A5:A54` |
| `DocDone` | `Documents!B5:B34` | `EventStatus` | `Events!E5:E54` |
| `Children` | `Settings!C9` | `HealthRange` | `Family Dashboard!C13:C18` |

---

## 3. Dashboard — the 12 KPIs

```sheets
Children         =Children
Schools          =Schools
First Day        =FirstDay
Supplies Bought  =IFERROR(COUNTIF(SupBought,"Yes")/COUNTA(SupItem),0)
Budget Spent     =BudgetSpent
Budget Left      =BudgetPlanTotal-BudgetSpent
Fees Paid        =IFERROR(FeePaidTotal/FeeAmtTotal,0)
Forms Done       =IFERROR(COUNTIF(DocDone,"Yes")/COUNTA(DocName),0)
Uniforms Ready   =IFERROR(COUNTIF(ClothReady,"Ready")/COUNTA(ClothItem),0)
Events (30d)     =COUNTIFS(EventDate,">="&TODAY(),EventDate,"<="&TODAY()+30)
To-Do Open       =COUNTIF(EventStatus,"To Do")+COUNTIF(EventStatus,"In Progress")
Readiness        =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Budget by Category (donut) from the Budget tab. Turn off auto data labels.

---

## 4. Readiness Score (6 dimensions)

```sheets
Supplies bought     =IFERROR(COUNTIF(SupBought,"Yes")/COUNTA(SupItem),0)
Clothing & uniforms =IFERROR(COUNTIF(ClothReady,"Ready")/COUNTA(ClothItem),0)
Fees paid           =IFERROR(FeePaidTotal/FeeAmtTotal,0)
Forms & documents   =IFERROR(COUNTIF(DocDone,"Yes")/COUNTA(DocName),0)
Overdue caught up   =IFERROR(COUNTIFS(EventDate,"<"&TODAY(),EventStatus,"Done")/COUNTIF(EventDate,"<"&TODAY()),0)
Budget on track     =IFERROR(1-MAX(BudgetSpent-BudgetPlanTotal,0)/BudgetPlanTotal,0)
Readiness Score     =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `COUNTIF`/`COUNTIFS`, `SUMIF`, `IFERROR`, `TODAY()` date math,
`QUERY` ("what's due this week"), `FILTER`/`SORT`. A profile per child, contact
list and calendar are plain tables — duplicate a child block for kid #7 and #8.

---

## 5. Printables

The 12-page PDF is print-ready as-is (US Letter). To edit, open the PNG page
files or re-run `build_pdf.py`. Google Sheets users can also print any tab:
File ▸ Print ▸ fit to width.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
