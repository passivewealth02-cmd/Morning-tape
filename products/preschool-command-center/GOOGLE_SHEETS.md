# Preschool Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Child Profiles, Skills & Milestones, ABC & 123, Weekly
Themes, Daily Rhythm, Activity Planner, Read-Aloud Log, Arts & Sensory, Field
Trips, Attendance, Portfolio, Supplies, Kindergarten Readiness, Goals, Settings**.

> Build **Settings** first (family + children + days goal + dropdown lists), then
> Skills & Milestones, ABC & 123, Weekly Themes, Activity Planner and Kindergarten
> Readiness, then the Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `FamilyName`, `SchoolYear`, `Children` (2), `DaysGoal` (150),
`BooksRead` (38), `SuppliesBudget` (400), `KinderDate`.

Lists: `DomainList, SkillStatusList, ActStatusList, ThemeStatusList,
SupplyStatusList, LovedList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `SkillDomain` | `'Skills & Milestones'!A5:A49` | `SkillName` | `'Skills & Milestones'!B5:B49` |
| `SkillStatus` | `'Skills & Milestones'!D5:D49` | `AbcRecognize` | `'ABC & 123'!B5:B30` |
| `AbcSound` | `'ABC & 123'!C5:C30` | `AbcWrite` | `'ABC & 123'!D5:D30` |
| `ThemeName` | `'Weekly Themes'!A5:A34` | `ThemeStatus` | `'Weekly Themes'!C5:C34` |
| `ActName` | `'Activity Planner'!B5:B49` | `ActStatus` | `'Activity Planner'!D5:D49` |
| `TripName` | `'Field Trips'!B5:B34` | `DaysDone` | `Attendance!B45` |
| `RdyName` | `'Kindergarten Readiness'!A5:A24` | `RdyDone` | `'Kindergarten Readiness'!B5:B24` |
| `GoalProgress` | `Goals!D5:D10` | `SuppliesSpent` | `Supplies!D13` |
| `DomSumVals` | `'Skills & Milestones'!C52:C58` | `HealthRange` | `Dashboard!C13:C18` |

---

## 3. Dashboard — the 12 KPIs

```sheets
Children         =Children
Preschool Days   =DaysDone
Letters Known    =COUNTIF(AbcRecognize,"Yes")
Skills Mastered  =COUNTIF(SkillStatus,"Mastered")
Themes Done      =COUNTIF(ThemeStatus,"Done")
Activities Done  =IFERROR(COUNTIF(ActStatus,"Done")/COUNTA(ActName),0)
Books Read       =BooksRead
Field Trips      =COUNTA(TripName)
Supplies Spent   =SuppliesSpent
Ready for K      =IFERROR(COUNTIF(RdyDone,"Yes")/COUNTA(RdyName),0)
Goals            =IFERROR(AVERAGE(GoalProgress),0)
Ready Score      =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Skills Mastered by Domain (bar) from the summary block on the Skills tab
(`=COUNTIFS(SkillDomain,<domain>,SkillStatus,"Mastered")`); Themes This Year
(donut). Turn off auto data labels.

---

## 4. Ready-for-Kindergarten Score (6 dimensions)

```sheets
Letters known (of 26)   =IFERROR(COUNTIF(AbcRecognize,"Yes")/26,0)
Skills mastered         =IFERROR(COUNTIF(SkillStatus,"Mastered")/COUNTA(SkillName),0)
Themes done             =IFERROR(COUNTIF(ThemeStatus,"Done")/COUNTA(ThemeName),0)
Activities done         =IFERROR(COUNTIF(ActStatus,"Done")/COUNTA(ActName),0)
Kindergarten readiness  =IFERROR(COUNTIF(RdyDone,"Yes")/COUNTA(RdyName),0)
Goals & milestones      =IFERROR(AVERAGE(GoalProgress),0)
Ready Score             =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `COUNTIF`/`COUNTIFS`, `AVERAGE`, `IFERROR`, plus `QUERY`
("skills still emerging", "activities by theme") and `FILTER`/`SORT`. Skills,
themes, ABC & activities are plain tables — add a row and the dashboard follows.

---

## 5. Printables

The 12-page PDF is print-ready as-is (US Letter). Google Sheets users can also
print any tab: File ▸ Print ▸ fit to width.

> Every child grows at their own pace — the readiness checklist is a gentle
> guide, not a test. This is a planning & keepsake tool.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
