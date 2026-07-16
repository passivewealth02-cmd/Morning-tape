# Tot-School Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Tot Profiles, Milestones, First Words & Concepts, Weekly
Themes, Daily Rhythm, Tot Trays, Board-Book Log, Sensory Play, Outings,
Attendance, Portfolio, Supplies, Ready for Preschool, Goals, Settings**.

> Build **Settings** first (family + tots + days goal + dropdown lists), then
> Milestones, First Words & Concepts, Weekly Themes, Tot Trays and Ready for
> Preschool, then the Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `FamilyName`, `SchoolYear`, `Tots` (2), `DaysGoal` (120),
`Words` (45), `BoardBooks` (24), `SuppliesBudget` (300), `PreschoolDate`.

Lists: `DomainList, MileStatusList, TrayStatusList, ThemeStatusList,
CategoryList, LovedList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `MileDomain` | `Milestones!A5:A49` | `MileName` | `Milestones!B5:B49` |
| `MileStatus` | `Milestones!D5:D49` | `ConceptName` | `'First Words & Concepts'!A5:A34` |
| `ConceptCat` | `'First Words & Concepts'!B5:B34` | `ConceptDone` | `'First Words & Concepts'!C5:C34` |
| `ThemeName` | `'Weekly Themes'!A5:A34` | `ThemeStatus` | `'Weekly Themes'!C5:C34` |
| `TrayName` | `'Tot Trays'!A5:A44` | `TrayStatus` | `'Tot Trays'!C5:C44` |
| `TripName` | `Outings!B5:B34` | `DaysDone` | `Attendance!B45` |
| `RdyName` | `'Ready for Preschool'!A5:A24` | `RdyDone` | `'Ready for Preschool'!B5:B24` |
| `GoalProgress` | `Goals!D5:D10` | `SuppliesSpent` | `Supplies!D13` |
| `DomSumVals` | `Milestones!C52:C57` | `HealthRange` | `Dashboard!C13:C18` |

---

## 3. Dashboard — the 12 KPIs

```sheets
Tots             =Tots
Tot-School Days  =DaysDone
Words Spoken     =Words
Milestones Met   =COUNTIF(MileStatus,"Met")
Themes Done      =COUNTIF(ThemeStatus,"Done")
Trays Done       =IFERROR(COUNTIF(TrayStatus,"Done")/COUNTA(TrayName),0)
Board Books      =BoardBooks
Outings          =COUNTA(TripName)
Supplies Spent   =SuppliesSpent
Preschool-Ready  =IFERROR(COUNTIF(RdyDone,"Yes")/COUNTA(RdyName),0)
Goals            =IFERROR(AVERAGE(GoalProgress),0)
Growing Score    =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Milestones Met by Domain (bar) from the summary block on the Milestones
tab (`=COUNTIFS(MileDomain,<domain>,MileStatus,"Met")`); Themes This Year
(donut). Turn off auto data labels.

---

## 4. Growing Score (6 dimensions)

```sheets
Milestones met        =IFERROR(COUNTIF(MileStatus,"Met")/COUNTA(MileName),0)
Words & concepts      =IFERROR(COUNTIF(ConceptDone,"Yes")/COUNTA(ConceptName),0)
Themes done           =IFERROR(COUNTIF(ThemeStatus,"Done")/COUNTA(ThemeName),0)
Trays & play done     =IFERROR(COUNTIF(TrayStatus,"Done")/COUNTA(TrayName),0)
Ready for preschool   =IFERROR(COUNTIF(RdyDone,"Yes")/COUNTA(RdyName),0)
Goals & little wins   =IFERROR(AVERAGE(GoalProgress),0)
Growing Score         =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `COUNTIF`/`COUNTIFS`, `AVERAGE`, `IFERROR`, plus `QUERY`
("milestones still emerging", "trays by skill area") and `FILTER`/`SORT`.
Milestones, concepts, themes & trays are plain tables — add a row and the
dashboard follows.

---

## 5. Printables

The 12-page PDF is print-ready as-is (US Letter). Google Sheets users can also
print any tab: File ▸ Print ▸ fit to width.

> Every toddler grows on their own timeline — the milestones & readiness lists
> are gentle guides, not a test. This is a planning & keepsake tool.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
