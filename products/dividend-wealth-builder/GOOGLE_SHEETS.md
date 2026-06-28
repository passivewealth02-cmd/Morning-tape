# Dividend Wealth Builder — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. All
ranges assume the same sheet layout: **Dashboard**, **Holdings**,
**Calculations**, **Settings**, **Projections**.

> Brand tokens are reused verbatim — palette and table rules live at the end
> of this doc so the Google Sheet matches the Excel file pixel-for-pixel.

---

## 1. Settings sheet

| Cell | Name | Default | Format |
| ---- | ---- | ------- | ------ |
| `C5` | `MonthlyContribution` | `500` | `"$"#,##0.00` |
| `C6` | `TargetMonthlyIncome` | `5000` | `"$"#,##0.00` |
| `C7` | `AnnualGrowthRate` | `0.08` | `0.00%` |
| `C8` | `DividendGrowthRate` | `0.06` | `0.00%` |
| `C9` | `ReinvestToggle` | `1` | `0` |
| `C10` | `TaxRate` | `0.15` | `0.00%` |
| `C11` | `YearsToProject` | `30` | `0` |

Define the named ranges via **Data → Named ranges**.

Dropdown lists (Settings `E6:E11`, `F6:F8`, `F9:F12`):

- Sectors: `Tech, Finance, Energy, REITs, Consumer, Healthcare`
- Risk: `Low, Medium, High`
- Frequency: `Monthly, Quarterly, Semi-Annual, Annual`

Frequency lookup table at `E18:F22`:

```
Frequency       Payments/Yr
Monthly         12
Quarterly       4
Semi-Annual     2
Annual          1
```

Name this range `FreqTable`.

---

## 2. Holdings sheet (rows 5–64)

Header row 4 — columns:

```
A Ticker   B Company   C Sector   D Shares   E Buy   F Current
G Value    H Div/Sh    I Freq     J Annual   K Yield L Risk
```

### Auto-fill calculated columns (use `ARRAYFORMULA` so adding rows just
### works):

**G5 — Total Value (whole column):**

```sheets
=ARRAYFORMULA(IF(D5:D64="","",D5:D64*F5:F64))
```

**J5 — Annual Income (whole column):**

```sheets
=ARRAYFORMULA(
  IF((D5:D64="")+(H5:H64="")+(I5:I64="")>0,"",
     D5:D64*H5:H64*IFERROR(VLOOKUP(I5:I64,FreqTable,2,FALSE),0)))
```

**K5 — Yield % (whole column):**

```sheets
=ARRAYFORMULA(
  IFERROR(
    IF((F5:F64="")+(H5:H64="")+(I5:I64="")>0,"",
       (H5:H64*VLOOKUP(I5:I64,FreqTable,2,FALSE))/F5:F64),
    ""))
```

### Data validation

- `C5:C64` — list from range `Settings!E6:E11`
- `I5:I64` — list from range `Settings!F9:F12`
- `L5:L64` — whole number between 1 and 5

### Conditional formatting

| Range | Rule | Style |
| ----- | ---- | ----- |
| `K5:K64` | Color scale `min → 50% → max` | `#E5D3BA → #FFE08C → #75E6C1` |
| `L5:L64` | Data bar 1–5 | `#C94C4C` |
| `A5:L64` | Custom formula `=AND($F5<>"",$F5<$E5)` | Background `#FBE6E6` |

Freeze rows 1–4.

---

## 3. Calculations sheet

```sheets
C5  =SUMPRODUCT((Holdings!G5:G64<>"")*IFERROR(Holdings!G5:G64,0))
C6  =SUMPRODUCT((Holdings!J5:J64<>"")*IFERROR(Holdings!J5:J64,0))
C7  =C6/12
C8  =IFERROR(C6/C5,0)
C9  =C7*(1-TaxRate)
C10 =SUMPRODUCT((Holdings!A5:A64<>"")*1)
C11 =IFERROR(C7/TargetMonthlyIncome,0)
```

Named ranges (Data → Named ranges):

- `TotalPortfolioValue → C5`
- `TotalAnnualIncome → C6`
- `TotalMonthlyIncome → C7`
- `PortfolioYield → C8`
- `AfterTaxMonthly → C9`
- `ProgressToGoal → C11`

### Sector breakdown (E5:H11)

`E6:E11` = sector list (typed or `=Settings!E6:E11`).

```sheets
F6  =ARRAYFORMULA(IF(E6:E11="","",SUMIFS(Holdings!G5:G64,Holdings!C5:C64,E6:E11)))
G6  =ARRAYFORMULA(IF(E6:E11="","",SUMIFS(Holdings!J5:J64,Holdings!C5:C64,E6:E11)))
H6  =ARRAYFORMULA(IFERROR(F6:F11/TotalPortfolioValue,0))
```

Alternative single-formula sector aggregation using `QUERY`:

```sheets
=QUERY(Holdings!A5:L64,
  "select C, sum(G), sum(J) where C is not null
   group by C label sum(G) 'Value', sum(J) 'Annual Income'",1)
```

### Monthly dividend distribution (B15:D27)

```sheets
B16:B27  Jan..Dec (typed)
C16      =ARRAYFORMULA(IF(B16:B27="","",TotalAnnualIncome/12))
D16      =C16
D17      =D16+C17  (drag down through D27)
```

---

## 4. Projections sheet

### FIRE summary (`C5:C8`)

```sheets
C5  =TargetMonthlyIncome*12*25
C6  =IFERROR(IF(TotalMonthlyIncome>=1000,0,
       LOG((1000/(TotalMonthlyIncome+0.0001)))/LOG(1+DividendGrowthRate)),"—")
C7  =IFERROR(IF(TotalMonthlyIncome>=5000,0,
       LOG((5000/(TotalMonthlyIncome+0.0001)))/LOG(1+DividendGrowthRate)),"—")
C8  =IFERROR(IF(TotalMonthlyIncome>=TargetMonthlyIncome,0,
       LOG((TargetMonthlyIncome/(TotalMonthlyIncome+0.0001)))/LOG(1+DividendGrowthRate)),"—")
```

### Year-by-year table (`B12:G42`)

Header row in row 12. Year 1 row (`B13`):

```sheets
B13 1
C13 =TotalPortfolioValue*(1+AnnualGrowthRate)+MonthlyContribution*12
D13 =TotalAnnualIncome*(1+DividendGrowthRate)
E13 =D13/12
F13 =MonthlyContribution*12
G13 =IFERROR(C13/FireNumber,0)
```

Recursive rows (`B14:G42`):

```sheets
B14 =B13+1
C14 =C13*(1+AnnualGrowthRate)+MonthlyContribution*12+IF(ReinvestToggle=1,D13,0)
D14 =D13*(1+DividendGrowthRate)+IF(ReinvestToggle=1,C13*PortfolioYield*0.05,0)
E14 =D14/12
F14 =F13+MonthlyContribution*12
G14 =IFERROR(C14/FireNumber,0)
```

Drag `B14:G14` down to row 42.

### Charts

| Chart | Type | Data | Categories |
| ----- | ---- | ---- | ---------- |
| Wealth Growth | Line | `Projections!C13:C42` | `Projections!B13:B42` |
| Annual Dividend Income | Line | `Projections!D13:D42` | `Projections!B13:B42` |
| Contributions vs Income | Column | `Projections!F13:F42`, `Projections!D13:D42` | `Projections!B13:B42` |
| Sector Allocation | Pie | `Calculations!F6:F11` | `Calculations!E6:E11` |
| Monthly Income | Column | `Calculations!C16:C27` | `Calculations!B16:B27` |
| Yield vs Risk | Scatter | x = `Holdings!L5:L64`, y = `Holdings!K5:K64` | — |

---

## 5. Dashboard sheet

### KPI cards (row 4 = label, row 5 = value)

```sheets
B5 =TotalPortfolioValue       [$0.00]
D5 =TotalMonthlyIncome        [$0.00]
F5 =TotalAnnualIncome         [$0.00]
H5 =PortfolioYield            [0.00%]
J5 =DividendGrowthRate        [0.00%]
L5 =ProgressToGoal            [0.00%]
```

Each label cell uses the brand section style; each value cell uses the
KPI value style.

### Filters (B9, D9, F9)

Data validation `list from range`:

- `B9` → `Settings!E6:E11`
- `D9` → `Settings!F6:F8`
- `F9` → `Settings!F9:F12`

Filters are visual cues — actual filtering happens via the AutoFilter on
Holdings row 4 (`Data → Create a filter`).

### Embed all six charts from sections 3 and 4 onto Dashboard.

---

## 6. Auto reset / rollover

Add an Apps Script (Extensions → Apps Script) for the monthly reset and
year rollover:

```javascript
function monthlySnapshot() {
  const ss = SpreadsheetApp.getActive();
  const calc = ss.getSheetByName('Calculations');
  const log = ss.getSheetByName('Reports') || ss.insertSheet('Reports');
  const month = Utilities.formatDate(new Date(),
                  Session.getScriptTimeZone(), 'yyyy-MM');
  log.appendRow([
    month,
    calc.getRange('C5').getValue(),  // Total Portfolio Value
    calc.getRange('C6').getValue(),  // Annual Income
    calc.getRange('C7').getValue(),  // Monthly Income
    calc.getRange('C8').getValue(),  // Yield
  ]);
}

function installTriggers() {
  ScriptApp.newTrigger('monthlySnapshot')
    .timeBased().onMonthDay(1).atHour(6).create();
}
```

Run `installTriggers` once. Snapshots will be logged on the first of each
month — drives historical trend lines on Reports without any manual work.

---

## 7. Brand palette (Format → Theme + custom)

| Token | Hex |
| ----- | --- |
| Primary | `#1B4F48` |
| Accent | `#937356` |
| Surface | `#E5D3BA` |
| Highlight | `#75E6C1` |
| Background | `#FFFFFF` |
| Text | `#333333` |
| Success | `#75E6C1` |
| Warning | `#937356` |
| Danger | `#C94C4C` |

Row banding (Format → Alternating colors): header `#1B4F48` white text,
odd `#FFFFFF`, even `#F4ECDE`.
