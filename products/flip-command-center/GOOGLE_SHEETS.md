# Flip Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Deal Analyzer, Property Details, Rehab Budget, Scope of
Work, Contractors, Draws & Payments, Materials, Timeline, Holding Costs,
Financing, Comps & ARV, Selling & Exit, Punch List, Photo Log, Settings**.

> Build the **Deal Analyzer** inputs first, then the **Rehab Budget** and
> **Holding Costs** (they feed the analyzer), then the Scope of Work, Timeline
> and Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Deal Analyzer — the inputs (named cells)

`ARV` (340000), `PurchasePrice` (185000), `BuyClosing` (3500),
`HoldMonths` (5), `SellCostPct` (0.07), `LoanLTV` (0.80), `LoanRate` (0.10),
`Rule70` (0.70). Settings adds `ProfitTarget` (0.15), `ROITarget` (0.30).

## 2. Deal Analyzer — the outputs (named cells)

```sheets
LoanAmount       =PurchasePrice*LoanLTV
DownPayment      =PurchasePrice-LoanAmount
SellingCosts     =ARV*SellCostPct
AllInCost        =PurchasePrice+RehabBudget+BuyClosing+HoldingTotal+SellingCosts
CashInvested     =DownPayment+RehabBudget+BuyClosing+HoldingTotal
ProjectedProfit  =ARV-AllInCost
CashOnCash       =IFERROR(ProjectedProfit/CashInvested,0)
ReturnOnCost     =IFERROR(ProjectedProfit/AllInCost,0)
MAO70            =Rule70*ARV-RehabBudget
Verdict          =IF(PurchasePrice<=MAO70,"BUY","PASS")
```

---

## 3. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `RehabCat` | `'Rehab Budget'!B5:B15` | `RehabPlanned` | `'Rehab Budget'!C5:C15` |
| `RehabActual` | `'Rehab Budget'!D5:D15` | `RehabBudget` | `'Rehab Budget'!C16` (total planned) |
| `RehabSpent` | `'Rehab Budget'!D16` (total actual) | `HoldingTotal` | `'Holding Costs'!D10` |
| `TaskName` | `'Scope of Work'!B5:B49` | `TaskStatus` | `'Scope of Work'!D5:D49` |
| `PhaseName` | `Timeline!A5:A24` | `PhaseStatus` | `Timeline!D5:D24` |
| `PayAmount` | `'Draws & Payments'!D5:D44` | `PayStatus` | `'Draws & Payments'!E5:E44` |
| `CompPrice` | `'Comps & ARV'!D5:D16` | `HealthRange` | `Dashboard!C13:C18` |

---

## 4. Holding Costs (feeds All-In Cost)

```sheets
Loan interest / mo   =LoanAmount*LoanRate/12
+ taxes, insurance, utilities, misc (per month)
Monthly total        =SUM(monthly items)
HoldingTotal         =Monthly total * HoldMonths        (named at the total cell)
```

## 5. Rehab Budget

Per row: `Remaining =C-D`, `% Used =IFERROR(D/C,0)` (conditional-format red > 100%).
Totals: `RehabBudget =SUM(planned)`, `RehabSpent =SUM(actual)`.

---

## 6. Dashboard — the 12 KPIs

```sheets
ARV               =ARV
Purchase          =PurchasePrice
Rehab Budget      =RehabBudget
All-In Cost       =AllInCost
Projected Profit  =ProjectedProfit
Cash-on-Cash ROI  =CashOnCash
70% Rule MAO      =MAO70
Verdict           =IF(PurchasePrice<=MAO70,"BUY","PASS")
Budget Used       =IFERROR(RehabSpent/RehabBudget,0)
Spent to Date     =RehabSpent
Tasks Done        =IFERROR(COUNTIF(TaskStatus,"Done")/COUNTA(TaskName),0)
Deal Score        =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Rehab — Planned vs Actual (clustered column) from `RehabPlanned` /
`RehabActual` with `RehabCat` categories.

## 7. Deal & Project Health (6 dimensions → Deal Score)

```sheets
Profit margin (of ARV)  =IFERROR(MIN((ProjectedProfit/ARV)/ProfitTarget,1),0)
Cash-on-cash ROI        =IFERROR(MIN(CashOnCash/ROITarget,1),0)
Meets 70% rule          =IF(PurchasePrice<=MAO70,1,0.6)
Rehab on budget         =IFERROR(IF(RehabSpent<=RehabBudget,1,MAX(0,2-RehabSpent/RehabBudget)),0)
Scope complete          =IFERROR(COUNTIF(TaskStatus,"Done")/COUNTA(TaskName),0)
Timeline progress       =IFERROR(COUNTIF(PhaseStatus,"Done")/COUNTA(PhaseName),0)
Deal Score              =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `IF`, `IFERROR`, `MIN`/`MAX`, `SUMIF`, `COUNTIF`, `AVERAGE`,
`TEXT` (verdict banner), plus `QUERY`/`FILTER` for "unpaid draws" or "open scope."

> Not financial advice — run your own numbers and confirm local costs & comps.

---

## 8. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
