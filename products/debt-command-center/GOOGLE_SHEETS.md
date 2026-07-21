# Debt Payoff Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Debts, Payoff Plan, Snowball vs Avalanche, Payment Log,
Balance History, Extra Payment, Milestones, Interest Tracker, Accelerators,
Settings**.

> Build **Debts** first (the engine), then Payoff Plan, the Snowball vs Avalanche
> comparison, Payment Log, Extra Payment, Milestones and Accelerators, then the
> Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Household`, `Method` (Snowball/Avalanche), `ExtraPay` ($300),
`ExtraTarget` ($300), `StarterEF` ($1,000), `EFSaved` ($1,000).

Lists: `TypeList, YesNoList, MethodList`.

---

## 2. Debts — the engine

For each debt (balance, APR, min payment, original balance):

```sheets
% Paid (H)  =IFERROR((Original-Balance)/Original,0)
```

Named: `DebtName` (B), `DebtType` (C), `DebtBalance` (D), `DebtAPR` (E),
`DebtMin` (F), `DebtOrig` (G), `DebtPct` (H), plus the totals `DebtTotal`,
`DebtMinTotal`, `DebtOrigTotal`.

---

## 3. Payoff projection (snowball & avalanche)

The month-by-month simulation is computed in `build_xlsx.py` and written into the
**Snowball vs Avalanche** tab as `SnowMonths`, `AvalMonths`, `SnowInterest`,
`AvalInterest`, `SnowDate` and `InterestSaved`, with the chosen-method values
`MyMonths` and `MyInterest`:

```sheets
Months (my method)   =IF(Method="Avalanche",AvalMonths,SnowMonths)
Interest (my method) =IF(Method="Avalanche",AvalInterest,SnowInterest)
Interest saved       =SnowInterest-AvalInterest
```

> To recompute for your own debts, re-run `build_xlsx.py` (it re-simulates), or
> keep the values and edit them by hand. Each accrues monthly interest
> (`APR/12`), pays every minimum, then rolls the remainder down a fixed attack
> order until each balance hits $0.

---

## 4. Dashboard — the 12 KPIs

```sheets
Total Debt       =DebtTotal
Paid Off         =DebtOrigTotal-DebtTotal
% Paid           =IFERROR((DebtOrigTotal-DebtTotal)/DebtOrigTotal,0)
Monthly Payment  =DebtMinTotal+ExtraPay
Extra Payment    =ExtraPay
Highest APR      =MAX(DebtAPR)
Debt-Free Date   =SnowDate
Months to Free   =MyMonths
Total Interest   =MyInterest
Interest Saved   =InterestSaved
Focus Debt       =INDEX(PlanName,1)
Momentum         =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Balance by Debt (bar) from the Debts balance column.

---

## 5. Payoff Momentum Score (6 dimensions)

```sheets
Debt reduced          =IFERROR((DebtOrigTotal-DebtTotal)/DebtOrigTotal,0)
Extra payment funded  =IFERROR(MIN(ExtraPay/ExtraTarget,1),0)
Payments on time      =IFERROR(COUNTIF(PayOnTime,"Yes")/COUNTA(PayOnTime),0)
Milestones hit        =IFERROR(COUNTIF(MilestoneDone,"Yes")/COUNTA(MilestoneDone),0)
Per-debt progress     =IFERROR(AVERAGE(DebtPct),0)
Starter fund ready    =IFERROR(MIN(EFSaved/StarterEF,1),0)
Momentum              =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `SUMIF`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`, `INDEX`,
`IFERROR`, data bars (per-debt % paid), color scales (APR by rate, momentum) and
conditional formatting (on-time = mint, missed = red; accelerators on = mint).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter) — including a **debt
thermometer** to color as you pay down. Print any tab: File ▸ Print ▸ fit to
width.

> A personal debt-payoff tool, not financial, tax or credit advice — for big
> decisions, talk to a qualified professional.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
