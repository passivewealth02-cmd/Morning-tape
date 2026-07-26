# Trucking Owner-Operator Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Trucking_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMIF`, `COUNTA`, `MIN`,
`AVERAGE`, `IF`, `IFERROR`, named ranges, data validation, conditional formatting,
charts). Confirm:

- [ ] **Fixed Costs** totals **$4,490** a month and **$0.401** fixed cost per mile.
- [ ] **Variable Costs** shows **$0.600** fuel per mile and **$0.860** variable per mile.
- [ ] **Cost Per Mile** shows **11,200** total miles, **10.7%** deadhead, **$1.261** cost
      per mile run, **$14,122** total cost and **$1.41** cost per loaded mile.
- [ ] The **deadhead block** at the bottom shows the $2.35 load is really **$2.10** per
      mile actually driven — **$2,518** a month. This is the block buyers screenshot;
      make sure it survives the import.
- [ ] **Loads** totals **10,000** loaded, **1,200** deadhead, **$23,500**, and flags any
      load whose all-in rate is below your cost per loaded mile in red.
- [ ] **Fuel Log** totals **1,723** gallons and computes **6.50** actual MPG.
- [ ] **IFTA & Miles** totals **11,200** — it should tie to your total miles run.
- [ ] **Monthly Summary** shows **$9,378** profit at a **39.9%** net margin.
- [ ] **Dashboard** shows all 12 KPIs, the Road Health bars, the where-the-$2.35-goes and
      revenue-by-month charts and **Road Score 90%**.
- [ ] Dropdowns work (load status, equipment type, maintenance type, yes/no) — from
      **Settings**.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`FixedTotal`, `FixedCPM`, `FuelCPM`, `VarCPM`, `TotalMiles`, `DeadheadPct`,
> `TotalCPM`, `TotalCostMonth`, `CostPerLoaded`, `RatePerMile`, `ProfitPerMile`,
> `CoverRatio`, `MonthlyRevenue`, `MonthlyProfit`, `ActualMPG`, `Reserve`,
> `HealthRange`, etc.) resolve — Sheets occasionally re-scopes on import.

> **Per-mile costs display to three decimals** ($0.401, $0.860, $1.261). If Sheets rounds
> them to two on import, select the cells and set the format to `0.000` — a penny of
> rounding on 100,000 miles a year is a thousand dollars of illusion.

---

## 3. Make the "Make a Copy" share link

1. **Share → General access → Anyone with the link → Viewer.**
2. Copy the link. It ends in `/edit?usp=sharing`.
3. **Replace `/edit?usp=sharing` with `/copy`.**

```
https://docs.google.com/spreadsheets/d/FILE_ID/copy
```

Now anyone who clicks it gets **"Make a copy"** — their own private copy, your master
untouched. Put this link in `GOOGLE_SHEETS_TEMPLATE_LINK.txt` in the buyer's download.

---

## 4. What buyers receive

- The `.xlsx` file (opens in Excel & Google Sheets)
- The **"Make a Copy"** Google Sheets link (1 click → their own copy)
- The 12-page printable PDF
- A Start-Here quick-start guide

> **A note on diesel and rates:** diesel price, insurance, permits and freight rates all
> move constantly, and they vary by lane, authority and equipment. Say so plainly in your
> listing — buyers should enter their own numbers in Settings. This protects you from
> "the fuel price was wrong" messages.

> **A note on compliance:** the pre-trip inspection page is an organizing aid, not a
> substitute for a required DVIR, and nothing here is DOT, tax or legal advice. Say that
> plainly too — professional drivers respect it.

> Keep the delivery a plain digital download. It's fine to answer a buyer's question
> in Messages after a sale, but never advertise support, setup, dispatch, coaching, or
> "free updates" as part of the listing — Etsy's Services policy treats that as selling
> a service.
