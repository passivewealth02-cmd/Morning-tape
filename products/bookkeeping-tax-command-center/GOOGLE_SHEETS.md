# Small Business Bookkeeping & Tax Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Bookkeeping_Tax_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMIF`, `COUNTA`, `COUNTIF`,
`AVERAGE`, `MIN`, `IF`, `IFERROR`, named ranges, data validation, conditional
formatting, charts). Confirm:

- [ ] **Income** totals **$96,000** gross revenue.
- [ ] **COGS & Inventory** totals **$28,800**; **Expenses** totals **$19,200**.
- [ ] **Schedule C P&L** shows **$67,200** gross profit and **$48,000** net profit
      (Line 31), a **50%** net margin, **$6,782** SE tax, **$5,353** income tax,
      **$12,135** total tax and **$3,034** per quarter.
- [ ] **Mileage** shows **4,000** miles and a **$2,800** deduction.
- [ ] **Dashboard** shows all 12 KPIs, the Books Health bars, the where-the-revenue-goes
      and net-profit-by-month charts and **Books Score 90%**.
- [ ] Dropdowns work (income source, expense category, invoice status, yes/no) — from
      **Settings**.
- [ ] The **Net Profit by Month** chart renders.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`GrossRevenue`, `COGSTotal`, `NetProfit`, `SETax`, `IncomeTax`, `TotalTax`,
> `QuarterlyTax`, `MileageDeduction`, `HealthRange`, etc.) resolve — Sheets
> occasionally re-scopes on import.

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

> Keep the delivery a plain digital download. It's fine to answer a buyer's question
> in Messages after a sale, but never advertise support, setup, bookkeeping services or
> "free updates" as part of the listing — Etsy's Services policy treats that as selling
> a service. This is especially important here: sell the *file*, never the bookkeeping.
