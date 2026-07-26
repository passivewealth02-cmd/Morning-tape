# First-Time Home Buyer & Mortgage Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Home_Buyer_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `PMT`, `COUNTA`, `COUNTIF`,
`AVERAGE`, `MIN`, `MAX`, `CEILING`, `IF`, `AND`, `IFERROR`, named ranges, data
validation, conditional formatting, charts). Confirm:

- [ ] **Affordability** shows a **$266,800** loan, **$1,686** principal & interest and a
      **$2,213** monthly payment, with **28.0%** front-end and **33.6%** back-end DTI
      and a **COMFORTABLE** verdict.
- [ ] **Closing Costs** totals **$8,700**; **Down Payment** shows **$31,900** cash to
      close against **$32,000** saved.
- [ ] **Lender Compare** recalculates a monthly payment for each lender's rate.
- [ ] **Dashboard** shows all 12 KPIs, the Buyer Health bars, the where-the-payment-goes
      and saved-by-month charts and **Buyer Score 90%**.
- [ ] Dropdowns work (loan type, home status, priority, yes/no) — from **Settings**.
- [ ] The **Saved by Month** chart renders.

> `PMT` is native to both Google Sheets and Excel and returns a negative number by
> convention — the workbook negates it (`=-PMT(...)`) so your payment displays positive.
> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`HomePrice`, `LoanAmount`, `PITI`, `FrontDTI`, `BackDTI`, `CashToClose`,
> `HealthRange`, etc.) resolve — Sheets occasionally re-scopes on import.

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
> in Messages after a sale, but never advertise support, setup, mortgage advice or
> "free updates" as part of the listing — Etsy's Services policy treats that as selling
> a service, and lending advice carries its own regulatory risk. Sell the *file*.
