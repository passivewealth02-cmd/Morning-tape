# Rental Property & Landlord Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Rental_Property_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMIF`, `INDEX`, `MATCH`,
`COUNTA`, `AVERAGE`, `MIN`, `IFERROR`, named ranges, data validation, conditional
formatting, charts). Confirm:

- [ ] **Deal Analyzer** shows NOI **$1,095**, cash flow **$360**, cap rate **6.0%**,
      cash-on-cash **8.0%** and DSCR **1.49**.
- [ ] **Rent Roll** totals **$7,000** portfolio rent.
- [ ] **Reserves & Escrow** shows each fund's funded %.
- [ ] **Dashboard** shows all 12 KPIs, the Landlord Health bars, the cash-flow-by-
      month chart and **Landlord Score 90%**.
- [ ] Dropdowns work (unit status, pay status, expense category, yes/no) — from
      **Settings**.
- [ ] The **Cash Flow by Month** chart renders.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`CashFlow`, `NOI`, `CapRate`, `CoC`, `DSCR`, `PortfolioRent`, `ReservePct`,
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
> in Messages after a sale, but never advertise support, setup, or "free updates" as
> part of the listing — Etsy's Services policy treats that as selling a service.
