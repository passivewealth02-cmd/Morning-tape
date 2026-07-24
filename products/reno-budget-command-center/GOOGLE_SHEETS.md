# Home Renovation & Remodel Budget Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Reno_Budget_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMPRODUCT`, `SUMIF`,
`COUNTA`, `COUNTIF`, `INDEX`, `AVERAGE`, `MIN`, `IF`, `IFERROR`, named ranges, data
validation, conditional formatting, charts). Confirm:

- [ ] **Rooms** totals a **$60,000** budget and **$42,000** spent, **$18,000**
      remaining (70% used).
- [ ] **Budget vs Actual** shows **$18,000** remaining and a **$9,000** contingency
      reserve, **$7,000** left.
- [ ] **Payments** shows **$38,000** paid and **$4,000** outstanding.
- [ ] **Change Orders** totals **$3,000**.
- [ ] **Dashboard** shows all 12 KPIs, the Reno Health bars, the budget-used and
      spent-by-month charts and **Reno Score 90%**.
- [ ] Dropdowns work (room, category, status, yes/no) — from **Settings**.
- [ ] The **Spent by Month** chart renders.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`TotalBudget`, `TotalSpent`, `Remaining`, `ContingencyReserve`, `ContingencyLeft`,
> `PaidToDate`, `Outstanding`, `ChangeOrders`, `HealthRange`, etc.) resolve — Sheets
> occasionally re-scopes on import. `SUMPRODUCT` and `INDEX` are native to Sheets.

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
