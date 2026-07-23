# Cleaning & Service Business Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Cleaning_Business_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `COUNTIF`, `COUNTA`,
`AVERAGE`, `MIN`, `IFERROR`, named ranges, data validation, conditional formatting,
charts). Confirm:

- [ ] **Job P&L** shows a **$100** job profit and a **$40.00** profit per hour.
- [ ] **Clients** rolls up to **$4,500** recurring and **$6,000** monthly revenue.
- [ ] **Expenses** shows a **$2,900** monthly profit.
- [ ] **Dashboard** shows all 12 KPIs, the Service Health bars, the profit-by-month
      chart and **Service Score 90%**.
- [ ] Dropdowns work (frequency, job status, lead stage, yes/no) — from **Settings**.
- [ ] The **Profit by Month** chart renders.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`JobProfit`, `ProfitHour`, `MRR`, `MonthlyRevenue`, `NetProfit`, `RecurringClients`,
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
