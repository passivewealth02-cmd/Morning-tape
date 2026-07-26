# Contractor Job Costing & Bidding Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Contractor_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMIF`, `SUMPRODUCT`,
`COUNTA`, `COUNTIF`, `AVERAGE`, `ROUND`, `IF`, `IFERROR`, `AND`, named ranges, data
validation, conditional formatting, charts). Confirm:

- [ ] **Bid Builder** shows **$38,100** direct cost, **$4,572** overhead, **$42,672**
      total cost and a **$53,340** bid price with **$10,668** planned profit.
- [ ] The **markup-is-not-margin** block underneath shows the wrong way — **$51,206** at
      a real **16.7%** margin, **$2,134** short. This is the block buyers screenshot;
      make sure it survives the import.
- [ ] **Job Costing** shows **$41,178** actual cost, **$12,162** actual profit and a
      **22.8%** actual margin, with overrun lines flagged red.
- [ ] **Jobs & Pipeline** shows **9** jobs and **$186,000** backlog; **Invoices** totals
      **$31,000** outstanding.
- [ ] **Bid Log** computes a **40.9%** win rate (9 of 22).
- [ ] **Dashboard** shows all 12 KPIs, the Builder Health bars, the where-the-bid-dollar-
      goes and revenue-by-month charts and **Builder Score 90%**.
- [ ] Dropdowns work (job status, trade, change-order status, paid/unpaid, bid outcome) —
      from **Settings**.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`DirectCost`, `Overhead`, `TotalCost`, `BidPrice`, `PlannedProfit`, `ActualCost`,
> `ActualProfit`, `ActualMargin`, `Backlog`, `Receivable`, `WinRate`, `MarginTarget`,
> `HealthRange`, etc.) resolve — Sheets occasionally re-scopes on import. `WinRateGoal`
> and `MarginTarget` are scalar names, so they appear as constants rather than cells.

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

> **A note on rates:** labor burden, workers' comp, liability insurance, bonding and
> licensing all vary by state and trade, and the sample uses a fully burdened $45/hr.
> Say so plainly in your listing — buyers should enter their own burdened rate and
> overhead percentage in Settings. This protects you from "the rate was wrong" messages.

> Keep the delivery a plain digital download. It's fine to answer a buyer's question
> in Messages after a sale, but never advertise support, setup, bid building, coaching,
> or "free updates" as part of the listing — Etsy's Services policy treats that as
> selling a service.
