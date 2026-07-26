# Amazon FBA & Online Seller Profit Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `FBA_Seller_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMIF`, `SUMPRODUCT`,
`COUNTA`, `COUNTIF`, `AVERAGE`, `MIN`, `ROUND`, `IF`, `IFERROR`, named ranges, data
validation, conditional formatting, charts). Confirm:

- [ ] **Profit Calculator** shows **$10.25** total Amazon fees, an **$8.70** landed
      cost, **$11.04** net per unit, **36.8%** net margin and **126.9%** ROI.
- [ ] Monthly: **420** units → **$12,596** revenue and **$4,637** profit.
- [ ] **Product Catalog** recalculates net/unit for all 8 SKUs and flags any negative
      net in red.
- [ ] **Inventory** shows **60** days of cover on the flagship and flags reorders.
- [ ] **PPC & ACoS** totals **20.0%** ACoS; TACoS shows against monthly revenue.
- [ ] **Dashboard** shows all 12 KPIs, the Seller Health bars, the where-the-$29.99-goes
      and net-profit-by-month charts and **Seller Score 90%**.
- [ ] Dropdowns work (channel, stock status, return reason, yes/no) — from **Settings**.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`SalePrice`, `TotalFees`, `LandedCost`, `NetPerUnit`, `NetMargin`, `ROI`, `ACoS`,
> `FlagshipCover`, `HealthRange`, etc.) resolve — Sheets occasionally re-scopes on
> import.

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

> **A note on fees:** Amazon's referral rates, FBA fees and storage rates change, and
> they vary by category and size tier. Say so plainly in your listing — buyers should
> pull their own numbers from the Seller Central fee preview and enter them in Settings.
> This protects you from "the fee was wrong" messages.

> Keep the delivery a plain digital download. It's fine to answer a buyer's question
> in Messages after a sale, but never advertise support, setup, coaching, or "free
> updates" as part of the listing — Etsy's Services policy treats that as selling a
> service.
