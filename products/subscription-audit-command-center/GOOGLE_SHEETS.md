# Subscription & Bills Audit Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Subscription_Audit_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMIF`, `COUNTA`, `COUNTIF`,
`AVERAGE`, `MIN`, `IF`, `IFERROR`, named ranges, data validation, conditional
formatting, charts). Confirm:

- [ ] **Subscriptions** totals **$216.77** monthly across **14** subs; 2 on annual
      billing.
- [ ] **Subscription Audit** shows **$2,601** annual, a **$105.92** monthly cancel
      total and **$1,271** annual savings, **$110.85** kept.
- [ ] **Cancel Finder** lists the 5 flagged subs and their annual savings.
- [ ] **Bills** totals **$468** monthly.
- [ ] **Dashboard** shows all 12 KPIs, the Audit Health bars, the keep-vs-cancel and
      recurring-by-month charts and **Audit Score 90%**.
- [ ] Dropdowns work (category, billing, action, yes/no) — from **Settings**.
- [ ] The **Recurring by Month** chart renders.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`SubMonthly`, `CancelMonthly`, `KeepMonthly`, `BillMonthly`, `AnnualBilled`,
> `TrialCount`, `HikeTotal`, `HealthRange`, etc.) resolve — Sheets occasionally
> re-scopes on import. `SUMIF` and `COUNTIF` are native to Sheets.

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
