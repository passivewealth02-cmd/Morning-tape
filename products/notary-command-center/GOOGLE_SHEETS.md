# Notary & Loan Signing Agent Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Notary_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `COUNTA`, `MIN`, `MAX`,
`ROUNDUP`, `AVERAGE`, `IF`, `IFERROR`, named ranges, data validation, conditional
formatting, charts). Confirm:

- [ ] **Signing Profit** shows **$138.89/hr** as what it feels like, **$6.30** printing,
      **$8.36** driving, **$110.34** net per signing, **2.5 hrs** door to door and
      **$44.14** real hourly. This is the block buyers screenshot; make sure it survives
      the import.
- [ ] Mileage deduction shows **$26.60** per signing and **$1,383** for the month.
- [ ] **Signings Log** computes cost and net per row, shades "Overdue" amber and "fell
      through" red, and shows a **96.2%** getting-paid rate.
- [ ] **Fee Schedule** lists 13 services including the trip fee.
- [ ] **Invoices** totals **52** signings, **$6,500** invoiced and **$2,140** outstanding.
- [ ] **Notarial Journal** shows **12** acts and the red bound-journal warning above the
      table.
- [ ] **Expenses** shows **$289** fixed per month.
- [ ] **Tax Set-Aside** shows **$2,100** saved against **$5,250** due.
- [ ] **Monthly Summary** shows **$6,500** revenue, **$5,449** profit, **83.8%** margin
      and a **3-signing** break-even covered **17.3×**.
- [ ] **Dashboard** shows all 12 KPIs, the Business Health bars, the
      where-the-2.5-hours-go donut and revenue chart, and **Signing Score 90%**.
- [ ] Dropdowns work (signing type, payment status, ID type, expense category, yes/no) —
      from **Settings**.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`LooksLike`, `PrintCost`, `DriveCost`, `NetPerSigning`, `TotalHours`, `RealHourly`,
> `MileageDeduction`, `MileageMonth`, `PaidRate`, `Invoiced`, `Receivable`, `FixedTotal`,
> `TaxReserve`, `Revenue`, `Profit`, `BreakEven`, `HealthRange`, etc.) resolve — Sheets
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

---

## 5. Three things to say plainly in the listing

> **Maximum fees are set by your state.** What a notary may charge per notarial act is
> statutory and varies widely; travel, printing and scanback charges are usually separate
> but not always. The fee schedule in this file is a starting point to edit, not a
> recommendation. Say so — experienced notaries respect that you know fees are statutory,
> and it protects you.

> **The Notarial Journal tab is a convenience record.** Many states require a **bound,
> sequential** journal with specific entries. This does not replace one. The warning is
> already on the tab and on printable page 5; repeat it in the listing.

> **The IRS mileage rate changes every year.** The sample uses a current-year figure and
> Settings has one input to update. Tell buyers to change it each January — it is the
> single most common reason a tool like this goes stale.

> Keep the delivery a plain digital download. It's fine to answer a buyer's question in
> Messages after a sale, but never advertise support, setup, **"become a signing agent"
> training, coaching or mentoring**, or "free updates" as part of the listing. The NSA
> space is saturated with coaching offers, which is exactly the category Etsy's Services
> policy removes.
