# Savings Challenge & Sinking Funds Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Savings_Challenge_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMIF`, `COUNTA`, `COUNTIF`,
`AVERAGE`, `MIN`, `MAX`, `IF`, `IFERROR`, named ranges, data validation, conditional
formatting, charts). Confirm:

- [ ] **Sinking Funds** totals a **$9,000** target and **$6,000** saved, with a **$750**
      monthly set-aside and **100%** on pace.
- [ ] **100 Envelope** shows the **$5,050** challenge total; **52-Week** shows **$1,378**.
- [ ] **Emergency Fund** shows **$6,000** and **100%** funded.
- [ ] **Savings Accounts** totals **$12,274**.
- [ ] **Dashboard** shows all 12 KPIs, the Savings Health bars, the funds-filled and
      saved-by-month charts and **Savings Score 90%**.
- [ ] Dropdowns work (fund category, account type, frequency, yes/no) — from **Settings**.
- [ ] The **Saved by Month** chart renders.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`FundsTarget`, `FundsSaved`, `MonthlyNeed`, `OnPace`, `EnvelopeTotal`, `Week52Total`,
> `EFCurrent`, `StreakDays`, `HealthRange`, etc.) resolve — Sheets occasionally
> re-scopes on import.

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
- The 12-page printable PDF — including the **100-envelope grid** and **52-week tracker**
- A Start-Here quick-start guide

> **Tip:** the printable 100-envelope grid is the single most shareable page in this
> product. Make it image 6 in the listing and expect it to drive saves and traffic.

> Keep the delivery a plain digital download. It's fine to answer a buyer's question
> in Messages after a sale, but never advertise support, setup, or "free updates" as
> part of the listing — Etsy's Services policy treats that as selling a service.
