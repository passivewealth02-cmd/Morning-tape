# Net Worth & FIRE Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Networth_FIRE_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMIF`, `COUNTIF`,
`COUNTA`, `AVERAGE`, `MIN`, `NPER`, `IFERROR`, named ranges, data validation,
conditional formatting, charts). Confirm:

- [ ] **Net Worth** shows **$250,000** (assets $530,000 − liabilities $280,000).
- [ ] **FIRE Number** shows a **$1,000,000** number, **25%** progress, a **$131,367**
      coast number and **10.5** years to FI.
- [ ] **Income & Expenses** shows a **50%** savings rate on **$40,000** saved.
- [ ] **Dashboard** shows all 12 KPIs, the FIRE Health bars, the progress-to-FIRE and
      net-worth-by-month charts and **FIRE Score 90%**.
- [ ] Dropdowns work (asset type, liability type, account type, yes/no) — from
      **Settings**.
- [ ] The **Net Worth by Month** chart renders.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`NetWorth`, `FIRENumber`, `FIREProgress`, `CoastNumber`, `YearsToFI`,
> `InvestedAssets`, `SavingsRate`, `HealthRange`, etc.) resolve — Sheets occasionally
> re-scopes on import. `NPER` is native to Sheets and Excel alike.

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
