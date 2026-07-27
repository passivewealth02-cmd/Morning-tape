# Estate & Emergency Organizer Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Estate_Organizer_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMIF`, `COUNTA`, `COUNTIF`,
`MIN`, `AVERAGE`, `IF`, `IFERROR`, named ranges, data validation, conditional formatting,
charts). Confirm:

- [ ] **Assets & Accounts** totals **$1,499,900**, with **$535,500** flagged as exposed to
      probate and those rows shaded red.
- [ ] **Debts & Bills** totals **$244,200**.
- [ ] **Estate Snapshot** shows **$1,255,700** net estate, **35.7%** probate share and a
      **$26,775** rough probate cost. This is the block buyers screenshot; make sure it
      survives the import.
- [ ] The survivor half shows **$45,400** reachable and **7.3 months** of runway, with
      life insurance covering the debts **2.05×**.
- [ ] **Beneficiaries** shows **7** eligible accounts with **7** named.
- [ ] **Legal Documents** shows **4** signed, **6** not started, **40%** ready — with
      signed rows mint and not-started rows red.
- [ ] **Digital Life** shows **9** services and the red warning line above the table.
- [ ] **Key Contacts** shows **12**.
- [ ] **Dashboard** shows all 12 KPIs, the Readiness bars, the probate donut, the
      assets-by-value chart and **Readiness Score 90%**.
- [ ] Dropdowns work (how it transfers, asset category, document status, digital action,
      yes/no) — from **Settings**.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`TotalAssets`, `ProbateExposed`, `TotalDebts`, `NetEstate`, `ProbateShare`,
> `ProbateCost`, `CashReachable`, `Runway`, `DebtCover`, `BeneNamed`, `DocsSigned`,
> `DigitalCount`, `ContactCount`, `HealthRange`, etc.) resolve — Sheets occasionally
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
- The 12-page printable PDF
- A Start-Here quick-start guide

---

## 5. Three things to tell buyers plainly

> **This is not legal advice and it is not a will.** Nothing written in this file
> transfers anything to anyone — only properly executed documents and beneficiary
> designations do that. Probate rules, costs and timelines vary enormously by state and
> country; the cost figure is a rough percentage the buyer sets themselves. Real documents
> should be prepared by a qualified attorney. Say all of this in the listing, not just in
> the file.

> **Never store real passwords in it.** The Digital Life tab is deliberately designed to
> record *where each login lives* — a password manager, a sealed envelope, a safe — and
> not the login itself. Buyers will ask you to add a password column. Don't. A shared
> Google Sheet full of real credentials is a genuine harm, and it would be your template
> it happened in.

> **A finished file nobody knows about helps nobody.** The single most useful instruction
> in the whole product is on the Start Here tab: tell one person it exists and where it
> is. Put that line in your listing description too — it's also the line that sells it.

> Keep the delivery a plain digital download. It's fine to answer a buyer's question in
> Messages after a sale, but never advertise support, setup, document review, coaching, or
> "free updates" as part of the listing — Etsy's Services policy treats that as selling a
> service, and in this niche it also edges toward practising law.
