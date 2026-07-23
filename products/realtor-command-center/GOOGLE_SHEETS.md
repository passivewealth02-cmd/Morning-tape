# Real Estate Agent Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Realtor_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `COUNTA`, `AVERAGE`, `MIN`,
`MAX`, `ROUNDUP`, `IFERROR`, named ranges, data validation, conditional formatting,
charts). Confirm:

- [ ] **Commission Calculator** shows GCI **$13,500**, agent commission **$9,450** and
      net per deal **$8,000**.
- [ ] **Closings** rolls up to **$75,600** GCI and **$3.6M** volume.
- [ ] **Business Expenses** shows a **$50,000** net income.
- [ ] **Dashboard** shows all 12 KPIs, the Agent Health bars, the GCI-by-month chart
      and **Agent Score 90%**.
- [ ] Dropdowns work (side, stage, source, yes/no) — from **Settings**.
- [ ] The **GCI by Month** chart renders.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`GCIDeal`, `AgentComm`, `NetDeal`, `GCI_YTD`, `NetIncome`, `PipelineGCI`,
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
