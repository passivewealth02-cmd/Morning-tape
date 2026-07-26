# Salon, Barber & Booth Renter Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Salon_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMPRODUCT`, `COUNTA`,
`COUNTIF`, `ROUNDUP`, `MIN`, `MAX`, `AVERAGE`, `IF`, `IFERROR`, named ranges, data
validation, conditional formatting, charts). Confirm:

- [ ] **Chair & Rent** shows **$1,150** fixed costs, **$7.19** rent per chair-hour,
      **70%** utilization and a **21-client** break-even at **4.4×** cover.
- [ ] **Service Pricing** shows **$56.31** service net, an **$8.99** rent load,
      **$47.32** you-actually-keep, **72.8%** true margin and **$37.86** an hour.
- [ ] **Services Menu** recalculates all 10 services and shows **$55.18** best per hour
      against **$37.35** worst. This is the tab buyers screenshot — make sure it survives
      the import.
- [ ] **Retail & Backbar** totals **28** units, **$672** revenue and an **11.2%** attach
      rate.
- [ ] **Income & Tips** shows **$6,652** revenue, **$4,339** profit, **65.2%** net margin
      and **$5,489** take-home with tips.
- [ ] **Rebooking & Retention** shows a **78%** rebooking rate and a **3.3%** no-show rate.
- [ ] **Dashboard** shows all 12 KPIs, the Chair Health bars, the where-the-$65-goes and
      revenue-by-month charts and **Chair Score 90%**.
- [ ] Dropdowns work (appointment status, service type, expense category, yes/no) — from
      **Settings**.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`ServicePrice`, `BackbarCost`, `CardFee`, `ServiceNet`, `RentPerHour`, `RentLoad`,
> `TrueNet`, `TrueMargin`, `BreakEven`, `CoverRatio`, `ClientsMonth`, `RebookRate`,
> `AttachRate`, `MonthlyProfit`, `TakeHome`, `HealthRange`, etc.) resolve — Sheets
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

> **A note on rates:** card processing rates vary by processor — Square, Stripe, Vagaro
> and GlossGenius all charge differently, and the sample uses 2.9% + 30¢. Say so plainly
> in your listing so buyers enter their own rate in Settings. Tips are taxable income;
> the printable pack includes a set-aside page for exactly that reason.

> Keep the delivery a plain digital download. It's fine to answer a buyer's question
> in Messages after a sale, but never advertise support, setup, pricing help, coaching,
> or "free updates" as part of the listing — Etsy's Services policy treats that as
> selling a service.
