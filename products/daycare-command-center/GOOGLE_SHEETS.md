# Daycare & Childcare Provider Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Daycare_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `SUMIF`, `COUNTA`, `COUNTIF`,
`ROUNDUP`, `MIN`, `MAX`, `AVERAGE`, `IF`, `IFERROR`, named ranges, data validation,
conditional formatting, charts). Confirm:

- [ ] **Costs & Expenses** totals **$3,825** fixed, **$157.00** per child and **$5,552**
      total costs.
- [ ] **Rate & Enrollment** shows **$1,060.85** tuition per child, **$903.85** net per
      child, **5** break-even children, **2.20×** cover and **91.7%** occupancy.
- [ ] The **empty-spot block** shows one open spot costing **$904 a month / $10,846 a
      year**. This is the block buyers screenshot; make sure it survives the import.
- [ ] **Children & Families** totals **$11,063.15** billed, and **Tuition & Payments**
      ties to the same figure with **$1,421.70** outstanding, flagged red.
- [ ] **Food Program** shows **$1,295.70** reimbursement covering **87%** of the food bill.
- [ ] **Ratios & Schedule** shows **5.5** children per caregiver and reads **COVERED**.
- [ ] **Compliance & Files** shows **7** complete and **4** missing something.
- [ ] **Tax Set-Aside** shows **$6,400** saved against **$16,000**.
- [ ] **Monthly Summary** shows **$12,539** revenue, **$6,987** your pay, **55.7%**
      margin, **$32.27** an hour and **$24.20** after the set-aside.
- [ ] **Dashboard** shows all 12 KPIs, the Program Health bars, the where-the-month-goes
      and revenue-by-month charts and **Care Score 90%**.
- [ ] Dropdowns work (schedule, payment status, age group, expense category, yes/no) —
      from **Settings**.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`TuitionPerChild`, `NetPerChild`, `BreakEven`, `CoverRatio`, `Occupancy`,
> `TuitionBilled`, `Outstanding`, `CACFP`, `FoodCoverage`, `ChildrenPerCaregiver`,
> `TaxReserve`, `Revenue`, `YourPay`, `PayPerHour`, `HealthRange`, etc.) resolve — Sheets
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

> **A note on state rules:** child-to-caregiver ratios, maximum group sizes, licensing
> requirements and CACFP reimbursement tiers all vary by state, and the CACFP rates
> change every July. The sample uses Tier I rates. Say so plainly in your listing so
> buyers enter their own in Settings — this is the single most important disclaimer on
> this product.

> **A note on the forms:** the enrollment form, emergency card and compliance checklist
> are organizing aids. They are not a substitute for the forms your licensing agency
> requires, and nothing here is legal, tax or licensing advice.

> Keep the delivery a plain digital download. It's fine to answer a buyer's question
> in Messages after a sale, but never advertise support, setup, rate-setting help,
> coaching, or "free updates" as part of the listing — Etsy's Services policy treats that
> as selling a service.
