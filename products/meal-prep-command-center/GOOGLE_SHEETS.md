# Meal Prep Business Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Meal_Prep_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`INDEX`, `MATCH`, `COUNTIF`,
`COUNTA`, `AVERAGE`, `MAX`, `MIN`, `IFERROR`, named ranges, data validation,
conditional formatting, charts). Confirm:

- [ ] **Meal Cost** totals **$5.00** per meal; ingredient-only cost is **$3.20**.
- [ ] **Meal Plans** show live margins (55% / 52% / 50% / 48%).
- [ ] **Subscribers** rolls up to **$4,475** weekly, **MRR $17,900**, **433** meals.
- [ ] **Dashboard** shows all 12 KPIs, the Prep Health bars, the MRR-by-month chart
      and **Prep Score 90%**.
- [ ] Dropdowns work (plan, unit, status, yes/no) — from the **Settings** lists.
- [ ] The **MRR by Month** chart renders.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`MealCost`, `IngredientCost`, `WeeklyRev`, `MRR`, `TotalSubs`, `MealsWeek`,
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
