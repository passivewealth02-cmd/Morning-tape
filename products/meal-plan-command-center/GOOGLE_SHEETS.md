# Meal Planning & Grocery Budget Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Meal_Plan_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `COUNTA`, `AVERAGE`, `MIN`,
`IF`, `IFERROR`, named ranges, data validation, conditional formatting, charts).
Confirm:

- [ ] **Cost Per Meal** shows a **$10.00** recipe cost, **4** servings, a **$2.50**
      cost per serving and a **$10.00** saved per serving.
- [ ] **Weekly Plan** totals **$66** across **7** planned dinners.
- [ ] **Budget** shows **$540** spent against a **$600** budget (90% used).
- [ ] **Savings** shows a **$960** monthly savings vs eating out.
- [ ] **Dashboard** shows all 12 KPIs, the Kitchen Health bars, the cook-vs-eat-out
      and saved-by-month charts and **Kitchen Score 90%**.
- [ ] Dropdowns work (category, store, meal type, yes/no) — from **Settings**.
- [ ] The **Saved by Month** chart renders.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`RecipeCost`, `CostPerServing`, `SavedServing`, `WeeklyPlanCost`, `SpentMonth`,
> `MonthlySavings`, `HealthRange`, etc.) resolve — Sheets occasionally re-scopes on
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

> Keep the delivery a plain digital download. It's fine to answer a buyer's question
> in Messages after a sale, but never advertise support, setup, or "free updates" as
> part of the listing — Etsy's Services policy treats that as selling a service.
