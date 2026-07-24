# Subscription & Bills Audit Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/subscription-audit-command-center/build
python3 build_xlsx.py      # -> ../Subscription_Audit_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial advice" note.
2. **Subscriptions** sums 14 subs to **MONTHLY SUBS $216.77** (`SubMonthly`);
   `SubCount` = 14; `AnnualBilled` = 2 (via `COUNTIF` on "Yearly").
3. **Subscription Audit** shows **$2,601** annual, a **$105.92** cancel total
   (`CancelMonthly` via `SUMIF` on "Cancel"), **$1,271** annual savings, and
   **$110.85** kept (`KeepMonthly`).
4. **Bills** totals **MONTHLY BILLS $468** (`BillMonthly`).
5. **Price Hikes** add **$132**/yr (`HikeTotal`); **Free Trials** count **3**
   (`TrialCount`); the **Categories** tab rolls up spend by category via `SUMIF`.
6. **Dashboard** fills 12 KPI cards + an Audit Health table + keep-vs-cancel &
   recurring-by-month charts. **Audit Score 90%** (billed-annually is the honest weak
   dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `COUNTA`, `COUNTIF`, `AVERAGE`, `MIN`, `IF`, `IFERROR` —
> opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Subscription_Audit_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: subscription audit, cancel list, bills list, renewal calendar, free trial
tracker, price hike log, spend by category, bill negotiation, savings log, monthly
summary, audit worksheet and an audit checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the keep-vs-cancel & recurring
charts), everything-inside (14 tabs), the cancel-savings finder, the subscriptions
list, the audit engine (finder + list), and the **12-page printables showcase**. Images
3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "basic tracker vs Command Center",
09 audit-your-spend in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($216.77 monthly · $2,601 annual
· 14 subs · $105.92 cancel · $1,271 saved · $468 bills · $110.85 kept · $132 hikes · 3
trials · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Subscription_Audit_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt           ← "Make a Copy" link
Subscription_Audit_Printables.pdf         ← 12-page print-ready pack
START_HERE.pdf                            ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| SAC-GS   | The Google Sheets / Excel file only | $19 |
| SAC-PDF  | The printable PDF only | $19 |
| SAC-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| SAC-COMM | The same files + a commercial-use file license | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, "done-for-you" builds, or "free
> updates / lifetime access" — offering a service (rather than a finished file) is
> what gets a listing removed and earns a strike. A commercial-use license is a
> permission term attached to the *file* and is allowed; doing the work *for* the
> buyer is not.

- **Strong year-round demand** with a January (new-year money-reset) peak. "Cancel your
  subscriptions" is an evergreen money-saving search — the cancel-savings finder is the
  hook.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **cancel-savings finder** and the **renewal calendar** are
  your strongest differentiators — most listings are just a blank list.
- Cross-sell the **Net Worth & FIRE** and **Meal Planning** templates — same
  budget-minded buyer.

---

## F. Maintenance

- Edit the `SUBS`, `BILLS`, `CAT_ROLLUP`, `RENEWALS`, `TRIALS`, `HIKES`, `SAVINGS`,
  `NEGOTIATION`, `MONTHS` constants and the `SUB_BUDGET`, `SAVE_GOAL`, `ANNUAL_GOAL`
  targets in `build_xlsx.py`; every KPI + the Audit Score recompute. Everything is
  cross-linked — flag a subscription "Cancel" or add one and the savings, keep-total
  and score follow.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
