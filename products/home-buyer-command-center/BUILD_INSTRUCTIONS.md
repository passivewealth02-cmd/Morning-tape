# First-Time Home Buyer & Mortgage Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/home-buyer-command-center/build
python3 build_xlsx.py      # -> ../Home_Buyer_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** shows the intro + a 6-step guide + a "not financial/lending advice"
   note.
2. **Affordability** takes a **$290,000** price at **8%** down → **$23,200**
   (`DownPayment`) and a **$266,800** loan (`LoanAmount`). `PMT` at 6.5% / 30 yr gives
   **$1,686** (`PrincipalInterest`); + tax **$266** + insurance **$150** + PMI **$111**
   = **MONTHLY PAYMENT $2,213** (`PITI`).
3. DTI: **FRONT-END 28.0%** (`FrontDTI`) and **BACK-END 33.6%** (`BackDTI`) against a
   $7,917 gross monthly income — verdict **COMFORTABLE**. Total interest **$340,289**.
4. **Closing Costs** totals **$8,700** (`ClosingTotal`); **Down Payment** shows **CASH
   TO CLOSE $31,900** (`CashToClose`) against **$32,000** saved (`Saved`).
5. **Home Comparison** scores 5 homes (best 92); **Lender Compare** recalculates the
   payment per lender rate via `PMT`.
6. **Dashboard** fills 12 KPI cards + a Buyer Health table + where-the-payment-goes &
   saved-by-month charts. **Buyer Score 90%** (20%-down is the honest weak dimension —
   the sample buyer pays PMI). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `PMT`, `COUNTA`, `COUNTIF`, `AVERAGE`, `MIN`, `MAX`, `CEILING`,
> `IF`, `AND`, `IFERROR` — opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Home_Buyer_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter.
Twelve pages: affordability worksheet, home comparison, closing-cost checklist, down
payment tracker, lender comparison, house hunting log, must-have list, credit prep, life
after buying, moving checklist, amortization snapshot and a closing-day checklist.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the payment-split & savings charts),
everything-inside (14 tabs), the affordability engine, the home comparison, the home
buyer engine (affordability + homes), and the **12-page printables showcase**. Images
3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "free calculator vs Command Center",
09 buy-your-first-home in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($290k price · $23.2k down ·
$266.8k loan · $1,686 P&I · $2,213 PITI · 28.0% / 33.6% DTI · $8,700 closing · $31,900
cash to close · $340,289 interest · 90% score) are verified against the workbook.

---

## D. Etsy delivery package

```
Home_Buyer_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt     ← "Make a Copy" link
Home_Buyer_Printables.pdf           ← 12-page print-ready pack
START_HERE.pdf                      ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| HBC-GS   | The Google Sheets / Excel file only | $19 |
| HBC-PDF  | The printable PDF only | $19 |
| HBC-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| HBC-COMM | The same files + a commercial-use file license (great for agents & lenders) | $49 |

> Etsy Services-policy note: keep every listing a plain digital download. Do not
> advertise setup help, custom work, consultations, mortgage or lending advice,
> "done-for-you" builds, or "free updates / lifetime access" — offering a service
> (rather than a finished file) is what gets a listing removed and earns a strike.
> Lending advice also carries regulatory risk of its own; the workbook and every doc
> state clearly that this is not financial or lending advice.

- **Spring/summer buying season is the peak**, with a January "this is the year we buy"
  surge. Emotional, high-stakes purchase — buyers happily pay $29 for clarity on the
  biggest cheque of their life.
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; the **all-in PITI** and the **DTI verdict** are your strongest
  differentiators — every free calculator omits taxes, insurance and PMI.
- The **HBC-COMM** license tier sells well to **real-estate agents and loan officers**
  who want to hand it to clients — same file, permission term attached.
- Cross-sell the **Net Worth & FIRE**, **Budget** and **Renovation Budget** templates.

---

## F. Maintenance

- Edit the `HOME_PRICE`, `DOWN_PCT`, `RATE`, `TERM_YEARS`, `TAX_RATE`,
  `INSURANCE_ANNUAL`, `PMI_RATE`, `ANNUAL_INCOME`, `OTHER_DEBTS`, `CLOSING_PCT`,
  `SAVED`, `CREDIT_SCORE` constants and the `FRONT_GOAL`, `BACK_GOAL`, `DOWN_GOAL`,
  `CREDIT_GOAL`, `EF_GOAL` targets plus the `CLOSING_ITEMS`, `HOMES`, `LENDERS`,
  `AMORT`, `AFTER_BUDGET`, `CREDIT_PREP`, `HUNT_LOG`, `MOVING`, `MONTHS` tables in
  `build_xlsx.py`; every KPI + the Buyer Score recompute.
- **Refresh `RATE` and the sample lender rates periodically** so screenshots stay
  credible as the market moves — a stale rate dates the listing instantly.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
