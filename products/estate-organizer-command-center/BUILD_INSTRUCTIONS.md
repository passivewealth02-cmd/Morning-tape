# Estate & Emergency Organizer Command Center™ — Build Instructions

---

## A. Google Sheets / Excel build

Requirements: Python 3.10+, `openpyxl >= 3.1.5`.

```bash
cd products/estate-organizer-command-center/build
python3 build_xlsx.py      # -> ../Estate_Organizer_Command_Center.xlsx  (14 tabs)
```

### Verifying
1. **Start Here** carries the full disclaimer block: an organizing tool, not legal/tax/
   financial advice, not a will, probate rules vary by state — **and never type real
   passwords into the file**. This block is not optional; it is why the product is safe
   to sell.
2. **Assets & Accounts** lists 12 assets summing to **$1,499,900** (`TotalAssets`), with
   a Probate? column; `SUMIF` gives **$535,500** exposed (`ProbateExposed`). Rows marked
   "Yes" turn red.
3. **Debts & Bills** sums to **$244,200** (`TotalDebts`) and **$2,400** monthly.
4. **Estate Snapshot**: **NET ESTATE $1,255,700** (`NetEstate`); **PROBATE SHARE 35.7%**
   (`ProbateShare`); × 5% = **ROUGH PROBATE COST $26,775** (`ProbateCost`).
5. Survivor half: joint + POD accounts = **CASH REACHABLE $45,400** (`CashReachable`) ÷
   $6,200/month = **7.3 MONTHS** (`Runway`). Life insurance $500,000 ÷ debts = **2.05×**
   (`DebtCover`). **This whole tab is the product's sales argument — check it renders.**
6. **Beneficiaries** lists 7 eligible accounts (`BeneEligible`) all with a primary named
   (`BeneNamed`); the Needs-action column flags red on "Yes".
7. **Legal Documents** counts **4 Signed** (`DocsSigned`) of 10, **6 Not started**, giving
   **40% readiness** (`DocReadiness`). Signed rows go mint, Not-started rows go red.
8. **Digital Life** documents **9** services (`DigitalCount`) and records *where each
   login lives* — a password manager, a sealed envelope, a safe — never the password. The
   warning line sits above the table on purpose.
9. **Key Contacts** lists **12** (`ContactCount`). **Medical & Care**, **Final Wishes** and
   **Household** are two-column reference tabs built by the shared `_two_col` helper.
10. **Dashboard** fills 12 KPI cards + a Readiness table + a probate donut and an
    assets-by-value bar chart. **Readiness Score 90%** (core documents signed is the
    honest weak dimension). No broken cells; custom tables start in column B.

> Note: uses `SUM`, `SUMIF`, `COUNTA`, `COUNTIF`, `MIN`, `AVERAGE`, `IF`, `IFERROR` —
> opens in Google Sheets or Excel 2019/365.

---

## B. Printable PDF build

```bash
python3 build_pdf.py       # -> ../Estate_Organizer_Printables.pdf  (12 pages, US Letter)
                           #    + page PNGs in ../marketing/print/
```

Ink-light forms on white with a forest-green header band. 300 DPI, US Letter. Twelve
pages: **if something happens — start here**, estate snapshot, assets & accounts, debts &
bills, beneficiary review, legal document checklist, insurance, digital life, key
contacts, medical & care, final wishes and household instructions.

**Page 1 is the most important page in the whole catalogue.** It is the page that sits at
the front of the binder: the first five calls in order, where the will and the fire safe
are, and a "take a breath first" note reminding whoever finds it that almost nothing is
urgent in the first week and to order ten certified death certificates. Feature it first
in the printables showcase.

Page 8 (Digital Life) prints the **do not write real passwords** warning in red above the
table. Keep it.

---

## C. Marketing images

```bash
python3 build_marketing.py         # -> ../marketing/01..06.png  (run build_pdf.py first)
python3 build_marketing_detail.py  # -> ../marketing/07..10.png
```

**Six app-screenshots**: hero (live dashboard with the probate donut and an
assets-by-value chart), everything-inside (14 tabs), the **estate snapshot / probate
engine**, the ten-documents checklist, the readiness engine (both), and the **12-page
printables showcase**. Images 3–5 each show a different tab.

**Four detailed images**: 07 feature spotlights, 08 "a printable binder vs Command
Center", 09 get-it-done in 4 steps, 10 what's-included / who-it's-for / works-with. Ten
images — fills all 10 Etsy slots. All headline numbers ($1,499,900 assets · $244,200 debts
· $1,255,700 net estate · $535,500 exposed · 35.7% share · $26,775 cost · $45,400
reachable · 7.3 months · $500,000 insurance · 4 of 10 documents · 21 accounts · 90% score)
are verified against the workbook. The probate donut splits 35.7% / 64.3% and the legend
figures ($535,500 / $964,400) sum to the $1,499,900 in assets.

The crest is a **lantern** — the light you leave on for the people who come after. It is
deliberately warm rather than sombre, and that tone should carry through every image: this
product sells on care, not fear.

---

## D. Etsy delivery package

```
Estate_Organizer_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
GOOGLE_SHEETS_TEMPLATE_LINK.txt        ← "Make a Copy" link
Estate_Organizer_Printables.pdf        ← 12-page print-ready pack
START_HERE.pdf                         ← onboarding quick-start
```

---

## E. Pricing  (all SKUs are the same digital files — a product, never a service)

| SKU | What the buyer downloads | Price |
| --- | ------ | ----- |
| EST-GS   | The Google Sheets / Excel file only | $22 |
| EST-PDF  | The printable PDF only | $19 |
| EST-BUNDLE | The spreadsheet + the printable PDF | **$29** |
| EST-COMM | The same files + a commercial-use file license | $49 |

> ⚠ **This listing needs more care than any other in the shop.** Two rules:
>
> 1. **Never imply legal validity.** Do not use the words "will kit", "legal documents",
>    "legally binding", "attorney-approved" or "estate plan". Say **organizer**,
>    **checklist**, **record**. Etsy will not pull you for organizing tools; it *will*
>    pull unlicensed legal-document sales, and so will several state bars.
> 2. **No services.** No setup help, consultations, **coaching**, document review, or
>    "free updates / lifetime access". Plain digital file only.
>
> The disclaimer belongs in the listing body *and* on the Start Here tab. It is already
> on the tab.

- **A large, calm, high-intent audience.** "In case of emergency binder" and "what my
  family should know" are established, high-volume Etsy searches, and the existing
  listings are almost all fill-in-the-blank PDFs at $8–$15. Yours computes.
- **Demand peaks in January** (new-year organizing) and again in **October–November**
  (open enrolment, and families together at the holidays deciding to sort this out).
- Use all 10 photos + a walkthrough video that shows the *file*. Lead photo = the
  feature-forward hero; **the probate snapshot is your single most persuasive image** —
  "$535,500 of your estate is heading for probate because nothing is named on it" is a
  fact almost no buyer has ever had put to them.
- **Sell it as kindness, not fear.** The best-performing angle here is "the kindest
  afternoon's work there is", not "if you died tomorrow". The lantern crest supports that.
- Cross-sell **Next Chapter™** (divorce organization) — adjacent life-transition buyer,
  no overlap.

---

## F. Maintenance

- Edit the `MONTHLY_NEED`, `PROBATE_PCT`, `LIFE_INSURANCE`, `CONTACT_GOAL`, `RUNWAY_GOAL`,
  `COVER_GOAL`, `DOC_GOAL` constants and the `ASSETS`, `DEBTS`, `BENEFICIARIES`,
  `DOCUMENTS`, `INSURANCE`, `DIGITAL`, `CONTACTS`, `MEDICAL`, `WISHES`, `HOUSEHOLD`,
  `REACHABLE` tables in `build_xlsx.py`; every KPI + the Readiness Score recompute.
- **Keep `REACHABLE` consistent with `ASSETS`** — it should be exactly the joint and
  payable-on-death rows. It currently is ($8,400 + $22,000 + $15,000 = $45,400).
- **Do not add a passwords column to the Digital Life tab**, however often buyers ask.
  Recording where a login lives is safe; recording the login is a liability you do not
  want attached to your shop.
- Keep `build_marketing.py`'s KPIs in sync with the workbook.
