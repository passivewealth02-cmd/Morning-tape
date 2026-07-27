# Daycare & Childcare Provider Command Center™ — You're Not Babysitting

> Not a sign-in sheet — a **complete price-it, fill-it, keep-it system**.
> One premium **Google Sheets + printable PDF** command center for home childcare
> providers: a rate & break-even engine (weekly rate → **net per child** → how many
> children it takes just to cover the house), children & families, tuition & payments with
> who's behind, attendance, costs, the CACFP food program, ratios by age group, compliance
> files, supplies, a tax set-aside and a monthly summary — everything cross-linked and
> live.

| | |
| - | - |
| **Product** | Daycare & Childcare Provider Command Center™ |
| **Target** | Licensed home daycare providers · in-home & family childcare · small centres and preschools · nannies & nanny-share providers · providers opening their first program · anyone raising their rates this year |
| **Angle** | The tuition that came in is not your pay. What's left after everything is your pay. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $24 single file · **$32 bundle** (Sheets + PDF) · $55 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/daycare-command-center/
├── README.md
├── Daycare_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Daycare_Printables.pdf          ← 12-page print-ready pack (US Letter)
├── GOOGLE_SHEETS.md
├── BUILD_INSTRUCTIONS.md
├── ETSY_LISTING.md
└── build/
    ├── build_xlsx.py
    ├── build_pdf.py
    ├── build_marketing.py
    └── build_marketing_detail.py
```

---

## The 14-tab system

| # | Tab | # | Tab |
| - | --- | - | --- |
| 1 | Start Here | 8 | Food Program |
| 2 | Dashboard | 9 | Ratios & Schedule |
| 3 | Rate & Enrollment | 10 | Compliance & Files |
| 4 | Children & Families | 11 | Supplies |
| 5 | Tuition & Payments | 12 | Tax Set-Aside |
| 6 | Attendance | 13 | Monthly Summary |
| 7 | Costs & Expenses | 14 | Settings |

## The 12 printable PDF pages

Rate & Break-Even Worksheet · Monthly Costs Worksheet · Enrollment Form · Emergency &
Allergy Card · Daily Sheet · Attendance Sheet · Tuition & Payment Log · CACFP Meal Count ·
Ratios & Daily Schedule · Compliance File Checklist · Tax Set-Aside · Monthly Summary &
Review.

---

## Signature automation — the rate & break-even engine

Home childcare providers are some of the hardest-working business owners there are, and
almost none of them know these two numbers:

```
Tuition per child = weekly rate × weeks per month
Net per child     = tuition per child − food − supplies & activities
BREAK-EVEN        = fixed costs ÷ net per child     ← children who only pay the bills
Your pay          = revenue − fixed costs − per-child costs
Pay per hour      = your pay ÷ hours actually open
Empty spot cost   = (capacity − enrolled) × net per child
```

On the sample program that maths out to something most providers have never seen written
down: **five children only pay the bills.** Everyone enrolled after that is the provider's
actual income — and **one empty spot costs $10,846 a year**, which reframes a "quiet
month" as the most expensive thing in the program.

The other line the workbook insists on: your pay is **$32.27 an hour** open, but a quarter
of that belongs to tax you haven't paid yet, so what you really keep is **$24.20**.

### The 12 dashboard KPIs
Weekly Rate · Tuition/Child · Cost/Child · Net Per Child · Enrolled · Occupancy · Monthly
Revenue · Monthly Costs · Your Pay · Break-Even Children · Your $/Hour · Care Score.
The **Care Score** blends spots filled, margin, break-even cover, ratios, your hourly rate
and your tax reserve into one 0–100% number.

**Verified sample program** (Little Acorns Home Childcare, provider Dana): 11 enrolled of
12 licensed (**91.7%**) · weekly rate **$245** → **$1,061**/month · cost per child
**$157** → **net per child $904** · fixed costs **$3,825** → **break-even 5 children**
(covered **2.2×**) · revenue **$12,539** (tuition $11,063 + fees $180 + CACFP $1,296) ·
costs **$5,552** · **your pay $6,987** (**55.7%** margin, **$83,842**/yr) · **$32.27/hour**
open → **$24.20** after tax set-aside · ratio **5.5:1** · **Care Score 90%** (the honest
weak spot: a tax reserve only 40% funded).

---

## Premium provider design

- A **break-even child count** — the number nobody has ever shown them
- **What one empty spot costs**, in dollars per year, not vibes
- **Your pay per hour**, then again after the tax set-aside
- **CACFP meal count** at Tier I rates, showing the food program covers **87%** of the food bill
- **Ratios by age group** with a live COVERED / SHORT STAFFED flag
- **Compliance files** flagged incomplete before an inspector finds them
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business & organizing tool, not financial, tax, legal or licensing advice.** Ratios,
> group sizes, licensing rules and CACFP reimbursement tiers vary by state — enter your
> own in Settings.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Daycare_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Daycare_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
