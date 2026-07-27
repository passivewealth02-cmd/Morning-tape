# Notary & Loan Signing Agent Command Center™ — That $125 Signing Is Not $138 an Hour

> Not a mileage log — a **complete price-it, run-it, keep-it system**.
> One premium **Google Sheets + printable PDF** command center for mobile notaries and
> loan signing agents: a real-hourly profit engine (fee − printing − driving, over the
> hours it **actually** takes door to door), a signings log, a written fee schedule, a
> mileage log with the IRS deduction, printing costs, invoices and who pays late, signing
> companies, a notarial journal, expenses, a tax set-aside and a monthly summary —
> everything cross-linked and live.

| | |
| - | - |
| **Product** | Notary & Loan Signing Agent Command Center™ |
| **Target** | Loan signing agents · mobile notaries · new NSAs pricing their first year · notaries adding signings to an existing business · agents on multiple signing platforms · anyone who suspects they're underpriced |
| **Angle** | Count the drive and the printer — they're the two costs nobody puts in the fee. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $19 single file · **$25 bundle** (Sheets + PDF) · $45 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/notary-command-center/
├── README.md
├── Notary_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Notary_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Invoices |
| 2 | Dashboard | 9 | Signing Companies |
| 3 | Signing Profit | 10 | Notarial Journal |
| 4 | Signings Log | 11 | Expenses |
| 5 | Fee Schedule | 12 | Tax Set-Aside |
| 6 | Mileage Log | 13 | Monthly Summary |
| 7 | Printing & Supplies | 14 | Settings |

## The 12 printable PDF pages

What This Signing Really Pays · Fee Schedule · Signings Log · Mileage Log · Notarial
Journal Sheet · Signing Day Checklist · Invoice · Who Owes You · Signing Companies ·
Printing & Supplies · Tax Set-Aside · Monthly Summary & Review.

---

## Signature automation — the real-hourly engine

Here is the arithmetic every new signing agent gets wrong, and it's the whole product:

```
What it feels like = fee ÷ appointment length
Printing           = pages × cost per page
Driving            = miles × vehicle cost per mile
Net per signing    = fee − printing − driving
Hours door to door = prep + drive (both ways) + appointment
REAL HOURLY        = net per signing ÷ hours door to door
Mileage deduction  = miles × IRS rate
```

A $125 refinance is a 45-minute appointment, so it feels like **$138.89 an hour**. But you
printed 180 pages before you left, drove 38 miles round trip, and the job took **2.5 hours
door to door**. The real number is **$44.14 an hour** — a **$94.75/hour** gap.

That's still a good business. It's just a completely different business from the one in
the buyer's head, and seeing it is what lets them price properly, add surcharges, and turn
down the $65 offers.

### The 12 dashboard KPIs
Average Fee · Looks Like/Hour · Printing · Driving · Net Per Signing · Really Earns/Hour ·
Signings/Month · Monthly Revenue · Monthly Profit · Break-Even Signings · Mileage
Deduction · Signing Score.
The **Signing Score** blends real hourly rate, margin, overhead cover, getting paid,
signings booked and the tax reserve into one 0–100% number.

**Verified sample business** (Quill & Seal Notary, agent Sloane): fee **$125.00** · feels
like **$138.89/hr** · printing **$6.30** (180 pages) · driving **$8.36** (38 miles) → **net
$110.34** over **2.5 hours** = **$44.14/hour** · 52 signings → revenue **$6,500** · profit
**$5,449** (**83.8%** margin, **$65,384**/yr) · fixed costs only **$289**, so **3 signings**
cover the whole overhead · mileage deduction **$1,383**/month · **Signing Score 90%** (the
honest weak spot: a tax reserve only 40% funded).

---

## Premium signing-agent design

- A **real hourly rate** that counts prep, both drive legs and the appointment
- **Printing and driving costed per signing**, the two lines nobody bills for
- A **written fee schedule** including the trip fee for a door cancellation
- **Mileage logged with the IRS deduction** — worth more than the gas costs you
- **Signing companies ranked** by average fee and payment terms, slow payers named
- A **notarial journal** tab, with a clear warning about state bound-journal rules
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business & organizing tool, not legal, tax or accounting advice.** Two things vary
> by state and the buyer must check theirs: the **maximum fee** allowed per notarial act,
> and whether a **bound sequential journal** is required. The journal tab here is a
> convenience record, not a substitute for one.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Notary_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Notary_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
