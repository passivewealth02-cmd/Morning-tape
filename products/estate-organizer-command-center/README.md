# Estate & Emergency Organizer Command Center™ — Everything They'd Need to Find

> Not a binder — a **complete list-it, name-it, tell-them system**.
> One premium **Google Sheets + printable PDF** command center for getting your affairs in
> order: an estate snapshot (assets − debts → net estate, and **the share heading for
> probate because nothing is named on it**), assets & accounts, debts & bills,
> beneficiaries, the ten legal documents, insurance, digital life, key contacts, medical &
> care wishes, final wishes and household instructions — everything cross-linked and live.

| | |
| - | - |
| **Product** | Estate & Emergency Organizer Command Center™ |
| **Target** | Anyone getting their affairs in order · retirees and near-retirees · adult children helping a parent · the newly widowed sorting an estate · anyone with a will and nothing else · caregivers & family executors |
| **Angle** | A beneficiary form takes ten minutes and skips probate entirely. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $22 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/estate-organizer-command-center/
├── README.md
├── Estate_Organizer_Command_Center.xlsx   ← Google Sheets / Excel master (14 tabs)
├── Estate_Organizer_Printables.pdf        ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Insurance |
| 2 | Dashboard | 9 | Digital Life |
| 3 | Estate Snapshot | 10 | Key Contacts |
| 4 | Assets & Accounts | 11 | Medical & Care |
| 5 | Debts & Bills | 12 | Final Wishes |
| 6 | Beneficiaries | 13 | Household |
| 7 | Legal Documents | 14 | Settings |

## The 12 printable PDF pages

If Something Happens — Start Here · Estate Snapshot · Assets & Accounts · Debts & Bills ·
Beneficiary Review · Legal Document Checklist · Insurance Policies · Digital Life · Key
Contacts · Medical & Care · Final Wishes · Household Instructions.

---

## Signature automation — the probate exposure engine

Every estate binder on the market is a place to *write things down*. This one does
arithmetic, and the arithmetic is the product:

```
Net estate       = total assets − total debts
Probate-exposed  = assets with no beneficiary, no joint owner, no trust
Probate share    = probate-exposed ÷ total assets
Rough cost       = probate-exposed × your probate % estimate
Cash reachable   = joint + payable-on-death accounts   ← available in days
Runway           = cash reachable ÷ monthly household cost
Debt cover       = life insurance ÷ total debts
```

On the sample household that produces the line nobody expects: **35.7% of the estate —
$535,500 — would go through probate**, at a rough cost of **$26,775** and months of delay,
purely because a brokerage account, the house and two vehicles have nothing named on them.
Meanwhile $964,400 skips probate entirely, because someone spent ten minutes on a
beneficiary form.

The second engine is the "emergency" half: **$45,400 is reachable in days, not months** —
enough to keep the household running **7.3 months** before anything else is released.

### The 12 dashboard KPIs
Total Assets · Total Debts · Net Estate · Exposed to Probate · Probate Share · Rough
Probate Cost · Cash They Can Reach · Months of Runway · Life Insurance · Documents Signed ·
Accounts Documented · Readiness Score.
The **Readiness Score** blends survivor runway, beneficiaries named, insurance cover, key
contacts, digital life documented and signed documents into one 0–100% number.

**Verified sample household** (Lantern & Oak, prepared by Margot): assets **$1,499,900** ·
debts **$244,200** · **net estate $1,255,700** · **probate-exposed $535,500 (35.7%)** ·
rough probate cost **$26,775** · cash reachable **$45,400** → **7.3 months** of runway ·
life insurance **$500,000** covering debts **2.05×** · beneficiaries named on **7 of 7**
eligible accounts · **21** accounts documented · **Readiness Score 90%** (the honest weak
spot: only **4 of the 10** core documents signed).

---

## Premium organizing design

- A **probate exposure engine** — the one thing no estate binder does
- **Cash a survivor can actually reach**, and how long it lasts
- **Beneficiary gaps flagged red**, because those forms beat your will
- The **ten documents scored honestly** — most people have a will and nothing else
- **Digital Life records where each login lives, never the password itself**
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **An organizing tool, not legal, tax or financial advice, and not a will.** Nothing
> written here transfers anything to anyone — only properly executed documents and
> beneficiary designations do that. Probate rules, costs and timelines vary enormously by
> state and country. Please have real documents prepared by a qualified attorney.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Estate_Organizer_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Estate_Organizer_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
