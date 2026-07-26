# Salon, Barber & Booth Renter Command Center™ — Price the Chair, Not Just the Cut

> Not a price list — a **complete price-it, book-it, keep-it system**.
> One premium **Google Sheets + printable PDF** command center for stylists, barbers and
> booth renters: a true-ticket engine (price − product − card fee − **the rent your chair
> charges by the hour** → what you actually keep), a chair break-even, your whole menu
> costed by the hour, a client book, appointments, retail & backbar, income & tips,
> expenses, rebooking & retention, product inventory and a monthly summary — everything
> cross-linked and live.

| | |
| - | - |
| **Product** | Salon, Barber & Booth Renter Command Center™ |
| **Target** | Booth & suite renters · hairstylists and colourists · barbers · lash, brow, nail & esthetics techs · salon owners with a team · home-studio stylists |
| **Angle** | An empty chair still charges rent — so every service owes its share. |
| **Formats** | Google Sheets (14-tab system) + Excel `.xlsx` edition + 12-page printable PDF |
| **Pricing** | $22 single file · **$29 bundle** (Sheets + PDF) · $49 with a commercial-use file license — all the *same digital files*, never a service |

---

## Contents

```
products/salon-command-center/
├── README.md
├── Salon_Command_Center.xlsx     ← Google Sheets / Excel master (14 tabs)
├── Salon_Printables.pdf          ← 12-page print-ready pack (US Letter)
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
| 1 | Start Here | 8 | Retail & Backbar |
| 2 | Dashboard | 9 | Income & Tips |
| 3 | Service Pricing | 10 | Expenses |
| 4 | Chair & Rent | 11 | Rebooking & Retention |
| 5 | Services Menu | 12 | Inventory |
| 6 | Client Book | 13 | Monthly Summary |
| 7 | Appointments | 14 | Settings |

## The 12 printable PDF pages

Service Pricing Sheet · Chair Cost & Break-Even · Service Menu Costed by the Hour · New
Client Consultation · Colour Formula Card · Client Card · Day Sheet · Rebooking &
Retention · Retail Sales Log · Product Inventory & Reorder · Income, Tips & Tax
Set-Aside · Monthly Summary & Chair Review.

---

## Signature automation — the true-ticket engine

Every stylist subtracts product. Almost none subtract **the chair**. Your booth rent runs
whether that seat is full or empty, so it is a cost **per hour**, and every service owes
its share of it:

```
Card fee     = price × card rate + per-swipe fee
Service net  = price − backbar − card fee
Rent/hour    = fixed monthly costs ÷ hours the chair is open
Rent load    = rent/hour × hours this service takes
YOU KEEP     = service net − rent load
True margin  = you keep ÷ price
Real $/hour  = you keep ÷ hours this service takes
Break-even   = fixed costs ÷ service net      ← clients/month just to cover the chair
```

The **Services Menu** tab runs that math down your whole price list, and the result is
usually a shock: on the sample menu the **$35 men's cut pays $55.18 an hour** while the
**$95 root touch-up pays $37.35** — the cheap service is the better service, and no price
list on earth tells you that.

### The 12 dashboard KPIs
Ticket Price · Backbar · Card Fee · Rent/Chair-Hour · You Actually Keep · True Margin ·
Clients/Month · Monthly Revenue · Monthly Profit · Break-Even Clients · Chair
Utilization · Chair Score.
The **Chair Score** blends true margin, rebooking, retail attach, break-even cover,
no-shows and new-client flow into one 0–100% number.

**Verified sample studio** (Gilded Chair Studio, stylist Remi, flagship women's cut &
style): ticket **$65.00** · backbar **$6.50** · card fee **$2.19** · rent/chair-hour
**$7.19** × 1.25 hrs = **$8.99** → **you keep $47.32** · true margin **72.8%** · real rate
**$37.86/hr** · break-even **21 clients** · served **92** · revenue **$6,652** · profit
**$4,339** · tips $1,150 → take-home **$5,489** · utilization **70%** · **Chair Score 90%**
(the honest weak spot: 4 new clients a month against a goal of 10).

---

## Premium chair design

- A **true-ticket engine** that charges rent by the hour, not just product
- Your **whole menu costed per hour** — losers flagged red, best/worst called out
- A **break-even client count**: how many people just to cover the chair
- **Retail attach rate** — the highest-margin thing in the room, and it costs no chair time
- **Rebooking, no-shows, tips and inventory** all rolling into one score
- Brand palette: Primary `#1B4F48`, Gold `#937356`, Surface `#E5D3BA`,
  Mint `#75E6C1`, Ivory `#FBF8F2`

> **A business & organizing tool, not financial, tax or accounting advice.** Card
> processing rates vary by processor, and tips are taxable income — set some aside.

---

## Build & ship

```bash
cd build && python3 build_xlsx.py          # -> ../Salon_Command_Center.xlsx
python3 build_pdf.py                        # -> ../Salon_Printables.pdf (+ page PNGs)
python3 build_marketing.py                  # -> ../marketing/01..06.png
python3 build_marketing_detail.py           # -> ../marketing/07..10.png
```

See `BUILD_INSTRUCTIONS.md`, `GOOGLE_SHEETS.md` and `ETSY_LISTING.md`.
