"""Printable PDF pack for Salon, Barber & Booth Renter Command Center™ (12 pages, US Letter)."""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFont

PRIMARY = (27, 79, 72); ACCENT = (147, 115, 86); GOLD_LT = (201, 168, 106); GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186); WHITE = (255, 255, 255); TEXT = (51, 51, 51); MUTED = (120, 114, 104)
LINE = (200, 194, 182); ROW_ALT = (247, 243, 236)
W, H = 2550, 3300
M = 190
SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"


def fs(s, bold=True):
    return ImageFont.truetype(SANS_B if bold else SANS_R, s)


def fserif(s):
    return ImageFont.truetype(SERIF_B, s)


def page(title, subtitle):
    img = Image.new("RGB", (W, H), WHITE); d = ImageDraw.Draw(img)
    bh = 340
    d.rectangle((0, 0, W, bh), fill=PRIMARY)
    d.rectangle((0, bh, W, bh + 10), fill=GOLD_LT); d.rectangle((0, bh + 10, W, bh + 14), fill=GOLD_HI)
    d.text((M, 90), "SALON, BARBER & BOOTH RENTER COMMAND CENTER™", font=fs(26), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Stylist / Date: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Salon Command Center™", font=fs(22, bold=False), fill=ACCENT, anchor="rt")
    return img, d


def field(d, x, y, w, label):
    d.text((x, y), label.upper(), font=fs(24), fill=ACCENT, anchor="lt")
    d.line((x, y + 62, x + w, y + 62), fill=LINE, width=2)


def section(d, x, y, w, text):
    d.rounded_rectangle((x, y, x + w, y + 56), radius=8, fill=SURFACE)
    d.text((x + 20, y + 28), text.upper(), font=fs(26), fill=PRIMARY, anchor="lm")


def checkbox(d, x, y, size=48):
    d.rounded_rectangle((x, y, x + size, y + size), radius=8, outline=PRIMARY, width=3)


def table(d, x, y, w, headers, colf, nrows, rowh=92, filled_rows=None):
    filled_rows = filled_rows or []
    colx = [x + w * f for f in colf]; hh = 78
    d.rounded_rectangle((x, y, x + w, y + hh), radius=8, fill=PRIMARY)
    for i, h in enumerate(headers):
        d.text((colx[i] + (18 if i == 0 else 0), y + hh / 2), h, font=fs(26), fill=WHITE, anchor="lm" if i == 0 else "mm")
    for r in range(nrows):
        ry = y + hh + r * rowh
        if r % 2:
            d.rectangle((x, ry, x + w, ry + rowh), fill=ROW_ALT)
        d.line((x, ry + rowh, x + w, ry + rowh), fill=LINE, width=2)
        if r < len(filled_rows):
            for ci, val in enumerate(filled_rows[r]):
                d.text((colx[ci] + (18 if ci == 0 else 0), ry + rowh / 2), str(val), font=fs(26, bold=(ci == 0)),
                       fill=PRIMARY if ci == 0 else TEXT, anchor="lm" if ci == 0 else "mm")
    d.rectangle((x, y, x + w, y + hh + nrows * rowh), outline=LINE, width=2)
    for cx in colx[1:]:
        d.line((cx, y, cx, y + hh + nrows * rowh), fill=LINE, width=1)


CW = W - 2 * M; TOP = 430


def p01(i):
    img, d = page("Service Pricing Sheet", "What the ticket says vs what you keep")
    field(d, M, TOP, CW / 2 - 30, "Service"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Date")
    table(d, M, TOP + 150, CW, ["THE TICKET", "AMOUNT"], [0.0, 0.78], 7,
          filled_rows=[["Service price", ""], ["− Backbar / product", ""], ["− Card processing fee", ""],
                       ["= SERVICE NET", ""], ["− Rent this service owes", ""], ["= YOU ACTUALLY KEEP", ""],
                       ["= TRUE MARGIN %", ""]], rowh=200)
    y = TOP + 150 + 78 + 7 * 200 + 40
    field(d, M, y, CW / 3 - 20, "Hours it takes"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Rent / chair-hour")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Your real $ / hour")
    y2 = y + 210
    section(d, M, y2, CW, "Before you price the next one")
    notes = ["Your chair charges rent by the hour, booked or not — every service owes its share.",
             "A $5 raise is almost pure profit. A $5 discount is almost pure loss.",
             "If the true margin is under your goal, the answer is price or time — never volume."]
    for j, n in enumerate(notes):
        d.text((M + 24, y2 + 110 + j * 84), "•  " + n, font=fs(30, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p02(i):
    img, d = page("Chair Cost & Break-Even", "An empty chair still charges rent")
    section(d, M, TOP, CW, "What the chair costs every month")
    table(d, M, TOP + 96, CW, ["FIXED COST LINE", "MONTHLY"], [0.0, 0.78], 10, 168)
    y = TOP + 96 + 78 + 10 * 168 + 36
    field(d, M, y, CW / 3 - 20, "Total fixed / month"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Hours chair is open")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Rent / chair-hour")
    y2 = y + 170
    field(d, M, y2, CW / 3 - 20, "Service net"); field(d, M + CW / 3 + 10, y2, CW / 3 - 20, "= Break-even clients")
    field(d, M + 2 * CW / 3 + 20, y2, CW / 3 - 20, "…per week")
    i.append(img)


def p03(i):
    img, d = page("Service Menu — Costed by the Hour", "Some services aren't worth the chair")
    table(d, M, TOP, CW, ["SERVICE", "PRICE", "MINS", "BACKBAR", "YOU KEEP", "PER HOUR"],
          [0.0, 0.36, 0.50, 0.62, 0.76, 0.90], 18, 128)
    i.append(img)


def p04(i):
    img, d = page("New Client Consultation", "Ask before you cut")
    field(d, M, TOP, CW / 2 - 30, "Client name"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Date / Referred by")
    field(d, M, TOP + 150, CW / 2 - 30, "Phone"); field(d, M + CW / 2 + 30, TOP + 150, CW / 2 - 30, "Email")
    section(d, M, TOP + 320, CW, "Hair history — last service, box color, chemical treatments")
    for j in range(5):
        d.line((M + 20, TOP + 460 + j * 92, W - M - 20, TOP + 460 + j * 92), fill=LINE, width=2)
    y = TOP + 460 + 5 * 92 + 40
    section(d, M, y, CW, "What they want today — and what they DON'T want")
    for j in range(5):
        d.line((M + 20, y + 140 + j * 92, W - M - 20, y + 140 + j * 92), fill=LINE, width=2)
    y2 = y + 140 + 5 * 92 + 40
    field(d, M, y2, CW / 3 - 20, "Allergies / patch test"); field(d, M + CW / 3 + 10, y2, CW / 3 - 20, "Home routine")
    field(d, M + 2 * CW / 3 + 20, y2, CW / 3 - 20, "Budget range")
    i.append(img)


def p05(i):
    img, d = page("Color Formula Card", "So the second visit matches the first")
    field(d, M, TOP, CW / 2 - 30, "Client"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Date")
    table(d, M, TOP + 150, CW, ["AREA", "FORMULA", "DEVELOPER", "TIME"], [0.0, 0.24, 0.62, 0.82], 9, 178)
    y = TOP + 150 + 78 + 9 * 178 + 36
    field(d, M, y, CW / 2 - 30, "Toner / gloss"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Processing notes")
    y2 = y + 170
    field(d, M, y2, CW / 3 - 20, "Result"); field(d, M + CW / 3 + 10, y2, CW / 3 - 20, "Adjust next time")
    field(d, M + 2 * CW / 3 + 20, y2, CW / 3 - 20, "Rebook in (weeks)")
    i.append(img)


def p06(i):
    img, d = page("Client Card", "What they're worth, at a glance")
    field(d, M, TOP, CW / 2 - 30, "Client"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Since")
    table(d, M, TOP + 150, CW, ["DATE", "SERVICE", "TICKET", "RETAIL", "TIP"],
          [0.0, 0.18, 0.56, 0.72, 0.88], 17, 132)
    y = TOP + 150 + 78 + 17 * 132 + 30
    field(d, M, y, CW / 3 - 20, "Visits this year"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Spend this year")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Average ticket")
    i.append(img)


def p07(i):
    img, d = page("Day Sheet", "Every hour of the chair, accounted for")
    field(d, M, TOP, CW / 2 - 30, "Date"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Hours open")
    table(d, M, TOP + 150, CW, ["TIME", "CLIENT", "SERVICE", "TICKET", "RETAIL", "TIP"],
          [0.0, 0.14, 0.38, 0.62, 0.76, 0.90], 15, 132)
    y = TOP + 150 + 78 + 15 * 132 + 30
    field(d, M, y, CW / 3 - 20, "Clients today"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Hours booked")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Day total")
    i.append(img)


def p08(i):
    img, d = page("Rebooking & Retention", "The cheapest marketing there is")
    table(d, M, TOP, CW, ["CLIENT", "LAST IN", "REBOOKED?", "NEXT DATE", "NOTES"],
          [0.0, 0.30, 0.44, 0.60, 0.76], 18, 128)
    y = TOP + 78 + 18 * 128 + 30
    field(d, M, y, CW / 3 - 20, "Clients served"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Rebooked")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Rebooking rate %")
    i.append(img)


def p09(i):
    img, d = page("Retail Sales Log", "Profit with no chair time attached")
    table(d, M, TOP, CW, ["DATE", "PRODUCT", "CLIENT", "PRICE", "PROFIT"],
          [0.0, 0.16, 0.46, 0.68, 0.84], 18, 128)
    y = TOP + 78 + 18 * 128 + 30
    field(d, M, y, CW / 3 - 20, "Retail revenue"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Service revenue")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Attach rate %")
    i.append(img)


def p10(i):
    img, d = page("Product Inventory & Reorder", "Never say “I'm out of that”")
    table(d, M, TOP, CW, ["ITEM", "ON HAND", "REORDER AT", "UNIT COST", "REORDER?"],
          [0.0, 0.40, 0.56, 0.72, 0.88], 18, 128)
    i.append(img)


def p11(i):
    img, d = page("Income, Tips & Tax Set-Aside", "Tips are income — set some aside")
    table(d, M, TOP, CW, ["WEEK", "SERVICES", "RETAIL", "TIPS", "TOTAL"],
          [0.0, 0.24, 0.44, 0.62, 0.82], 6, 150)
    y = TOP + 78 + 6 * 150 + 40
    section(d, M, y, CW, "Set aside for tax")
    y2 = y + 110
    field(d, M, y2, CW / 3 - 20, "Month total"); field(d, M + CW / 3 + 10, y2, CW / 3 - 20, "Set-aside %")
    field(d, M + 2 * CW / 3 + 20, y2, CW / 3 - 20, "= Move to savings")
    y3 = y2 + 180
    section(d, M, y3, CW, "Monthly expenses — what the chair cost")
    table(d, M, y3 + 96, CW, ["EXPENSE", "CATEGORY", "AMOUNT"], [0.0, 0.46, 0.76], 7, 132)
    i.append(img)


def p12(i):
    img, d = page("Monthly Summary & Chair Review", "Is the chair actually working?")
    table(d, M, TOP, CW, ["MONTH", "REVENUE", "COSTS", "PROFIT", "TIPS"],
          [0.0, 0.26, 0.44, 0.62, 0.82], 8, 168)
    y = TOP + 78 + 8 * 168 + 40
    section(d, M, y, CW, "Six checks before next month")
    checks = ["True margin at or above goal", "Most clients rebooked before leaving",
              "Retail attaching to services", "Chair covered several times over",
              "No-shows under control", "New clients still coming in"]
    for j, c in enumerate(checks):
        cy = y + 110 + j * 100
        checkbox(d, M + 20, cy)
        d.text((M + 110, cy + 24), c, font=fs(30, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def main():
    out_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print_dir = os.path.join(out_dir, "marketing", "print"); os.makedirs(print_dir, exist_ok=True)
    imgs = []
    for fn in [p01, p02, p03, p04, p05, p06, p07, p08, p09, p10, p11, p12]:
        fn(imgs)
    for i, im in enumerate(imgs, 1):
        im.save(os.path.join(print_dir, f"page_{i:02d}.png"), "PNG", optimize=True)
    pdf_path = os.path.join(out_dir, "Salon_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
