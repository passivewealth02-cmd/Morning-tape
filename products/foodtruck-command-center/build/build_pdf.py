"""Build the printable PDF pack for Food Truck Command Center™ (12 pages, US Letter).

  1  Event P&L Sheet          7  Fuel & Mileage Log
  2  Break-Even Worksheet     8  Permit Tracker
  3  Daily Sales Log          9  Supplies / Shopping List
  4  Menu & Cost Card        10  Bookings Calendar
  5  Inventory & Par         11  Cash & Tips Reconciliation
  6  Prep List               12  Monthly P&L Summary

Outputs ../Food_Truck_Printables.pdf and page PNGs in ../marketing/print/.
Run: python3 build_pdf.py
"""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFont

PRIMARY = (27, 79, 72); PRIMARY_DK = (18, 56, 51); ACCENT = (147, 115, 86)
GOLD = (180, 145, 90); GOLD_LT = (201, 168, 106); GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186); MINT = (227, 248, 239); WARN = (251, 240, 226)
WHITE = (255, 255, 255); TEXT = (51, 51, 51); MUTED = (120, 114, 104)
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
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)
    bh = 340
    d.rectangle((0, 0, W, bh), fill=PRIMARY)
    d.rectangle((0, bh, W, bh + 10), fill=GOLD_LT)
    d.rectangle((0, bh + 10, W, bh + 14), fill=GOLD_HI)
    d.text((M, 90), "FOOD TRUCK COMMAND CENTER™", font=fs(32), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Truck / Date: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Food Truck Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
    return img, d


def field(d, x, y, w, label, lab_font=None, line=True):
    lab_font = lab_font or fs(24)
    d.text((x, y), label.upper(), font=lab_font, fill=ACCENT, anchor="lt")
    if line:
        d.line((x, y + 62, x + w, y + 62), fill=LINE, width=2)


def section(d, x, y, w, text):
    d.rounded_rectangle((x, y, x + w, y + 56), radius=8, fill=SURFACE)
    d.text((x + 20, y + 28), text.upper(), font=fs(26), fill=PRIMARY, anchor="lm")


def checkbox(d, x, y, label, size=44, font=None, fill=TEXT):
    font = font or fs(30, bold=False)
    d.rounded_rectangle((x, y, x + size, y + size), radius=8, outline=PRIMARY, width=3)
    d.text((x + size + 24, y + size / 2), label, font=font, fill=fill, anchor="lm")


def table(d, x, y, w, headers, colf, nrows, rowh=92, filled_rows=None):
    filled_rows = filled_rows or []
    colx = [x + w * f for f in colf]
    hh = 78
    d.rounded_rectangle((x, y, x + w, y + hh), radius=8, fill=PRIMARY)
    for i, h in enumerate(headers):
        d.text((colx[i] + (18 if i == 0 else 0), y + hh / 2), h, font=fs(26),
               fill=WHITE, anchor="lm" if i == 0 else "mm")
    for r in range(nrows):
        ry = y + hh + r * rowh
        if r % 2:
            d.rectangle((x, ry, x + w, ry + rowh), fill=ROW_ALT)
        d.line((x, ry + rowh, x + w, ry + rowh), fill=LINE, width=2)
        if r < len(filled_rows):
            for ci, val in enumerate(filled_rows[r]):
                d.text((colx[ci] + (18 if ci == 0 else 0), ry + rowh / 2), str(val),
                       font=fs(26, bold=(ci == 0)), fill=PRIMARY if ci == 0 else TEXT,
                       anchor="lm" if ci == 0 else "mm")
    d.rectangle((x, y, x + w, y + hh + nrows * rowh), outline=LINE, width=2)
    for cx in colx[1:]:
        d.line((cx, y, cx, y + hh + nrows * rowh), fill=LINE, width=1)


CW = W - 2 * M
TOP = 430


def p01(imgs):
    img, d = page("Event P&L Sheet", "One gig, one page — did it make money?")
    field(d, M, TOP, CW / 2 - 30, "Event / location")
    field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Date")
    y = TOP + 150
    for lab in ["Sales", "Food cost", "Fuel", "Event fee", "Staff", "Other"]:
        field(d, M, y, CW / 2 - 30, lab); y += 150
    d.rounded_rectangle((M + CW / 2 + 30, TOP + 150, W - M, TOP + 150 + 380), radius=14, fill=MINT)
    d.text((M + CW / 2 + 70, TOP + 210), "NET PROFIT", font=fs(34), fill=PRIMARY, anchor="lt")
    d.text((M + CW / 2 + 70, TOP + 300), "$ ______________", font=fserif(54), fill=PRIMARY, anchor="lt")
    d.text((M + CW / 2 + 70, TOP + 420), "sales − all costs", font=fs(24, bold=False), fill=TEXT, anchor="lt")
    imgs.append(img)


def p02(imgs):
    img, d = page("Break-Even Worksheet", "The sales you need to cover fixed costs")
    y = TOP + 20
    for lab in ["Monthly fixed overhead  $ ____________", "Average net profit per event  $ __________",
                "Break-even events = overhead ÷ avg net  =  __________", "Average sales per event  $ __________",
                "Break-even sales / month  $ ____________", "Events booked this month  __________"]:
        d.text((M, y), lab, font=fs(34, bold=False), fill=TEXT, anchor="lt"); y += 170
    y += 10
    section(d, M, y, CW, "Every gig past break-even is profit in your pocket")
    imgs.append(img)


def p03(imgs):
    img, d = page("Daily Sales Log", "The pulse of the week")
    table(d, M, TOP, CW, ["DATE", "LOCATION", "SALES", "COGS", "FOOD %"],
          [0.0, 0.20, 0.50, 0.66, 0.84], 16, rowh=130)
    imgs.append(img)


def p04(imgs):
    img, d = page("Menu & Cost Card", "Cost & price every item on the truck")
    table(d, M, TOP, CW, ["ITEM", "PLATE COST", "PRICE", "FOOD %", "MARGIN $"],
          [0.0, 0.40, 0.56, 0.72, 0.86], 16, rowh=130)
    imgs.append(img)


def p05(imgs):
    img, d = page("Inventory & Par", "What to restock before the next service")
    table(d, M, TOP, CW, ["ITEM", "PAR", "ON HAND", "UNIT", "TO BUY"],
          [0.0, 0.42, 0.56, 0.72, 0.86], 18, rowh=120)
    imgs.append(img)


def p06(imgs):
    img, d = page("Prep List", "The day's mise en place")
    half = CW / 2 - 30
    for i, st in enumerate(["Smoker / Grill", "Fryer", "Cold / Prep", "Restock"]):
        x = M + (i % 2) * (half + 60)
        y = TOP + (i // 2) * 1180
        section(d, x, y, half, st)
        for k in range(9):
            checkbox(d, x, y + 90 + k * 110, "", size=44)
            d.line((x + 68, y + 90 + k * 110 + 44, x + half, y + 90 + k * 110 + 44), fill=LINE, width=2)
    imgs.append(img)


def p07(imgs):
    img, d = page("Fuel & Mileage Log", "Track every fill-up")
    table(d, M, TOP, CW, ["DATE", "ODOMETER", "MILES", "GALLONS", "COST"],
          [0.0, 0.22, 0.44, 0.62, 0.82], 20, rowh=118)
    imgs.append(img)


def p08(imgs):
    img, d = page("Permit Tracker", "Never get shut down for a lapsed license")
    table(d, M, TOP, CW, ["PERMIT / LICENSE", "NUMBER", "EXPIRES", "STATUS"],
          [0.0, 0.42, 0.60, 0.82], 14, rowh=150)
    imgs.append(img)


def p09(imgs):
    img, d = page("Supplies / Shopping List", "The next restock run")
    half = CW / 2 - 30
    for col in range(2):
        x = M + col * (half + 60)
        for k in range(16):
            checkbox(d, x, TOP + k * 130, "", size=48)
            d.line((x + 72, TOP + k * 130 + 48, x + half, TOP + k * 130 + 48), fill=LINE, width=2)
    imgs.append(img)


def p10(imgs):
    img, d = page("Bookings Calendar", "The right gigs on the schedule")
    table(d, M, TOP, CW, ["EVENT", "DATE", "LOCATION", "DEPOSIT", "STATUS"],
          [0.0, 0.32, 0.48, 0.70, 0.86], 16, rowh=130)
    imgs.append(img)


def p11(imgs):
    img, d = page("Cash & Tips Reconciliation", "Make the drawer balance every day")
    table(d, M, TOP, CW, ["DAY", "CASH", "CARD", "TIPS", "TOTAL"],
          [0.0, 0.24, 0.42, 0.60, 0.80], 16, rowh=130)
    imgs.append(img)


def p12(imgs):
    img, d = page("Monthly P&L Summary", "Add it all up — the month at a glance")
    y = TOP + 10
    for lab in ["Total event sales  $ ____________", "Total food cost  $ ____________",
                "Total fuel  $ ____________", "Total staff  $ ____________",
                "Fixed overhead  $ ____________", "NET PROFIT FOR THE MONTH  $ ____________",
                "Events run  __________     Break-even events  __________"]:
        big = lab.startswith("NET")
        d.text((M, y), lab, font=fs(38 if big else 34, bold=big), fill=PRIMARY if big else TEXT, anchor="lt")
        y += 180
    imgs.append(img)


def main():
    out_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print_dir = os.path.join(out_dir, "marketing", "print")
    os.makedirs(print_dir, exist_ok=True)
    imgs = []
    for fn in [p01, p02, p03, p04, p05, p06, p07, p08, p09, p10, p11, p12]:
        fn(imgs)
    for i, im in enumerate(imgs, 1):
        im.save(os.path.join(print_dir, f"page_{i:02d}.png"), "PNG", optimize=True)
    pdf_path = os.path.join(out_dir, "Food_Truck_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
