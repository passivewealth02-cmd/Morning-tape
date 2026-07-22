"""Build the printable PDF pack for Cafe & Coffee Shop Command Center™ (12 pages).

  1  Cup Cost Card            7  Inventory & Par
  2  Menu Board              8  Waste Log
  3  Daypart Sales Log       9  Ordering Sheet
  4  Weekly Sales Log       10  Cash & Tips
  5  Labor & Prime Cost     11  Open / Close Checklist
  6  Bean & Milk Usage      12  Regulars & Loyalty

Outputs ../Cafe_Printables.pdf and page PNGs in ../marketing/print/.
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
    d.text((M, 90), "CAFE & COFFEE SHOP COMMAND CENTER™", font=fs(30), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Cafe / Date: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Cafe Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def p01(imgs):
    img, d = page("Cup Cost Card", "Cost a drink to the cup — beans, milk, cup & lid")
    field(d, M, TOP, CW / 2 - 30, "Drink"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Size")
    table(d, M, TOP + 130, CW, ["COMPONENT", "QTY", "UNIT", "COST/UNIT", "EXT. COST"],
          [0.0, 0.44, 0.58, 0.72, 0.86], 11, rowh=140)
    y = TOP + 130 + 78 + 11 * 140 + 30
    field(d, M, y, CW / 2 - 30, "Cup cost"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Menu price")
    imgs.append(img)


def p02(imgs):
    img, d = page("Menu Board", "Cup cost, price, beverage-cost % & margin")
    table(d, M, TOP, CW, ["ITEM", "CUP COST", "PRICE", "BEV %", "MARGIN $"],
          [0.0, 0.40, 0.56, 0.72, 0.86], 18, rowh=120)
    imgs.append(img)


def p03(imgs):
    img, d = page("Daypart Sales Log", "Where the money is by time of day")
    table(d, M, TOP, CW, ["DAYPART", "TRANSACTIONS", "SALES", "AVG TICKET"],
          [0.0, 0.38, 0.60, 0.82], 10, rowh=200,
          filled_rows=[["Morning (6-10)"], ["Midday (10-2)"], ["Afternoon (2-6)"], ["Evening (6-close)"]])
    imgs.append(img)


def p04(imgs):
    img, d = page("Weekly Sales Log", "The week at a glance")
    table(d, M, TOP, CW, ["DAY", "SALES", "TRANSACTIONS", "AVG TICKET"],
          [0.0, 0.30, 0.56, 0.80], 8, rowh=170, filled_rows=[[dn] for dn in DAYS])
    imgs.append(img)


def p05(imgs):
    img, d = page("Labor & Prime Cost", "Beverage % + labor % = prime cost")
    table(d, M, TOP, CW, ["DAY", "LABOR $", "SALES", "LABOR %"],
          [0.0, 0.30, 0.54, 0.80], 8, rowh=160, filled_rows=[[dn] for dn in DAYS])
    y = TOP + 78 + 8 * 160 + 30
    section(d, M, y, CW, "Prime cost  =  beverage %  +  labor %  =  __________ %  (keep under 60%)")
    imgs.append(img)


def p06(imgs):
    img, d = page("Bean & Milk Usage", "Where beverage cost really comes from")
    table(d, M, TOP, CW, ["ITEM", "QTY", "UNIT", "WEEKLY COST"],
          [0.0, 0.42, 0.58, 0.80], 14, rowh=155)
    imgs.append(img)


def p07(imgs):
    img, d = page("Inventory & Par", "Order before you 86 the oat milk")
    table(d, M, TOP, CW, ["ITEM", "PAR", "ON HAND", "UNIT", "TO ORDER"],
          [0.0, 0.42, 0.56, 0.72, 0.86], 18, rowh=120)
    imgs.append(img)


def p08(imgs):
    img, d = page("Waste Log", "Every dumped shot & spoiled gallon")
    table(d, M, TOP, CW, ["DATE", "ITEM", "REASON", "COST"],
          [0.0, 0.16, 0.44, 0.82], 20, rowh=118)
    imgs.append(img)


def p09(imgs):
    img, d = page("Ordering Sheet", "Your standing order, by supplier")
    table(d, M, TOP, CW, ["ITEM", "SUPPLIER", "PAR ORDER", "COST"],
          [0.0, 0.38, 0.62, 0.82], 18, rowh=120)
    imgs.append(img)


def p10(imgs):
    img, d = page("Cash & Tips", "Balance the drawer every day")
    table(d, M, TOP, CW, ["DAY", "CASH", "CARD", "TIPS", "TOTAL"],
          [0.0, 0.24, 0.42, 0.60, 0.80], 8, rowh=170, filled_rows=[[dn] for dn in DAYS])
    imgs.append(img)


def p11(imgs):
    img, d = page("Open / Close Checklist", "Same standard, every shift")
    half = CW / 2 - 30
    for i, st in enumerate(["Opening", "Closing"]):
        x = M + i * (half + 60)
        section(d, x, TOP, half, st)
        for k in range(12):
            checkbox(d, x, TOP + 96 + k * 110, "", size=48)
            d.line((x + 72, TOP + 96 + k * 110 + 48, x + half, TOP + 96 + k * 110 + 48), fill=LINE, width=2)
    imgs.append(img)


def p12(imgs):
    img, d = page("Regulars & Loyalty", "The backbone of a café")
    table(d, M, TOP, CW, ["REGULAR", "USUAL ORDER", "VISITS/WK", "SPEND/WK"],
          [0.0, 0.30, 0.58, 0.80], 16, rowh=130)
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
    pdf_path = os.path.join(out_dir, "Cafe_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
