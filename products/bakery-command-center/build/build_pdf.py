"""Build the printable PDF pack for Bakery Command Center™ (12 pages, US Letter).

  1  Recipe Cost Card         7  Waste & Day-Old Log
  2  Product Price List       8  Sales Log
  3  Pre-Order Form           9  Ordering Sheet
  4  Wholesale Order Sheet   10  Cash & Deposits
  5  Production Plan         11  Market Day Sheet
  6  Inventory & Par         12  Bake-Day Checklist

Outputs ../Bakery_Printables.pdf and page PNGs in ../marketing/print/.
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
    d.text((M, 90), "BAKERY COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Bakery / Date: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Bakery Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Recipe Cost Card", "Cost by the batch, divide by yield")
    field(d, M, TOP, CW / 2 - 30, "Recipe"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Batch yield")
    table(d, M, TOP + 130, CW, ["INGREDIENT", "QTY", "UNIT", "COST/UNIT", "EXT. COST"],
          [0.0, 0.44, 0.58, 0.72, 0.86], 12, rowh=140)
    y = TOP + 130 + 78 + 12 * 140 + 30
    field(d, M, y, CW / 2 - 30, "Batch cost"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Cost per unit")
    imgs.append(img)


def p02(imgs):
    img, d = page("Product Price List", "Unit cost, retail & wholesale, margin & food %")
    table(d, M, TOP, CW, ["PRODUCT", "UNIT COST", "RETAIL", "WHOLESALE", "MARGIN", "FOOD %"],
          [0.0, 0.34, 0.50, 0.64, 0.80, 0.90], 18, rowh=120)
    imgs.append(img)


def p03(imgs):
    img, d = page("Pre-Order Form", "Custom cakes & special orders")
    for i, lab in enumerate(["Customer name", "Phone / email", "Order details", "Flavor / notes",
                             "Pickup date & time", "Quantity", "Price quoted", "Deposit paid", "Balance due"]):
        field(d, M, TOP + i * 200, CW, lab)
    imgs.append(img)


def p04(imgs):
    img, d = page("Wholesale Order Sheet", "Standing accounts & weekly quantities")
    table(d, M, TOP, CW, ["ACCOUNT", "ITEM", "WEEKLY QTY", "UNIT PRICE", "WEEKLY REV"],
          [0.0, 0.32, 0.52, 0.70, 0.86], 16, rowh=130)
    imgs.append(img)


def p05(imgs):
    img, d = page("Production Plan", "The weekly bake schedule")
    table(d, M, TOP, CW, ["DAY", "MORNING BAKE", "AFTERNOON BAKE"],
          [0.0, 0.18, 0.60], 7, rowh=290, filled_rows=[[dn] for dn in DAYS])
    imgs.append(img)


def p06(imgs):
    img, d = page("Inventory & Par", "Order before the Saturday bake")
    table(d, M, TOP, CW, ["INGREDIENT", "PAR", "ON HAND", "UNIT", "TO ORDER"],
          [0.0, 0.42, 0.56, 0.72, 0.86], 18, rowh=120)
    imgs.append(img)


def p07(imgs):
    img, d = page("Waste & Day-Old Log", "Track it, donate it, shrink it")
    table(d, M, TOP, CW, ["DATE", "ITEM", "REASON", "COST"],
          [0.0, 0.16, 0.44, 0.82], 20, rowh=118)
    imgs.append(img)


def p08(imgs):
    img, d = page("Sales Log", "The pulse of the counter")
    table(d, M, TOP, CW, ["DAY", "RETAIL SALES", "UNITS", "AVG SALE"],
          [0.0, 0.30, 0.56, 0.80], 8, rowh=170, filled_rows=[[dn] for dn in DAYS])
    imgs.append(img)


def p09(imgs):
    img, d = page("Ordering Sheet", "Your standing supplier order")
    table(d, M, TOP, CW, ["ITEM", "SUPPLIER", "PAR ORDER", "COST"],
          [0.0, 0.38, 0.62, 0.82], 18, rowh=120)
    imgs.append(img)


def p10(imgs):
    img, d = page("Cash & Deposits", "Reconcile the till daily")
    table(d, M, TOP, CW, ["DAY", "CASH", "CARD", "DEPOSITS", "TOTAL"],
          [0.0, 0.24, 0.42, 0.60, 0.80], 8, rowh=170, filled_rows=[[dn] for dn in DAYS])
    imgs.append(img)


def p11(imgs):
    img, d = page("Market Day Sheet", "Farmers markets & fairs")
    for i, lab in enumerate(["Market & date", "Booth fee", "Items & quantities brought", "Sales total",
                             "Items sold / left", "Net (sales − fee)", "Notes for next time"]):
        field(d, M, TOP + i * 220, CW, lab)
    imgs.append(img)


def p12(imgs):
    img, d = page("Bake-Day Checklist", "Same standard, every morning")
    half = CW / 2 - 30
    for i, st in enumerate(["Opening / Mise", "Bake & Display", "Wholesale & Pre-Orders", "Closing"]):
        x = M + (i % 2) * (half + 60)
        y = TOP + (i // 2) * 1180
        section(d, x, y, half, st)
        for k in range(8):
            checkbox(d, x, y + 90 + k * 118, "", size=48)
            d.line((x + 72, y + 90 + k * 118 + 48, x + half, y + 90 + k * 118 + 48), fill=LINE, width=2)
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
    pdf_path = os.path.join(out_dir, "Bakery_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
