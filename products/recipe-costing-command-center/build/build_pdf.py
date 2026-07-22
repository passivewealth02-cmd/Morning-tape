"""Build the printable PDF pack for Recipe Costing & Menu Engineering Command Center™.

  1  Recipe Cost Card             7  Menu Item P&L
  2  Prep List                    8  Specials & LTO Planner
  3  Menu Engineering Worksheet   9  Batch Recipe Scaler
  4  Food-Cost Pricing Guide     10  Vendor Price Tracker
  5  Ingredient Price Log        11  Waste Log
  6  Portion & Yield Worksheet   12  Weekly Food-Cost Tracker

Outputs ../Recipe_Costing_Printables.pdf and page PNGs in ../marketing/print/.
Run: python3 build_pdf.py
"""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFont

PRIMARY = (27, 79, 72); PRIMARY_DK = (18, 56, 51); ACCENT = (147, 115, 86)
GOLD = (180, 145, 90); GOLD_LT = (201, 168, 106); GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186); MINT = (227, 248, 239); WARN = (251, 240, 226)
WHITE = (255, 255, 255); TEXT = (51, 51, 51); MUTED = (120, 114, 104)
LINE = (200, 194, 182); ROW_ALT = (247, 243, 236); RED_BG = (251, 230, 230)

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
    d.text((M, 90), "RECIPE COSTING & MENU ENGINEERING COMMAND CENTER™", font=fs(28), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Restaurant: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Recipe Costing Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
    return img, d


def field(d, x, y, w, label, lab_font=None, line=True):
    lab_font = lab_font or fs(24)
    d.text((x, y), label.upper(), font=lab_font, fill=ACCENT, anchor="lt")
    if line:
        d.line((x, y + 62, x + w, y + 62), fill=LINE, width=2)


def section(d, x, y, w, text):
    d.rounded_rectangle((x, y, x + w, y + 56), radius=8, fill=SURFACE)
    d.text((x + 20, y + 28), text.upper(), font=fs(26), fill=PRIMARY, anchor="lm")


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


# ---------------------------------------------------------------- pages
def p01(imgs):
    img, d = page("Recipe Cost Card", "Cost a plate line by line — the foundation of a profitable menu")
    field(d, M, TOP, CW / 2 - 30, "Recipe name")
    field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Yield (servings)")
    table(d, M, TOP + 130, CW, ["INGREDIENT", "QTY", "UNIT", "COST/UNIT", "EXT. COST"],
          [0.0, 0.44, 0.58, 0.72, 0.86], 13, rowh=130)
    y = TOP + 130 + 78 + 13 * 130 + 30
    field(d, M, y, CW / 2 - 30, "Total recipe cost")
    field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Cost per serving")
    imgs.append(img)


def p02(imgs):
    img, d = page("Prep List", "The day's mise en place — by station")
    half = CW / 2 - 30
    for i, st in enumerate(["Grill", "Sauté", "Fry", "Cold / Pantry"]):
        x = M + (i % 2) * (half + 60)
        y = TOP + (i // 2) * 1180
        section(d, x, y, half, st)
        for k in range(9):
            d.rounded_rectangle((x, y + 90 + k * 110, x + 44, y + 90 + k * 110 + 44), radius=8, outline=PRIMARY, width=3)
            d.line((x + 68, y + 90 + k * 110 + 44, x + half, y + 90 + k * 110 + 44), fill=LINE, width=2)
    imgs.append(img)


def p03(imgs):
    img, d = page("Menu Engineering Worksheet", "Place each item in its quadrant — then act")
    gx0, gy0 = M + 120, TOP + 60
    gx1, gy1 = W - M, H - 460
    midx = (gx0 + gx1) / 2; midy = (gy0 + gy1) / 2
    quads = [("STARS", "feature & protect", MINT, gx0, gy0),
             ("PUZZLES", "reposition & upsell", SURFACE, midx, gy0),
             ("PLOWHORSES", "cut cost / raise price", WARN, gx0, midy),
             ("DOGS", "rework or remove", RED_BG, midx, midy)]
    for name, act, col, qx, qy in quads:
        d.rectangle((qx, qy, qx + (gx1 - gx0) / 2, qy + (gy1 - gy0) / 2), fill=col, outline=PRIMARY, width=2)
        d.text((qx + 30, qy + 30), name, font=fs(34), fill=PRIMARY, anchor="lt")
        d.text((qx + 30, qy + 84), act, font=fs(24, bold=False), fill=TEXT, anchor="lt")
    d.rectangle((gx0, gy0, gx1, gy1), outline=PRIMARY, width=4)
    d.line((midx, gy0, midx, gy1), fill=PRIMARY, width=4)
    d.line((gx0, midy, gx1, midy), fill=PRIMARY, width=4)
    d.text((gx0 - 70, midy), "PROFIT", font=fs(26), fill=ACCENT, anchor="mm")
    d.text((midx, gy1 + 40), "POPULARITY  →", font=fs(26), fill=ACCENT, anchor="mm")
    imgs.append(img)


def p04(imgs):
    img, d = page("Food-Cost Pricing Guide", "Turn plate cost into a target-margin price")
    d.text((M, TOP + 10), "SUGGESTED PRICE  =  PLATE COST  ÷  TARGET FOOD-COST %", font=fs(34), fill=PRIMARY, anchor="lt")
    table(d, M, TOP + 110, CW, ["PLATE COST", "@ 25%", "@ 28%", "@ 30%", "@ 33%", "@ 35%"],
          [0.0, 0.24, 0.40, 0.56, 0.72, 0.86], 12, rowh=150,
          filled_rows=[[f"${c:.2f}"] for c in [2.00, 2.50, 3.00, 3.50, 4.00, 4.50, 5.00, 6.00, 7.00, 8.00, 10.00, 12.00]])
    imgs.append(img)


def p05(imgs):
    img, d = page("Ingredient Price Log", "Pack size & price → cost per unit")
    table(d, M, TOP, CW, ["INGREDIENT", "PACK SIZE", "UNIT", "PACK PRICE", "COST/UNIT"],
          [0.0, 0.42, 0.58, 0.72, 0.86], 22, rowh=112)
    imgs.append(img)


def p06(imgs):
    img, d = page("Portion & Yield Worksheet", "Edible-portion cost is your true cost")
    d.text((M, TOP + 10), "TRUE (EP) COST  =  AS-PURCHASED COST  ÷  YIELD %", font=fs(32), fill=PRIMARY, anchor="lt")
    table(d, M, TOP + 100, CW, ["ITEM", "AP COST", "YIELD %", "EP COST"],
          [0.0, 0.40, 0.60, 0.80], 16, rowh=150)
    imgs.append(img)


def p07(imgs):
    img, d = page("Menu Item P&L", "One item, all the numbers that matter")
    y = TOP
    for lab in ["Item name", "Plate cost", "Menu price", "Food cost % (cost ÷ price)",
                "Gross margin $ (price − cost)", "Units sold / month", "Monthly profit (margin × units)",
                "Menu-engineering class", "Action to take"]:
        field(d, M, y, CW, lab); y += 170
    imgs.append(img)


def p08(imgs):
    img, d = page("Specials & LTO Planner", "Cost the special before you run it")
    table(d, M, TOP, CW, ["SPECIAL", "PLATE COST", "PRICE", "PROJ. UNITS", "PROJ. PROFIT"],
          [0.0, 0.36, 0.54, 0.70, 0.86], 14, rowh=150)
    imgs.append(img)


def p09(imgs):
    img, d = page("Batch Recipe Scaler", "Batch cost ÷ yield = cost per portion")
    table(d, M, TOP, CW, ["COMPONENT", "BATCH YIELD", "BATCH COST", "COST/SERVING"],
          [0.0, 0.40, 0.60, 0.82], 16, rowh=150)
    imgs.append(img)


def p10(imgs):
    img, d = page("Vendor Price Tracker", "Spot price creep before it eats your margin")
    table(d, M, TOP, CW, ["INGREDIENT (PACK)", "LAST PRICE", "THIS PRICE", "CHANGE %"],
          [0.0, 0.42, 0.60, 0.82], 18, rowh=128)
    imgs.append(img)


def p11(imgs):
    img, d = page("Waste Log", "Waste is pure profit leaving the building")
    table(d, M, TOP, CW, ["DATE", "ITEM", "REASON", "COST"],
          [0.0, 0.16, 0.44, 0.82], 20, rowh=118)
    imgs.append(img)


def p12(imgs):
    img, d = page("Weekly Food-Cost Tracker", "Purchases ÷ sales — keep the number honest")
    table(d, M, TOP, CW, ["WEEK", "FOOD PURCHASES", "FOOD SALES", "FOOD COST %"],
          [0.0, 0.34, 0.58, 0.82], 8, rowh=170,
          filled_rows=[[f"Week {i+1}"] for i in range(6)])
    y = TOP + 78 + 8 * 170 + 30
    field(d, M, y, CW, "Target food-cost % for the period")
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
    pdf_path = os.path.join(out_dir, "Recipe_Costing_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
