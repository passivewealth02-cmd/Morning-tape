"""Printable PDF pack for Trucking Owner-Operator Command Center™ (12 pages, US Letter)."""
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
    d.text((M, 90), "TRUCKING OWNER-OPERATOR COMMAND CENTER™", font=fs(26), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Truck / Date: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Trucking Command Center™", font=fs(22, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Cost Per Mile Worksheet", "The one number that decides everything")
    field(d, M, TOP, CW / 2 - 30, "Month"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Truck / Unit #")
    table(d, M, TOP + 150, CW, ["THE MATH", "AMOUNT"], [0.0, 0.78], 8,
          filled_rows=[["Fixed costs this month", ""], ["÷ Total miles run", ""],
                       ["= Fixed cost per mile", ""], ["+ Variable cost per mile", ""],
                       ["= COST PER MILE RUN", ""], ["× Total miles = total cost", ""],
                       ["÷ LOADED miles", ""], ["= COST PER LOADED MILE", ""]], rowh=200)
    y = TOP + 150 + 78 + 8 * 200 + 36
    field(d, M, y, CW / 3 - 20, "Your rate / loaded mile"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "= Profit / mile")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Rate floor (never below)")
    i.append(img)


def p02(i):
    img, d = page("Fixed Costs Worksheet", "What the truck costs standing still")
    table(d, M, TOP, CW, ["FIXED COST LINE", "MONTHLY", "YEARLY"], [0.0, 0.58, 0.80], 14, 165)
    y = TOP + 78 + 14 * 165 + 36
    field(d, M, y, CW / 3 - 20, "Total fixed / month"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Total miles run")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Fixed cost per mile")
    i.append(img)


def p03(i):
    img, d = page("Should I Take This Load?", "Run the numbers before you call the broker")
    field(d, M, TOP, CW / 2 - 30, "Broker / Load #"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Lane")
    table(d, M, TOP + 150, CW, ["THE LOAD", "AMOUNT"], [0.0, 0.78], 8,
          filled_rows=[["Loaded miles", ""], ["+ Deadhead to pickup", ""], ["= Total miles you'll run", ""],
                       ["Rate offered (total $)", ""], ["÷ Loaded miles = rate / loaded mile", ""],
                       ["÷ TOTAL miles = all-in rate / mile", ""], ["Your cost per loaded mile", ""],
                       ["= PROFIT ON THIS LOAD", ""]], rowh=178)
    y = TOP + 150 + 78 + 8 * 178 + 36
    field(d, M, y, CW / 2 - 30, "Detention / layover terms"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Take it?  Y / N")
    i.append(img)


def p04(i):
    img, d = page("Trip Sheet", "Every leg, logged")
    field(d, M, TOP, CW / 3 - 20, "Date"); field(d, M + CW / 3 + 10, TOP, CW / 3 - 20, "Driver")
    field(d, M + 2 * CW / 3 + 20, TOP, CW / 3 - 20, "Unit #")
    table(d, M, TOP + 150, CW, ["FROM → TO", "STATE", "START ODO", "END ODO", "MILES"],
          [0.0, 0.42, 0.56, 0.72, 0.88], 15, 132)
    y = TOP + 150 + 78 + 15 * 132 + 30
    field(d, M, y, CW / 3 - 20, "Loaded miles"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Deadhead miles")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Total miles")
    i.append(img)


def p05(i):
    img, d = page("Fuel Log", "Half an MPG is worth thousands")
    table(d, M, TOP, CW, ["DATE", "STATE", "GALLONS", "$/GAL", "COST", "ODOMETER"],
          [0.0, 0.16, 0.32, 0.48, 0.62, 0.80], 18, 128)
    y = TOP + 78 + 18 * 128 + 30
    field(d, M, y, CW / 3 - 20, "Total gallons"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Total miles")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Actual MPG")
    i.append(img)


def p06(i):
    img, d = page("Maintenance & PM Schedule", "By the odometer, not by hope")
    table(d, M, TOP, CW, ["DATE", "TYPE", "WORK DONE", "COST", "ODOMETER", "NEXT DUE"],
          [0.0, 0.14, 0.30, 0.58, 0.72, 0.88], 17, 132)
    y = TOP + 78 + 17 * 132 + 30
    field(d, M, y, CW / 2 - 30, "Spent this month"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "= Per mile run")
    i.append(img)


def p07(i):
    img, d = page("Settlement Reconciliation", "Check every deduction")
    table(d, M, TOP, CW, ["WEEK", "GROSS", "DEDUCTIONS", "NET", "WHAT FOR"],
          [0.0, 0.20, 0.38, 0.56, 0.72], 14, 152)
    y = TOP + 78 + 14 * 152 + 36
    field(d, M, y, CW / 3 - 20, "Total gross"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Total deductions")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Total net")
    i.append(img)


def p08(i):
    img, d = page("IFTA Miles by State", "Quarterly filing, made painless")
    half = CW / 2 - 30
    for k, st in enumerate(["Quarter miles — part 1", "Quarter miles — part 2"]):
        x = M + k * (half + 60)
        table(d, x, TOP, half, ["STATE", "MILES", "GALLONS"], [0.0, 0.40, 0.72], 16, 138)
    y = TOP + 78 + 16 * 138 + 36
    field(d, M, y, CW / 3 - 20, "Total miles"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Total gallons")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Fleet MPG")
    i.append(img)


def p09(i):
    img, d = page("Pre-Trip Inspection", "Sign it before you roll")
    field(d, M, TOP, CW / 3 - 20, "Date"); field(d, M + CW / 3 + 10, TOP, CW / 3 - 20, "Unit #")
    field(d, M + 2 * CW / 3 + 20, TOP, CW / 3 - 20, "Odometer")
    half = CW / 2 - 30
    groups = [("Tractor", ["Tires & wheels", "Brakes & air lines", "Lights & reflectors", "Fluids & leaks",
                           "Steering & suspension", "Mirrors & glass", "Horn & wipers", "Coupling & fifth wheel",
                           "Emergency equipment", "Seat belt & cab"]),
              ("Trailer & paperwork", ["Tires & wheels", "Brakes & air lines", "Lights & reflectors",
                                       "Landing gear", "Doors & seals", "Load secured", "Cargo temp / seal #",
                                       "Bills of lading", "Permits & registration", "ELD logged in"])]
    for k, (title, items) in enumerate(groups):
        x = M + k * (half + 60)
        section(d, x, TOP + 180, half, title)
        for j, it in enumerate(items):
            cy = TOP + 276 + j * 160
            checkbox(d, x, cy)
            d.text((x + 72, cy + 24), it, font=fs(28, bold=False), fill=TEXT, anchor="lt")
            d.line((x + 72, cy + 96, x + half, cy + 96), fill=LINE, width=2)
    y = TOP + 276 + 10 * 160 + 40
    field(d, M, y, CW / 2 - 30, "Defects found"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Driver signature")
    i.append(img)


def p10(i):
    img, d = page("Maintenance Reserve Fund", "Because engines don't ask first")
    table(d, M, TOP, CW, ["FUND", "TARGET", "SAVED", "FUNDED %"], [0.0, 0.44, 0.62, 0.82], 10, 165)
    y = TOP + 78 + 10 * 165 + 36
    field(d, M, y, CW / 3 - 20, "Total target"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Total saved")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Set aside per mile")
    y2 = y + 200
    section(d, M, y2, CW, "What a blown engine costs versus what you've saved")
    for j in range(4):
        d.line((M + 20, y2 + 150 + j * 100, W - M - 20, y2 + 150 + j * 100), fill=LINE, width=2)
    i.append(img)


def p11(i):
    img, d = page("Monthly Summary", "Is the year working?")
    table(d, M, TOP, CW, ["MONTH", "LOADED MI", "REVENUE", "COST", "PROFIT", "$ / MILE"],
          [0.0, 0.22, 0.40, 0.58, 0.74, 0.89], 14, 165)
    y = TOP + 78 + 14 * 165 + 36
    field(d, M, y, CW / 3 - 20, "Year revenue"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Year profit")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Net margin %")
    i.append(img)


def p12(i):
    img, d = page("Rate Floor Card", "Tape this to the dash")
    section(d, M, TOP, CW, "Never book a load below this")
    y = TOP + 130
    for lab in ["My cost per loaded mile", "My rate floor (cost + minimum profit)", "My target rate per loaded mile"]:
        d.rounded_rectangle((M, y, W - M, y + 240), radius=12, outline=LINE, width=3)
        d.text((M + 40, y + 60), lab.upper(), font=fs(30), fill=ACCENT, anchor="lt")
        d.line((M + 40, y + 176, W - M - 40, y + 176), fill=LINE, width=3)
        y += 280
    y += 40
    section(d, M, y, CW, "Before you say yes")
    checks = ["Did I add the deadhead miles to get there?",
              "Is the all-in rate per TOTAL mile still above my cost?",
              "Detention and layover terms in writing?",
              "Is this broker's payment history acceptable?",
              "Does this load put me somewhere with freight out?"]
    for j, c in enumerate(checks):
        cy = y + 140 + j * 130
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
    pdf_path = os.path.join(out_dir, "Trucking_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
