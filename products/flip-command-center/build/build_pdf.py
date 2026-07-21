"""Build the printable PDF pack for Flip Command Center™ (12 pages, US Letter).

Ink-light, print-ready job-site forms on white with a forest-green header band:

  1  Deal Analyzer Worksheet      7  Project Timeline
  2  Rehab Budget                 8  Holding-Cost Worksheet
  3  Scope of Work                9  Comps & ARV Worksheet
  4  Contractor & Bid Sheet      10  Selling & Exit / Net Sheet
  5  Draw Schedule               11  Punch List
  6  Materials Shopping List     12  Before & After Photo Log

Outputs ../Flip_Printables.pdf and page PNGs in ../marketing/print/.
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
    d.text((M, 90), "FLIP COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(78), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "214 Maple St  ·  Fix & Flip", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Flip Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
    return img, d


def checkbox(d, x, y, label, size=44, font=None, fill=TEXT):
    font = font or fs(30, bold=False)
    d.rounded_rectangle((x, y, x + size, y + size), radius=8, outline=PRIMARY, width=3)
    d.text((x + size + 24, y + size / 2), label, font=font, fill=fill, anchor="lm")


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
    img, d = page("Deal Analyzer Worksheet", "Run the numbers before you make the offer — know your max")
    y = TOP
    half = CW / 2 - 30
    section(d, M, y, half, "The numbers (enter)")
    section(d, M + CW / 2 + 30, y, half, "The deal (calculate)")
    y += 96
    left = ["After-Repair Value (ARV)", "Purchase Price", "Rehab Budget", "Buy-Side Closing",
            "Holding Months", "Selling Cost %", "Loan-to-Value", "Loan Rate"]
    right = ["Loan Amount", "Down Payment", "Holding Costs (total)", "Selling Costs",
             "All-In Cost", "Cash Invested", "Projected Profit", "Cash-on-Cash ROI"]
    for i in range(8):
        field(d, M, y + i * 118, half, left[i])
        field(d, M + CW / 2 + 30, y + i * 118, half, right[i])
    y2 = y + 8 * 118 + 20
    section(d, M, y2, CW, "The 70% rule"); y2 += 96
    d.text((M, y2), "Max Allowable Offer  =  (0.70 × ARV)  −  Rehab Budget", font=fs(34), fill=PRIMARY, anchor="lt")
    y2 += 80
    field(d, M, y2, half, "Max Allowable Offer (MAO)")
    checkbox(d, M + CW / 2 + 30, y2 - 6, "BUY — at or under MAO", size=44, font=fs(28, bold=False))
    checkbox(d, M + CW / 2 + 30, y2 + 66, "PASS — over MAO, renegotiate", size=44, font=fs(28, bold=False))
    imgs.append(img)


def p02(imgs):
    img, d = page("Rehab Budget", "Every category — planned vs actual, so overruns show up early")
    table(d, M, TOP, CW, ["CATEGORY", "PLANNED", "ACTUAL", "REMAINING", "% USED"],
          [0.0, 0.42, 0.58, 0.74, 0.90], 16, rowh=130,
          filled_rows=[["Demo & dumpster"], ["Roof"], ["Kitchen"], ["Bathrooms"], ["Flooring"],
                       ["Paint"], ["HVAC"], ["Electrical"], ["Plumbing"], ["Landscaping"], ["Permits"]])
    imgs.append(img)


def p03(imgs):
    img, d = page("Scope of Work", "Room by room, task by task — the full punch of the rehab")
    table(d, M, TOP, CW, ["ROOM / AREA", "TASK", "TRADE", "DONE"],
          [0.0, 0.24, 0.72, 0.90], 20, rowh=130)
    imgs.append(img)


def p04(imgs):
    img, d = page("Contractor & Bid Sheet", "Your crew & bids — trade, contact, bid and who's hired")
    table(d, M, TOP, CW, ["TRADE", "COMPANY", "CONTACT / PHONE", "BID", "HIRED?"],
          [0.0, 0.20, 0.46, 0.74, 0.90], 18, rowh=120)
    imgs.append(img)


def p05(imgs):
    img, d = page("Draw Schedule & Payment Log", "Every dollar out the door — draws, deposits & finals")
    table(d, M, TOP, CW, ["DATE", "PAYEE", "CATEGORY", "AMOUNT", "PAID?"],
          [0.0, 0.16, 0.50, 0.74, 0.90], 18, rowh=120)
    imgs.append(img)


def p06(imgs):
    img, d = page("Materials Shopping List", "Every material, where it's from & what's still needed")
    table(d, M, TOP, CW, ["ITEM", "ROOM", "STORE", "COST", "GOT IT?"],
          [0.0, 0.36, 0.54, 0.74, 0.90], 20, rowh=130)
    imgs.append(img)


def p07(imgs):
    img, d = page("Project Timeline", "The whole job on a calendar — each phase & its dates")
    table(d, M, TOP, CW, ["PHASE", "START", "TARGET END", "DAYS", "STATUS"],
          [0.0, 0.40, 0.56, 0.72, 0.86], 12, rowh=170,
          filled_rows=[["Acquisition & permits"], ["Demo"], ["Rough-in (MEP)"], ["Kitchen & baths"],
                       ["Finishes"], ["Punch list & staging"], ["List & sell"]])
    imgs.append(img)


def p08(imgs):
    img, d = page("Holding-Cost Worksheet", "Every month you hold, itemized — speed protects profit")
    table(d, M, TOP, CW, ["LINE ITEM", "PER MONTH", "× MONTHS", "TOTAL"],
          [0.0, 0.44, 0.64, 0.84], 8, rowh=150,
          filled_rows=[["Loan interest"], ["Property taxes"], ["Insurance"], ["Utilities"],
                       ["Lawn / security"], ["Misc"]])
    y = TOP + 78 + 8 * 150 + 60
    section(d, M, y, CW, "The rule of thumb"); y += 96
    d.text((M, y), "Every extra month of holding comes straight out of your profit. Sell fast.",
           font=fs(30, bold=False), fill=TEXT, anchor="lt")
    imgs.append(img)


def p09(imgs):
    img, d = page("Comps & ARV Worksheet", "The sold comps behind your ARV — the most important number")
    table(d, M, TOP, CW, ["COMP ADDRESS", "SQ FT", "BEDS/BATHS", "SOLD PRICE", "$/SF"],
          [0.0, 0.36, 0.52, 0.70, 0.88], 10, rowh=150)
    y = TOP + 78 + 10 * 150 + 50
    field(d, M, y, CW / 2 - 30, "Average comp price")
    field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Your ARV estimate")
    imgs.append(img)


def p10(imgs):
    img, d = page("Selling & Exit / Net Sheet", "The closing table — the profit you actually pocket")
    y = TOP
    labels = ["Sale Price (= ARV)", "Agent Commission + Closing", "Net Sale Proceeds",
              "Less: Loan Payoff", "Less: Cash Invested", "Projected Profit",
              "Cash-on-Cash ROI", "Profit Margin (of ARV)"]
    for i, lab in enumerate(labels):
        field(d, M, y + i * 150, CW, lab)
    imgs.append(img)


def p11(imgs):
    img, d = page("Punch List & Final Walkthrough", "The last 2% that sells the house — before photos")
    cols = [
        ["Touch-up paint on cabinets", "Caulk tub & re-seat toilet", "Replace outlet covers",
         "Fix sticking front door", "Adjust closet doors", "Final deep clean", "Blow off driveway"],
        ["Test all appliances", "Check every light & switch", "Confirm smoke/CO detectors",
         "Clean windows inside & out", "Stage key rooms", "Yard sign & lockbox", "Pro listing photos"],
    ]
    for c, items in enumerate(cols):
        x = M + c * (CW / 2 + 20)
        for i, it in enumerate(items):
            checkbox(d, x, TOP + i * 170, it, size=54, font=fs(28, bold=False))
    y = TOP + 7 * 170 + 30
    section(d, M, y, CW, "Final notes before listing"); y += 100
    for i in range(3):
        d.line((M, y + i * 84 + 30, M + CW, y + i * 84 + 30), fill=LINE, width=2)
    imgs.append(img)


def p12(imgs):
    img, d = page("Before & After Photo Log", "Document the transformation — for the listing & your brand")
    y = TOP
    field(d, M, y, CW / 2 - 30, "Room / area"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Date")
    y += 130
    caps = ["Before", "After", "Before", "After"]
    bw = CW / 2 - 30
    for i, cap in enumerate(caps):
        col = i % 2
        x = M + col * (CW / 2 + 30)
        yy = y + (i // 2) * 780
        d.rounded_rectangle((x, yy, x + bw, yy + 620), radius=18, outline=GOLD_LT, width=4)
        d.text((x + bw / 2, yy + 300), "tape / photo here", font=fs(28, bold=False), fill=MUTED, anchor="mm")
        d.text((x + 10, yy + 650), cap.upper(), font=fs(24), fill=ACCENT, anchor="lt")
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
    pdf_path = os.path.join(out_dir, "Flip_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
