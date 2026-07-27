"""Printable PDF pack for Notary & Loan Signing Agent Command Center™ (12 pages, US Letter)."""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFont

PRIMARY = (27, 79, 72); ACCENT = (147, 115, 86); GOLD_LT = (201, 168, 106); GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186); WHITE = (255, 255, 255); TEXT = (51, 51, 51); MUTED = (120, 114, 104)
LINE = (200, 194, 182); ROW_ALT = (247, 243, 236); DANGER = (201, 76, 76)
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
    d.text((M, 90), "NOTARY & LOAN SIGNING AGENT COMMAND CENTER™", font=fs(26), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Agent / Date: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Notary Command Center™", font=fs(22, bold=False), fill=ACCENT, anchor="rt")
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
    edges = colx + [x + w]
    d.rounded_rectangle((x, y, x + w, y + hh), radius=8, fill=PRIMARY)
    for i, h in enumerate(headers):
        if i == 0:
            d.text((colx[i] + 18, y + hh / 2), h, font=fs(26), fill=WHITE, anchor="lm")
        else:
            d.text(((edges[i] + edges[i + 1]) / 2, y + hh / 2), h, font=fs(26), fill=WHITE, anchor="mm")
    for r in range(nrows):
        ry = y + hh + r * rowh
        if r % 2:
            d.rectangle((x, ry, x + w, ry + rowh), fill=ROW_ALT)
        d.line((x, ry + rowh, x + w, ry + rowh), fill=LINE, width=2)
        if r < len(filled_rows):
            for ci, val in enumerate(filled_rows[r]):
                if ci == 0:
                    d.text((colx[ci] + 18, ry + rowh / 2), str(val), font=fs(26, bold=True), fill=PRIMARY, anchor="lm")
                else:
                    d.text(((edges[ci] + edges[ci + 1]) / 2, ry + rowh / 2), str(val), font=fs(26), fill=TEXT, anchor="mm")
    d.rectangle((x, y, x + w, y + hh + nrows * rowh), outline=LINE, width=2)
    for cx in colx[1:]:
        d.line((cx, y, cx, y + hh + nrows * rowh), fill=LINE, width=1)


CW = W - 2 * M; TOP = 430


def p01(i):
    img, d = page("What This Signing Really Pays", "Count the drive and the printer")
    field(d, M, TOP, CW / 2 - 30, "Company / order #"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Date")
    table(d, M, TOP + 150, CW, ["THE MATH", "AMOUNT"], [0.0, 0.78], 8,
          filled_rows=[["The fee", ""], ["÷ appointment length = what it FEELS like", ""],
                       ["− Printing (pages × cost per page)", ""],
                       ["− Driving (miles × vehicle cost)", ""],
                       ["= NET PER SIGNING", ""], ["Hours door to door (prep + drive + appt)", ""],
                       ["= WHAT YOU ACTUALLY EARN / HOUR", ""],
                       ["Mileage deduction (miles × IRS rate)", ""]], rowh=190)
    y = TOP + 150 + 78 + 8 * 190 + 36
    field(d, M, y, CW / 3 - 20, "Pages printed"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Round-trip miles")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Total hours")
    i.append(img)


def p02(i):
    img, d = page("Fee Schedule", "Stop quoting from memory")
    table(d, M, TOP, CW, ["SERVICE", "FEE", "NOTES"], [0.0, 0.46, 0.62], 16, 148)
    y = TOP + 78 + 16 * 148 + 36
    section(d, M, y, CW, "Before you quote")
    notes = ["Check your state's MAXIMUM fee per notarial act. It is not optional.",
             "Travel, printing and scanbacks are usually billed separately \\u2014 confirm for your state.",
             "Bill the trip fee when a signing cancels at the door. Every time."]
    for j, n in enumerate(notes):
        d.text((M + 24, y + 110 + j * 84), "•  " + n, font=fs(28, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p03(i):
    img, d = page("Signings Log", "Every job, what it netted")
    table(d, M, TOP, CW, ["DATE", "COMPANY", "TYPE", "FEE", "PAGES", "MILES", "STATUS"],
          [0.0, 0.12, 0.34, 0.52, 0.62, 0.71, 0.82], 20, 118)
    y = TOP + 78 + 20 * 118 + 30
    field(d, M, y, CW / 3 - 20, "Signings this period"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Total fees")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Total miles")
    i.append(img)


def p04(i):
    img, d = page("Mileage Log", "The deduction is worth more than the gas")
    field(d, M, TOP, CW / 2 - 30, "Month"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "IRS rate this year")
    table(d, M, TOP + 150, CW, ["DATE", "TRIP (FROM → TO)", "MILES", "PURPOSE"],
          [0.0, 0.12, 0.52, 0.64], 18, 128)
    y = TOP + 150 + 78 + 18 * 128 + 30
    field(d, M, y, CW / 2 - 30, "Total miles"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "= Deduction")
    i.append(img)


def p05(i):
    img, d = page("Notarial Journal Sheet", "Fill it in at the table, not later")
    d.text((M, TOP - 40), "CHECK YOUR STATE. Many require a BOUND, SEQUENTIAL journal — this is a convenience record.",
           font=fs(25), fill=DANGER, anchor="lt")
    table(d, M, TOP + 40, CW, ["DATE", "DOCUMENT TYPE", "SIGNER", "ID PRESENTED", "FEE", "SIG"],
          [0.0, 0.12, 0.36, 0.56, 0.76, 0.85], 18, 128)
    y = TOP + 40 + 78 + 18 * 128 + 30
    field(d, M, y, CW / 2 - 30, "Acts recorded"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Fees collected")
    i.append(img)


def p06(i):
    img, d = page("Signing Day Checklist", "Before you leave the house")
    half = CW / 2 - 30
    groups = [("Before you print", ["Confirmed appointment time & address", "Confirmed borrower names & spelling",
                                    "Checked package page count", "Checked for a second package (purchase)",
                                    "Confirmed printing instructions", "Confirmed scanback requirement",
                                    "Confirmed fee IN WRITING", "Asked about stairs, pets, parking"]),
              ("In the bag", ["Both packages, correctly sorted", "Journal & seal", "Two blue and two black pens",
                              "Thumbprint pad", "Phone charger & hotspot", "Shipping label & envelope",
                              "Backup ID verification method", "Business cards"])]
    for k, (title, items) in enumerate(groups):
        x = M + k * (half + 60)
        section(d, x, TOP, half, title)
        for j, it in enumerate(items):
            cy = TOP + 96 + j * 200
            checkbox(d, x, cy)
            d.text((x + 72, cy + 24), it, font=fs(26, bold=False), fill=TEXT, anchor="lt")
            d.line((x + 72, cy + 110, x + half, cy + 110), fill=LINE, width=2)
    y = TOP + 96 + 8 * 200 + 40
    field(d, M, y, CW / 2 - 30, "Anything unusual about this one"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Fee agreed")
    i.append(img)


def p07(i):
    img, d = page("Invoice", "Send it the same day")
    field(d, M, TOP, CW / 2 - 30, "Bill to"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Invoice # / date")
    field(d, M, TOP + 150, CW / 2 - 30, "Your business & address"); field(d, M + CW / 2 + 30, TOP + 150, CW / 2 - 30, "Payment terms")
    table(d, M, TOP + 320, CW, ["DATE", "SIGNING / SERVICE", "BORROWER", "AMOUNT"],
          [0.0, 0.12, 0.46, 0.78], 10, 148)
    y = TOP + 320 + 78 + 10 * 148 + 40
    field(d, M, y, CW / 3 - 20, "Subtotal"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Surcharges")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "TOTAL DUE")
    y2 = y + 180
    field(d, M, y2, CW / 2 - 30, "Remit to"); field(d, M + CW / 2 + 30, y2, CW / 2 - 30, "Due date")
    i.append(img)


def p08(i):
    img, d = page("Who Owes You", "Net 45 is a loan you gave them")
    table(d, M, TOP, CW, ["COMPANY", "SIGNINGS", "AMOUNT", "INVOICED", "DAYS OUT", "STATUS"],
          [0.0, 0.32, 0.46, 0.60, 0.74, 0.86], 18, 128)
    y = TOP + 78 + 18 * 128 + 30
    field(d, M, y, CW / 3 - 20, "Total invoiced"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Total outstanding")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Oldest unpaid (days)")
    i.append(img)


def p09(i):
    img, d = page("Signing Companies", "Who is actually worth saying yes to")
    table(d, M, TOP, CW, ["COMPANY", "CONTACT", "PAY TERMS", "AVG FEE", "NOTES"],
          [0.0, 0.24, 0.48, 0.62, 0.74], 18, 128)
    y = TOP + 78 + 18 * 128 + 30
    field(d, M, y, CW / 2 - 30, "Best average fee"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Slowest payer")
    i.append(img)


def p10(i):
    img, d = page("Printing & Supplies", "What a page actually costs")
    table(d, M, TOP, CW, ["ITEM", "COST", "YIELDS", "UNIT", "COST PER UNIT"],
          [0.0, 0.38, 0.52, 0.66, 0.80], 14, 158)
    y = TOP + 78 + 14 * 158 + 36
    field(d, M, y, CW / 3 - 20, "Cost per page"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Pages per signing")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Printing per signing")
    i.append(img)


def p11(i):
    img, d = page("Tax Set-Aside", "Nobody withholds anything from a 1099")
    table(d, M, TOP, CW, ["QUARTER", "INCOME", "SET ASIDE", "ESTIMATED DUE", "SHORT BY"],
          [0.0, 0.26, 0.44, 0.62, 0.83], 5, 175)
    y = TOP + 78 + 5 * 175 + 40
    field(d, M, y, CW / 3 - 20, "Year income"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Year set aside")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Still to save")
    y2 = y + 200
    section(d, M, y2, CW, "Deductions signing agents forget")
    items = ["Every mile — to signings, the post office, the supply store",
             "Toner, paper, pens, seals and thumbprint pads",
             "Printer lease, repairs and the maintenance kit",
             "E&O insurance, bond, commission and background checks",
             "Phone and internet, at the business-use percentage",
             "NNA membership, training and continuing education"]
    for j, it in enumerate(items):
        cy = y2 + 130 + j * 116
        checkbox(d, M + 20, cy)
        d.text((M + 110, cy + 24), it, font=fs(28, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p12(i):
    img, d = page("Monthly Summary & Review", "Is this actually paying?")
    table(d, M, TOP, CW, ["MONTH", "SIGNINGS", "REVENUE", "COSTS", "PROFIT", "$ / HOUR"],
          [0.0, 0.24, 0.42, 0.58, 0.74, 0.89], 8, 168)
    y = TOP + 78 + 8 * 168 + 40
    section(d, M, y, CW, "Six checks before next month")
    checks = ["Real hourly rate at or above my goal", "Margin healthy after printing and driving",
              "Overhead more than covered", "Everyone who owes me has actually paid",
              "Enough signings booked for next month", "Tax set aside, not spent"]
    for j, c in enumerate(checks):
        cy = y + 130 + j * 116
        checkbox(d, M + 20, cy)
        d.text((M + 110, cy + 24), c, font=fs(29, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def main():
    out_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print_dir = os.path.join(out_dir, "marketing", "print"); os.makedirs(print_dir, exist_ok=True)
    imgs = []
    for fn in [p01, p02, p03, p04, p05, p06, p07, p08, p09, p10, p11, p12]:
        fn(imgs)
    for i, im in enumerate(imgs, 1):
        im.save(os.path.join(print_dir, f"page_{i:02d}.png"), "PNG", optimize=True)
    pdf_path = os.path.join(out_dir, "Notary_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
