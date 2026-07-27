"""Printable PDF pack for Daycare & Childcare Provider Command Center™ (12 pages, US Letter)."""
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
    d.text((M, 90), "DAYCARE & CHILDCARE PROVIDER COMMAND CENTER™", font=fs(26), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Provider / Date: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Daycare Command Center™", font=fs(22, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Rate & Break-Even Worksheet", "What one child actually nets you")
    field(d, M, TOP, CW / 2 - 30, "Program"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Month")
    table(d, M, TOP + 150, CW, ["THE MATH", "AMOUNT"], [0.0, 0.78], 7,
          filled_rows=[["Full-time weekly rate", ""], ["× weeks per month", ""],
                       ["= Tuition per child / month", ""], ["− Food, supplies & activities", ""],
                       ["= NET PER CHILD", ""], ["Fixed costs ÷ net per child", ""],
                       ["= BREAK-EVEN CHILDREN", ""]], rowh=200)
    y = TOP + 150 + 78 + 7 * 200 + 36
    field(d, M, y, CW / 3 - 20, "Licensed capacity"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Enrolled now")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Empty spots × net")
    y2 = y + 210
    section(d, M, y2, CW, "Before you set next year's rate")
    notes = ["Your fixed costs run at eleven children and at six. Only enrolment changes.",
             "An empty spot is not a light day — it is the most expensive thing in the program.",
             "Raise the rate for new families first. Give current families notice, not a discount."]
    for j, n in enumerate(notes):
        d.text((M + 24, y2 + 110 + j * 84), "•  " + n, font=fs(29, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p02(i):
    img, d = page("Monthly Costs Worksheet", "What it costs before a child walks in")
    table(d, M, TOP, CW, ["COST LINE", "CATEGORY", "MONTHLY"], [0.0, 0.52, 0.80], 15, 155)
    y = TOP + 78 + 15 * 155 + 36
    field(d, M, y, CW / 3 - 20, "Total fixed / month"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Cost per child")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Total costs")
    i.append(img)


def p03(i):
    img, d = page("Enrollment Form", "Everything you need on file")
    field(d, M, TOP, CW / 2 - 30, "Child's full name"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Date of birth")
    field(d, M, TOP + 150, CW / 3 - 20, "Start date"); field(d, M + CW / 3 + 10, TOP + 150, CW / 3 - 20, "Schedule")
    field(d, M + 2 * CW / 3 + 20, TOP + 150, CW / 3 - 20, "Weekly rate")
    section(d, M, TOP + 320, CW, "Parents & guardians")
    table(d, M, TOP + 400, CW, ["NAME", "RELATIONSHIP", "PHONE", "EMAIL"], [0.0, 0.32, 0.52, 0.72], 4, 140)
    y = TOP + 400 + 78 + 4 * 140 + 40
    section(d, M, y, CW, "Authorized for pickup")
    table(d, M, y + 80, CW, ["NAME", "RELATIONSHIP", "PHONE"], [0.0, 0.40, 0.70], 4, 140)
    y2 = y + 80 + 78 + 4 * 140 + 40
    field(d, M, y2, CW / 2 - 30, "Parent signature"); field(d, M + CW / 2 + 30, y2, CW / 2 - 30, "Date")
    i.append(img)


def p04(i):
    img, d = page("Emergency & Allergy Card", "The page an inspector asks for first")
    field(d, M, TOP, CW / 2 - 30, "Child"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Date of birth")
    section(d, M, TOP + 160, CW, "Emergency contacts (in order)")
    table(d, M, TOP + 240, CW, ["NAME", "RELATIONSHIP", "PHONE"], [0.0, 0.40, 0.70], 4, 140)
    y = TOP + 240 + 78 + 4 * 140 + 40
    section(d, M, y, CW, "Allergies, medications & medical conditions")
    for j in range(5):
        d.line((M + 20, y + 150 + j * 100, W - M - 20, y + 150 + j * 100), fill=LINE, width=2)
    y2 = y + 150 + 5 * 100 + 40
    field(d, M, y2, CW / 3 - 20, "Doctor"); field(d, M + CW / 3 + 10, y2, CW / 3 - 20, "Doctor phone")
    field(d, M + 2 * CW / 3 + 20, y2, CW / 3 - 20, "Insurance / policy #")
    y3 = y2 + 190
    field(d, M, y3, CW / 2 - 30, "Consent to treat — signature"); field(d, M + CW / 2 + 30, y3, CW / 2 - 30, "Date")
    i.append(img)


def p05(i):
    img, d = page("Daily Sheet", "What the parent takes home")
    field(d, M, TOP, CW / 2 - 30, "Child"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Date")
    section(d, M, TOP + 160, CW, "Meals & snacks")
    table(d, M, TOP + 240, CW, ["MEAL", "WHAT THEY ATE", "ALL / SOME / NONE"], [0.0, 0.26, 0.74], 4, 140)
    y = TOP + 240 + 78 + 4 * 140 + 36
    section(d, M, y, CW, "Naps & nappies / bathroom")
    table(d, M, y + 80, CW, ["TIME", "NAP / CHANGE", "NOTES"], [0.0, 0.18, 0.44], 5, 130)
    y2 = y + 80 + 78 + 5 * 130 + 36
    section(d, M, y2, CW, "Today we played, learned & loved")
    for j in range(4):
        d.line((M + 20, y2 + 150 + j * 100, W - M - 20, y2 + 150 + j * 100), fill=LINE, width=2)
    y3 = y2 + 150 + 4 * 100 + 30
    field(d, M, y3, CW / 2 - 30, "Please send tomorrow"); field(d, M + CW / 2 + 30, y3, CW / 2 - 30, "Mood today")
    i.append(img)


def p06(i):
    img, d = page("Attendance Sheet", "Signed in, signed out")
    field(d, M, TOP, CW / 2 - 30, "Week beginning"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Provider")
    table(d, M, TOP + 150, CW, ["CHILD", "IN", "PARENT SIGN", "OUT", "PARENT SIGN"],
          [0.0, 0.30, 0.40, 0.62, 0.72], 15, 132)
    y = TOP + 150 + 78 + 15 * 132 + 30
    field(d, M, y, CW / 3 - 20, "Children present"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Absent")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Late pickups")
    i.append(img)


def p07(i):
    img, d = page("Tuition & Payment Log", "Who has paid, who is behind")
    table(d, M, TOP, CW, ["CHILD", "DUE", "PAID", "DATE", "BALANCE"],
          [0.0, 0.34, 0.50, 0.66, 0.84], 17, 132)
    y = TOP + 78 + 17 * 132 + 30
    field(d, M, y, CW / 3 - 20, "Total due"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Total collected")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Still outstanding")
    i.append(img)


def p08(i):
    img, d = page("CACFP Meal Count", "Claim every meal you served")
    field(d, M, TOP, CW / 2 - 30, "Week beginning"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Program / claim #")
    table(d, M, TOP + 150, CW, ["CHILD", "MON", "TUE", "WED", "THU", "FRI"],
          [0.0, 0.34, 0.47, 0.60, 0.73, 0.86], 14, 138)
    y = TOP + 150 + 78 + 14 * 138 + 30
    field(d, M, y, CW / 3 - 20, "Breakfasts"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Lunches")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Snacks")
    i.append(img)


def p09(i):
    img, d = page("Ratios & Daily Schedule", "The rule that closes programs")
    table(d, M, TOP, CW, ["AGE GROUP", "CHILDREN", "STATE RATIO", "CAREGIVERS NEEDED"],
          [0.0, 0.40, 0.58, 0.78], 6, 155)
    y = TOP + 78 + 6 * 155 + 36
    field(d, M, y, CW / 3 - 20, "Caregivers on shift"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Children per caregiver")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Covered?  Y / N")
    y2 = y + 190
    section(d, M, y2, CW, "Daily schedule")
    table(d, M, y2 + 80, CW, ["TIME", "ACTIVITY", "WHERE"], [0.0, 0.20, 0.70], 10, 132)
    i.append(img)


def p10(i):
    img, d = page("Compliance File Checklist", "Complete, before they ask")
    field(d, M, TOP, CW / 2 - 30, "Child"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "File reviewed")
    half = CW / 2 - 30
    groups = [("Child's file", ["Enrollment form signed", "Emergency & allergy card", "Immunization record",
                                "Consent to treat", "Photo & field trip consent", "Signed tuition contract",
                                "Authorized pickup list", "Infant feeding plan"]),
              ("Program file", ["Current licence displayed", "Insurance certificate", "CPR & first aid current",
                                "Background checks on file", "Emergency evacuation plan", "Fire drill log",
                                "Menu posted & CACFP records", "Incident & injury log"])]
    for k, (title, items) in enumerate(groups):
        x = M + k * (half + 60)
        section(d, x, TOP + 180, half, title)
        for j, it in enumerate(items):
            cy = TOP + 276 + j * 190
            checkbox(d, x, cy)
            d.text((x + 72, cy + 24), it, font=fs(27, bold=False), fill=TEXT, anchor="lt")
            d.line((x + 72, cy + 108, x + half, cy + 108), fill=LINE, width=2)
    y = TOP + 276 + 8 * 190 + 40
    field(d, M, y, CW / 2 - 30, "Missing items"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Follow up by")
    i.append(img)


def p11(i):
    img, d = page("Tax Set-Aside", "A quarter of it is not yours")
    table(d, M, TOP, CW, ["QUARTER", "INCOME", "SET ASIDE", "ESTIMATED DUE", "SHORT BY"],
          [0.0, 0.26, 0.44, 0.62, 0.83], 5, 175)
    y = TOP + 78 + 5 * 175 + 40
    field(d, M, y, CW / 3 - 20, "Year income"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Year set aside")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Still to save")
    y2 = y + 200
    section(d, M, y2, CW, "Deductions providers forget")
    items = ["Time-space percentage of the home", "Food actually served (beyond CACFP)",
             "Toys, curriculum, books & craft supplies", "Mileage to shops, trips & training",
             "Licensing, insurance & professional dues", "Depreciation on furniture & play equipment"]
    for j, it in enumerate(items):
        cy = y2 + 130 + j * 116
        checkbox(d, M + 20, cy)
        d.text((M + 110, cy + 24), it, font=fs(29, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p12(i):
    img, d = page("Monthly Summary & Review", "Is the program actually working?")
    table(d, M, TOP, CW, ["MONTH", "ENROLLED", "REVENUE", "COSTS", "YOUR PAY", "$ / HOUR"],
          [0.0, 0.24, 0.42, 0.58, 0.74, 0.89], 8, 168)
    y = TOP + 78 + 8 * 168 + 40
    section(d, M, y, CW, "Six checks before next month")
    checks = ["Spots filled at or above my goal", "Margin healthy after every cost",
              "Fixed costs more than covered", "Ratios comfortable, not just legal",
              "My hourly at or above my goal", "Tax set aside, not spent"]
    for j, c in enumerate(checks):
        cy = y + 130 + j * 116
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
    pdf_path = os.path.join(out_dir, "Daycare_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
