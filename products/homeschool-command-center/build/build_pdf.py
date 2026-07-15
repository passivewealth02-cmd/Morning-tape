"""Build the printable PDF pack for Homeschool Command Center™ (12 pages, US Letter).

Ink-light, print-ready forms on white with a forest-green header band & gold rules:

  1  Student-at-a-Glance          7  Field-Trip Log
  2  Weekly Lesson Plan           8  Assignment & Grade Sheet
  3  Daily Homeschool Schedule    9  Book List & Wish List
  4  Attendance & Hours Log      10  Homeschool Goals
  5  Subject / Course Plan       11  Records & Compliance Checklist
  6  Reading Log                 12  Report Card / Transcript

Outputs ../Homeschool_Printables.pdf and page PNGs in ../marketing/print/.
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
    d.text((M, 90), "HOMESCHOOL COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(78), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "The Bennett Family  ·  2026–2027", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Homeschool Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Student-at-a-Glance", "One page per student — the plan, the goals & what makes them tick")
    d.rounded_rectangle((W - M - 420, TOP, W - M, TOP + 520), radius=18, outline=GOLD_LT, width=4)
    d.text((W - M - 210, TOP + 260), "photo", font=fs(30, bold=False), fill=MUTED, anchor="mm")
    lw = CW - 480
    for i, lab in enumerate(["Student name", "Grade & age", "Approach / method", "Learning style",
                             "Strengths", "Working on"]):
        field(d, M, TOP + i * 92, lw, lab)
    y = TOP + 560
    section(d, M, y, CW, "This year's subjects & curriculum"); y += 96
    for i in range(6):
        field(d, M, y, CW / 2 - 30, "Subject") if i % 2 == 0 else field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Curriculum")
        if i % 2 == 1:
            y += 92
    y += 20
    section(d, M, y, CW, "Goals for the year"); y += 96
    for i in range(3):
        checkbox(d, M, y + i * 84, "", size=40)
        d.line((M + 66, y + i * 84 + 40, M + CW, y + i * 84 + 40), fill=LINE, width=2)
    y += 3 * 84 + 40
    section(d, M, y, CW, "Notes"); y += 96
    for i in range(2):
        d.line((M, y + i * 80 + 34, M + CW, y + i * 80 + 34), fill=LINE, width=2)
    imgs.append(img)


def p02(imgs):
    img, d = page("Weekly Lesson Plan", "Map the week, subject by subject — then check it off")
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    subs = ["Math", "Language Arts", "Science", "History", "Extras / Electives", "Read-aloud"]
    gx, gy = M, TOP + 10
    gw = CW; gh = 2500
    colw = gw / (len(days) + 1)
    rowh = gh / (len(subs) + 1)
    d.rounded_rectangle((gx, gy, gx + gw, gy + rowh), radius=8, fill=PRIMARY)
    for i, dn in enumerate(days):
        d.text((gx + colw * (i + 1) + colw / 2, gy + rowh / 2), dn, font=fs(28), fill=WHITE, anchor="mm")
    for r, rn in enumerate(subs):
        ry = gy + rowh * (r + 1)
        d.rectangle((gx, ry, gx + colw, ry + rowh), fill=SURFACE)
        d.text((gx + 20, ry + rowh / 2), rn, font=fs(24), fill=PRIMARY, anchor="lm")
    for i in range(len(days) + 2):
        d.line((gx + colw * i, gy, gx + colw * i, gy + gh), fill=LINE, width=2)
    for r in range(len(subs) + 2):
        d.line((gx, gy + rowh * r, gx + gw, gy + rowh * r), fill=LINE, width=2)
    imgs.append(img)


def p03(imgs):
    img, d = page("Daily Homeschool Schedule", "Your family's rhythm — a loop, not a rigid clock")
    times = ["Morning basket", "Chores & breakfast", "Block 1", "Block 2", "Break / recess",
             "Block 3", "Lunch", "Read-aloud / rest", "Block 4", "Afternoon / activities", "Wrap-up & tidy"]
    y = TOP
    field(d, M, y, CW / 2 - 30, "Student / whole family"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Day")
    y += 130
    for t in times:
        d.rounded_rectangle((M, y, M + 340, y + 66), radius=8, fill=SURFACE)
        d.text((M + 20, y + 33), t, font=fs(24), fill=PRIMARY, anchor="lm")
        d.line((M + 380, y + 50, M + CW, y + 50), fill=LINE, width=2)
        y += 92
    imgs.append(img)


def p04(imgs):
    img, d = page("Attendance & Hours Log", "Month by month — the days & hours your state may ask for")
    d.text((M, TOP), "MONTH:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 160, TOP + 34, M + 700, TOP + 34), fill=LINE, width=2)
    d.text((M + 760, TOP), "GOAL DAYS:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1000, TOP + 34, M + 1300, TOP + 34), fill=LINE, width=2)
    weeks = ["Week 1", "Week 2", "Week 3", "Week 4", "Week 5"]
    table(d, M, TOP + 90, CW, ["WEEK", "M", "T", "W", "TH", "F", "DAYS", "HOURS"],
          [0.0, 0.18, 0.30, 0.42, 0.54, 0.66, 0.78, 0.89], 5, rowh=180,
          filled_rows=[[w] for w in weeks])
    y = TOP + 90 + 78 + 5 * 180 + 60
    section(d, M, y, CW, "Month totals"); y += 100
    for lab in ["Days this month", "Hours this month", "Running total — days", "Running total — hours"]:
        field(d, M, y, CW / 2 - 30, lab)
        if "hours" in lab.lower() and "Running" not in lab:
            pass
        y += 92 if lab in ("Hours this month", "Running total — hours") else 0
        if lab in ("Days this month", "Running total — days"):
            field(d, M + CW / 2 + 30, y, CW / 2 - 30, lab.replace("Days", "Hours").replace("days", "hours"))
            y += 92
    imgs.append(img)


def p05(imgs):
    img, d = page("Subject / Course Plan", "Your course of study — scope & how you'll cover each subject")
    table(d, M, TOP, CW, ["SUBJECT", "SCOPE / GOALS THIS YEAR", "CURRICULUM", "CADENCE"],
          [0.0, 0.24, 0.62, 0.84], 11, rowh=180,
          filled_rows=[["Math"], ["Lang Arts"], ["Science"], ["History"], ["Geography"], ["Art"]])
    imgs.append(img)


def p06(imgs):
    img, d = page("Reading Log", "Every book, every child — reading is the heart of it")
    d.text((M, TOP), "STUDENT:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 200, TOP + 34, M + 900, TOP + 34), fill=LINE, width=2)
    d.text((M + 1000, TOP), "GOAL:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1140, TOP + 34, M + 1500, TOP + 34), fill=LINE, width=2)
    table(d, M, TOP + 90, CW, ["#", "DATE", "TITLE", "AUTHOR", "PAGES", "★"],
          [0.0, 0.08, 0.24, 0.62, 0.80, 0.92], 18, rowh=120,
          filled_rows=[[str(i + 1)] for i in range(18)])
    imgs.append(img)


def p07(imgs):
    img, d = page("Field-Trip Log", "Learning beyond the table — trips, tie-ins & what they cost")
    table(d, M, TOP, CW, ["DATE", "PLACE", "SUBJECT TIE-IN", "WHO WENT", "COST", "★"],
          [0.0, 0.16, 0.44, 0.64, 0.82, 0.92], 16, rowh=130)
    imgs.append(img)


def p08(imgs):
    img, d = page("Assignment & Grade Sheet", "A clean record of graded work — the source of report cards")
    d.text((M, TOP), "STUDENT:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 200, TOP + 34, M + 900, TOP + 34), fill=LINE, width=2)
    d.text((M + 1000, TOP), "SUBJECT:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1200, TOP + 34, M + 1600, TOP + 34), fill=LINE, width=2)
    table(d, M, TOP + 90, CW, ["DATE", "ASSIGNMENT", "POSSIBLE", "SCORE", "GRADE"],
          [0.0, 0.16, 0.58, 0.72, 0.86], 18, rowh=120)
    imgs.append(img)


def p09(imgs):
    img, d = page("Book List & Wish List", "What you're reading & what to borrow, buy or save for later")
    table(d, M, TOP, CW, ["TITLE", "AUTHOR", "FOR", "HAVE", "WANT"],
          [0.0, 0.40, 0.62, 0.78, 0.89], 20, rowh=130)
    imgs.append(img)


def p10(imgs):
    img, d = page("Homeschool Goals", "The wins that matter — for each child & the whole family")
    y = TOP
    for who in ["Student 1", "Student 2", "Student 3", "Family"]:
        section(d, M, y, CW, f"{who} — goals this year"); y += 96
        for i in range(3):
            checkbox(d, M, y + i * 88, "", size=44)
            d.line((M + 70, y + i * 88 + 44, M + CW, y + i * 88 + 44), fill=LINE, width=2)
        y += 3 * 88 + 50
    imgs.append(img)


def p11(imgs):
    img, d = page("Records & Compliance Checklist", "Keep the paper trail tidy — confirm your own state's rules")
    cols = [
        ["Letter of intent / notification filed", "Immunization records or exemption",
         "Attendance log kept current", "Instructional-hours log kept", "Curriculum plan documented",
         "Portfolio started & maintained", "Work samples saved per subject"],
        ["Standardized test / assessment scheduled", "Report cards / transcript updated",
         "Evaluator or umbrella-school contact", "End-of-year evaluation booked",
         "Copies backed up (cloud / binder)", "Next-year plan drafted", "Kept: this is not legal advice"],
    ]
    for c, items in enumerate(cols):
        x = M + c * (CW / 2 + 20)
        for i, it in enumerate(items):
            checkbox(d, x, TOP + i * 170, it, size=54, font=fs(28, bold=False))
    y = TOP + 7 * 170 + 30
    section(d, M, y, CW, "My state's specific requirements (write them here)"); y += 100
    for i in range(3):
        d.line((M, y + i * 84 + 30, M + CW, y + i * 84 + 30), fill=LINE, width=2)
    imgs.append(img)


def p12(imgs):
    img, d = page("Report Card / Transcript", "A tidy term-by-term record — print, sign & file")
    y = TOP
    field(d, M, y, CW / 2 - 30, "Student"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Grade & year")
    y += 130
    table(d, M, y, CW, ["SUBJECT", "TERM 1", "TERM 2", "TERM 3", "TERM 4", "FINAL"],
          [0.0, 0.40, 0.52, 0.64, 0.76, 0.88], 9, rowh=120)
    y2 = y + 78 + 9 * 120 + 60
    section(d, M, y2, CW, "Comments"); y2 += 96
    for i in range(3):
        d.line((M, y2 + i * 80 + 30, M + CW, y2 + i * 80 + 30), fill=LINE, width=2)
    y2 += 3 * 80 + 50
    field(d, M, y2, CW / 2 - 30, "Parent / teacher signature")
    field(d, M + CW / 2 + 30, y2, CW / 2 - 30, "Date")
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
    pdf_path = os.path.join(out_dir, "Homeschool_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
