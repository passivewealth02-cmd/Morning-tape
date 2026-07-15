"""Build the printable PDF pack for Back-to-School Command Center™ (12 pages, US Letter).

Ink-light, print-ready forms on white with a forest-green header band & gold rules:

  1  Child Information Sheet        7  Clothing-Size Tracker
  2  School Contact Page           8  Lunchbox Planner
  3  First-Day Prep Checklist       9  Field-Trip & Payment Log
  4  Backpack Checklist            10  Parent-Teacher Meeting Notes
  5  Weekly Family Schedule        11  School-Year Goals
  6  School Supply Checklist       12  First-Day & Last-Day Memory Pages

Outputs ../Back_to_School_Printables.pdf and page PNGs in ../marketing/print/.
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

W, H = 2550, 3300      # US Letter @ 300 dpi
M = 190                # margin

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
    # header band
    bh = 340
    d.rectangle((0, 0, W, bh), fill=PRIMARY)
    d.rectangle((0, bh, W, bh + 10), fill=GOLD_LT)
    d.rectangle((0, bh + 10, W, bh + 14), fill=GOLD_HI)
    d.text((M, 90), "BACK-TO-SCHOOL COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(78), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    # footer
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "The Rivera Family  ·  2026–2027", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Back-to-School Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Child Information Sheet", "One page per child — everything a sitter, nurse or office might ask")
    # photo box
    d.rounded_rectangle((W - M - 420, TOP, W - M, TOP + 520), radius=18, outline=GOLD_LT, width=4)
    d.text((W - M - 210, TOP + 260), "photo", font=fs(30, bold=False), fill=MUTED, anchor="mm")
    lw = CW - 480
    labels1 = ["Full name", "Nickname", "Date of birth", "Grade & teacher", "School", "Bus / route"]
    for i, lab in enumerate(labels1):
        field(d, M, TOP + i * 92, lw, lab)
    y = TOP + 560
    section(d, M, y, CW, "Sizes & health"); y += 96
    grid = [("Shirt / dress", "Pants"), ("Shoe", "Coat"), ("Allergies", "Medications"), ("Doctor", "Doctor phone")]
    for a, b in grid:
        field(d, M, y, CW / 2 - 30, a); field(d, M + CW / 2 + 30, y, CW / 2 - 30, b); y += 92
    y += 20
    section(d, M, y, CW, "Emergency contacts"); y += 96
    for a, b in [("Parent 1 / phone", "Parent 2 / phone"), ("Backup contact / phone", "Relationship")]:
        field(d, M, y, CW / 2 - 30, a); field(d, M + CW / 2 + 30, y, CW / 2 - 30, b); y += 92
    y += 20
    section(d, M, y, CW, "Notes (what helps this child thrive)"); y += 96
    for i in range(3):
        d.line((M, y + i * 84 + 40, M + CW, y + i * 84 + 40), fill=LINE, width=2)
    imgs.append(img)


def p02(imgs):
    img, d = page("School Contact Page", "Every school, office & teacher number in one place")
    rows = [
        ("Lincoln High", "Front Office", "(555) 201-4000", "office@lincoln.edu"),
        ("Lincoln High", "Advisor", "(555) 201-4021", "alvarez@lincoln.edu"),
        ("Jefferson Middle", "Front Office", "(555) 201-3000", "office@jeffms.edu"),
        ("Jefferson Middle", "Band", "(555) 201-3044", "doyle@jeffms.edu"),
        ("Oakwood Elem.", "Front Office", "(555) 201-2000", "office@oakwood.edu"),
        ("Oakwood Elem.", "Nurse", "(555) 201-2050", "nurse@oakwood.edu"),
        ("Little Sprouts", "Ms. Dana", "(555) 201-1000", "hello@littlesprouts.com"),
        ("District", "Transportation", "(555) 201-9000", "bus@district.org"),
    ]
    table(d, M, TOP, CW, ["SCHOOL", "CONTACT", "PHONE", "EMAIL"], [0.0, 0.28, 0.52, 0.74], 14, rowh=104,
          filled_rows=rows)
    y = TOP + 78 + 14 * 104 + 60
    section(d, M, y, CW, "Also good to have")
    y += 96
    for lab in ["Attendance / absence line", "After-care / aftercare phone", "PTA / room parent", "Pediatrician"]:
        field(d, M, y, CW, lab); y += 92
    imgs.append(img)


def p03(imgs):
    img, d = page("First-Day Preparation Checklist", "Do these before the first bell — and breathe")
    cols = [
        ["Forms signed & returned", "Immunizations up to date", "Supplies bought & labeled",
         "Backpack packed", "First-day outfit ready", "Shoes fit & laced", "Haircuts done",
         "Water bottle & lunchbox", "Bus route / drop-off confirmed", "Emergency contacts updated"],
        ["Bedtime routine restarted", "Alarm clocks set", "Breakfast plan for week 1",
         "Lunch plan for week 1", "Teacher names learned", "Classroom / room numbers noted",
         "After-care arranged", "Photo consent decided", "First-day photo spot picked",
         "Something special for the fridge"],
    ]
    for c, items in enumerate(cols):
        x = M + c * (CW / 2 + 20)
        for i, it in enumerate(items):
            checkbox(d, x, TOP + i * 128, it, size=52, font=fs(30, bold=False))
    imgs.append(img)


def p04(imgs):
    img, d = page("Backpack Checklist", "Tape it inside the closet door — every morning, every kid")
    cols = [
        ["Homework folder", "Reading book / log", "Signed notes & forms", "Water bottle",
         "Lunch or lunch money", "Snack", "Glasses / case", "Inhaler / meds (if needed)"],
        ["Pencil pouch", "Charged device (if allowed)", "Library books (due day)", "Gym clothes / shoes",
         "Instrument (band days)", "Permission slip + money", "Jacket / weather gear", "Show & tell item"],
    ]
    for c, items in enumerate(cols):
        x = M + c * (CW / 2 + 20)
        for i, it in enumerate(items):
            checkbox(d, x, TOP + i * 150, it, size=56, font=fs(32, bold=False))
    y = TOP + 8 * 150 + 40
    section(d, M, y, CW, "Weekly rhythm")
    y += 100
    for lab in ["Library day", "Gym / PE day", "Band / music day", "Early release / half day"]:
        field(d, M, y, CW / 2 - 30, lab)
        if lab == "Gym / PE day":
            pass
        y += 92
    imgs.append(img)


def p05(imgs):
    img, d = page("Weekly Family Schedule", "The whole crew, one grid — practices, pickups & who drives")
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Sat / Sun"]
    rows = ["Morning", "After school", "Evening", "Dinner", "Who drives"]
    gx, gy = M, TOP + 20
    gw = CW; gh = 2380
    colw = gw / (len(days) + 1)
    rowh = gh / (len(rows) + 1)
    # header row
    d.rounded_rectangle((gx, gy, gx + gw, gy + rowh), radius=8, fill=PRIMARY)
    for i, dname in enumerate(days):
        d.text((gx + colw * (i + 1) + colw / 2, gy + rowh / 2), dname, font=fs(28), fill=WHITE, anchor="mm")
    for r, rname in enumerate(rows):
        ry = gy + rowh * (r + 1)
        d.rectangle((gx, ry, gx + colw, ry + rowh), fill=SURFACE)
        d.text((gx + 24, ry + rowh / 2), rname, font=fs(26), fill=PRIMARY, anchor="lm")
    for i in range(len(days) + 2):
        d.line((gx + colw * i, gy, gx + colw * i, gy + gh), fill=LINE, width=2)
    for r in range(len(rows) + 2):
        d.line((gx, gy + rowh * r, gx + gw, gy + rowh * r), fill=LINE, width=2)
    imgs.append(img)


def p06(imgs):
    img, d = page("School Supply Checklist", "Shop once, for everyone — grouped so nothing's forgotten")
    groups = [
        ("The basics (each child)", ["Pencils & erasers", "Pens (blue/black/red)", "Highlighters",
                                     "Glue sticks", "Scissors", "Folders & binders", "Notebooks", "Ruler"]),
        ("By grade", ["Graphing calculator (HS)", "Composition books (MS)", "Crayons & markers (elem)",
                      "Watercolors (elem)", "Rest / nap mat (PreK-K)", "Headphones / earbuds"]),
        ("Household restock", ["Backpacks", "Lunchboxes & water bottles", "Label maker / name labels",
                               "Hand sanitizer & tissues", "Ziploc bags", "Printer paper & ink"]),
    ]
    x = M; y = TOP
    for gi, (title, items) in enumerate(groups):
        section(d, x, y, CW, title); y += 96
        for i, it in enumerate(items):
            col = i % 2
            cx = x + col * (CW / 2 + 20)
            checkbox(d, cx, y + (i // 2) * 110, it, size=46, font=fs(28, bold=False))
        y += ((len(items) + 1) // 2) * 110 + 60
    imgs.append(img)


def p07(imgs):
    img, d = page("Clothing-Size Tracker", "Every size in one glance — no more guessing in the store")
    kids = ["Mateo", "Sofia", "Liam", "Ava", "Noah", "Emma", "", ""]
    table(d, M, TOP, CW, ["CHILD", "SHIRT", "PANTS", "SHOE", "COAT", "DRESS"],
          [0.0, 0.22, 0.38, 0.54, 0.70, 0.85], 8, rowh=150,
          filled_rows=[[k] for k in kids if k])
    y = TOP + 78 + 8 * 150 + 60
    section(d, M, y, CW, "Notes")
    y += 96
    for i in range(2):
        d.line((M, y + i * 84 + 30, M + CW, y + i * 84 + 30), fill=LINE, width=2)
    imgs.append(img)


def p08(imgs):
    img, d = page("Lunchbox Planner", "A week of lunches everyone will actually eat")
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    table(d, M, TOP, CW, ["DAY", "MAIN", "FRUIT / VEG", "SNACK", "DRINK"],
          [0.0, 0.16, 0.44, 0.66, 0.85], 5, rowh=200,
          filled_rows=[[dn] for dn in days])
    y = TOP + 78 + 5 * 200 + 60
    section(d, M, y, CW, "Grocery list this makes")
    y += 100
    for i in range(2):
        cx = M + i * (CW / 2 + 20)
        for j in range(5):
            checkbox(d, cx, y + j * 96, "", size=44)
            d.line((cx + 70, y + j * 96 + 44, cx + CW / 2 - 40, y + j * 96 + 44), fill=LINE, width=2)
    imgs.append(img)


def p09(imgs):
    img, d = page("Field-Trip & Payment Log", "Permission slips & money in — nothing lost in a backpack")
    table(d, M, TOP, CW, ["DATE", "CHILD", "TRIP / ITEM", "AMOUNT", "SLIP IN", "PAID"],
          [0.0, 0.16, 0.34, 0.66, 0.80, 0.90], 16, rowh=130)
    imgs.append(img)


def p10(imgs):
    img, d = page("Parent-Teacher Meeting Notes", "Walk in prepared, walk out with a plan")
    for lab, w in [("Child", CW / 2 - 30), ("Date", CW / 2 - 30)]:
        pass
    field(d, M, TOP, CW / 2 - 30, "Child"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Teacher & date")
    y = TOP + 120
    section(d, M, y, CW, "Questions I want to ask"); y += 100
    for i in range(4):
        checkbox(d, M, y + i * 92, "", size=40)
        d.line((M + 66, y + i * 92 + 40, M + CW, y + i * 92 + 40), fill=LINE, width=2)
    y += 4 * 92 + 40
    section(d, M, y, CW, "What the teacher shared"); y += 100
    for i in range(5):
        d.line((M, y + i * 84 + 30, M + CW, y + i * 84 + 30), fill=LINE, width=2)
    y += 5 * 84 + 60
    section(d, M, y, CW, "Our plan / follow-up"); y += 100
    for i in range(4):
        d.line((M, y + i * 84 + 30, M + CW, y + i * 84 + 30), fill=LINE, width=2)
    imgs.append(img)


def p11(imgs):
    img, d = page("School-Year Goals", "One page to dream a little — for each child & the whole family")
    boxes = ["This year I want to learn…", "A habit I'll build…", "A kindness goal…", "Something I'll try that's new…"]
    y = TOP
    field(d, M, y, CW / 2 - 30, "Child"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Grade & year")
    y += 130
    for b in boxes:
        section(d, M, y, CW, b); y += 92
        for i in range(2):
            d.line((M, y + i * 78 + 34, M + CW, y + i * 78 + 34), fill=LINE, width=2)
        y += 2 * 78 + 50
    section(d, M, y, CW, "Family goal for the year"); y += 92
    for i in range(2):
        d.line((M, y + i * 78 + 34, M + CW, y + i * 78 + 34), fill=LINE, width=2)
    imgs.append(img)


def p12(imgs):
    img, d = page("First-Day & Last-Day Memory Page", "Same page in September and June — watch them grow")
    halfw = CW / 2 - 30
    for c, (when, tint) in enumerate([("FIRST DAY", MINT), ("LAST DAY", WARN)]):
        x = M + c * (halfw + 60)
        d.rounded_rectangle((x, TOP, x + halfw, TOP + 1120), radius=20, outline=GOLD_LT, width=4)
        d.rounded_rectangle((x, TOP, x + halfw, TOP + 70), radius=20, fill=PRIMARY)
        d.text((x + halfw / 2, TOP + 35), when + " OF SCHOOL", font=fs(30), fill=GOLD_HI, anchor="mm")
        # photo box
        d.rounded_rectangle((x + 40, TOP + 110, x + halfw - 40, TOP + 640), radius=14, outline=LINE, width=3)
        d.text((x + halfw / 2, TOP + 375), "tape a photo here", font=fs(26, bold=False), fill=MUTED, anchor="mm")
        prompts = ["Name", "Grade", "Teacher", "Age & height", "Favorite thing", "Wants to be…", "Best friend"]
        yy = TOP + 690
        for p in prompts:
            d.text((x + 46, yy), p.upper(), font=fs(20), fill=ACCENT, anchor="lt")
            d.line((x + 46 + 300, yy + 34, x + halfw - 46, yy + 34), fill=LINE, width=2)
            yy += 60
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
    pdf_path = os.path.join(out_dir, "Back_to_School_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
