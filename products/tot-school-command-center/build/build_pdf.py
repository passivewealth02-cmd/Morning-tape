"""Build the printable PDF pack for Tot-School Command Center™ (12 pages, US Letter).

Ink-light, print-ready forms on white with a forest-green header band & gold rules:

  1  Tot-at-a-Glance              7  Board-Book Log
  2  Weekly Theme Plan            8  Outings & Nature
  3  Daily Rhythm                 9  Sensory Play Bank
  4  Milestones Checklist         10 Ready-for-Preschool Checklist
  5  First Words & Concepts       11 Tot Goals
  6  Tot-Tray Planner             12 Portfolio of Firsts

Outputs ../Tot_School_Printables.pdf and page PNGs in ../marketing/print/.
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
    d.text((M, 90), "TOT-SCHOOL COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(78), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "The Bennett Family  ·  2026–2027", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Tot-School Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Tot-at-a-Glance", "One sweet page per tot — who they are & what they love right now")
    d.rounded_rectangle((W - M - 420, TOP, W - M, TOP + 520), radius=18, outline=GOLD_LT, width=4)
    d.text((W - M - 210, TOP + 260), "photo", font=fs(30, bold=False), fill=MUTED, anchor="mm")
    lw = CW - 480
    for i, lab in enumerate(["Tot's name", "Age & stage", "Personality", "Comfort item",
                             "Nap schedule", "Words so far"]):
        field(d, M, TOP + i * 92, lw, lab)
    y = TOP + 560
    section(d, M, y, CW, "Favorites right now"); y += 96
    labs = ["Favorite book", "Favorite song", "Favorite toy", "Favorite food", "Favorite person", "Silliest thing"]
    for i, lab in enumerate(labs):
        col = i % 2
        field(d, M + col * (CW / 2 + 30), y, CW / 2 - 30, lab)
        if col == 1:
            y += 92
    y += 20
    section(d, M, y, CW, "This year we hope to…"); y += 96
    for i in range(3):
        checkbox(d, M, y + i * 84, "", size=40)
        d.line((M + 66, y + i * 84 + 40, M + CW, y + i * 84 + 40), fill=LINE, width=2)
    y += 3 * 84 + 40
    section(d, M, y, CW, "Notes"); y += 96
    for i in range(2):
        d.line((M, y + i * 80 + 34, M + CW, y + i * 80 + 34), fill=LINE, width=2)
    imgs.append(img)


def p02(imgs):
    img, d = page("Weekly Theme Plan", "One gentle theme a week — a few books, songs & tot trays")
    y = TOP
    field(d, M, y, CW / 2 - 30, "This week's theme"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Focus / a word or two")
    y += 130
    section(d, M, y, CW, "Board books to read this week"); y += 96
    for i in range(3):
        d.line((M, y + i * 78 + 34, M + CW, y + i * 78 + 34), fill=LINE, width=2)
    y += 3 * 78 + 30
    section(d, M, y, CW, "Tot trays & play by day"); y += 96
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    for dn in days:
        d.rounded_rectangle((M, y, M + 340, y + 66), radius=8, fill=SURFACE)
        d.text((M + 20, y + 33), dn, font=fs(26), fill=PRIMARY, anchor="lm")
        d.line((M + 380, y + 50, M + CW, y + 50), fill=LINE, width=2)
        y += 96
    y += 10
    section(d, M, y, CW, "Song / fingerplay  ·  Sensory idea  ·  Snack"); y += 96
    for i in range(2):
        d.line((M, y + i * 80 + 34, M + CW, y + i * 80 + 34), fill=LINE, width=2)
    imgs.append(img)


def p03(imgs):
    img, d = page("Daily Rhythm", "Your day's loop around meals & naps — rhythm, not a rigid clock")
    times = ["Wake & breakfast", "Tot time (song, book, tray)", "Snack & board book",
             "Outside / movement", "Sensory or messy play", "Lunch", "Nap / quiet time",
             "Open play (blocks, pretend)", "Snack", "Outing or backyard", "Tidy-up & wind-down"]
    y = TOP
    field(d, M, y, CW / 2 - 30, "Tot / whole family"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Day")
    y += 130
    for t in times:
        d.rounded_rectangle((M, y, M + 720, y + 66), radius=8, fill=SURFACE)
        d.text((M + 20, y + 33), t, font=fs(23), fill=PRIMARY, anchor="lm")
        d.line((M + 760, y + 50, M + CW, y + 50), fill=LINE, width=2)
        y += 92
    imgs.append(img)


def p04(imgs):
    img, d = page("Milestones Checklist", "The whole-child picture — check them off as they emerge & are met")
    cols = [
        ("Gross & Fine Motor", ["Walks well", "Runs", "Kicks a ball", "Climbs stairs w/ rail",
                                 "Jumps two feet", "Stacks 4+ blocks", "Scribbles", "Uses a spoon"]),
        ("Language", ["Says 20+ words", "Two-word phrases", "Names body parts", "Follows a direction",
                       "Waves bye-bye", "Says mama / dada", "Points to ask", "Enjoys songs"]),
        ("Social & Self-Help", ["Plays alongside others", "Shows affection", "Imitates grown-ups",
                                 "Pretend play", "Open cup", "Finger-feeds", "Helps dress", "Potty interest"]),
        ("Cognitive", ["Sorts shapes", "Cause & effect", "Points to pictures", "Finds hidden object",
                        "Object permanence", "Simple sorting"]),
    ]
    y0 = TOP
    for ci, (head, items) in enumerate(cols):
        x = M + (ci % 2) * (CW / 2 + 20)
        y = y0 + (ci // 2) * 1120
        section(d, x, y, CW / 2 - 20, head)
        for i, it in enumerate(items):
            checkbox(d, x, y + 90 + i * 118, it, size=48, font=fs(27, bold=False))
    imgs.append(img)


def p05(imgs):
    img, d = page("First Words & Concepts", "The early basics — tick each one off as it clicks")
    cols = [
        ("Colors", ["Red", "Blue", "Yellow", "Green", "Orange", "Purple"]),
        ("Shapes & Counting", ["Circle", "Square", "Triangle", "Counts to 3", "Counts to 5", "Big / little"]),
        ("Body Parts", ["Nose", "Eyes", "Mouth", "Ears", "Hands", "Feet", "Tummy", "Hair"]),
        ("Animal Sounds", ["Dog — woof", "Cat — meow", "Cow — moo", "Duck — quack", "Sheep — baa", "Pig — oink"]),
    ]
    y0 = TOP
    for ci, (head, items) in enumerate(cols):
        x = M + (ci % 2) * (CW / 2 + 20)
        y = y0 + (ci // 2) * 1120
        section(d, x, y, CW / 2 - 20, head)
        for i, it in enumerate(items):
            checkbox(d, x, y + 90 + i * 118, it, size=48, font=fs(27, bold=False))
    y = y0 + 2 * 1120 - 120
    field(d, M, y, CW / 2 - 30, "Words spoken so far")
    field(d, M + CW / 2 + 30, y, CW / 2 - 30, "New words this month")
    imgs.append(img)


def p06(imgs):
    img, d = page("Tot-Tray Planner", "Simple invitations to play — one tray at a time, tied to a skill")
    d.text((M, TOP), "THEME:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 160, TOP + 34, M + 900, TOP + 34), fill=LINE, width=2)
    d.text((M + 1000, TOP), "WEEK OF:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1220, TOP + 34, M + 1600, TOP + 34), fill=LINE, width=2)
    table(d, M, TOP + 90, CW, ["#", "TOT TRAY / ACTIVITY", "SKILL AREA", "DONE"],
          [0.0, 0.08, 0.58, 0.86], 16, rowh=130,
          filled_rows=[[str(i + 1)] for i in range(16)])
    imgs.append(img)


def p07(imgs):
    img, d = page("Board-Book Log", "The heart of tot-school — every favorite & how many times")
    d.text((M, TOP), "TOT:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 120, TOP + 34, M + 900, TOP + 34), fill=LINE, width=2)
    d.text((M + 1000, TOP), "ROTATION:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1240, TOP + 34, M + 1600, TOP + 34), fill=LINE, width=2)
    table(d, M, TOP + 90, CW, ["#", "TITLE", "AUTHOR", "TIMES", "★"],
          [0.0, 0.08, 0.56, 0.78, 0.91], 18, rowh=120,
          filled_rows=[[str(i + 1)] for i in range(18)])
    imgs.append(img)


def p08(imgs):
    img, d = page("Outings & Nature", "Little adventures — every outing, what it tied to & a memory")
    table(d, M, TOP, CW, ["DATE", "PLACE", "THEME TIE-IN", "MEMORY / NOTE", "★"],
          [0.0, 0.15, 0.42, 0.63, 0.92], 16, rowh=130)
    imgs.append(img)


def p09(imgs):
    img, d = page("Sensory Play Bank", "Grab-and-go messy play for busy little hands — jot your favorites")
    table(d, M, TOP, CW, ["ACTIVITY", "AREA", "MATERIALS", "BUILDS"],
          [0.0, 0.34, 0.52, 0.78], 20, rowh=130)
    imgs.append(img)


def p10(imgs):
    img, d = page("Ready for Preschool", "A gentle 'getting there' guide — every tot grows on their own timeline")
    cols = [
        ["Separates from caregiver briefly", "Follows a simple routine", "Uses words to ask", "Points to communicate",
         "Feeds self finger foods", "Drinks from a cup", "Shows potty interest", "Plays near other children"],
        ["Sits for a short story", "Stacks a few blocks", "Scribbles with a crayon", "Waves & greets people",
         "Names a few body parts", "Follows a one-step direction", "Naps on a schedule", "Beginning to choose / say no"],
    ]
    for c, items in enumerate(cols):
        x = M + c * (CW / 2 + 20)
        for i, it in enumerate(items):
            checkbox(d, x, TOP + i * 150, it, size=54, font=fs(28, bold=False))
    y = TOP + 8 * 150 + 30
    section(d, M, y, CW, "What we're gently working on next"); y += 100
    for i in range(3):
        d.line((M, y + i * 84 + 30, M + CW, y + i * 84 + 30), fill=LINE, width=2)
    imgs.append(img)


def p11(imgs):
    img, d = page("Tot Goals", "The little wins that matter — for each tot & the whole family")
    y = TOP
    for who in ["Tot 1", "Tot 2", "Together as a family", "A goal for me (the grown-up)"]:
        section(d, M, y, CW, f"{who} — goals this year"); y += 96
        for i in range(3):
            checkbox(d, M, y + i * 88, "", size=44)
            d.line((M + 70, y + i * 88 + 44, M + CW, y + i * 88 + 44), fill=LINE, width=2)
        y += 3 * 88 + 50
    imgs.append(img)


def p12(imgs):
    img, d = page("Portfolio of Firsts", "Tape in a first each month — these tiny years go by so fast")
    y = TOP
    field(d, M, y, CW / 2 - 30, "Tot"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Month & year")
    y += 130
    caps = ["First steps", "First scribble", "A proud milestone", "A messy-play grin"]
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
    pdf_path = os.path.join(out_dir, "Tot_School_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
