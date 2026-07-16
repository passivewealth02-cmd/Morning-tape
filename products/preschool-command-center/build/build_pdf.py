"""Build the printable PDF pack for Preschool Command Center™ (12 pages, US Letter).

Ink-light, print-ready forms on white with a forest-green header band & gold rules:

  1  Child-at-a-Glance            7  Read-Aloud Log
  2  Weekly Theme Plan            8  Nature Walk & Field-Trip Log
  3  Daily Rhythm                 9  Arts & Sensory Idea Bank
  4  Skills & Milestones          10 Kindergarten Readiness Checklist
  5  ABC & 123 Chart              11 Preschool Goals
  6  Activity Planner             12 Portfolio & Keepsakes

Outputs ../Preschool_Printables.pdf and page PNGs in ../marketing/print/.
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
    d.text((M, 90), "PRESCHOOL COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(78), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "The Bennett Family  ·  2026–2027", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Preschool Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Child-at-a-Glance", "One sweet page per child — who they are & what they're working on")
    d.rounded_rectangle((W - M - 420, TOP, W - M, TOP + 520), radius=18, outline=GOLD_LT, width=4)
    d.text((W - M - 210, TOP + 260), "photo", font=fs(30, bold=False), fill=MUTED, anchor="mm")
    lw = CW - 480
    for i, lab in enumerate(["Child's name", "Age & stage", "Personality", "Loves",
                             "Working on", "Comfort item"]):
        field(d, M, TOP + i * 92, lw, lab)
    y = TOP + 560
    section(d, M, y, CW, "Favorites right now"); y += 96
    labs = ["Favorite book", "Favorite song", "Favorite play", "Favorite food", "Best friend", "Silliest word"]
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
    img, d = page("Weekly Theme Plan", "One gentle theme a week — a letter, some books & a little play")
    y = TOP
    field(d, M, y, CW / 2 - 30, "This week's theme"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Letter & number focus")
    y += 130
    section(d, M, y, CW, "Books to read this week"); y += 96
    for i in range(3):
        d.line((M, y + i * 78 + 34, M + CW, y + i * 78 + 34), fill=LINE, width=2)
    y += 3 * 78 + 30
    section(d, M, y, CW, "Activities by day"); y += 96
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    for dn in days:
        d.rounded_rectangle((M, y, M + 340, y + 66), radius=8, fill=SURFACE)
        d.text((M + 20, y + 33), dn, font=fs(26), fill=PRIMARY, anchor="lm")
        d.line((M + 380, y + 50, M + CW, y + 50), fill=LINE, width=2)
        y += 96
    y += 10
    section(d, M, y, CW, "Song / fingerplay  ·  Craft or sensory  ·  Snack idea"); y += 96
    for i in range(2):
        d.line((M, y + i * 80 + 34, M + CW, y + i * 80 + 34), fill=LINE, width=2)
    imgs.append(img)


def p03(imgs):
    img, d = page("Daily Rhythm", "Your day's loop — little ones thrive on rhythm, not a rigid clock")
    times = ["Welcome & free play", "Circle time (calendar, weather, song)", "Letter & theme intro",
             "Theme activity", "Snack & story", "Outside / gross motor", "Table time (ABC/123, fine motor)",
             "Lunch", "Rest / quiet time", "Open play", "Tidy-up & wind-down"]
    y = TOP
    field(d, M, y, CW / 2 - 30, "Child / whole family"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Day")
    y += 130
    for t in times:
        d.rounded_rectangle((M, y, M + 720, y + 66), radius=8, fill=SURFACE)
        d.text((M + 20, y + 33), t, font=fs(23), fill=PRIMARY, anchor="lm")
        d.line((M + 760, y + 50, M + CW, y + 50), fill=LINE, width=2)
        y += 92
    imgs.append(img)


def p04(imgs):
    img, d = page("Skills & Milestones", "The whole-child picture — check them off as they emerge & master")
    cols = [
        ("Fine & Gross Motor", ["Tripod crayon grip", "Cuts on a line", "Buttons & zips", "Strings beads",
                                 "Hops on one foot", "Pedals a trike", "Throws & catches", "Balances"]),
        ("Language & Pre-Reading", ["Full sentences", "Follows 2-step directions", "Names 8+ colors",
                                     "Rhymes words", "Recognizes own name", "Knows most letters",
                                     "Retells a story", "Print awareness"]),
        ("Pre-Math", ["Counts to 20", "Counts to 10", "Sorts by color/size", "Numerals 0–10",
                       "Copies AB patterns", "Names shapes"]),
        ("Social & Self-Help", ["Takes turns", "Shares", "Names feelings", "Separates calmly",
                                 "Dresses self", "Toilet independent", "Washes hands", "Cleans up"]),
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
    img, d = page("ABC & 123 Chart", "Letter by letter — recognizes, says the sound & writes it")
    half = CW / 2 - 30
    # left: A-M, right: N-Z
    letters = [chr(65 + i) for i in range(26)]
    table(d, M, TOP, half, ["LETTER", "SEES", "SAYS", "WRITES"],
          [0.0, 0.40, 0.60, 0.80], 13, rowh=140,
          filled_rows=[[L] for L in letters[:13]])
    table(d, M + CW / 2 + 30, TOP, half, ["LETTER", "SEES", "SAYS", "WRITES"],
          [0.0, 0.40, 0.60, 0.80], 13, rowh=140,
          filled_rows=[[L] for L in letters[13:]])
    y = TOP + 78 + 13 * 140 + 60
    section(d, M, y, CW, "Numbers & counting"); y += 100
    labs = ["Counts aloud to ___", "Recognizes numerals 0–___", "1-to-1 counting to ___", "Writes numerals 0–___"]
    for i, lab in enumerate(labs):
        col = i % 2
        field(d, M + col * (CW / 2 + 30), y, CW / 2 - 30, lab)
        if col == 1:
            y += 92
    imgs.append(img)


def p06(imgs):
    img, d = page("Activity Planner", "Play with a purpose — every activity tied to a theme & a skill")
    d.text((M, TOP), "THEME:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 160, TOP + 34, M + 900, TOP + 34), fill=LINE, width=2)
    d.text((M + 1000, TOP), "WEEK OF:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1220, TOP + 34, M + 1600, TOP + 34), fill=LINE, width=2)
    table(d, M, TOP + 90, CW, ["#", "ACTIVITY", "SKILL DOMAIN", "DONE"],
          [0.0, 0.08, 0.58, 0.86], 16, rowh=130,
          filled_rows=[[str(i + 1)] for i in range(16)])
    imgs.append(img)


def p07(imgs):
    img, d = page("Read-Aloud Log", "The heart of preschool — every favorite & how many times")
    d.text((M, TOP), "CHILD:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 160, TOP + 34, M + 900, TOP + 34), fill=LINE, width=2)
    d.text((M + 1000, TOP), "BOOKS GOAL:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1280, TOP + 34, M + 1600, TOP + 34), fill=LINE, width=2)
    table(d, M, TOP + 90, CW, ["#", "TITLE", "AUTHOR", "TIMES", "★"],
          [0.0, 0.08, 0.56, 0.78, 0.91], 18, rowh=120,
          filled_rows=[[str(i + 1)] for i in range(18)])
    imgs.append(img)


def p08(imgs):
    img, d = page("Nature Walk & Field-Trip Log", "Learning beyond the table — outings, tie-ins & memories")
    table(d, M, TOP, CW, ["DATE", "PLACE", "THEME TIE-IN", "MEMORY / NOTE", "★"],
          [0.0, 0.15, 0.42, 0.63, 0.92], 16, rowh=130)
    imgs.append(img)


def p09(imgs):
    img, d = page("Arts & Sensory Idea Bank", "Grab-and-go play for busy little hands — jot your family's favorites")
    table(d, M, TOP, CW, ["ACTIVITY", "AREA", "MATERIALS", "BUILDS"],
          [0.0, 0.34, 0.52, 0.78], 20, rowh=130)
    imgs.append(img)


def p10(imgs):
    img, d = page("Kindergarten Readiness", "A gentle guide, never a test — every child grows at their own pace")
    cols = [
        ["Writes first name", "Recognizes 20+ letters", "Says most letter sounds", "Counts to 20",
         "Recognizes numerals 0–10", "Knows 8+ colors", "Names basic shapes", "Holds a pencil correctly"],
        ["Cuts with scissors on a line", "Uses the bathroom independently", "Dresses self",
         "Follows 2-step directions", "Takes turns & shares", "Separates calmly", "Sits for a short lesson",
         "Identifies rhyming words"],
    ]
    for c, items in enumerate(cols):
        x = M + c * (CW / 2 + 20)
        for i, it in enumerate(items):
            checkbox(d, x, TOP + i * 150, it, size=54, font=fs(28, bold=False))
    y = TOP + 8 * 150 + 30
    section(d, M, y, CW, "Skills we're focusing on next"); y += 100
    for i in range(3):
        d.line((M, y + i * 84 + 30, M + CW, y + i * 84 + 30), fill=LINE, width=2)
    imgs.append(img)


def p11(imgs):
    img, d = page("Preschool Goals", "The little wins that matter — for each child & the whole family")
    y = TOP
    for who in ["Child 1", "Child 2", "Together as a family", "A goal for me (the grown-up)"]:
        section(d, M, y, CW, f"{who} — goals this year"); y += 96
        for i in range(3):
            checkbox(d, M, y + i * 88, "", size=44)
            d.line((M + 70, y + i * 88 + 44, M + CW, y + i * 88 + 44), fill=LINE, width=2)
        y += 3 * 88 + 50
    imgs.append(img)


def p12(imgs):
    img, d = page("Portfolio & Keepsakes", "Tape in a favorite each month — the art, the milestones, the sweet wins")
    y = TOP
    field(d, M, y, CW / 2 - 30, "Child"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Month & year")
    y += 130
    caps = ["First name written", "Best painting / drawing", "A proud milestone", "A nature find or photo"]
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
    pdf_path = os.path.join(out_dir, "Preschool_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
