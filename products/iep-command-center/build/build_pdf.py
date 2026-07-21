"""Build the printable PDF pack for IEP Command Center™ (12 pages, US Letter).

The advocacy binder:
  1  Student Profile & Team       7  Behavior Tracker (ABC)
  2  IEP Goals & Progress         8  Meeting Notes / Prep
  3  Progress Monitoring          9  Communication Log
  4  Services & Minutes          10  Strengths & Interests
  5  Accommodations Checklist    11  Records & Documents
  6  Therapy / Session Log       12  Wins & Milestones

Outputs ../IEP_Printables.pdf and page PNGs in ../marketing/print/.
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
    d.text((M, 90), "IEP COMMAND CENTER™", font=fs(34), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(76), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Confidential — keep private", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "IEP Command Center™", font=fs(24, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Student Profile & Team", "The one-page snapshot you bring to every meeting")
    lw = CW / 2 - 30
    left = ["Child's name", "Grade & school", "Date of birth", "Plan type (IEP / 504)",
            "Eligibility category", "Case manager"]
    right = ["Next annual review", "Diagnoses (if shared)", "Parent contact", "Provider contacts",
             "Strengths", "What helps"]
    for i in range(6):
        field(d, M, TOP + i * 118, lw, left[i])
        field(d, M + CW / 2 + 30, TOP + i * 118, lw, right[i])
    y = TOP + 6 * 118 + 30
    section(d, M, y, CW, "The whole child — strengths, interests & what makes them light up"); y += 96
    for i in range(4):
        d.line((M, y + i * 84 + 30, M + CW, y + i * 84 + 30), fill=LINE, width=2)
    imgs.append(img)


def p02(imgs):
    img, d = page("IEP Goals & Progress", "Baseline, target & current — progress toward each goal")
    table(d, M, TOP, CW, ["AREA", "GOAL", "BASE", "TARGET", "NOW", "%"],
          [0.0, 0.16, 0.60, 0.70, 0.81, 0.91], 10, rowh=180,
          filled_rows=[["Reading"], ["Math"], ["Writing"], ["Speech"], ["Social/Emo"]])
    imgs.append(img)


def p03(imgs):
    img, d = page("Progress Monitoring", "Every data point, dated — the evidence behind each goal")
    d.text((M, TOP), "GOAL AREA:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 260, TOP + 34, M + 1100, TOP + 34), fill=LINE, width=2)
    d.text((M + 1200, TOP), "MEASURE:", font=fs(26), fill=ACCENT, anchor="lt")
    d.line((M + 1420, TOP + 34, W - M, TOP + 34), fill=LINE, width=2)
    table(d, M, TOP + 90, CW, ["#", "DATE", "MEASURE", "VALUE", "NOTE"],
          [0.0, 0.08, 0.28, 0.48, 0.64], 18, rowh=120,
          filled_rows=[[str(i + 1)] for i in range(18)])
    imgs.append(img)


def p04(imgs):
    img, d = page("Services & Minutes", "The supports your child is owed — scheduled vs delivered")
    table(d, M, TOP, CW, ["SERVICE", "PROVIDER", "FREQUENCY", "MIN/WK", "DELIVERED"],
          [0.0, 0.26, 0.48, 0.68, 0.84], 12, rowh=170,
          filled_rows=[["Special Ed"], ["Speech"], ["Occupational"], ["Physical"], ["Counseling"]])
    imgs.append(img)


def p05(imgs):
    img, d = page("Accommodations Checklist", "Check that each accommodation is actually in place")
    items = ["Extended time (1.5×) on tests", "Preferential seating", "Assignments chunked into steps",
             "Frequent movement / sensory breaks", "Text-to-speech for reading", "Reduced written output / scribe",
             "Visual schedule & checklists", "Quiet, low-distraction test space", "Directions repeated & checked",
             "Modified spelling / reduced list", "Noise-reducing headphones", "Calm-down space available",
             "Copy of class notes provided", "Check-in / check-out routine"]
    for i, it in enumerate(items):
        col = i % 2
        x = M + col * (CW / 2 + 30)
        checkbox(d, x, TOP + (i // 2) * 170, it, size=54, font=fs(28, bold=False))
    imgs.append(img)


def p06(imgs):
    img, d = page("Therapy / Session Log", "Every OT, PT, speech & counseling session — & home carryover")
    table(d, M, TOP, CW, ["DATE", "TYPE", "PROVIDER", "FOCUS / NOTE"],
          [0.0, 0.14, 0.34, 0.54], 16, rowh=130)
    imgs.append(img)


def p07(imgs):
    img, d = page("Behavior Tracker (ABC)", "Situation, behavior & the support that helped — patterns emerge")
    table(d, M, TOP, CW, ["DATE", "SITUATION (BEFORE)", "BEHAVIOR", "SUPPORT THAT HELPED"],
          [0.0, 0.14, 0.42, 0.68], 15, rowh=140)
    imgs.append(img)


def p08(imgs):
    img, d = page("Meeting Notes / Prep", "Walk in prepared — questions to ask & what was decided")
    y = TOP
    field(d, M, y, CW / 2 - 30, "Meeting date & type"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Who's attending")
    y += 130
    section(d, M, y, CW, "My top questions & requests"); y += 96
    for i in range(5):
        checkbox(d, M, y + i * 84, "", size=40)
        d.line((M + 66, y + i * 84 + 40, M + CW, y + i * 84 + 40), fill=LINE, width=2)
    y += 5 * 84 + 40
    section(d, M, y, CW, "Decisions & agreements (get it in writing)"); y += 96
    for i in range(5):
        d.line((M, y + i * 80 + 30, M + CW, y + i * 80 + 30), fill=LINE, width=2)
    imgs.append(img)


def p09(imgs):
    img, d = page("Communication Log", "Every email, call & note — a paper trail protects your child")
    table(d, M, TOP, CW, ["DATE", "WITH", "TOPIC", "FOLLOW-UP"],
          [0.0, 0.14, 0.40, 0.76], 18, rowh=120)
    imgs.append(img)


def p10(imgs):
    img, d = page("Strengths & Interests", "Start every meeting here — so much more than a list of needs")
    cols = [("STRENGTHS", 6), ("INTERESTS & MOTIVATORS", 6)]
    for c, (head, n) in enumerate(cols):
        x = M + c * (CW / 2 + 30)
        section(d, x, TOP, CW / 2 - 30, head)
        for i in range(n):
            d.text((x, TOP + 96 + i * 110), "•", font=fs(34), fill=GOLD_LT, anchor="lt")
            d.line((x + 44, TOP + 96 + i * 110 + 44, x + CW / 2 - 30, TOP + 96 + i * 110 + 44), fill=LINE, width=2)
    y = TOP + 96 + 6 * 110 + 40
    section(d, M, y, CW, "What I want the team to know about my child"); y += 96
    for i in range(3):
        d.line((M, y + i * 84 + 30, M + CW, y + i * 84 + 30), fill=LINE, width=2)
    imgs.append(img)


def p11(imgs):
    img, d = page("Records & Documents", "The binder checklist — everything meeting-ready")
    items = ["Current IEP / 504 on file", "Signed consent forms", "Most recent evaluation report",
             "Progress reports (each period)", "Meeting notices & minutes", "Work samples for the binder",
             "Prior written notices (PWN)", "Data charts printed", "Report cards", "Communication log copies",
             "Medical / therapy reports", "Contact list for the team"]
    for i, it in enumerate(items):
        col = i % 2
        x = M + col * (CW / 2 + 30)
        checkbox(d, x, TOP + (i // 2) * 190, it, size=54, font=fs(28, bold=False))
    y = TOP + 6 * 190 + 20
    section(d, M, y, CW, "Where things live"); y += 96
    for i in range(2):
        d.line((M, y + i * 80 + 30, M + CW, y + i * 80 + 30), fill=LINE, width=2)
    imgs.append(img)


def p12(imgs):
    img, d = page("Wins & Milestones", "The progress that gets lost between meetings — celebrate it here")
    y = TOP
    for i in range(9):
        d.rounded_rectangle((M, y, M + 110, y + 110), radius=14, outline=GOLD_LT, width=4)
        d.text((M + 55, y + 55), "★", font=fs(50), fill=GOLD_LT, anchor="mm")
        d.line((M + 160, y + 70, W - M - 380, y + 70), fill=LINE, width=2)
        d.text((W - M, y + 70), "date: __________", font=fs(24, bold=False), fill=MUTED, anchor="rt")
        y += 180
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
    pdf_path = os.path.join(out_dir, "IEP_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
