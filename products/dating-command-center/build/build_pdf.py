"""Printable PDF pack for Dating Life Command Center™ (12 pages, US Letter)."""
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
    d.text((M, 90), "DATING LIFE COMMAND CENTER™", font=fs(26), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Month: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Dating Command Center™", font=fs(22, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Non-Negotiables", "Write these before you meet anyone")
    d.text((M, TOP - 30), "Read this page when you are tempted to bend one. That is what it is for.",
           font=fs(27, bold=False), fill=MUTED, anchor="lt")
    table(d, M, TOP + 60, CW, ["WHAT", "WEIGHT", "WHY IT MATTERS TO ME"], [0.0, 0.34, 0.52], 14, 168)
    y = TOP + 60 + 78 + 14 * 168 + 40
    section(d, M, y, CW, "The rule")
    d.text((M + 24, y + 110), "•  If it is a dealbreaker, it is a dealbreaker on date one \\u2014 not date twelve.",
           font=fs(29, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p02(i):
    img, d = page("The Month, Counted", "Matches to conversations to dates")
    table(d, M, TOP, CW, ["THE FUNNEL", "COUNT"], [0.0, 0.78], 8,
          filled_rows=[["Matches", ""], ["Conversations that actually started", ""],
                       ["First dates", ""], ["Second dates", ""], ["Still seeing", ""],
                       ["= Matches that became a date  (%)", ""],
                       ["= First dates that became a second  (%)", ""],
                       ["Days from match to first date", ""]], rowh=190)
    y = TOP + 78 + 8 * 190 + 40
    field(d, M, y, CW / 3 - 20, "Hours this month"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Money this month")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Per second date")
    i.append(img)


def p03(i):
    img, d = page("Effort & Reciprocity", "Score yourself, then score them")
    d.text((M, TOP - 30), "Out of ten, on what actually happened this month. Be fair to both of you.",
           font=fs(27, bold=False), fill=MUTED, anchor="lt")
    field(d, M, TOP + 50, CW / 2 - 30, "Who this is about"); field(d, M + CW / 2 + 30, TOP + 50, CW / 2 - 30, "Month")
    table(d, M, TOP + 200, CW, ["WHAT", "YOU", "THEM", "DIFFERENCE"], [0.0, 0.52, 0.68, 0.84], 8, 168,
          filled_rows=[["Initiating contact", "", "", ""], ["Planning the dates", "", "", ""],
                       ["Paying / splitting fairly", "", "", ""], ["Being flexible on timing", "", "", ""],
                       ["Following through on plans", "", "", ""], ["Asking about your life", "", "", ""],
                       ["Making you feel wanted", "", "", ""], ["TOTAL", "", "", ""]])
    y = TOP + 200 + 78 + 8 * 168 + 40
    field(d, M, y, CW / 2 - 30, "Your total ÷ their total ="); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "1.0 is even")
    y2 = y + 170
    d.text((M + 24, y2), "A low score doesn't make someone a bad person. It means you're doing more of the work.",
           font=fs(28, bold=False), fill=TEXT, anchor="lt")
    d.text((M + 24, y2 + 70), "What you do with that is entirely yours to decide.",
           font=fs(28, bold=False), fill=ACCENT, anchor="lt")
    y3 = y2 + 180
    section(d, M, y3, CW, "Before you score them low, ask")
    qs = ["Have I actually told them what I need, out loud, once?",
          "Am I scoring what they do, or what I'm afraid they think?",
          "If a friend showed me this page, what would I say to her?"]
    for j, q in enumerate(qs):
        d.text((M + 24, y3 + 110 + j * 130), "\u2022  " + q, font=fs(28, bold=False), fill=TEXT, anchor="lt")
        d.line((M + 60, y3 + 176 + j * 130, W - M, y3 + 176 + j * 130), fill=LINE, width=2)
    i.append(img)


def p04(i):
    img, d = page("Green Flags", "Tick what you have actually SEEN")
    flags = ["Texts back in a normal amount of time", "Asks questions about your life",
             "Plans an actual date, with a time and a place", "Is kind to waiting staff",
             "Talks about their friends warmly", "Handles a small disagreement calmly",
             "Respects a no the first time", "Is where they said they'd be",
             "Introduces you to people in their life", "Says what they want out loud",
             "Remembers what you told them", "You feel calm afterwards, not anxious"]
    for j, f in enumerate(flags):
        cy = TOP + j * 200
        checkbox(d, M + 20, cy)
        d.text((M + 110, cy + 20), f, font=fs(30, bold=False), fill=TEXT, anchor="lt")
        d.line((M + 110, cy + 110, W - M, cy + 110), fill=LINE, width=2)
    i.append(img)


def p05(i):
    img, d = page("Red Flags", "The ones we talk ourselves out of")
    flags = ["Only texts late at night", "Vague about plans until the last minute",
             "Speaks badly about every ex", "Pushes past a no",
             "Love-bombs, then goes quiet", "Won't be seen with you in public",
             "Makes you feel like you're too much", "Rude to staff or strangers",
             "Story details keep changing", "Won't discuss what this is",
             "Only reaches out when they need something", "You feel anxious after seeing them"]
    for j, f in enumerate(flags):
        cy = TOP + j * 200
        checkbox(d, M + 20, cy)
        d.text((M + 110, cy + 20), f, font=fs(30, bold=False), fill=TEXT, anchor="lt")
        d.line((M + 110, cy + 110, W - M, cy + 110), fill=LINE, width=2)
    i.append(img)


def p06(i):
    img, d = page("Date Log", "What it cost, how long, how it felt")
    table(d, M, TOP, CW, ["DATE", "WHO", "WHAT", "COST", "HOURS", "HOW IT FELT"],
          [0.0, 0.10, 0.24, 0.58, 0.68, 0.79], 20, 118)
    y = TOP + 78 + 20 * 118 + 30
    field(d, M, y, CW / 3 - 20, "Dates this month"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Total spent")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "How many were 'great'?")
    i.append(img)


def p07(i):
    img, d = page("People", "Where it actually stands")
    table(d, M, TOP, CW, ["NAME", "MET ON", "FIRST DATE", "STATUS", "HOW IT FELT"],
          [0.0, 0.22, 0.40, 0.56, 0.74], 18, 128)
    y = TOP + 78 + 18 * 128 + 30
    field(d, M, y, CW / 3 - 20, "In the picture"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Actually seeing")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Slow fading")
    i.append(img)


def p08(i):
    img, d = page("Conversations", "Who's dating you and who's just texting you")
    table(d, M, TOP, CW, ["WHO", "DAYS TALKING", "YOU SENT", "THEY SENT", "MET YET?"],
          [0.0, 0.24, 0.44, 0.62, 0.80], 18, 128)
    y = TOP + 78 + 18 * 128 + 36
    section(d, M, y, CW, "The rule")
    d.text((M + 24, y + 110), "•  If they've been talking to you for three weeks and haven't asked you out, that IS the answer.",
           font=fs(28, bold=False), fill=TEXT, anchor="lt")
    d.text((M + 24, y + 190), "•  You are allowed to ask them out. You are also allowed to stop replying.",
           font=fs(28, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p09(i):
    img, d = page("Time & Money", "What it costs in both currencies")
    section(d, M, TOP, CW, "Money")
    table(d, M, TOP + 80, CW, ["WHERE IT GOES", "MONTHLY", "YEARLY"], [0.0, 0.58, 0.80], 8, 148)
    y = TOP + 80 + 78 + 8 * 148 + 36
    section(d, M, y, CW, "Time")
    table(d, M, y + 80, CW, ["WHERE IT GOES", "HOURS"], [0.0, 0.78], 5, 148,
          filled_rows=[["Swiping", ""], ["Messaging", ""], ["On dates", ""],
                       ["Getting ready", ""], ["= TOTAL HOURS", ""]])
    y2 = y + 80 + 78 + 5 * 148 + 30
    field(d, M, y2, CW / 2 - 30, "Per first date"); field(d, M + CW / 2 + 30, y2, CW / 2 - 30, "Per SECOND date")
    i.append(img)


def p10(i):
    img, d = page("Safety Plan", "Fill this in before every first date")
    steps = ["First date is somewhere public, always", "A friend has the name, photo and location",
             "Share live location for the first few dates", "Arrange your own transport there and back",
             "Video-call before meeting in person", "Reverse-image-search the photos",
             "Don't share your home address early", "Have a check-in text time agreed",
             "Trust the feeling and leave if you want to"]
    for j, s in enumerate(steps):
        cy = TOP + j * 150
        checkbox(d, M + 20, cy)
        d.text((M + 110, cy + 20), s, font=fs(30, bold=False), fill=TEXT, anchor="lt")
    y = TOP + 9 * 150 + 40
    section(d, M, y, CW, "Who knows where you are tonight")
    for j, lab in enumerate(["Friend's name & number", "Check-in time", "Where you're going",
                             "Their name & profile link", "How you're getting home"]):
        field(d, M + 24, y + 120 + j * 150, CW - 48, lab)
    y2 = y + 120 + 5 * 150 + 40
    d.text((M + 24, y2), "You never owe anyone a second more of your evening. Leaving early is always allowed.",
           font=fs(28), fill=PRIMARY, anchor="lt")
    i.append(img)


def p11(i):
    img, d = page("Reflection", "The page that actually changes things")
    qs = ["Who did I feel most like myself with?", "Who did I feel anxious around, and why?",
          "What pattern keeps repeating?", "What did I do differently this month?",
          "What am I tolerating that I shouldn't?", "What went better than last month?",
          "What do I want more of?", "What would I tell a friend in my position?"]
    y = TOP
    for q in qs:
        d.text((M, y), q.upper(), font=fs(25), fill=ACCENT, anchor="lt")
        for k in range(2):
            d.line((M, y + 80 + k * 78, W - M, y + 80 + k * 78), fill=LINE, width=2)
        y += 250
    i.append(img)


def p12(i):
    img, d = page("Month in Review", "And what you want the next one to look like")
    table(d, M, TOP, CW, ["MONTH", "MATCHES", "DATES", "2ND DATES", "SPENT", "HOURS"],
          [0.0, 0.24, 0.42, 0.56, 0.72, 0.86], 8, 168)
    y = TOP + 78 + 8 * 168 + 40
    section(d, M, y, CW, "Six honest checks")
    checks = ["I'm meeting people, not just texting them", "Second dates are actually happening",
              "Green flags outweigh red", "I'm spending inside my budget",
              "I did every safety step, every time", "The effort is going both ways"]
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
    pdf_path = os.path.join(out_dir, "Dating_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
