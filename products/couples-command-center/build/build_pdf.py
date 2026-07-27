"""Printable PDF pack for Relationship & Couples Command Center™ (12 pages, US Letter)."""
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
    d.text((M, 90), "RELATIONSHIP & COUPLES COMMAND CENTER™", font=fs(26), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Month: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Couples Command Center™", font=fs(22, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("Fair Share Worksheet", "50/50 is not the same as fair")
    field(d, M, TOP, CW / 2 - 30, "Partner A"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Partner B")
    table(d, M, TOP + 150, CW, ["THE MATH", "A", "B"], [0.0, 0.56, 0.78], 9, 178,
          filled_rows=[["Monthly income after tax", "", ""], ["Share of combined income", "", ""],
                       ["", "", ""],
                       ["50/50: each pays half the bills", "", ""], ["= what's left", "", ""],
                       ["", "", ""],
                       ["Proportional: your share of the bills", "", ""], ["= what's left", "", ""],
                       ["= as a % of your OWN income", "", ""]])
    y = TOP + 150 + 78 + 9 * 178 + 36
    field(d, M, y, CW / 2 - 30, "Shared bills, total"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "The 50/50 gap between you")
    i.append(img)


def p02(i):
    img, d = page("Shared Bills", "Everything you pay for together")
    table(d, M, TOP, CW, ["BILL", "CATEGORY", "MONTHLY", "WHO PAYS IT"], [0.0, 0.38, 0.56, 0.72], 20, 118)
    y = TOP + 78 + 20 * 118 + 30
    field(d, M, y, CW / 2 - 30, "Total each month"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Total each year")
    i.append(img)


def p03(i):
    img, d = page("Invisible Labour", "Fill it in separately. Then compare.")
    d.text((M, TOP - 30), "Include the remembering, not just the doing. Knowing the dog is due at the vet IS the work.",
           font=fs(26, bold=False), fill=MUTED, anchor="lt")
    table(d, M, TOP + 50, CW, ["TASK \\u2014 HOURS A WEEK", "A", "B", "GAP"], [0.0, 0.56, 0.70, 0.85], 13, 158)
    y = TOP + 50 + 78 + 13 * 158 + 36
    field(d, M, y, CW / 3 - 20, "A's hours a week"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "B's hours a week")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "= Extra hours a YEAR")
    i.append(img)


def p04(i):
    img, d = page("The Conversation", "After you've both filled in the last page")
    section(d, M, TOP, CW, "Read these before you start")
    notes = ["Nobody is lying when the two columns don't match.",
             "The work you don't do is genuinely hard to see. That's why it's called invisible.",
             "This is a conversation, not a verdict, and not a scoreboard to win.",
             "Pick ONE row to move this month. Not all of them. One."]
    for j, n in enumerate(notes):
        d.text((M + 24, TOP + 110 + j * 92), "\\u2022  " + n, font=fs(28, bold=False), fill=TEXT, anchor="lt")
    y = TOP + 110 + 4 * 92 + 60
    qs = ["Which row surprised you most?", "Which row did you not know the other one was doing?",
          "What is the one thing that would make the biggest difference?",
          "What will actually change, and starting when?",
          "When will we look at this page again?"]
    for q in qs:
        d.text((M, y), q.upper(), font=fs(25), fill=ACCENT, anchor="lt")
        for k in range(2):
            d.line((M, y + 80 + k * 78, W - M, y + 80 + k * 78), fill=LINE, width=2)
        y += 250
    i.append(img)


def p05(i):
    img, d = page("Money Goals", "What you're building together")
    table(d, M, TOP, CW, ["GOAL", "TARGET", "SAVED", "BY WHEN", "STATUS"],
          [0.0, 0.34, 0.50, 0.64, 0.80], 16, 148)
    y = TOP + 78 + 16 * 148 + 36
    field(d, M, y, CW / 3 - 20, "Total targets"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "Total saved")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Saved each month")
    i.append(img)


def p06(i):
    img, d = page("Weekly Check-In", "Fifteen minutes. Same time every week.")
    section(d, M, TOP, CW, "The four questions")
    qs = ["Anything on your mind about money this week?",
          "What's in the diary that I should know about?",
          "Is anything feeling uneven right now?",
          "What do you need from me this week?"]
    y = TOP + 110
    for j, q in enumerate(qs):
        d.text((M + 24, y), f"{j+1}.  {q}", font=fs(29), fill=PRIMARY, anchor="lt")
        for k in range(3):
            d.line((M + 60, y + 90 + k * 78, W - M, y + 90 + k * 78), fill=LINE, width=2)
        y += 400
    y2 = y + 20
    field(d, M, y2, CW / 2 - 30, "Anything to carry into next week"); field(d, M + CW / 2 + 30, y2, CW / 2 - 30, "Date")
    i.append(img)


def p07(i):
    img, d = page("The Big Conversations", "The ones couples avoid until they can't")
    talks = ["Do we both want children, and when", "Where we want to live in five years",
             "How we handle a big unexpected bill", "What we each need when we're stressed",
             "Money we each spend without asking", "How we split things if incomes change",
             "What we want the next ten years to feel like", "Wills, beneficiaries & who decides if",
             "What retirement looks like for both of us", "How we'd handle caring for a parent"]
    for j, t in enumerate(talks):
        cy = TOP + j * 240
        checkbox(d, M + 20, cy)
        d.text((M + 110, cy + 18), t, font=fs(29), fill=PRIMARY, anchor="lt")
        d.text((M + 110, cy + 82), "WHERE WE LANDED", font=fs(20), fill=ACCENT, anchor="lt")
        d.line((M + 110, cy + 148, W - M, cy + 148), fill=LINE, width=2)
    i.append(img)


def p08(i):
    img, d = page("Date Nights", "It stops happening by accident")
    table(d, M, TOP, CW, ["DATE", "WHAT YOU DID", "WHO PLANNED IT", "COST", "HOW IT WAS"],
          [0.0, 0.10, 0.44, 0.62, 0.74], 18, 128)
    y = TOP + 78 + 18 * 128 + 36
    section(d, M, y, CW, "One thing to notice")
    d.text((M + 24, y + 110), "\\u2022  Look down the 'who planned it' column. If it's always the same name, that's invisible labour too.",
           font=fs(27, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p09(i):
    img, d = page("Household Admin", "Who actually does what")
    table(d, M, TOP, CW, ["ITEM", "WHOSE JOB", "HOW OFTEN", "NOTES"], [0.0, 0.34, 0.50, 0.64], 20, 118)
    y = TOP + 78 + 20 * 118 + 30
    d.text((M + 24, y + 20), "Anything with nobody's name on it is being done by whoever notices first. That is not nothing.",
           font=fs(27, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p10(i):
    img, d = page("Individual Money", "What's yours stays yours")
    table(d, M, TOP, CW, ["", "A", "B"], [0.0, 0.56, 0.78], 8, 168,
          filled_rows=[["Income", "", ""], ["Your share of the bills", "", ""],
                       ["Into shared savings", "", ""], ["Personal savings", "", ""],
                       ["Debt payments", "", ""], ["Guilt-free spending", "", ""],
                       ["", "", ""], ["= What's genuinely yours", "", ""]])
    y = TOP + 78 + 8 * 168 + 40
    section(d, M, y, CW, "The rule that stops most money arguments")
    field(d, M + 24, y + 110, CW - 48, "Under this amount, either of us can spend without asking")
    d.text((M + 24, y + 260), "It genuinely does not matter what the number is. It matters that you both know it.",
           font=fs(27, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p11(i):
    img, d = page("Month in Review", "Money, hours, and each other")
    table(d, M, TOP, CW, ["MONTH", "SAVED", "CHORE SPLIT", "DATE NIGHTS", "CHECK-INS"],
          [0.0, 0.26, 0.44, 0.64, 0.82], 8, 168)
    y = TOP + 78 + 8 * 168 + 40
    section(d, M, y, CW, "Six honest checks")
    checks = ["The bills are split fairly, not just evenly", "We had our weekly check-ins",
              "We're saving what we said we would", "Our goals are on track",
              "We had our date nights", "The housework is actually shared"]
    for j, c in enumerate(checks):
        cy = y + 130 + j * 116
        checkbox(d, M + 20, cy)
        d.text((M + 110, cy + 24), c, font=fs(29, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p12(i):
    img, d = page("If Incomes Change", "Come back to this page. It takes five minutes.")
    section(d, M, TOP, CW, "Recalculate together")
    table(d, M, TOP + 90, CW, ["THE MATH", "A", "B"], [0.0, 0.56, 0.78], 6, 190,
          filled_rows=[["New monthly income", "", ""], ["New share of combined income", "", ""],
                       ["Shared bills total", "", ""], ["= New fair share each", "", ""],
                       ["= What's left", "", ""], ["= As a % of your own income", "", ""]])
    y = TOP + 90 + 78 + 6 * 190 + 40
    section(d, M, y, CW, "And ask")
    qs = ["Does the chore split need to change too?",
          "Does anyone's guilt-free spending number change?",
          "Do any of the savings amounts change?"]
    for j, q in enumerate(qs):
        d.text((M + 24, y + 110 + j * 130), "\\u2022  " + q, font=fs(28, bold=False), fill=TEXT, anchor="lt")
        d.line((M + 60, y + 176 + j * 130, W - M, y + 176 + j * 130), fill=LINE, width=2)
    i.append(img)


def main():
    out_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print_dir = os.path.join(out_dir, "marketing", "print"); os.makedirs(print_dir, exist_ok=True)
    imgs = []
    for fn in [p01, p02, p03, p04, p05, p06, p07, p08, p09, p10, p11, p12]:
        fn(imgs)
    for i, im in enumerate(imgs, 1):
        im.save(os.path.join(print_dir, f"page_{i:02d}.png"), "PNG", optimize=True)
    pdf_path = os.path.join(out_dir, "Couples_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
