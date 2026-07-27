"""Printable PDF pack for Estate & Emergency Organizer Command Center™ (12 pages, US Letter)."""
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
    d.text((M, 90), "ESTATE & EMERGENCY ORGANIZER COMMAND CENTER™", font=fs(26), fill=GOLD_HI, anchor="lt")
    d.text((M, 150), title, font=fserif(74), fill=WHITE, anchor="lt")
    if subtitle:
        d.text((M, 258), subtitle, font=fs(30, bold=False), fill=(214, 226, 222), anchor="lt")
    d.line((M, H - 130, W - M, H - 130), fill=LINE, width=2)
    d.text((M, H - 108), "Prepared by / Date: ______________", font=fs(24, bold=False), fill=MUTED, anchor="lt")
    d.text((W - M, H - 108), "Estate Organizer™", font=fs(22, bold=False), fill=ACCENT, anchor="rt")
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
    img, d = page("If Something Happens — Start Here", "The first page anyone should read")
    field(d, M, TOP, CW / 2 - 30, "Prepared by"); field(d, M + CW / 2 + 30, TOP, CW / 2 - 30, "Last updated")
    section(d, M, TOP + 160, CW, "The first five calls, in order")
    table(d, M, TOP + 240, CW, ["#", "WHO", "WHY", "PHONE"], [0.0, 0.10, 0.44, 0.76], 5, 150,
          filled_rows=[["1", "", "", ""], ["2", "", "", ""], ["3", "", "", ""], ["4", "", "", ""], ["5", "", "", ""]])
    y = TOP + 240 + 78 + 5 * 150 + 40
    section(d, M, y, CW, "Where the important things are")
    for lab in ["This organizer lives", "Original will & legal documents", "Fire safe / safe deposit box & key",
                "Password manager master key", "Deeds, titles & policies"]:
        field(d, M + 24, y + 100, CW - 48, lab)
        y += 186
    y += 60
    section(d, M, y, CW, "Take a breath first")
    notes = ["Nothing has to be decided today. Almost none of this is urgent in the first week.",
             "Order 10 certified death certificates — nearly every institution wants an original.",
             "Do not close accounts or sell anything until the executor and attorney say so."]
    for j, n in enumerate(notes):
        d.text((M + 24, y + 110 + j * 84), "\u2022  " + n, font=fs(29, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p02(i):
    img, d = page("Estate Snapshot", "What you own, what you owe, what's exposed")
    table(d, M, TOP, CW, ["THE MATH", "AMOUNT"], [0.0, 0.78], 8,
          filled_rows=[["Total assets", ""], ["− Total debts", ""], ["= NET ESTATE", ""],
                       ["Assets with no beneficiary / joint owner", ""], ["= SHARE EXPOSED TO PROBATE", ""],
                       ["× probate cost estimate", ""], ["= ROUGH COST OF PROBATE", ""],
                       ["Cash a survivor could reach in days", ""]], rowh=190)
    y = TOP + 78 + 8 * 190 + 40
    field(d, M, y, CW / 3 - 20, "Monthly household cost"); field(d, M + CW / 3 + 10, y, CW / 3 - 20, "= Months of runway")
    field(d, M + 2 * CW / 3 + 20, y, CW / 3 - 20, "Life insurance in force")
    i.append(img)


def p03(i):
    img, d = page("Assets & Accounts", "Everything you own, and how it transfers")
    table(d, M, TOP, CW, ["ASSET", "VALUE", "HOW IT TRANSFERS", "PROBATE?", "WHERE TO FIND IT"],
          [0.0, 0.26, 0.38, 0.62, 0.72], 20, 118)
    y = TOP + 78 + 20 * 118 + 30
    field(d, M, y, CW / 2 - 30, "Total assets"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Of which exposed to probate")
    i.append(img)


def p04(i):
    img, d = page("Debts, Bills & Subscriptions", "What is owed and who to call")
    table(d, M, TOP, CW, ["DEBT OR BILL", "HELD BY", "BALANCE", "MONTHLY", "WHO TO CALL"],
          [0.0, 0.28, 0.48, 0.62, 0.74], 20, 118)
    y = TOP + 78 + 20 * 118 + 30
    field(d, M, y, CW / 2 - 30, "Total owed"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Total monthly")
    i.append(img)


def p05(i):
    img, d = page("Beneficiary Review", "The forms that override your will")
    table(d, M, TOP, CW, ["ACCOUNT", "PRIMARY", "CONTINGENT", "LAST REVIEWED"],
          [0.0, 0.30, 0.54, 0.78], 16, 148)
    y = TOP + 78 + 16 * 148 + 36
    section(d, M, y, CW, "Remember")
    notes = ["A beneficiary designation beats your will. If they disagree, the form wins.",
             "Anything with no beneficiary and no joint owner goes through probate.",
             "Review these after every marriage, divorce, birth and death in the family."]
    for j, n in enumerate(notes):
        d.text((M + 24, y + 110 + j * 84), "•  " + n, font=fs(29, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def p06(i):
    img, d = page("Legal Document Checklist", "The ten that decide everything")
    docs = ["Will", "Living trust", "Financial power of attorney", "Healthcare power of attorney",
            "Living will / advance directive", "HIPAA release", "Beneficiary designations reviewed",
            "Letter of instruction", "Digital asset authorization", "Funeral & burial wishes"]
    y = TOP
    for j, doc in enumerate(docs):
        cy = y + j * 230
        checkbox(d, M + 20, cy)
        d.text((M + 110, cy + 18), doc, font=fs(32), fill=PRIMARY, anchor="lt")
        d.text((M + 110, cy + 84), "WHERE IT IS KEPT", font=fs(21), fill=ACCENT, anchor="lt")
        d.line((M + 110, cy + 140, M + CW * 0.62, cy + 140), fill=LINE, width=2)
        d.text((M + CW * 0.66, cy + 84), "PREPARED BY / DATE", font=fs(21), fill=ACCENT, anchor="lt")
        d.line((M + CW * 0.66, cy + 140, W - M, cy + 140), fill=LINE, width=2)
    i.append(img)


def p07(i):
    img, d = page("Insurance Policies", "Nobody claims a policy they don't know exists")
    table(d, M, TOP, CW, ["POLICY", "COMPANY", "COVERAGE", "POLICY #", "CONTACT"],
          [0.0, 0.26, 0.48, 0.62, 0.80], 16, 148)
    y = TOP + 78 + 16 * 148 + 36
    field(d, M, y, CW / 2 - 30, "Total coverage listed"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Agent / broker")
    i.append(img)


def p08(i):
    img, d = page("Digital Life", "Where the login lives — never the password")
    d.text((M, TOP - 40), "DO NOT WRITE ACTUAL PASSWORDS ON THIS PAGE. Record where each one lives.",
           font=fs(26), fill=DANGER, anchor="lt")
    table(d, M, TOP + 40, CW, ["SERVICE", "WHAT IT IS", "WHERE THE LOGIN LIVES", "WHAT TO DO"],
          [0.0, 0.24, 0.46, 0.78], 18, 128)
    y = TOP + 40 + 78 + 18 * 128 + 30
    field(d, M, y, CW / 2 - 30, "Password manager & master key"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Phone PIN — where kept")
    i.append(img)


def p09(i):
    img, d = page("Key Contacts", "The calls someone would need to make")
    table(d, M, TOP, CW, ["NAME", "ROLE", "PHONE", "EMAIL"], [0.0, 0.30, 0.52, 0.70], 20, 118)
    y = TOP + 78 + 20 * 118 + 30
    field(d, M, y, CW / 2 - 30, "Executor named in the will"); field(d, M + CW / 2 + 30, y, CW / 2 - 30, "Attorney")
    i.append(img)


def p10(i):
    img, d = page("Medical & Care", "Print this one for the fridge")
    field(d, M, TOP, CW / 3 - 20, "Name"); field(d, M + CW / 3 + 10, TOP, CW / 3 - 20, "Date of birth")
    field(d, M + 2 * CW / 3 + 20, TOP, CW / 3 - 20, "Blood type")
    table(d, M, TOP + 160, CW, ["ITEM", "DETAIL"], [0.0, 0.34], 14, 158,
          filled_rows=[["Allergies", ""], ["Ongoing conditions", ""], ["Current medications", ""],
                       ["Pharmacy", ""], ["Primary physician", ""], ["Specialists", ""],
                       ["Health insurance & member #", ""], ["Resuscitation wishes", ""],
                       ["Organ donation", ""], ["Preferred hospital", ""],
                       ["Who decides if I can't", ""], ["Mobility & daily care needs", ""],
                       ["Vision, hearing & dental", ""], ["Anything else they should know", ""]])
    i.append(img)


def p11(i):
    img, d = page("Final Wishes", "In your words, so nobody has to guess")
    table(d, M, TOP, CW, ["ITEM", "WHAT I WANT"], [0.0, 0.34], 12, 190,
          filled_rows=[["Burial or cremation", ""], ["Service", ""], ["Where", ""],
                       ["Readings or music", ""], ["Who should speak", ""],
                       ["Flowers or donations", ""], ["Obituary", ""],
                       ["Who to notify first", ""], ["Pre-paid arrangements", ""],
                       ["Estimated cost & how to pay it", ""], ["What I'd like said", ""],
                       ["Anything else", ""]])
    i.append(img)


def p12(i):
    img, d = page("Household Instructions", "The small practical things")
    table(d, M, TOP, CW, ["ITEM", "INSTRUCTION"], [0.0, 0.30], 11, 190,
          filled_rows=[["Water shut-off", ""], ["Electrical panel", ""], ["Furnace & filter size", ""],
                       ["Spare keys", ""], ["Fire safe combination is with", ""],
                       ["Pets — who takes them & the vet", ""], ["Lawn & snow", ""],
                       ["Utilities & autopay", ""], ["Mail — stop or forward", ""],
                       ["Safe deposit box & key", ""], ["Anything that would confuse someone", ""]])
    y = TOP + 78 + 11 * 190 + 36
    section(d, M, y, CW, "One last thing")
    d.text((M + 24, y + 110), "•  Tell one person this file exists and where it is. That is the whole point.",
           font=fs(29, bold=False), fill=TEXT, anchor="lt")
    d.text((M + 24, y + 194), "•  Review it every year, and after any birth, death, marriage or move.",
           font=fs(29, bold=False), fill=TEXT, anchor="lt")
    i.append(img)


def main():
    out_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print_dir = os.path.join(out_dir, "marketing", "print"); os.makedirs(print_dir, exist_ok=True)
    imgs = []
    for fn in [p01, p02, p03, p04, p05, p06, p07, p08, p09, p10, p11, p12]:
        fn(imgs)
    for i, im in enumerate(imgs, 1):
        im.save(os.path.join(print_dir, f"page_{i:02d}.png"), "PNG", optimize=True)
    pdf_path = os.path.join(out_dir, "Estate_Organizer_Printables.pdf")
    imgs[0].save(pdf_path, "PDF", resolution=300.0, save_all=True, append_images=imgs[1:])
    print(f"Wrote {pdf_path}  ({len(imgs)} pages) + page PNGs in {print_dir}")


if __name__ == "__main__":
    main()
