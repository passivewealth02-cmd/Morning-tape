"""Marketing image set for Budget & Money Command Center™ (6 images, 2000x2000)."""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFilter, ImageFont

PRIMARY = (27, 79, 72); PRIMARY_DK = (18, 56, 51); PRIMARY_LT = (33, 92, 83)
ACCENT = (147, 115, 86); GOLD = (180, 145, 90); GOLD_LT = (201, 168, 106); GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186); HIGHLIGHT = (117, 230, 193); BG = (251, 248, 242)
BG_TOP = (253, 250, 246); BG_BOT = (242, 235, 223); WHITE = (255, 255, 255)
TEXT = (51, 51, 51); TEXT_MUTED = (132, 126, 116); DANGER = (201, 76, 76)
MINT_BG = (227, 248, 239); WARN_BG = (251, 240, 226); RED_BG = (251, 230, 230)
GRID = (228, 222, 210); ROW_ALT = (246, 241, 232); DOT = (228, 220, 206); GREEN = (32, 120, 96)
SIZE = 2000

SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"

TABS = ["Dashboard", "Income", "Monthly Budget", "Bills", "Expense Log", "Savings Goals",
        "Sinking Funds", "Debt Snapshot", "Net Worth", "Subscriptions", "Year View",
        "No-Spend", "Settings"]

FILE_LABEL = "Budget_Command_Center.xlsx — The Bennett Household · This Month"


def fs(s, bold=True):
    return ImageFont.truetype(SANS_B if bold else SANS_R, s)


def fserif(s):
    return ImageFont.truetype(SERIF_B, s)


def vgradient(w, h, top, bottom):
    col = Image.new("RGB", (1, h)); px = col.load()
    for y in range(h):
        t = y / max(h - 1, 1)
        px[0, y] = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3))
    return col.resize((w, h)).convert("RGBA")


def grad_round(c, box, radius, top, bottom, outline=None, width=0):
    x0, y0, x1, y1 = [int(v) for v in box]
    w, h = x1 - x0, y1 - y0
    if w <= 0 or h <= 0:
        return
    g = vgradient(w, h, top, bottom)
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, w - 1, h - 1), radius=radius, fill=255)
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0)); layer.paste(g, (x0, y0), mask)
    c.alpha_composite(layer)
    if outline and width:
        ImageDraw.Draw(c).rounded_rectangle((x0, y0, x1, y1), radius=radius, outline=outline, width=width)


def radial_glow(c, cx, cy, r, color, strength=120):
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).ellipse((cx - r, cy - r, cx + r, cy + r), fill=color + (strength,))
    c.alpha_composite(layer.filter(ImageFilter.GaussianBlur(r // 2)))


def premium_bg(c, band_h=0):
    c.alpha_composite(vgradient(c.width, c.height, BG_TOP, BG_BOT))
    dots = Image.new("RGBA", c.size, (0, 0, 0, 0)); dd = ImageDraw.Draw(dots)
    sp, r = 50, 3
    for y in range(sp // 2, c.height, sp):
        for x in range(sp // 2, c.width, sp):
            dd.ellipse((x - r, y - r, x + r, y + r), fill=DOT + (140,))
    c.alpha_composite(dots)
    if band_h:
        hero_band(c, band_h)


def hero_band(c, band_h):
    c.alpha_composite(vgradient(c.width, band_h, PRIMARY_LT, PRIMARY_DK), (0, 0))
    radial_glow(c, c.width // 2, band_h // 2 - 30, 520, (60, 130, 118), 70)
    wm = Image.new("RGBA", c.size, (0, 0, 0, 0)); wd = ImageDraw.Draw(wm)
    for rr in (300, 230, 160):
        wd.ellipse((c.width - 120 - rr, band_h - 60 - rr, c.width - 120 + rr, band_h - 60 + rr), outline=(255, 255, 255, 22), width=3)
    c.alpha_composite(wm)
    d = ImageDraw.Draw(c)
    d.rectangle((0, band_h - 5, c.width, band_h), fill=GOLD_LT)
    d.rectangle((0, band_h - 5, c.width, band_h - 2), fill=GOLD_HI)


def shadow(c, box, radius, blur=24, alpha=70, dy=18):
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).rounded_rectangle(box, radius=radius, fill=(18, 50, 45, alpha))
    c.alpha_composite(layer.filter(ImageFilter.GaussianBlur(blur)), (0, dy))


def tc(d, xy, t, f, fill, anchor="mm"):
    d.text(xy, t, font=f, fill=fill, anchor=anchor)


def wordmark(c, cx, cy, text, size, max_w=None):
    d = ImageDraw.Draw(c)
    if max_w:
        while size > 20 and d.textlength(text, font=fserif(size)) > max_w:
            size -= 2
    f = fserif(size)
    sh = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).text((cx + 5, cy + 7), text, font=f, fill=(8, 30, 27, 160), anchor="mm")
    c.alpha_composite(sh.filter(ImageFilter.GaussianBlur(7)))
    bb = d.textbbox((cx, cy), text, font=f, anchor="mm"); pad = 30
    bx0, by0, bx1, by1 = bb[0] - pad, bb[1] - pad, bb[2] + pad, bb[3] + pad
    w, h = bx1 - bx0, by1 - by0
    grad = vgradient(w, h, GOLD_HI, GOLD)
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).text((cx - bx0, cy - by0), text, font=f, fill=255, anchor="mm")
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0)); layer.paste(grad, (bx0, by0), mask)
    c.alpha_composite(layer)


def pill(c, cx, cy, text, font, pad_x=60, pad_y=26, star=False, fg=WHITE, grad=(GOLD_LT, GOLD), outline=GOLD_HI):
    d = ImageDraw.Draw(c)
    label = f"★  {text}" if star else text
    tw = d.textlength(label, font=font); th = font.size
    w, h = tw + pad_x * 2, th + pad_y * 2
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, h // 2, 22, 70, 12)
    grad_round(c, box, h // 2, grad[0], grad[1])
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle((box[0] + 5, box[1] + 5, box[2] - 5, box[3] - 5), radius=(h - 10) // 2, outline=outline, width=2)
    od.text((cx, cy), label, font=font, fill=fg, anchor="mm")
    c.alpha_composite(ov)


def gold_divider(c, cx, cy, width=560, color=GOLD_HI):
    d = ImageDraw.Draw(c)
    d.line((cx - width // 2, cy, cx - 30, cy), fill=color, width=3)
    d.line((cx + 30, cy, cx + width // 2, cy), fill=color, width=3)
    d.polygon([(cx, cy - 12), (cx + 16, cy), (cx, cy + 12), (cx - 16, cy)], fill=color)


def coin_crest(c, cx, cy, r=56, glow=True):
    """Brand crest — a stacked gold coin marked '$': the icon of money & budgeting."""
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    c.alpha_composite(ov)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    # coin stack shadow discs behind
    cr = r * 0.50
    for dy in (r * 0.34, r * 0.17):
        d.ellipse((cx - cr, cy - cr + dy, cx + cr, cy + cr + dy), fill=(150, 120, 70))
    # top coin face
    grad_round(c, (cx - cr, cy - cr, cx + cr, cy + cr), int(cr), GOLD_HI, GOLD, outline=(150, 120, 70), width=3)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.ellipse((cx - cr * 0.74, cy - cr * 0.74, cx + cr * 0.74, cy + cr * 0.74), outline=PRIMARY_DK, width=2)
    d.text((cx, cy - 2), "$", font=fserif(int(cr * 1.05)), fill=PRIMARY_DK, anchor="mm")
    c.alpha_composite(ov)


book_crest = coin_crest


def stat_chip(c, cx, cy, big, small, w=400, h=150):
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, 20, 24, 80, 16)
    grad_round(c, box, 20, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=3)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((box[0] + 18, box[1] + 12, box[2] - 18, box[1] + 18), radius=3, fill=GOLD_HI)
    d.text((cx, cy - h * 0.16), big, font=fserif(46), fill=GOLD_HI, anchor="mm")
    d.text((cx, cy + h * 0.28), small, font=fs(20), fill=WHITE, anchor="mm")
    c.alpha_composite(ov)


def donut(d, cx, cy, r, segs, center_top=None, center_sub=None, hole=0.55):
    ang = -90
    for pct, col in segs:
        s = pct * 3.6
        d.pieslice((cx - r, cy - r, cx + r, cy + r), ang, ang + s, fill=col); ang += s
    hr = r * hole
    d.ellipse((cx - hr, cy - hr, cx + hr, cy + hr), fill=WHITE)
    if center_top:
        d.text((cx, cy - (10 if center_sub else 0)), center_top, font=fserif(int(r * 0.34)), fill=PRIMARY, anchor="mm")
    if center_sub:
        d.text((cx, cy + int(r * 0.26)), center_sub, font=fs(int(r * 0.13)), fill=TEXT_MUTED, anchor="mm")


def legend(d, x, y, items, fsz=20, gap=40):
    for i, (col, lab) in enumerate(items):
        yy = y + i * gap
        d.rounded_rectangle((x, yy, x + 24, yy + 24), radius=5, fill=col)
        d.text((x + 36, yy + 12), lab, font=fs(fsz, bold=False), fill=TEXT, anchor="lm")


def hbars(img, d, box, items, color=(HIGHLIGHT, (70, 200, 165)), labelcol=ACCENT):
    x0, y0, x1, y1 = box; n = len(items); rowh = min((y1 - y0) / n, 110); bw_max = (x1 - x0) - 150
    y0 = y0 + max(((y1 - y0) - rowh * n) / 2, 0)
    for i, (lab, frac, vlabel) in enumerate(items):
        yy = y0 + rowh * i + rowh * 0.16
        d.text((x0, yy), lab, font=fs(18, bold=False), fill=TEXT, anchor="lt")
        d.rounded_rectangle((x0, yy + 30, x0 + bw_max, yy + 56), radius=13, fill=(236, 230, 220))
        grad_round(img, (x0, yy + 30, x0 + bw_max * frac, yy + 56), 13, color[0], color[1])
        d.text((x0 + bw_max + 14, yy + 43), vlabel, font=fs(18), fill=labelcol, anchor="lm")


def fit_font(d, text, max_w, start, serif=True):
    s = start; f = fserif(s) if serif else fs(s)
    while s > 12 and d.textlength(text, font=f) > max_w:
        s -= 1; f = fserif(s) if serif else fs(s)
    return f


def app_window(img, box, active_idx, content_fn, file_label=FILE_LABEL):
    x0, y0, x1, y1 = box
    shadow(img, box, 26, 40, 95, 22)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle(box, radius=24, fill=WHITE, outline=(210, 203, 190), width=2)
    img.alpha_composite(ov)
    tb_h = 58
    grad_round(img, (x0, y0, x1, y0 + tb_h + 24), 24, (54, 56, 60), (44, 46, 50))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rectangle((x0, y0 + tb_h, x1, y0 + tb_h + 4), fill=(36, 38, 42))
    for i, col in enumerate([(237, 106, 94), (245, 191, 79), (98, 197, 84)]):
        od.ellipse((x0 + 30 + i * 36, y0 + tb_h // 2 - 11, x0 + 52 + i * 36, y0 + tb_h // 2 + 11), fill=col)
    od.text(((x0 + x1) / 2, y0 + tb_h // 2), file_label, font=fs(18, bold=False), fill=(225, 222, 215), anchor="mm")
    img.alpha_composite(ov)
    sb_w = int((x1 - x0) * 0.205)
    sb = (x0, y0 + tb_h, x0 + sb_w, y1)
    grad_round(img, (sb[0], sb[1], sb[2] + 24, sb[3]), 0, PRIMARY_LT, PRIMARY_DK)
    grad_round(img, (sb[0], y1 - 24, sb[2], y1), 24, PRIMARY_DK, PRIMARY_DK)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    bx = sb[0] + 26
    od.text((bx, sb[1] + 26), "MONEY", font=fs(16), fill=GOLD_HI, anchor="lt")
    od.text((bx, sb[1] + 49), "14-tab system", font=fs(14, bold=False), fill=(170, 200, 192), anchor="lt")
    od.line((sb[0] + 20, sb[1] + 76, sb[2] - 16, sb[1] + 76), fill=(255, 255, 255, 40), width=1)
    list_top = sb[1] + 84
    rowh = (y1 - 18 - list_top) / len(TABS)
    palette = [HIGHLIGHT, GOLD_HI, SURFACE, (150, 200, 190)]
    for i, name in enumerate(TABS):
        ry = list_top + i * rowh
        if i == active_idx:
            od.rounded_rectangle((sb[0] + 12, ry + 1, sb[2] - 10, ry + rowh - 1), radius=6, fill=(255, 255, 255, 235))
            od.rounded_rectangle((sb[0] + 12, ry + 1, sb[0] + 18, ry + rowh - 1), radius=3, fill=GOLD_HI)
            dotc = PRIMARY; txtc = PRIMARY; font = fs(14)
        else:
            dotc = palette[i % len(palette)]; txtc = (214, 226, 222); font = fs(13, bold=False)
        cyr = ry + rowh / 2
        od.ellipse((sb[0] + 26, cyr - 4, sb[0] + 34, cyr + 4), fill=dotc)
        od.text((sb[0] + 46, cyr), name, font=font, fill=txtc, anchor="lm")
    img.alpha_composite(ov)
    cbox = (sb[2] + 1, y0 + tb_h + 4, x1, y1)
    content_fn(img, cbox)


KPIS = [
    ("INCOME", "$5,200", "this month"),
    ("SPENT", "$5,200", "actual"),
    ("LEFT TO BUDGET", "$0", "zero-based ✓"),
    ("SAVED", "$850", "this month"),
    ("SAVINGS RATE", "16%", "of income"),
    ("BILLS PAID", "80%", "8 of 10"),
    ("NET WORTH", "$111,300", "assets − debt"),
    ("TOTAL DEBT", "$30,200", "3 debts"),
    ("SAVINGS GOALS", "53%", "avg progress"),
    ("SINKING FUNDS", "$1,650", "set aside"),
    ("SUBSCRIPTIONS", "$80", "per month"),
    ("HEALTH SCORE", "80%", "blended"),
]


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Money Dashboard", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 64), "The Bennett Household · This Month  ·  give every dollar a job, then watch it grow", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 200, y0 + 26, x1 - pad, y0 + 62), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 100, y0 + 44), "● $0 left — balanced", font=fs(15), fill=PRIMARY, anchor="mm")
    gx = x0 + pad; gy = y0 + 98; gw = (x1 - x0 - 2 * pad); gap = 14
    kw = (gw - 5 * gap) / 6; kh = 116
    for i, (lab, val, sub) in enumerate(KPIS):
        r, ci = divmod(i, 6)
        kx = gx + ci * (kw + gap); ky = gy + r * (kh + gap)
        d.rounded_rectangle((kx, ky, kx + kw, ky + kh), radius=12, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((kx + 12, ky, kx + kw - 12, ky + 5), radius=2, fill=GOLD_LT)
        d.text((kx + 14, ky + 16), lab, font=fs(12), fill=ACCENT, anchor="lt")
        vf = fit_font(d, val, kw - 28, 30)
        d.text((kx + 14, ky + 58), val, font=vf, fill=PRIMARY, anchor="lm")
        d.text((kx + 14, ky + 96), sub, font=fs(12, bold=False), fill=TEXT_MUTED, anchor="lm")
    cy_top = gy + 2 * (kh + gap) + 18
    d.text((gx, cy_top), "FINANCIAL HEALTH · SPENDING · BILLS", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 34; panel_h = (y1 - panels_y - pad); pw = (gw - 3 * gap) / 4
    px = gx
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Financial Health", font=fs(17), fill=ACCENT, anchor="lt")
    hbars(img, d, (px + 20, panels_y + 50, px + pw - 16, panels_y + panel_h - 16),
          [("On budget", 1.0, "100%"), ("Savings rate", 0.82, "82%"), ("Bills on time", 0.80, "80%"),
           ("Emergency fund", 0.63, "63%"), ("Savings goals", 0.53, "53%"), ("Sinking funds", 1.0, "100%")],
          color=(GOLD_HI, GOLD))
    px = gx + (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Spending by Group", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.44, panels_y + panel_h * 0.50, min(panel_h * 0.27, pw * 0.27),
          [(28.8, PRIMARY), (23.2, GOLD_LT), (25.0, HIGHLIGHT), (13.8, SURFACE), (9.2, ACCENT)], "$5,200", "spent")
    legend(d, px + pw * 0.04, panels_y + panel_h - 118, [(PRIMARY, "Housing $1,500"), (GOLD_LT, "Food $1,207"),
           (HIGHLIGHT, "Save+Debt $1,300"), (SURFACE, "Bills $718"), (ACCENT, "Other $475")], 14, 26)
    px = gx + 2 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Bills & Money Moves", font=fs(17), fill=ACCENT, anchor="lt")
    tasks = [("Student loan — day 28", WARN_BG), ("Gym — day 25", WARN_BG), ("Rent — paid ✓", MINT_BG),
             ("Car insurance — paid", MINT_BG), ("Internet — paid", MINT_BG), ("8 of 10 bills paid", MINT_BG)]
    ty = panels_y + 58
    for lab, dot in tasks:
        d.ellipse((px + 20, ty + 6, px + 40, ty + 26), fill=dot, outline=PRIMARY, width=2)
        d.text((px + 54, ty + 16), lab, font=fs(16, bold=False), fill=TEXT, anchor="lm")
        ty += 46
    px = gx + 3 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Health Score", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.5, panels_y + panel_h * 0.54, min(panel_h * 0.30, pw * 0.30),
          [(80, PRIMARY), (20, SURFACE)], "80%", "on track")


def _table(img, cbox, title, subtitle, headers, colf, rows, total_row=None,
           status_col=None, status_map=None, hdr_top=104):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), title, font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), subtitle, font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    tx0, tx1 = x0 + pad, x1 - pad
    ty = y0 + hdr_top
    colx = [tx0 + (tx1 - tx0) * f for f in colf]
    hdr_h = 42
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else "mm"
        d.text((colx[i] + (14 if i == 0 else 0), ty + hdr_h / 2), h, font=fs(15), fill=WHITE, anchor=anc)
    nrows = len(rows) + (1 if total_row else 0)
    rh = (y1 - pad - (ty + hdr_h)) / nrows
    for i, row in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        for ci, val in enumerate(row):
            anc = "lm" if ci == 0 else "mm"
            hx = colx[ci] + (14 if ci == 0 else 0)
            sval = str(val)
            if status_map is not None and ci == status_col:
                bg, fg = status_map.get(sval, ((235, 230, 222), TEXT_MUTED))
                d.rounded_rectangle((hx - 60, ry + rh / 2 - 15, hx + 60, ry + rh / 2 + 15), radius=14, fill=bg)
                d.text((hx, ry + rh / 2), sval, font=fs(14), fill=fg, anchor="mm")
            else:
                col = PRIMARY if ci == 0 else TEXT
                d.text((hx, ry + rh / 2), sval, font=fs(16) if ci == 0 else fs(15, bold=False), fill=col, anchor=anc)
    if total_row:
        ry = ty + hdr_h + len(rows) * rh
        d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
        for ci, val in enumerate(total_row):
            anc = "lm" if ci == 0 else "mm"
            hx = colx[ci] + (14 if ci == 0 else 0)
            if val != "":
                d.text((hx, ry + rh / 2), str(val), font=fs(16), fill=PRIMARY, anchor=anc)


def content_budget(img, cbox):
    rows = [
        ("Rent / Mortgage", "$1,500", "$1,500", "On track"),
        ("Groceries", "$600", "$642", "Over"),
        ("Dining & takeout", "$250", "$310", "Over"),
        ("Transportation", "$300", "$280", "Under"),
        ("Savings & investing", "$800", "$800", "On track"),
        ("Extra debt payment", "$450", "$450", "On track"),
    ]
    _table(img, cbox, "Monthly Budget (Zero-Based)",
           "Give every dollar a job — planned vs actual, until 'Left to Budget' hits $0",
           ["CATEGORY", "PLANNED", "ACTUAL", "STATUS"],
           [0.0, 0.44, 0.62, 0.84], rows,
           total_row=("TOTAL", "$5,200", "$5,200", ""),
           status_col=3, status_map={"On track": (MINT_BG, PRIMARY), "Under": (SURFACE, (110, 88, 58)),
                                     "Over": (RED_BG, DANGER)})


def content_networth(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Net Worth", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "Assets minus liabilities — the one number that tells the truth. Watch it grow.",
           font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 250, y0 + 24, x1 - pad, y0 + 62), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 125, y0 + 43), "● net worth $111,300", font=fs(15), fill=PRIMARY, anchor="mm")
    items = [
        ("Home equity", 60000, "$60,000", False),
        ("Retirement (401k/IRA)", 45000, "$45,000", False),
        ("Car value", 15000, "$15,000", False),
        ("Cash & checking", 12000, "$12,000", False),
        ("Emergency fund", 9500, "$9,500", False),
        ("Liabilities (debt)", 30200, "$30,200", True),
    ]
    mx = max(v for _, v, _, _ in items)
    gx = x0 + pad; gy = y0 + 118; gw = x1 - x0 - 2 * pad
    rowh = (y1 - gy - 118 - pad) / len(items)
    bw_max = gw - 240
    for i, (lab, val, vlab, is_liab) in enumerate(items):
        yy = gy + i * rowh
        d.text((gx, yy + rowh * 0.20), lab, font=fs(21, bold=False), fill=TEXT, anchor="lt")
        d.rounded_rectangle((gx, yy + rowh * 0.50, gx + bw_max, yy + rowh * 0.50 + 32), radius=16, fill=(236, 230, 220))
        col = (DANGER, (150, 50, 50)) if is_liab else (HIGHLIGHT, (70, 200, 165))
        grad_round(img, (gx, yy + rowh * 0.50, gx + bw_max * (val / mx), yy + rowh * 0.50 + 32), 16, col[0], col[1])
        d.text((gx + bw_max + 18, yy + rowh * 0.50 + 16), vlab, font=fs(22), fill=DANGER if is_liab else PRIMARY, anchor="lm")
    # net worth banner
    by = y1 - 118 + 6
    grad_round(img, (gx, by, x1 - pad, y1 - pad), 16, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=3)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((gx + 28, (by + y1 - pad) / 2), "NET WORTH  (assets − liabilities)", font=fs(22), fill=GOLD_HI, anchor="lm")
    od.text((x1 - pad - 28, (by + y1 - pad) / 2), "$111,300", font=fserif(46), fill=WHITE, anchor="rm")
    img.alpha_composite(ov)


def content_savings(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Savings Goals", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "Every goal with its target & progress — a full bar is a very good day.",
           font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    items = [
        ("Emergency Fund (6 mo)", 9500, 15000, "$9,500 / $15,000", "63%"),
        ("Home Projects", 3400, 5000, "$3,400 / $5,000", "68%"),
        ("Vacation Fund", 1800, 3000, "$1,800 / $3,000", "60%"),
        ("New Car Fund", 2200, 10000, "$2,200 / $10,000", "22%"),
    ]
    gx = x0 + pad; gy = y0 + 118; gw = x1 - x0 - 2 * pad
    rowh = (y1 - gy - pad) / len(items)
    bw_max = gw - 150
    for i, (lab, saved, target, vlab, pct) in enumerate(items):
        yy = gy + i * rowh
        d.text((gx, yy + rowh * 0.14), lab, font=fs(23), fill=PRIMARY, anchor="lt")
        d.text((gx + bw_max + 20, yy + rowh * 0.14), vlab, font=fs(17, bold=False), fill=TEXT_MUTED, anchor="lt")
        d.rounded_rectangle((gx, yy + rowh * 0.50, gx + bw_max, yy + rowh * 0.50 + 40), radius=20, fill=(236, 230, 220))
        frac = saved / target
        grad_round(img, (gx, yy + rowh * 0.50, gx + bw_max * frac, yy + rowh * 0.50 + 40), 20, HIGHLIGHT, (70, 200, 165))
        d.text((gx + bw_max + 20, yy + rowh * 0.50 + 20), pct, font=fs(26), fill=PRIMARY, anchor="lm")


def content_bills(img, cbox):
    rows = [
        ("Rent / Mortgage", "$1,500", "Day 1", "Paid"),
        ("Electric", "$120", "Day 8", "Paid"),
        ("Phone", "$90", "Day 15", "Paid"),
        ("Car insurance", "$140", "Day 18", "Paid"),
        ("Gym", "$20", "Day 25", "Due"),
        ("Student loan", "$210", "Day 28", "Due"),
    ]
    _table(img, cbox, "Bill Tracker",
           "Every recurring bill — amount, due day & status. Never a late fee again.  8 of 10 paid",
           ["BILL", "AMOUNT", "DUE", "STATUS"],
           [0.0, 0.46, 0.64, 0.84], rows,
           status_col=3, status_map={"Paid": (MINT_BG, PRIMARY), "Due": (WARN_BG, ACCENT),
                                     "Overdue": (RED_BG, DANGER)})


# ---------- renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    coin_crest(img, SIZE // 2, 132, r=56)
    pill(img, SIZE // 2, 256, "THE COMPLETE PERSONAL-FINANCE SYSTEM", font=fs(20), pad_x=40, pad_y=20)
    wordmark(img, SIZE // 2, 400, "BUDGET COMMAND CENTER", 84, max_w=1780)
    gold_divider(img, SIZE // 2, 500, width=520)
    tc(d, (SIZE // 2, 550), "Zero-based budget, bills, savings, debt & net worth — your whole money life, one system.",
       fs(20, bold=False), (224, 213, 190))
    chips = [("14", "CONNECTED TABS"), ("12", "PRINTABLE PAGES"), ("ZERO", "BASED BUDGET")]
    cw = 440
    total = len(chips) * cw + (len(chips) - 1) * 28
    startx = (SIZE - total) // 2 + cw // 2
    for i, (b, s) in enumerate(chips):
        stat_chip(img, startx + i * (cw + 28), 704, b, s, w=cw)
    app_window(img, (70, 800, SIZE - 70, 1900), 0, content_dashboard)
    pill(img, SIZE // 2, SIZE - 52, "GOOGLE SHEETS + PRINTABLE PDF · INSTANT DOWNLOAD",
         font=fs(33), pad_x=50, pad_y=24, star=True, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_inside(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 120, "EVERYTHING INSIDE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 238), "14 Connected Tabs", fserif(58), WHITE)
    gold_divider(img, SIZE // 2, 308, width=520)
    tc(d, (SIZE // 2, 352), "Not a spreadsheet — a complete give-every-dollar-a-job money system",
       fs(24, bold=False), (226, 214, 190))
    cards = [
        ("Start Here", "how it all works"), ("Money Dashboard", "income, spend, save & Health"),
        ("Income", "every paycheck & source"), ("Monthly Budget", "zero-based, planned vs actual"),
        ("Bills", "amounts, due days & status"), ("Expense Log", "every transaction, tagged"),
        ("Savings Goals", "targets & progress bars"), ("Sinking Funds", "for the big irregular costs"),
        ("Debt Snapshot", "balance, rate & minimum"), ("Net Worth", "assets − liabilities"),
        ("Subscriptions", "the quiet budget-killer"), ("Year View", "12 months at a glance"),
        ("No-Spend", "the streak challenge"), ("Settings", "set it once"),
        ("+ 12 Printable Pages", "print & keep"), ("Live Health Score", "it all rolls up"),
    ]
    cols = 4
    margin = 88
    gx, gy = 22, 20
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    top = 452
    rows_n = 4
    ch = (SIZE - top - 60 - (rows_n - 1) * gy) // rows_n
    for i, (title, sub) in enumerate(cards):
        r, ccol = divmod(i, cols)
        x = margin + ccol * (cw + gx); y = top + r * (ch + gy)
        shadow(img, (x, y, x + cw, y + ch), 14, 12, 40, 8)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        last = (i >= len(cards) - 2)
        od.rounded_rectangle((x, y, x + cw, y + ch), radius=13, fill=(MINT_BG if last else WHITE), outline=(GOLD_LT if last else (232, 224, 208)), width=2)
        od.rectangle((x + 3, y + 8, x + 7, y + ch - 8), fill=GOLD_LT)
        cyc = y + ch // 2; bx = x + 42
        od.ellipse((bx - 20, cyc - 20, bx + 20, cyc + 20), fill=PRIMARY)
        od.text((bx, cyc), str(i + 1) if not last else "★", font=fs(19), fill=GOLD_HI, anchor="mm")
        tf = fit_font(od, title, cw - 78, 18, serif=False)
        od.text((x + 76, cyc - 15), title, font=tf, fill=PRIMARY, anchor="lm")
        od.text((x + 76, cyc + 18), sub, font=fs(13, bold=False), fill=TEXT_MUTED, anchor="lm")
        img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_budget(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "THE ZERO-BASED BUDGET", font=fs(30), pad_x=48, pad_y=22)
    tc(d, (SIZE // 2, 232), "Give Every Dollar a Job", fserif(46), WHITE)
    tc(d, (SIZE // 2, 300), "Planned vs actual for every category — overspending flags itself the day it happens",
       fs(22, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 2, content_budget)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_networth(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "WATCH IT GROW", font=fs(30), pad_x=48, pad_y=22)
    tc(d, (SIZE // 2, 232), "One Number That Tells the Truth", fserif(44), WHITE)
    tc(d, (SIZE // 2, 300), "Assets minus liabilities, updated monthly — the real scoreboard for your money",
       fs(22, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 8, content_networth)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_save(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "SAVE IT · PAY IT", font=fs(32), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "Goals & Bills, Handled", fserif(44), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1150), 5, content_savings)
    app_window(img, (60, 1180, SIZE - 60, SIZE - 60), 3, content_bills)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_printables(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=470)
    d = ImageDraw.Draw(img)
    coin_crest(img, SIZE // 2, 110, r=48)
    pill(img, SIZE // 2, 214, "PLUS 12 PRINTABLE PAGES", font=fs(30), pad_x=48, pad_y=20)
    tc(d, (SIZE // 2, 300), "Print, Fill & Keep — No Screen Needed", fserif(44), WHITE)
    gold_divider(img, SIZE // 2, 362, width=480)
    tc(d, (SIZE // 2, 404), "A matching print-ready PDF pack for the budget binder & the fridge",
       fs(22, bold=False), (226, 214, 190))
    labels = ["Monthly Budget", "Bill Tracker", "Expense Log", "Savings Goals",
              "Sinking Funds", "Debt Payoff", "Net Worth Sheet", "Subscriptions Audit",
              "Income Tracker", "No-Spend Challenge", "Year-at-a-Glance", "Money Goals"]
    print_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing", "print")
    cols, rows_n = 4, 3
    margin = 96; gapx, gapy = 34, 66
    cw = (SIZE - 2 * margin - (cols - 1) * gapx) // cols
    ch = int(cw * 11 / 8.5)
    top = 560
    for i in range(12):
        r, c = divmod(i, cols)
        x = margin + c * (cw + gapx); y = top + r * (ch + gapy)
        shadow(img, (x, y, x + cw, y + ch), 12, 16, 60, 10)
        thumb_path = os.path.join(print_dir, f"page_{i+1:02d}.png")
        if os.path.exists(thumb_path):
            th = Image.open(thumb_path).convert("RGB").resize((cw, ch))
            img.paste(th, (x, y))
            ImageDraw.Draw(img).rounded_rectangle((x, y, x + cw, y + ch), radius=6, outline=(210, 203, 190), width=2)
        else:
            ImageDraw.Draw(img).rounded_rectangle((x, y, x + cw, y + ch), radius=6, fill=WHITE, outline=GRID, width=2)
        d = ImageDraw.Draw(img)
        d.rounded_rectangle((x + cw // 2 - 26, y - 20, x + cw // 2 + 26, y + 20), radius=13, fill=PRIMARY)
        d.text((x + cw // 2, y), str(i + 1), font=fs(20), fill=GOLD_HI, anchor="mm")
        d.text((x + cw // 2, y + ch + 30), labels[i], font=fs(18), fill=PRIMARY, anchor="mm")
    pill(img, SIZE // 2, SIZE - 46, "PRINT AT HOME · US LETTER · READY IN ONE CLICK",
         font=fs(28), pad_x=44, pad_y=22, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def main():
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("01_hero.png", render_hero),
        ("02_inside.png", render_inside),
        ("03_budget.png", render_budget),
        ("04_networth.png", render_networth),
        ("05_save.png", render_save),
        ("06_printables.png", render_printables),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()
