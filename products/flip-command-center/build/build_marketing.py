"""Marketing image set for Flip Command Center™ (6 images, 2000x2000).

  01_hero.png       - branded hero + live flip dashboard
  02_inside.png     - "everything inside — 17 connected tabs"
  03_deal.png       - the Deal Analyzer money shot (70% rule + profit)
  04_rehab.png      - rehab budget, planned vs actual (catch overruns)
  05_plan.png       - scope of work + timeline
  06_printables.png - the 12-page printable PDF pack

Run: python3 build_marketing.py   (run build_pdf.py first for the printables grid)
"""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFilter, ImageFont

PRIMARY = (27, 79, 72)
PRIMARY_DK = (18, 56, 51)
PRIMARY_LT = (33, 92, 83)
ACCENT = (147, 115, 86)
GOLD = (180, 145, 90)
GOLD_LT = (201, 168, 106)
GOLD_HI = (224, 196, 140)
SURFACE = (229, 211, 186)
HIGHLIGHT = (117, 230, 193)
BG = (251, 248, 242)
BG_TOP = (253, 250, 246)
BG_BOT = (242, 235, 223)
WHITE = (255, 255, 255)
TEXT = (51, 51, 51)
TEXT_MUTED = (132, 126, 116)
DANGER = (201, 76, 76)
MINT_BG = (227, 248, 239)
WARN_BG = (251, 240, 226)
RED_BG = (251, 230, 230)
GRID = (228, 222, 210)
ROW_ALT = (246, 241, 232)
DOT = (228, 220, 206)
GREEN = (32, 120, 96)
SIZE = 2000

SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"

TABS = ["Dashboard", "Deal Analyzer", "Property Details", "Rehab Budget", "Scope of Work",
        "Contractors", "Draws & Payments", "Materials", "Timeline", "Holding Costs",
        "Financing", "Comps & ARV", "Selling & Exit", "Punch List", "Photo Log", "Settings"]

FILE_LABEL = "Flip_Command_Center.xlsx — 214 Maple St · Fix & Flip · Maplewood"


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


def house_crest(c, cx, cy, r=56, glow=True):
    """Brand crest — a freshly-renovated little house: the icon of the flip."""
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    c.alpha_composite(ov)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    bw = r * 0.62          # body half-width
    eaves = cy - r * 0.02
    base = cy + r * 0.5
    # roof
    d.polygon([(cx, cy - r * 0.5), (cx - bw - r * 0.12, eaves), (cx + bw + r * 0.12, eaves)],
              fill=GOLD_HI, outline=(150, 120, 70))
    # body
    d.rectangle((cx - bw, eaves, cx + bw, base), fill=SURFACE, outline=(150, 120, 70), width=2)
    # door
    dw = r * 0.2
    d.rectangle((cx - dw, base - r * 0.42, cx + dw, base), fill=PRIMARY, outline=(150, 120, 70))
    d.ellipse((cx + dw * 0.35, base - r * 0.24, cx + dw * 0.55, base - r * 0.06), fill=GOLD_HI)
    # window
    wy = eaves + r * 0.12
    d.rectangle((cx + bw * 0.28, wy, cx + bw * 0.78, wy + r * 0.2), fill=HIGHLIGHT, outline=(150, 120, 70))
    # chimney
    d.rectangle((cx - bw * 0.72, cy - r * 0.34, cx - bw * 0.5, eaves - r * 0.06), fill=PRIMARY, outline=(150, 120, 70))
    c.alpha_composite(ov)


# keep the name used by the detail-image module
book_crest = house_crest


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
    od.text((bx, sb[1] + 26), "FIX & FLIP", font=fs(16), fill=GOLD_HI, anchor="lt")
    od.text((bx, sb[1] + 49), "17-tab system", font=fs(14, bold=False), fill=(170, 200, 192), anchor="lt")
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
    ("ARV", "$340,000", "after repair"),
    ("PURCHASE", "$185,000", "the offer"),
    ("REHAB BUDGET", "$45,000", "all categories"),
    ("ALL-IN COST", "$266,802", "everything in"),
    ("PROJECTED PROFIT", "$73,198", "at sale"),
    ("CASH-ON-CASH ROI", "77%", "on cash in"),
    ("70% RULE MAO", "$193,000", "max offer"),
    ("VERDICT", "BUY", "under MAO"),
    ("BUDGET USED", "76%", "$34K of $45K"),
    ("SPENT TO DATE", "$34,000", "so far"),
    ("TASKS DONE", "56%", "of scope"),
    ("DEAL SCORE", "83%", "blended"),
]


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Flip Dashboard", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 64), "214 Maple St · fix & flip  ·  know your number, protect it to the closing table", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 150, y0 + 26, x1 - pad, y0 + 62), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 75, y0 + 44), "● BUY", font=fs(16), fill=PRIMARY, anchor="mm")
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
    d.text((gx, cy_top), "WHERE THE SALE GOES · SCORE · WHAT'S LEFT", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 34; panel_h = (y1 - panels_y - pad); pw = (gw - 3 * gap) / 4
    px = gx
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Where the Sale Goes", font=fs(17), fill=ACCENT, anchor="lt")
    hbars(img, d, (px + 20, panels_y + 50, px + pw - 16, panels_y + panel_h - 16),
          [("Purchase", 0.54, "$185K"), ("Rehab", 0.13, "$45K"), ("Holding", 0.03, "$9.5K"),
           ("Selling", 0.07, "$24K"), ("Profit", 0.215, "$73K")], color=(GOLD_HI, GOLD))
    px = gx + (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Deal Score", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.5, panels_y + panel_h * 0.54, min(panel_h * 0.30, pw * 0.30),
          [(83, PRIMARY), (17, SURFACE)], "83%", "strong deal")
    px = gx + 2 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "What's Left", font=fs(17), fill=ACCENT, anchor="lt")
    tasks = [("Kitchen final & backsplash", WARN_BG), ("Appliance install", WARN_BG), ("Exterior paint", WARN_BG),
             ("Landscaping & curb", WARN_BG), ("Punch list", MINT_BG), ("Stage & list", MINT_BG)]
    ty = panels_y + 58
    for lab, dot in tasks:
        d.ellipse((px + 20, ty + 6, px + 40, ty + 26), fill=dot, outline=PRIMARY, width=2)
        d.text((px + 54, ty + 16), lab, font=fs(16, bold=False), fill=TEXT, anchor="lm")
        ty += 46
    px = gx + 3 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Cash-on-Cash ROI", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.5, panels_y + panel_h * 0.54, min(panel_h * 0.30, pw * 0.30),
          [(77, PRIMARY), (23, SURFACE)], "77%", "on cash in")


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
                d.rounded_rectangle((hx - 74, ry + rh / 2 - 15, hx + 74, ry + rh / 2 + 15), radius=14, fill=bg)
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


def content_deal(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Deal Analyzer", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "Enter the numbers once — the 70% rule, profit & ROI compute instantly",
           font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    gw = x1 - x0 - 2 * pad
    colw = (gw - 30) / 2
    ix = x0 + pad; ox = x0 + pad + colw + 30
    top = y0 + 116
    d.text((ix, top), "THE NUMBERS", font=fs(16), fill=ACCENT, anchor="lt")
    d.text((ox, top), "THE DEAL", font=fs(16), fill=ACCENT, anchor="lt")
    inputs = [("After-Repair Value", "$340,000"), ("Purchase Price", "$185,000"),
              ("Rehab Budget", "$45,000"), ("Buy-Side Closing", "$3,500"),
              ("Holding (5 months)", "$9,502"), ("Selling Costs (7%)", "$23,800")]
    outs = [("All-In Cost", "$266,802", False), ("Cash Invested", "$95,002", False),
            ("Projected Profit", "$73,198", True), ("Cash-on-Cash ROI", "77%", True),
            ("Return on Cost", "27%", False), ("70% Rule — Max Offer", "$193,000", False)]
    ry = top + 40; rh = 62
    for i, (lab, val) in enumerate(inputs):
        yy = ry + i * rh
        d.rounded_rectangle((ix, yy, ix + colw, yy + rh - 12), radius=10, fill=WHITE, outline=GRID, width=2)
        d.text((ix + 16, yy + (rh - 12) / 2), lab, font=fs(18, bold=False), fill=TEXT, anchor="lm")
        d.text((ix + colw - 16, yy + (rh - 12) / 2), val, font=fs(19), fill=PRIMARY, anchor="rm")
    for i, (lab, val, hot) in enumerate(outs):
        yy = ry + i * rh
        d.rounded_rectangle((ox, yy, ox + colw, yy + rh - 12), radius=10, fill=(MINT_BG if hot else WHITE),
                            outline=(GOLD_LT if hot else GRID), width=2)
        d.text((ox + 16, yy + (rh - 12) / 2), lab, font=fs(18, bold=(hot)), fill=(PRIMARY if hot else TEXT), anchor="lm")
        d.text((ox + colw - 16, yy + (rh - 12) / 2), val, font=fs(20 if hot else 19), fill=PRIMARY, anchor="rm")
    by = ry + 6 * rh + 8
    grad_round(img, (ix, by, ox + colw, by + 96), 16, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=3)
    d = ImageDraw.Draw(img)
    d.text(((ix + ox + colw) / 2, by + 32), "✓  BUY — under the 70% rule", font=fs(28), fill=GOLD_HI, anchor="mm")
    d.text(((ix + ox + colw) / 2, by + 70), "$73,198 projected profit  ·  77% cash-on-cash ROI",
           font=fs(20, bold=False), fill=WHITE, anchor="mm")


def content_rehab(img, cbox):
    rows = [
        ("Demo & dumpster", "$2,800", "$3,100", "111%"),
        ("Kitchen", "$10,500", "$5,500", "52%"),
        ("Bathrooms (2)", "$6,500", "$3,900", "60%"),
        ("Flooring", "$4,500", "$4,500", "100%"),
        ("Electrical", "$2,500", "$2,650", "106%"),
        ("Paint", "$3,200", "$900", "28%"),
        ("Permits & fees", "$1,000", "$1,050", "105%"),
    ]
    _table(img, cbox, "Rehab Budget",
           "Planned vs actual by category — overruns flag red the day they happen.  $34K of $45K used",
           ["CATEGORY", "PLANNED", "ACTUAL", "% USED"],
           [0.0, 0.42, 0.62, 0.85], rows,
           status_col=3, status_map={"111%": (RED_BG, DANGER), "106%": (RED_BG, DANGER), "105%": (RED_BG, DANGER),
                                     "100%": (WARN_BG, ACCENT), "60%": (MINT_BG, PRIMARY), "52%": (MINT_BG, PRIMARY),
                                     "28%": (MINT_BG, PRIMARY)},
           total_row=("TOTAL", "$45,000", "$34,000", "76%"))


def content_scope(img, cbox):
    rows = [
        ("Whole House", "Rewire & new panel", "Electrical", "Done"),
        ("Whole House", "LVP flooring throughout", "Flooring", "Done"),
        ("Kitchen", "Cabinets & quartz counters", "Kitchen", "In Progress"),
        ("Primary Bath", "New vanity, tile & tub", "Bath", "In Progress"),
        ("Kitchen", "Stainless appliance package", "Kitchen", "Not Started"),
        ("Yard", "Sod, mulch & landscaping", "Landscaping", "Not Started"),
    ]
    _table(img, cbox, "Scope of Work",
           "Room by room, task by task — 56% of the scope checked off",
           ["ROOM", "TASK", "TRADE", "STATUS"],
           [0.0, 0.22, 0.60, 0.84], rows,
           status_col=3, status_map={"Done": (MINT_BG, PRIMARY), "In Progress": (WARN_BG, ACCENT),
                                     "Not Started": ((235, 230, 222), TEXT_MUTED)})


def content_timeline(img, cbox):
    rows = [
        ("Acquisition & permits", "Done"),
        ("Demo", "Done"),
        ("Rough-in (MEP)", "Done"),
        ("Kitchen & baths", "In Progress"),
        ("Finishes (floor/paint/trim)", "In Progress"),
        ("Punch list & staging", "Not Started"),
    ]
    _table(img, cbox, "Timeline & Phases",
           "The whole job on a calendar — 3 of 7 phases complete, on pace to list",
           ["PHASE", "STATUS"],
           [0.0, 0.74], rows,
           status_col=1, status_map={"Done": (MINT_BG, PRIMARY), "In Progress": (WARN_BG, ACCENT),
                                     "Not Started": ((235, 230, 222), TEXT_MUTED)})


# ---------- renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    house_crest(img, SIZE // 2, 132, r=56)
    pill(img, SIZE // 2, 256, "THE ULTIMATE HOUSE-FLIPPING SYSTEM", font=fs(20), pad_x=40, pad_y=20)
    wordmark(img, SIZE // 2, 400, "FLIP COMMAND CENTER", 92, max_w=1700)
    gold_divider(img, SIZE // 2, 500, width=520)
    tc(d, (SIZE // 2, 550), "Deal analyzer, rehab budget, contractors, timeline & exit — from offer to sold, one system.",
       fs(20, bold=False), (224, 213, 190))
    chips = [("17", "CONNECTED TABS"), ("70%", "RULE + PROFIT CALC"), ("12", "PRINTABLE PAGES")]
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
    tc(d, (SIZE // 2, 238), "17 Connected Tabs", fserif(58), WHITE)
    gold_divider(img, SIZE // 2, 308, width=520)
    tc(d, (SIZE // 2, 352), "Not a checklist — a complete analyze-it, fund-it, flip-it operating system",
       fs(24, bold=False), (226, 214, 190))
    cards = [
        ("Start Here", "how it all works"), ("Flip Dashboard", "live profit & Deal Score"),
        ("Deal Analyzer", "70% rule + ROI engine"), ("Property Details", "the subject property"),
        ("Rehab Budget", "planned vs actual"), ("Scope of Work", "room-by-room tasks"),
        ("Contractors", "crew & bids"), ("Draws & Payments", "every dollar out"),
        ("Materials", "the shopping list"), ("Timeline", "phases & dates"),
        ("Holding Costs", "the silent profit-killer"), ("Financing", "the capital stack"),
        ("Comps & ARV", "back up your value"), ("Selling & Exit", "the net sheet"),
        ("Punch List", "the last 2%"), ("Photo Log", "before & after"),
        ("Settings", "set it once"), ("+ 12 Printable Pages", "job-site binder"),
    ]
    cols = 4
    margin = 88
    gx, gy = 22, 20
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    top = 452
    rows_n = 5
    ch = (SIZE - top - 60 - (rows_n - 1) * gy) // rows_n
    for i, (title, sub) in enumerate(cards):
        r, ccol = divmod(i, cols)
        x = margin + ccol * (cw + gx); y = top + r * (ch + gy)
        shadow(img, (x, y, x + cw, y + ch), 14, 12, 40, 8)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        last = (i == len(cards) - 1)
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


def render_deal(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "THE DEAL ANALYZER", font=fs(30), pad_x=48, pad_y=22)
    tc(d, (SIZE // 2, 232), "Know Your Number Before You Buy", fserif(46), WHITE)
    tc(d, (SIZE // 2, 300), "The 70% rule, all-in cost, profit & ROI — and a straight BUY / PASS verdict",
       fs(22, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 1, content_deal)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_rehab(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "REHAB BUDGET", font=fs(30), pad_x=48, pad_y=22)
    tc(d, (SIZE // 2, 232), "Catch Overruns the Day They Happen", fserif(44), WHITE)
    tc(d, (SIZE // 2, 300), "Planned vs actual for every category — over-budget lines flag red automatically",
       fs(22, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 3, content_rehab)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_plan(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "RUN THE JOB · HIT THE DATE", font=fs(32), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "Scope & Timeline, Handled", fserif(44), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1150), 4, content_scope)
    app_window(img, (60, 1180, SIZE - 60, SIZE - 60), 8, content_timeline)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_printables(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=470)
    d = ImageDraw.Draw(img)
    house_crest(img, SIZE // 2, 110, r=48)
    pill(img, SIZE // 2, 214, "PLUS 12 PRINTABLE PAGES", font=fs(30), pad_x=48, pad_y=20)
    tc(d, (SIZE // 2, 300), "The Job-Site Binder — Print & Go", fserif(44), WHITE)
    gold_divider(img, SIZE // 2, 362, width=480)
    tc(d, (SIZE // 2, 404), "A matching print-ready PDF pack for the truck, the site & the closing table",
       fs(22, bold=False), (226, 214, 190))
    labels = ["Deal Analyzer", "Rehab Budget", "Scope of Work", "Contractor & Bids",
              "Draw Schedule", "Materials List", "Project Timeline", "Holding Costs",
              "Comps & ARV", "Net / Exit Sheet", "Punch List", "Before & After"]
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
        ("03_deal.png", render_deal),
        ("04_rehab.png", render_rehab),
        ("05_plan.png", render_plan),
        ("06_printables.png", render_printables),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()
