"""Marketing image set for Trading Card Collection Command Center™ (6 images, 2000x2000).

Dense app-screenshot marketing mirroring the real workbook: a left sidebar of
all 13 tabs, the REAL computed KPI numbers, and fully populated tables/charts.

  01_hero.png        - branded hero + live collection dashboard
  02_inside.png      - "everything inside — 13 powerful tabs"
  03_dashboard.png   - full collection dashboard
  04_collection.png  - master collection (valued, with P/L)
  05_value.png       - grading center + trade tracker
  06_mobile.png      - mobile preview

Run: python3 build_marketing.py
"""
from __future__ import annotations
import datetime as dt
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

TABS = ["Dashboard", "Collection", "Purchases", "Sales", "Trades", "Grading",
        "Value Analytics", "Wishlist", "Duplicates", "Deck Builder",
        "Card Vault", "Analytics", "Settings"]

FILE_LABEL = "Trading_Card_Command_Center.xlsx — The Vaulted Collection · 146 cards"


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


def _tilted_card(c, cx, cy, w, h, ang, fillc, outlinec, sparkle=None):
    s = int(max(w, h) * 2)
    lay = Image.new("RGBA", (s, s), (0, 0, 0, 0)); ld = ImageDraw.Draw(lay)
    x0 = (s - w) / 2; y0 = (s - h) / 2
    ld.rounded_rectangle((x0, y0, x0 + w, y0 + h), radius=int(w * 0.14), fill=fillc, outline=outlinec, width=3)
    ld.rounded_rectangle((x0 + 5, y0 + 5, x0 + w - 5, y0 + h - 5), radius=int(w * 0.10), outline=outlinec, width=1)
    if sparkle:
        mx, my = s / 2, s / 2; rr = w * 0.22
        ld.polygon([(mx, my - rr), (mx + rr * 0.55, my), (mx, my + rr), (mx - rr * 0.55, my)], fill=sparkle)
    lay = lay.rotate(ang, resample=Image.BICUBIC, expand=False)
    c.alpha_composite(lay, (int(cx - s / 2), int(cy - s / 2)))


def card_crest(c, cx, cy, r=60, glow=True):
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    c.alpha_composite(ov)
    # fanned trading cards
    cw, ch = r * 0.62, r * 0.9
    _tilted_card(c, cx - r * 0.14, cy + r * 0.04, cw, ch, 14, GOLD, PRIMARY_DK)
    _tilted_card(c, cx + r * 0.16, cy + r * 0.02, cw, ch, -9, GOLD_HI, PRIMARY_DK, sparkle=PRIMARY_DK)
    # mint sparkle accent
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    sx, sy, sr = cx + r * 0.52, cy - r * 0.52, r * 0.14
    d.polygon([(sx, sy - sr), (sx + sr * 0.4, sy - sr * 0.4), (sx + sr, sy), (sx + sr * 0.4, sy + sr * 0.4),
               (sx, sy + sr), (sx - sr * 0.4, sy + sr * 0.4), (sx - sr, sy), (sx - sr * 0.4, sy - sr * 0.4)], fill=HIGHLIGHT)
    c.alpha_composite(ov)


def stat_chip(c, cx, cy, big, small, w=400, h=150):
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, 20, 24, 80, 16)
    grad_round(c, box, 20, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=3)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((box[0] + 18, box[1] + 12, box[2] - 18, box[1] + 18), radius=3, fill=GOLD_HI)
    d.text((cx, cy - h * 0.16), big, font=fserif(48), fill=GOLD_HI, anchor="mm")
    d.text((cx, cy + h * 0.28), small, font=fs(21), fill=WHITE, anchor="mm")
    c.alpha_composite(ov)


def donut(d, cx, cy, r, segs, center_top=None, center_sub=None, hole=0.55):
    ang = -90
    for pct, col in segs:
        s = pct * 3.6
        d.pieslice((cx - r, cy - r, cx + r, cy + r), ang, ang + s, fill=col); ang += s
    hr = r * hole
    d.ellipse((cx - hr, cy - hr, cx + hr, cy + hr), fill=WHITE)
    if center_top:
        d.text((cx, cy - (10 if center_sub else 0)), center_top, font=fserif(int(r * 0.32)), fill=PRIMARY, anchor="mm")
    if center_sub:
        d.text((cx, cy + int(r * 0.26)), center_sub, font=fs(int(r * 0.13)), fill=TEXT_MUTED, anchor="mm")


def legend(d, x, y, items, fsz=20, gap=40):
    for i, (col, lab) in enumerate(items):
        yy = y + i * gap
        d.rounded_rectangle((x, yy, x + 24, yy + 24), radius=5, fill=col)
        d.text((x + 36, yy + 12), lab, font=fs(fsz, bold=False), fill=TEXT, anchor="lm")


def hbars_money(img, d, box, items):
    x0, y0, x1, y1 = box; n = len(items); rowh = min((y1 - y0) / n, 110); bw_max = (x1 - x0) - 190
    y0 = y0 + max(((y1 - y0) - rowh * n) / 2, 0)
    for i, (lab, frac, vlabel) in enumerate(items):
        yy = y0 + rowh * i + rowh * 0.16
        d.text((x0, yy), lab, font=fs(18, bold=False), fill=TEXT, anchor="lt")
        d.rounded_rectangle((x0, yy + 30, x0 + bw_max, yy + 56), radius=13, fill=(236, 230, 220))
        grad_round(img, (x0, yy + 30, x0 + bw_max * frac, yy + 56), 13, PRIMARY_LT, PRIMARY_DK)
        d.text((x0 + bw_max + 14, yy + 43), vlabel, font=fs(18), fill=PRIMARY, anchor="lm")


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
    od.text(((x0 + x1) / 2, y0 + tb_h // 2), file_label, font=fs(20, bold=False), fill=(225, 222, 215), anchor="mm")
    img.alpha_composite(ov)
    sb_w = int((x1 - x0) * 0.205)
    sb = (x0, y0 + tb_h, x0 + sb_w, y1)
    grad_round(img, (sb[0], sb[1], sb[2] + 24, sb[3]), 0, PRIMARY_LT, PRIMARY_DK)
    grad_round(img, (sb[0], y1 - 24, sb[2], y1), 24, PRIMARY_DK, PRIMARY_DK)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    bx = sb[0] + 26
    od.text((bx, sb[1] + 30), "CARD VAULT", font=fs(19), fill=GOLD_HI, anchor="lt")
    od.text((bx, sb[1] + 56), "13-tab system", font=fs(15, bold=False), fill=(170, 200, 192), anchor="lt")
    od.line((sb[0] + 20, sb[1] + 88, sb[2] - 16, sb[1] + 88), fill=(255, 255, 255, 40), width=1)
    list_top = sb[1] + 100
    rowh = (y1 - 24 - list_top) / len(TABS)
    palette = [HIGHLIGHT, GOLD_HI, SURFACE, (150, 200, 190)]
    for i, name in enumerate(TABS):
        ry = list_top + i * rowh
        if i == active_idx:
            od.rounded_rectangle((sb[0] + 12, ry + 2, sb[2] - 10, ry + rowh - 2), radius=8, fill=(255, 255, 255, 235))
            od.rounded_rectangle((sb[0] + 12, ry + 2, sb[0] + 19, ry + rowh - 2), radius=3, fill=GOLD_HI)
            dotc = PRIMARY; txtc = PRIMARY; font = fs(18)
        else:
            dotc = palette[i % len(palette)]; txtc = (214, 226, 222); font = fs(17, bold=False)
        cyr = ry + rowh / 2
        od.ellipse((sb[0] + 30, cyr - 6, sb[0] + 42, cyr + 6), fill=dotc)
        od.text((sb[0] + 56, cyr), name, font=font, fill=txtc, anchor="lm")
    img.alpha_composite(ov)
    cbox = (sb[2] + 1, y0 + tb_h + 4, x1, y1)
    content_fn(img, cbox)


KPIS = [
    ("TOTAL CARDS", "146", "18 unique titles"),
    ("COLLECTION VALUE", "$2,840", "est. market"),
    ("PURCHASE COST", "$1,124", "all-in"),
    ("PROFIT / LOSS", "+$1,716", "unrealized"),
    ("GROWTH", "+153%", "value vs cost"),
    ("TOP CARD", "$620", "Ancient Colossus"),
    ("PHOTOS ON FILE", "67%", "12 of 18"),
    ("CARDS GRADED", "4", "PSA · BGS"),
    ("AWAITING GRADING", "2", "at PSA"),
    ("TRADES DONE", "5", "+$18 net"),
    ("WISHLIST", "40%", "4 of 10"),
    ("HEALTH SCORE", "77%", "collection health"),
]


def _months6():
    today = dt.date.today().replace(day=1)
    y, m = today.year, today.month
    seq = []
    for _ in range(6):
        seq.append(dt.date(y, m, 1).strftime("%b"))
        m -= 1
        if m == 0:
            m = 12; y -= 1
    return list(reversed(seq))


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Collection Dashboard", font=fs(33), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 64), "The Vaulted Collection · 146 cards  ·  your whole vault, automatically organized", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 150, y0 + 26, x1 - pad, y0 + 62), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 75, y0 + 44), "● live", font=fs(18), fill=PRIMARY, anchor="mm")
    gx = x0 + pad; gy = y0 + 98; gw = (x1 - x0 - 2 * pad); gap = 14
    kw = (gw - 5 * gap) / 6; kh = 116
    for i, (lab, val, sub) in enumerate(KPIS):
        r, ci = divmod(i, 6)
        kx = gx + ci * (kw + gap); ky = gy + r * (kh + gap)
        d.rounded_rectangle((kx, ky, kx + kw, ky + kh), radius=12, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((kx + 12, ky, kx + kw - 12, ky + 5), radius=2, fill=GOLD_LT)
        d.text((kx + 14, ky + 16), lab, font=fs(12), fill=ACCENT, anchor="lt")
        vf = fit_font(d, val, kw - 28, 30)
        col = GREEN if val.startswith("+") else PRIMARY
        d.text((kx + 14, ky + 58), val, font=vf, fill=col, anchor="lm")
        d.text((kx + 14, ky + 96), sub, font=fs(12, bold=False), fill=TEXT_MUTED, anchor="lm")
    cy_top = gy + 2 * (kh + gap) + 18
    d.text((gx, cy_top), "VALUE · GROWTH · GRADING", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 34; panel_h = (y1 - panels_y - pad); pw = (gw - 3 * gap) / 4
    # value by rarity donut
    px = gx
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Value by Rarity", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(51, PRIMARY), (33, ACCENT), (5, HIGHLIGHT), (4, SURFACE), (7, (170, 150, 120))], "$2.8K", "est. value")
    legend(d, px + pw * 0.04, panels_y + panel_h - 104, [(PRIMARY, "Secret Rare 51%"), (ACCENT, "Ultra Rare 33%"), (HIGHLIGHT, "Other 16%")], 15, 30)
    # value by set
    px = gx + (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Value by Set", font=fs(17), fill=ACCENT, anchor="lt")
    hbars_money(img, d, (px + 20, panels_y + 56, px + pw - 16, panels_y + panel_h - 20),
                [("Obsidian Flames", 1.0, "$1.1K"), ("Genesis", 0.56, "$620"),
                 ("Astral Dawn", 0.36, "$396"), ("Deep Currents", 0.27, "$298")])
    # 6-month growth
    px = gx + 2 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Collection Growth", font=fs(17), fill=ACCENT, anchor="lt")
    months = _months6()
    vals = [("$1.5K", 0.51), ("$1.7K", 0.61), ("$2.0K", 0.70), ("$2.3K", 0.80), ("$2.5K", 0.89), ("$2.8K", 1.0)]
    hbars_money(img, d, (px + 20, panels_y + 50, px + pw - 16, panels_y + panel_h - 16),
                [(m, f, v) for m, (v, f) in zip(months, vals)])
    # grading results donut
    px = gx + 3 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Grading Results", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(25, PRIMARY), (25, ACCENT), (25, HIGHLIGHT), (25, SURFACE)], "4", "graded")
    legend(d, px + pw * 0.04, panels_y + panel_h - 104,
           [(PRIMARY, "PSA 10 → $480"), (ACCENT, "BGS 9.5 → $620"), (HIGHLIGHT, "PSA 9 → $265")], 15, 30)


def _table(img, cbox, title, subtitle, headers, colf, rows, total_row=None,
           status_col=None, status_map=None, hdr_top=104, money_green=None):
    x0, y0, x1, y1 = cbox; pad = 30
    money_green = money_green or set()
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
                if ci in money_green and sval.startswith("+"):
                    col = GREEN
                elif ci in money_green and sval.startswith("-"):
                    col = DANGER
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
                col = GREEN if str(val).startswith("+") else PRIMARY
                d.text((hx, ry + rh / 2), str(val), font=fs(16), fill=col, anchor=anc)


def content_collection(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 20), "Master Collection", font=fs(31), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 60), "Every card valued — with live profit/loss and photo tracking per card", font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    strip = [("TOTAL CARDS", "146"), ("COLLECTION VALUE", "$2,840"), ("PURCHASE COST", "$1,124"), ("PROFIT / LOSS", "+$1,716")]
    sy = y0 + 92
    sw = (x1 - x0 - 2 * pad - 3 * 14) / 4
    for i, (lab, val) in enumerate(strip):
        sx = x0 + pad + i * (sw + 14)
        d.rounded_rectangle((sx, sy, sx + sw, sy + 90), radius=10, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((sx + 12, sy, sx + sw - 12, sy + 5), radius=2, fill=GOLD_LT)
        d.text((sx + sw / 2, sy + 30), lab, font=fs(15), fill=ACCENT, anchor="mm")
        col = GREEN if val.startswith("+") else PRIMARY
        d.text((sx + sw / 2, sy + 64), val, font=fserif(28), fill=col, anchor="mm")
    rows = [
        ("Ancient Colossus", "Genesis", "Secret · BGS 9.5", "1", "$210", "$620", "+$410"),
        ("Golden Phoenix (Alt)", "Obsidian Flames", "Secret · PSA 10", "1", "$150", "$480", "+$330"),
        ("Emberfang Dragon", "Obsidian Flames", "Ultra Rare", "1", "$120", "$340", "+$220"),
        ("Twin Flames Duo", "Obsidian Flames", "Ultra · PSA 9", "1", "$88", "$265", "+$177"),
        ("Celestial Guardian", "Astral Dawn", "Secret Rare · JP", "1", "$95", "$210", "+$115"),
        ("Royal Kraken", "Deep Currents", "Secret · PSA 8", "1", "$70", "$150", "+$80"),
        ("Shadow Assassin", "Midnight Veil", "Ultra Rare", "1", "$75", "$145", "+$70"),
        ("Thunderclap Roc", "Tempest Rising", "Ultra Rare", "1", "$55", "$96", "+$41"),
        ("Tidecaller Leviathan", "Deep Currents", "Ultra Rare", "1", "$60", "$88", "+$28"),
        ("Arcane Scholar", "Astral Dawn", "Uncommon", "24", "$20", "$72", "+$52"),
        ("Starlight Fairy", "Astral Dawn", "Common", "66", "$25", "$66", "+$41"),
        ("Mystic Fountain", "Deep Currents", "Uncommon", "30", "$18", "$60", "+$42"),
    ]
    ty = sy + 112
    tx0, tx1 = x0 + pad, x1 - pad
    headers = ["CARD", "SET", "RARITY / GRADE", "QTY", "PAID", "VALUE", "P/L"]
    colf = [0.0, 0.22, 0.40, 0.60, 0.68, 0.78, 0.90]
    colx = [tx0 + (tx1 - tx0) * f for f in colf]
    hdr_h = 40
    grad_round(img, (tx0, ty, tx1, ty + hdr_h), 8, PRIMARY_LT, PRIMARY_DK)
    for i, h in enumerate(headers):
        anc = "lm" if i == 0 else "mm"
        d.text((colx[i] + (14 if i == 0 else 0), ty + hdr_h / 2), h, font=fs(15), fill=WHITE, anchor=anc)
    rh = (y1 - pad - (ty + hdr_h)) / (len(rows) + 1)
    for i, row in enumerate(rows):
        ry = ty + hdr_h + i * rh
        if i % 2:
            d.rectangle((tx0, ry, tx1, ry + rh), fill=ROW_ALT)
        for ci, val in enumerate(row):
            anc = "lm" if ci == 0 else "mm"
            hx = colx[ci] + (14 if ci == 0 else 0)
            if ci == 6:
                col = GREEN
                f = fs(15)
            elif ci == 0:
                col = PRIMARY; f = fs(16)
            else:
                col = TEXT; f = fs(15, bold=False)
            d.text((hx, ry + rh / 2), val, font=f, fill=col, anchor=anc)
    ry = ty + hdr_h + len(rows) * rh
    d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
    for ci, val in [(0, "TOTAL — FULL COLLECTION"), (3, "146"), (4, "$1,124"), (5, "$2,840"), (6, "+$1,716")]:
        anc = "lm" if ci == 0 else "mm"
        hx = colx[ci] + (14 if ci == 0 else 0)
        col = GREEN if str(val).startswith("+") else PRIMARY
        d.text((hx, ry + rh / 2), str(val), font=fs(16), fill=col, anchor=anc)


def content_grading(img, cbox):
    rows = [
        ("Golden Phoenix (Alt Art)", "PSA", "PSA 10", "$150", "$480", "+$330"),
        ("Ancient Colossus", "BGS", "BGS 9.5", "$210", "$620", "+$410"),
        ("Twin Flames Duo", "PSA", "PSA 9", "$88", "$265", "+$177"),
        ("Royal Kraken", "PSA", "PSA 8", "$70", "$150", "+$80"),
        ("Emberfang Dragon (Holo)", "PSA", "Pending", "$340", "—", "—"),
        ("Celestial Guardian", "PSA", "Pending", "$210", "—", "—"),
    ]
    _table(img, cbox, "Grading Center",
           "Every submission tracked — grades in, and +$997 of value added so far",
           ["CARD", "COMPANY", "GRADE", "VALUE BEFORE", "VALUE AFTER", "GAIN"],
           [0.0, 0.28, 0.42, 0.58, 0.74, 0.89], rows,
           status_col=2, status_map={"PSA 10": (MINT_BG, PRIMARY), "BGS 9.5": (MINT_BG, PRIMARY),
                                      "PSA 9": (MINT_BG, PRIMARY), "PSA 8": (MINT_BG, PRIMARY),
                                      "Pending": (WARN_BG, ACCENT)},
           money_green={5})


def content_trades(img, cbox):
    rows = [
        ("Alex M.", "Storm Herald + Hoard", "Lunar Priestess + lot", "$45", "$52", "+$7"),
        ("Sam K.", "Tidecaller Leviathan", "Thunderclap Roc", "$88", "$96", "+$8"),
        ("League — Riley", "Uncommon lot ×6", "Storm Herald (Holo)", "$12", "$15", "+$3"),
        ("Jordan P.", "Inferno Titan ×2", "Verdant Sprite ×2", "$30", "$26", "-$4"),
        ("Riley T.", "Frostbite Wolf ×2", "Mystic Fountain lot", "$20", "$24", "+$4"),
    ]
    _table(img, cbox, "Trade Tracker",
           "Win every swap — value out vs value in, profit/loss per trade",
           ["PARTNER", "CARDS GIVEN", "CARDS RECEIVED", "VALUE OUT", "VALUE IN", "TRADE +/-"],
           [0.0, 0.17, 0.41, 0.64, 0.76, 0.89], rows, money_green={5},
           total_row=("TOTAL", "", "", "$195", "$213", "+$18"))


# ---------- renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    card_crest(img, SIZE // 2, 132, r=56)
    pill(img, SIZE // 2, 256, "THE ULTIMATE CARD COLLECTION, VALUE & INVENTORY SYSTEM", font=fs(25), pad_x=42, pad_y=20)
    wordmark(img, SIZE // 2, 400, "TRADING CARD COMMAND CENTER", 96, max_w=1840)
    gold_divider(img, SIZE // 2, 500, width=520)
    tc(d, (SIZE // 2, 550), "Collection, value, grading, trades & wishlist — know what your cards are worth, always.",
       fs(24, bold=False), (224, 213, 190))
    chips = [("13", "POWERFUL TABS"), ("AUTO", "VALUE + P/L"), ("2-in-1", "EXCEL + SHEETS")]
    cw = 420
    total = len(chips) * cw + (len(chips) - 1) * 32
    startx = (SIZE - total) // 2 + cw // 2
    for i, (b, s) in enumerate(chips):
        stat_chip(img, startx + i * (cw + 32), 704, b, s, w=cw)
    app_window(img, (70, 800, SIZE - 70, 1900), 0, content_dashboard)
    pill(img, SIZE // 2, SIZE - 52, "13 SHEETS · INSTANT DOWNLOAD · EXCEL + GOOGLE SHEETS",
         font=fs(33), pad_x=50, pad_y=24, star=True, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_inside(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 120, "EVERYTHING INSIDE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 238), "13 Powerful, Connected Tabs", fserif(58), WHITE)
    gold_divider(img, SIZE // 2, 308, width=520)
    tc(d, (SIZE // 2, 352), "Not an inventory sheet — a complete collection management system for serious collectors",
       fs(24, bold=False), (226, 214, 190))
    cards = [
        ("Collection Dashboard", "12 KPIs + live charts"), ("Master Collection", "every card, valued"),
        ("Purchase Tracker", "true all-in cost"), ("Sales Tracker", "fees & net proceeds"),
        ("Trade Tracker", "win every swap"), ("Grading Center", "PSA · BGS · CGC"),
        ("Value Analytics", "growth over time"), ("Wishlist", "grails & targets"),
        ("Duplicates Manager", "trade · sell · gift"), ("Deck Builder", "win/loss record"),
        ("Card Image Vault", "upload card photos"), ("Analytics", "collection health score"),
        ("Settings", "any card game"),
    ]
    cols = 4
    margin = 90
    gx, gy = 22, 22
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    top = 440
    rows_n = 4
    ch = (SIZE - top - 60 - (rows_n - 1) * gy) // rows_n
    for i, (title, sub) in enumerate(cards):
        r, ccol = divmod(i, cols)
        x = margin + ccol * (cw + gx); y = top + r * (ch + gy)
        shadow(img, (x, y, x + cw, y + ch), 14, 13, 42, 9)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        od.rounded_rectangle((x, y, x + cw, y + ch), radius=14, fill=WHITE, outline=(232, 224, 208), width=2)
        od.rounded_rectangle((x, y, x + 7, y + ch), radius=3, fill=GOLD_LT)
        od.rectangle((x + 3, y, x + 7, y + ch), fill=GOLD_LT)
        cyc = y + ch // 2; bx = x + 42
        od.ellipse((bx - 23, cyc - 23, bx + 23, cyc + 23), fill=PRIMARY)
        od.text((bx, cyc), str(i + 1), font=fs(21), fill=GOLD_HI, anchor="mm")
        od.text((x + 82, cyc - 17), title, font=fs(20), fill=PRIMARY, anchor="lm")
        od.text((x + 82, cyc + 22), sub, font=fs(15, bold=False), fill=TEXT_MUTED, anchor="lm")
        img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_dashboard(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "COLLECTION DASHBOARD", font=fs(34), pad_x=50, pad_y=22)
    tc(d, (SIZE // 2, 232), "Your Whole Collection, At A Glance", fserif(50), WHITE)
    tc(d, (SIZE // 2, 300), "12 live KPIs + value, growth & grading charts — all auto-updating",
       fs(24, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 0, content_dashboard)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_collection(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=320)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 112, "MASTER COLLECTION", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 230), "Every Card Valued — Profit On Every Row", fserif(46), WHITE)
    tc(d, (SIZE // 2, 290), "Set, rarity, condition, storage location, photo status & live P/L per card",
       fs(23, bold=False), (226, 214, 190))
    app_window(img, (80, 360, SIZE - 80, SIZE - 70), 1, content_collection)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_value(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "GRADING & TRADES", font=fs(36), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "Slabs Add Value. Trades Keep Score.", fserif(46), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1150), 5, content_grading)
    app_window(img, (60, 1180, SIZE - 60, SIZE - 60), 4, content_trades)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=400)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 130, "WORKS EVERYWHERE", font=fs(38), pad_x=54, pad_y=22)
    tc(d, (SIZE // 2, 250), "Excel · Google Sheets · Mobile", fserif(56), WHITE)
    tc(d, (SIZE // 2, 320), "Check a card's value at the card show — your whole collection in your pocket",
       fs(23, bold=False), (226, 214, 190))
    px, py = SIZE // 2, 1300
    pw, ph = 640, 1230
    phone = (px - pw // 2, py - ph // 2, px + pw // 2, py + ph // 2)
    shadow(img, phone, 64, 50, 110, 24)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle(phone, radius=64, fill=(26, 26, 30))
    bez = 22
    screen = (phone[0] + bez, phone[1] + bez + 30, phone[2] - bez, phone[3] - bez - 30)
    od.rounded_rectangle(screen, radius=44, fill=BG)
    od.rounded_rectangle((px - 95, phone[1] + 16, px + 95, phone[1] + 50), radius=18, fill=(14, 14, 18))
    img.alpha_composite(ov)
    sx0, sy0, sx1, sy1 = screen
    grad_round(img, (sx0, sy0, sx1, sy0 + 110), 44, PRIMARY_LT, PRIMARY_DK)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rectangle((sx0, sy0 + 106, sx1, sy0 + 110), fill=GOLD_LT)
    od.text(((sx0 + sx1) // 2, sy0 + 56), "Card Vault", font=fserif(34), fill=GOLD_HI, anchor="mm")
    y = sy0 + 150
    cards = [("COLLECTION VALUE", "$2,840", PRIMARY), ("PROFIT / LOSS", "+$1,716", GREEN),
             ("TOP CARD", "$620", ACCENT), ("HEALTH SCORE", "77%", PRIMARY)]
    for lab, val, col in cards:
        cb = (sx0 + 30, y, sx1 - 30, y + 135)
        od.rounded_rectangle(cb, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.rounded_rectangle((cb[0] + 20, y, cb[2] - 20, y + 5), radius=2, fill=GOLD_LT)
        od.text((cb[0] + 26, y + 34), lab, font=fs(22), fill=ACCENT, anchor="lt")
        vf = fit_font(od, val, sx1 - sx0 - 110, 46)
        od.text((cb[0] + 26, y + 94), val, font=vf, fill=col, anchor="lm")
        y += 152
    od.text((sx0 + 40, y + 16), "COLLECTOR TASKS", font=fs(22), fill=ACCENT, anchor="lt")
    y += 52
    for lab, state in [("Photograph new pickups (67%)", False), ("List Frostbite dupes for sale", False),
                       ("Check PSA submission status", True), ("Update market values", True)]:
        col = HIGHLIGHT if state else BG
        od.ellipse((sx0 + 40, y + 6, sx0 + 78, y + 44), fill=col, outline=PRIMARY, width=3)
        if state:
            od.text((sx0 + 59, y + 24), "✓", font=fs(24), fill=PRIMARY, anchor="mm")
        od.text((sx0 + 96, y + 24), lab, font=fs(22, bold=False), fill=TEXT, anchor="lm")
        y += 64
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def main():
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("01_hero.png", render_hero),
        ("02_inside.png", render_inside),
        ("03_dashboard.png", render_dashboard),
        ("04_collection.png", render_collection),
        ("05_value.png", render_value),
        ("06_mobile.png", render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()
