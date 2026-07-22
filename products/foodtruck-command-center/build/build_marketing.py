"""Marketing image set for Food Truck Command Center™ (6 images, 2000x2000)."""
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

TABS = ["Dashboard", "Menu & Cost", "Events", "Break-Even", "Daily Sales", "Commissary & Overhead",
        "Inventory & Par", "Fuel & Mileage", "Permits", "Supplies", "Bookings", "Cash & Tips", "Settings"]

FILE_LABEL = "Food_Truck_Command_Center.xlsx — Rolling Smoke BBQ · Dana"


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


def truck_crest(c, cx, cy, r=56, glow=True):
    """Brand crest — a food truck with a serving window."""
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    c.alpha_composite(ov)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    # body
    bx0, by0, bx1, by1 = cx - r * 0.60, cy - r * 0.34, cx + r * 0.58, cy + r * 0.22
    d.rounded_rectangle((bx0, by0, bx1, by1), radius=8, fill=GOLD_HI, outline=(150, 120, 70))
    # cab (front, right)
    d.polygon([(bx1, by0 + r * 0.10), (bx1 + r * 0.22, by0 + r * 0.10), (bx1 + r * 0.30, by1), (bx1, by1)], fill=GOLD_HI, outline=(150, 120, 70))
    # serving window
    d.rounded_rectangle((bx0 + r * 0.10, by0 + r * 0.12, bx1 - r * 0.10, cy - r * 0.02), radius=4, fill=PRIMARY_DK)
    # awning
    d.polygon([(bx0 + r * 0.06, by0 + r * 0.12), (bx1 - r * 0.06, by0 + r * 0.12), (bx1 - r * 0.02, by0 - r * 0.06), (bx0 + r * 0.02, by0 - r * 0.06)], fill=(150, 120, 70))
    # wheels
    for wx in (cx - r * 0.34, cx + r * 0.40):
        d.ellipse((wx - r * 0.13, by1 - r * 0.02, wx + r * 0.13, by1 + r * 0.24), fill=PRIMARY_DK, outline=GOLD_HI)
    c.alpha_composite(ov)


book_crest = truck_crest


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


def vbars(img, d, box, items, maxv, color=(PRIMARY_LT, PRIMARY)):
    x0, y0, x1, y1 = box; n = len(items); bw = (x1 - x0) / n * 0.62; gap = (x1 - x0) / n
    for i, (lab, val) in enumerate(items):
        bx = x0 + gap * i + (gap - bw) / 2
        bh = (y1 - y0 - 40) * (val / maxv)
        grad_round(img, (bx, y1 - 40 - bh, bx + bw, y1 - 40), 6, color[0], color[1])
        d.text((bx + bw / 2, y1 - 20), lab, font=fs(13, bold=False), fill=TEXT_MUTED, anchor="mm")
        d.text((bx + bw / 2, y1 - 52 - bh), f"${val/1000:.1f}k" if val >= 1000 else f"${val}", font=fs(13), fill=PRIMARY, anchor="mb")


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
    od.text((bx, sb[1] + 26), "TRUCK", font=fs(16), fill=GOLD_HI, anchor="lt")
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
    ("EVENTS", "8", "this month"),
    ("TOTAL SALES", "$15,230", "all gigs"),
    ("TOTAL PROFIT", "$7,566", "after costs"),
    ("AVG PROFIT/EVENT", "$946", "per gig"),
    ("FOOD COST", "29%", "of sales"),
    ("AVG SALES/EVENT", "$1,904", "per gig"),
    ("BREAK-EVEN", "2.3", "events/mo"),
    ("OVERHEAD/MO", "$2,140", "fixed cost"),
    ("TOP EVENT", "Corporate", "$1,695 net"),
    ("BEST SALES", "Festival", "$3,400"),
    ("PERMITS OK", "4/5", "1 to renew"),
    ("TRUCK SCORE", "93%", "on track"),
]

NETS = [("Downtown", 685), ("Brewery", 1095), ("Farmers", 421), ("Office", 600),
        ("Festival", 1480), ("Corporate", 1695), ("Sat Night", 940), ("Sunday", 650)]


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Truck Dashboard", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 64), "Rolling Smoke BBQ · Dana  ·  know your profit per event & cover your overhead", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 200, y0 + 26, x1 - pad, y0 + 62), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 100, y0 + 44), "● $7,566 profit", font=fs(15), fill=PRIMARY, anchor="mm")
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
    d.text((gx, cy_top), "TRUCK HEALTH · NET PROFIT BY EVENT · BOOKINGS", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 34; panel_h = (y1 - panels_y - pad); pw = (gw - 3 * gap) / 4
    px = gx
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Truck Health", font=fs(17), fill=ACCENT, anchor="lt")
    hbars(img, d, (px + 20, panels_y + 50, px + pw - 16, panels_y + panel_h - 16),
          [("Profit margin", 1.0, "100%"), ("Food cost on target", 1.0, "100%"), ("Events booked", 0.80, "80%"),
           ("Permits current", 0.80, "80%"), ("Profitable events", 1.0, "100%"), ("Overhead covered", 1.0, "100%")],
          color=(GOLD_HI, GOLD))
    px = gx + (pw + gap)
    d.rounded_rectangle((px, panels_y, px + 2 * pw + gap, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Net Profit by Event", font=fs(17), fill=ACCENT, anchor="lt")
    vbars(img, d, (px + 24, panels_y + 60, px + 2 * pw + gap - 20, panels_y + panel_h - 10), NETS, 1800)
    px = gx + 3 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Truck Score", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.5, panels_y + panel_h * 0.54, min(panel_h * 0.30, pw * 0.30),
          [(93, PRIMARY), (7, SURFACE)], "93%", "on track")


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
                d.rounded_rectangle((hx - 66, ry + rh / 2 - 15, hx + 66, ry + rh / 2 + 15), radius=14, fill=bg)
                d.text((hx, ry + rh / 2), sval, font=fs(14), fill=fg, anchor="mm")
            else:
                col = PRIMARY if ci == 0 else TEXT
                d.text((hx, ry + rh / 2), sval, font=fs(15) if ci == 0 else fs(15, bold=False), fill=col, anchor=anc)
    if total_row:
        ry = ty + hdr_h + len(rows) * rh
        d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
        for ci, val in enumerate(total_row):
            anc = "lm" if ci == 0 else "mm"
            hx = colx[ci] + (14 if ci == 0 else 0)
            if val != "":
                d.text((hx, ry + rh / 2), str(val), font=fs(15), fill=PRIMARY, anchor=anc)


def content_events(img, cbox):
    rows = [
        ("Corporate Catering", "$2,800", "$700", "$1,695"),
        ("Music Festival", "$3,400", "$1,020", "$1,480"),
        ("Brewery Friday", "$2,100", "$630", "$1,095"),
        ("Saturday Night Lot", "$1,900", "$570", "$940"),
        ("Downtown Lunch", "$1,450", "$435", "$685"),
        ("Sunday Brunch", "$1,350", "$405", "$650"),
    ]
    _table(img, cbox, "Event P&L",
           "Every gig — sales minus food, fuel, fees & staff = the net profit of each event",
           ["EVENT", "SALES", "FOOD", "NET PROFIT"],
           [0.0, 0.42, 0.60, 0.82], rows,
           total_row=("TOTAL / MONTH", "$15,230", "$4,429", "$7,566"))


def content_breakeven(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Break-Even", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "The sales you need to cover your fixed costs — know it before you park.",
           font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    # big number
    by = y0 + 130
    grad_round(img, (x0 + pad, by, x1 - pad, by + 200), 16, (70, 200, 165), PRIMARY, outline=GOLD_LT, width=3)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((x0 + pad + 40, by + 56), "BREAK-EVEN", font=fs(22), fill=GOLD_HI, anchor="lm")
    od.text((x0 + pad + 40, by + 130), "2.3 events / month", font=fserif(52), fill=WHITE, anchor="lm")
    od.text((x1 - pad - 40, by + 100), "covers $2,140 overhead", font=fs(19, bold=False), fill=(220, 235, 228), anchor="rm")
    img.alpha_composite(ov)
    # rows
    rows = [("Monthly fixed overhead", "$2,140"), ("Avg net profit / event", "$946"),
            ("Break-even events / month", "2.3"), ("Avg sales / event", "$1,904"),
            ("Events booked this month", "8"), ("Gigs past break-even → profit", "5.7")]
    ry = by + 250; rh = (y1 - pad - ry) / len(rows)
    for i, (lab, val) in enumerate(rows):
        yy = ry + i * rh
        if i % 2:
            d.rounded_rectangle((x0 + pad, yy, x1 - pad, yy + rh - 8), radius=8, fill=ROW_ALT)
        d.text((x0 + pad + 20, yy + rh / 2 - 4), lab, font=fs(20, bold=False), fill=TEXT, anchor="lm")
        d.text((x1 - pad - 24, yy + rh / 2 - 4), val, font=fs(24), fill=PRIMARY, anchor="rm")


def content_menu(img, cbox):
    rows = [
        ("Brisket Plate", "$5.50", "$18.00", "31%"),
        ("Combo Plate", "$6.20", "$20.00", "31%"),
        ("Pulled Pork Sandwich", "$2.80", "$12.00", "23%"),
        ("Smoked Wings", "$3.20", "$11.00", "29%"),
        ("Loaded Fries", "$2.10", "$9.00", "23%"),
        ("Mac & Cheese", "$1.40", "$6.00", "23%"),
    ]
    _table(img, cbox, "Menu & Cost",
           "Cost & price every item — keep a healthy margin on the whole board",
           ["ITEM", "COST", "PRICE", "FOOD %"],
           [0.0, 0.46, 0.64, 0.85], rows)


def content_bookings(img, cbox):
    rows = [
        ("Farmers Market", "in 4 days", "Town Square", "Confirmed"),
        ("Brewery Anniversary", "in 6 days", "Hops & Barrel", "Confirmed"),
        ("Corporate Lunch", "in 11 days", "TechPark", "Confirmed"),
        ("Food Truck Rally", "in 18 days", "Riverside Lot", "Pending"),
        ("Wedding Catering", "in 26 days", "Oak Hollow", "Confirmed"),
    ]
    _table(img, cbox, "Bookings",
           "Upcoming gigs on the calendar — deposits, locations & confirmation status",
           ["EVENT", "WHEN", "LOCATION", "STATUS"],
           [0.0, 0.34, 0.52, 0.82], rows,
           status_col=3, status_map={"Confirmed": (MINT_BG, PRIMARY), "Pending": (WARN_BG, ACCENT)})


# ---------- renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    truck_crest(img, SIZE // 2, 132, r=56)
    pill(img, SIZE // 2, 256, "THE COMPLETE MOBILE-FOOD BUSINESS SYSTEM", font=fs(20), pad_x=40, pad_y=20)
    wordmark(img, SIZE // 2, 400, "FOOD TRUCK", 96, max_w=1600)
    tc(d, (SIZE // 2, 480), "COMMAND CENTER", fs(30), GOLD_HI)
    gold_divider(img, SIZE // 2, 524, width=520)
    tc(d, (SIZE // 2, 568), "Profit per event, break-even, bookings, permits & inventory — your whole truck, one system.",
       fs(20, bold=False), (224, 213, 190))
    chips = [("14", "CONNECTED TABS"), ("12", "PRINTABLE PAGES"), ("EVENT", "P&L ENGINE")]
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
    tc(d, (SIZE // 2, 352), "Not a sales log — a complete know-your-numbers mobile-food system",
       fs(24, bold=False), (226, 214, 190))
    cards = [
        ("Start Here", "how it all works"), ("Truck Dashboard", "profit, break-even & score"),
        ("Menu & Cost", "cost & price every item"), ("Event P&L", "net profit per gig"),
        ("Break-Even", "the sales to cover fixed costs"), ("Daily Sales", "the pulse of the week"),
        ("Commissary & Overhead", "your monthly fixed costs"), ("Inventory & Par", "restock before service"),
        ("Fuel & Mileage", "every fill-up"), ("Permits", "never a lapsed license"),
        ("Supplies", "the next restock run"), ("Bookings", "the right gigs, scheduled"),
        ("Cash & Tips", "balance the drawer"), ("Settings", "set it once"),
        ("+ 12 Printable Pages", "for the truck"), ("Live Truck Score", "it all rolls up"),
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


def render_events(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "PROFIT PER GIG", font=fs(30), pad_x=48, pad_y=22)
    tc(d, (SIZE // 2, 232), "Did That Event Make Money?", fserif(44), WHITE)
    tc(d, (SIZE // 2, 300), "Sales minus food, fuel, fees & staff — the net profit of every gig, ranked",
       fs(22, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 2, content_events)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_breakeven(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "COVER YOUR OVERHEAD", font=fs(30), pad_x=48, pad_y=22)
    tc(d, (SIZE // 2, 232), "Know Your Break-Even", fserif(46), WHITE)
    tc(d, (SIZE // 2, 300), "How many gigs it takes to cover rent, insurance & the loan — then it's all profit",
       fs(22, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 3, content_breakeven)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_ops(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "PRICE IT · BOOK IT", font=fs(32), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "The Menu & the Schedule", fserif(44), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1150), 1, content_menu)
    app_window(img, (60, 1180, SIZE - 60, SIZE - 60), 10, content_bookings)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_printables(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=470)
    d = ImageDraw.Draw(img)
    truck_crest(img, SIZE // 2, 110, r=48)
    pill(img, SIZE // 2, 214, "PLUS 12 PRINTABLE PAGES", font=fs(30), pad_x=48, pad_y=20)
    tc(d, (SIZE // 2, 300), "Print, Fill & Keep — For the Truck", fserif(44), WHITE)
    gold_divider(img, SIZE // 2, 362, width=480)
    tc(d, (SIZE // 2, 404), "A matching print-ready pack — event P&L, prep list, cash-&-tips & permit tracker",
       fs(22, bold=False), (226, 214, 190))
    labels = ["Event P&L Sheet", "Break-Even", "Daily Sales Log", "Menu & Cost Card",
              "Inventory & Par", "Prep List", "Fuel & Mileage", "Permit Tracker",
              "Shopping List", "Bookings Calendar", "Cash & Tips", "Monthly P&L"]
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
        d.text((x + cw // 2, y + ch + 30), labels[i], font=fs(17), fill=PRIMARY, anchor="mm")
    pill(img, SIZE // 2, SIZE - 46, "PRINT AT HOME · US LETTER · READY IN ONE CLICK",
         font=fs(28), pad_x=44, pad_y=22, grad=(PRIMARY_LT, PRIMARY_DK))
    img.convert("RGB").save(out, "PNG", optimize=True)


def main():
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("01_hero.png", render_hero),
        ("02_inside.png", render_inside),
        ("03_events.png", render_events),
        ("04_breakeven.png", render_breakeven),
        ("05_ops.png", render_ops),
        ("06_printables.png", render_printables),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()
