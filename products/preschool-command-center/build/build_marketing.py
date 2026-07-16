"""Marketing image set for Preschool Command Center™ (6 images, 2000x2000).

  01_hero.png       - branded hero + live preschool dashboard
  02_inside.png     - "everything inside — 17 connected tabs"
  03_skills.png     - the whole-child skills & milestones tracker (distinct sheet)
  04_abc.png        - ABC & 123 mastery (the letter-tile shot)
  05_plan.png       - activity planner + kindergarten readiness
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
CORAL = (224, 122, 95)
SKY = (108, 164, 196)
LEAF = (70, 150, 80)
SIZE = 2000

SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"

TABS = ["Dashboard", "Child Profiles", "Skills & Milestones", "ABC & 123", "Weekly Themes",
        "Daily Rhythm", "Activity Planner", "Read-Aloud Log", "Arts & Sensory", "Field Trips",
        "Attendance", "Portfolio", "Supplies", "Kindergarten Readiness", "Goals", "Settings"]

FILE_LABEL = "Preschool_Command_Center.xlsx — The Bennett Family · 2026–2027 · Millie & Owen"


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


def block_crest(c, cx, cy, r=56, glow=True):
    """Brand crest — a little stack of A-B-C toy blocks: the icon of playful early learning."""
    if glow:
        radial_glow(c, cx, cy, int(r * 2.1), GOLD_HI, 90)
    grad_round(c, (cx - r, cy - r, cx + r, cy + r), 22, PRIMARY_LT, PRIMARY_DK, outline=GOLD_LT, width=4)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rounded_rectangle((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), radius=16, outline=GOLD_HI, width=2)
    c.alpha_composite(ov)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    bs = r * 0.52          # block size
    gap = r * 0.10
    lf = fserif(int(bs * 0.68))
    # bottom row: A (gold), B (mint); top: C (surface) centered
    bottom_y = cy + r * 0.06
    bx_a = cx - bs - gap * 0.5
    bx_b = cx + gap * 0.5
    top_y = bottom_y - bs - gap
    top_x = cx - bs * 0.5
    blocks = [(bx_a, bottom_y, GOLD_HI, "A"), (bx_b, bottom_y, HIGHLIGHT, "B"), (top_x, top_y, SURFACE, "C")]
    for bx, by, col, ch in blocks:
        d.rounded_rectangle((bx, by, bx + bs, by + bs), radius=int(bs * 0.18), fill=col, outline=(150, 120, 70), width=2)
        d.text((bx + bs / 2, by + bs / 2), ch, font=lf, fill=PRIMARY, anchor="mm")
    c.alpha_composite(ov)


# keep the name used by the detail-image module
book_crest = block_crest


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
    od.text((bx, sb[1] + 26), "PRESCHOOL", font=fs(16), fill=GOLD_HI, anchor="lt")
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
    ("CHILDREN", "2", "in the system"),
    ("PRESCHOOL DAYS", "96", "of 150"),
    ("LETTERS KNOWN", "24", "of 26"),
    ("SKILLS MASTERED", "25", "of 35"),
    ("THEMES DONE", "7", "of 14"),
    ("ACTIVITIES DONE", "62%", "checked off"),
    ("BOOKS READ", "38", "this year"),
    ("FIELD TRIPS", "7", "this year"),
    ("SUPPLIES SPENT", "$350", "of $400"),
    ("READY FOR K", "75%", "checklist"),
    ("GOALS", "71%", "avg progress"),
    ("READY SCORE", "70%", "blended"),
]


def content_dashboard(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "Preschool Dashboard", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 64), "The Bennett Family · Millie (4) & Owen (3)  ·  play with a plan, and watch them grow", font=fs(19, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 170, y0 + 26, x1 - pad, y0 + 62), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 85, y0 + 44), "● 70% ready", font=fs(16), fill=PRIMARY, anchor="mm")
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
    d.text((gx, cy_top), "GROWING WELL · THEMES · COMING UP", font=fs(20), fill=ACCENT, anchor="lt")
    panels_y = cy_top + 34; panel_h = (y1 - panels_y - pad); pw = (gw - 3 * gap) / 4
    px = gx
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Growing Well?", font=fs(17), fill=ACCENT, anchor="lt")
    hbars(img, d, (px + 20, panels_y + 50, px + pw - 16, panels_y + panel_h - 16),
          [("Letters", 0.92, "92%"), ("Skills", 0.71, "71%"), ("Themes", 0.50, "50%"),
           ("Activities", 0.62, "62%"), ("Readiness", 0.75, "75%")], color=(GOLD_HI, GOLD))
    px = gx + (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Themes This Year", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.42, panels_y + panel_h * 0.52, min(panel_h * 0.28, pw * 0.27),
          [(50, PRIMARY), (7, ACCENT), (43, SURFACE)], "7/14", "done")
    legend(d, px + pw * 0.04, panels_y + panel_h - 104, [(PRIMARY, "Done 7"), (ACCENT, "In progress 1"), (SURFACE, "Upcoming 6")], 15, 30)
    px = gx + 2 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Coming Up", font=fs(17), fill=ACCENT, anchor="lt")
    tasks = [("Weather week", WARN_BG), ("Under the Sea", MINT_BG), ("Aquarium field trip", MINT_BG),
             ("Outer Space", MINT_BG), ("Dinosaurs", MINT_BG), ("Spring & bugs", MINT_BG)]
    ty = panels_y + 58
    for lab, dot in tasks:
        d.ellipse((px + 20, ty + 6, px + 40, ty + 26), fill=dot, outline=PRIMARY, width=2)
        d.text((px + 54, ty + 16), lab, font=fs(16, bold=False), fill=TEXT, anchor="lm")
        ty += 46
    px = gx + 3 * (pw + gap)
    d.rounded_rectangle((px, panels_y, px + pw, panels_y + panel_h), radius=12, fill=WHITE, outline=GRID, width=2)
    d.text((px + 16, panels_y + 14), "Ready-for-K Score", font=fs(17), fill=ACCENT, anchor="lt")
    donut(d, px + pw * 0.5, panels_y + panel_h * 0.54, min(panel_h * 0.30, pw * 0.30),
          [(70, PRIMARY), (30, SURFACE)], "70%", "ready")


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
                d.text((hx, ry + rh / 2), sval, font=fs(16) if ci == 0 else fs(15, bold=False), fill=col, anchor=anc)
    if total_row:
        ry = ty + hdr_h + len(rows) * rh
        d.rectangle((tx0, ry, tx1, ry + rh), fill=SURFACE)
        for ci, val in enumerate(total_row):
            anc = "lm" if ci == 0 else "mm"
            hx = colx[ci] + (14 if ci == 0 else 0)
            if val != "":
                d.text((hx, ry + rh / 2), str(val), font=fs(16), fill=PRIMARY, anchor=anc)


def content_skills(img, cbox):
    rows = [
        ("Fine Motor", "Tripod crayon grip", "Millie", "Mastered"),
        ("Gross Motor", "Throws & catches a ball", "Millie", "Emerging"),
        ("Language", "Names 8+ colors", "Millie", "Mastered"),
        ("Pre-Math", "Counts to 20", "Millie", "Mastered"),
        ("Pre-Reading", "Retells a simple story", "Millie", "Emerging"),
        ("Social-Emotional", "Takes turns", "Millie", "Mastered"),
        ("Self-Help", "Cleans up toys", "Owen", "Emerging"),
    ]
    _table(img, cbox, "Skills & Milestones",
           "The whole-child picture — 25 of 35 skills mastered across 7 domains",
           ["DOMAIN", "SKILL / MILESTONE", "CHILD", "STATUS"],
           [0.0, 0.24, 0.62, 0.82], rows,
           status_col=3, status_map={"Mastered": (MINT_BG, PRIMARY), "Emerging": (WARN_BG, ACCENT),
                                     "Not Yet": ((235, 230, 222), TEXT_MUTED)})


def content_activities(img, cbox):
    rows = [
        ("Colors & Shapes", "Rainbow AB pattern beads", "Fine Motor", "Done"),
        ("Pumpkins", "Pumpkin seed counting", "Pre-Math", "Done"),
        ("Farm Animals", "Animal-sound matching", "Language", "Done"),
        ("Weather", "Cloud-dough sensory play", "Fine Motor", "Planned"),
        ("Under the Sea", "Fish counting clip cards", "Pre-Math", "Planned"),
        ("Outer Space", "Build a cardboard rocket", "Gross Motor", "Planned"),
    ]
    _table(img, cbox, "Activity Planner",
           "Play with a purpose — every activity tied to a theme & a skill.  62% done",
           ["THEME", "ACTIVITY", "SKILL DOMAIN", "STATUS"],
           [0.0, 0.24, 0.60, 0.83], rows,
           status_col=3, status_map={"Done": (MINT_BG, PRIMARY), "Planned": (WARN_BG, ACCENT),
                                     "Skipped": (RED_BG, DANGER)})


def content_readiness(img, cbox):
    rows = [
        ("Writes first name", "Ready"),
        ("Recognizes 20+ letters", "Ready"),
        ("Says most letter sounds", "Not yet"),
        ("Counts to 20", "Ready"),
        ("Knows 8+ colors & shapes", "Ready"),
        ("Cuts with scissors on a line", "Not yet"),
    ]
    _table(img, cbox, "Kindergarten Readiness",
           "The gentle guide, never a test — 12 of 16 skills ready, a milestone at a time",
           ["READINESS SKILL", "READY?"],
           [0.0, 0.74], rows,
           status_col=1, status_map={"Ready": (MINT_BG, PRIMARY), "Not yet": (WARN_BG, ACCENT)})


def content_abc(img, cbox):
    x0, y0, x1, y1 = cbox; pad = 30
    d = ImageDraw.Draw(img); d.rectangle(cbox, fill=BG)
    d.text((x0 + pad, y0 + 22), "ABC & 123 Mastery", font=fs(32), fill=PRIMARY, anchor="lt")
    d.text((x0 + pad, y0 + 62), "Letter by letter — recognizes, says the sound & writes it.  24 of 26 letters known",
           font=fs(18, bold=False), fill=TEXT_MUTED, anchor="lt")
    d.rounded_rectangle((x1 - pad - 300, y0 + 24, x1 - pad, y0 + 60), radius=18, fill=MINT_BG)
    d.text((x1 - pad - 150, y0 + 42), "● 92% of the alphabet", font=fs(15), fill=PRIMARY, anchor="mm")
    letters = [chr(65 + i) for i in range(26)]
    notknown = {"Q", "X"}
    cols = 7
    gx = x0 + pad; gy = y0 + 116; gw = x1 - x0 - 2 * pad; gap = 14
    tile = (gw - (cols - 1) * gap) / cols
    for i, L in enumerate(letters):
        r, c = divmod(i, cols)
        tx = gx + c * (tile + gap); ty = gy + r * (tile + gap)
        known = L not in notknown
        bg = MINT_BG if known else (238, 233, 225)
        d.rounded_rectangle((tx, ty, tx + tile, ty + tile), radius=14,
                            fill=bg, outline=(GOLD_LT if known else GRID), width=3)
        d.text((tx + tile / 2, ty + tile / 2 + 4), L, font=fserif(int(tile * 0.5)),
               fill=(PRIMARY if known else TEXT_MUTED), anchor="mm")
        if known:
            d.ellipse((tx + tile - 40, ty + 10, tx + tile - 10, ty + 40), fill=GREEN)
            d.text((tx + tile - 25, ty + 24), "✓", font=fs(int(tile * 0.20)), fill=WHITE, anchor="mm")
    cy = gy + 4 * (tile + gap) + 10
    d.text((gx, cy), "NUMBERS & COUNTING", font=fs(18), fill=ACCENT, anchor="lt")
    chips = [("to 20", "counts aloud"), ("0–8", "recognizes numerals"), ("to 12", "1-to-1 counting"), ("0–5", "writes numerals")]
    cw = (gw - 3 * 18) / 4
    for i, (big, small) in enumerate(chips):
        cx0 = gx + i * (cw + 18); cyy = cy + 40
        d.rounded_rectangle((cx0, cyy, cx0 + cw, cyy + 120), radius=14, fill=WHITE, outline=GRID, width=2)
        d.rounded_rectangle((cx0 + 14, cyy, cx0 + cw - 14, cyy + 5), radius=2, fill=GOLD_LT)
        d.text((cx0 + cw / 2, cyy + 46), big, font=fserif(40), fill=PRIMARY, anchor="mm")
        d.text((cx0 + cw / 2, cyy + 92), small, font=fs(16, bold=False), fill=TEXT_MUTED, anchor="mm")


# ---------- renders ----------

def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=640)
    d = ImageDraw.Draw(img)
    block_crest(img, SIZE // 2, 132, r=56)
    pill(img, SIZE // 2, 256, "THE COMPLETE AT-HOME PRESCHOOL SYSTEM", font=fs(20), pad_x=40, pad_y=20)
    wordmark(img, SIZE // 2, 400, "PRESCHOOL COMMAND CENTER", 84, max_w=1860)
    gold_divider(img, SIZE // 2, 500, width=520)
    tc(d, (SIZE // 2, 550), "Themes, skills, letters & sweet memories — one calm system, play-based, ages 3 to 5.",
       fs(21, bold=False), (224, 213, 190))
    chips = [("17", "CONNECTED TABS"), ("12", "PRINTABLE PAGES"), ("AGES", "3–5 · ANY PACE")]
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
    tc(d, (SIZE // 2, 352), "Not a printable pack — a complete at-home preschool planning & keepsake system",
       fs(24, bold=False), (226, 214, 190))
    cards = [
        ("Start Here", "how it all works"), ("Preschool Dashboard", "live Ready-for-K Score"),
        ("Child Profiles", "a profile per child"), ("Skills & Milestones", "the whole-child tracker"),
        ("ABC & 123", "letters & numbers"), ("Weekly Themes", "a gentle theme a week"),
        ("Daily Rhythm", "circle-time flow"), ("Activity Planner", "play with a purpose"),
        ("Read-Aloud Log", "the heart of it"), ("Arts & Sensory", "grab-and-go idea bank"),
        ("Nature & Field Trips", "learning outdoors"), ("Attendance & Days", "a light record"),
        ("Portfolio", "photos & keepsakes"), ("Supplies & Budget", "low-cost, low-stress"),
        ("Kindergarten Readiness", "the gentle checklist"), ("Goals & Milestones", "the little wins"),
        ("Settings", "set it once"), ("+ 12 Printable Pages", "print & keep"),
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


def render_skills(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "THE WHOLE-CHILD SKILLS TRACKER", font=fs(28), pad_x=48, pad_y=22)
    tc(d, (SIZE // 2, 232), "Every Milestone, One Place", fserif(46), WHITE)
    tc(d, (SIZE // 2, 300), "Fine & gross motor, language, pre-math, pre-reading, social-emotional & self-help",
       fs(22, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 2, content_skills)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_abc(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=360)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 116, "ABC & 123 MASTERY", font=fs(30), pad_x=48, pad_y=22)
    tc(d, (SIZE // 2, 232), "Letter by Letter, Number by Number", fserif(46), WHITE)
    tc(d, (SIZE // 2, 300), "See at a glance which letters they know — and celebrate every new one",
       fs(22, bold=False), (226, 214, 190))
    app_window(img, (70, 400, SIZE - 70, SIZE - 70), 3, content_abc)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_plan(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=300)
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 110, "PLAN THE PLAY · SEE THEM GROW", font=fs(32), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 224), "Activities & Readiness, Handled", fserif(44), WHITE)
    app_window(img, (60, 330, SIZE - 60, 1150), 6, content_activities)
    app_window(img, (60, 1180, SIZE - 60, SIZE - 60), 13, content_readiness)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_printables(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    premium_bg(img, band_h=470)
    d = ImageDraw.Draw(img)
    block_crest(img, SIZE // 2, 110, r=48)
    pill(img, SIZE // 2, 214, "PLUS 12 PRINTABLE PAGES", font=fs(30), pad_x=48, pad_y=20)
    tc(d, (SIZE // 2, 300), "Print, Play & Keep — No Screen Needed", fserif(44), WHITE)
    gold_divider(img, SIZE // 2, 362, width=480)
    tc(d, (SIZE // 2, 404), "A matching print-ready PDF pack for the wall, the binder & your keepsake box",
       fs(22, bold=False), (226, 214, 190))
    labels = ["Child-at-a-Glance", "Weekly Theme Plan", "Daily Rhythm", "Skills & Milestones",
              "ABC & 123 Chart", "Activity Planner", "Read-Aloud Log", "Nature & Field Trips",
              "Arts & Sensory", "Readiness Checklist", "Preschool Goals", "Portfolio"]
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
        ("03_skills.png", render_skills),
        ("04_abc.png", render_abc),
        ("05_plan.png", render_plan),
        ("06_printables.png", render_printables),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()
