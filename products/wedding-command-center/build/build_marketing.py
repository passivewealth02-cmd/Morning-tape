"""Generate the ELEVATED Etsy marketing image set for Wedding Command Center™.

Flagship product — 10 listing images at 2000x2000:
  01_hero.png        - main thumbnail (Driver-Budget format, luxury elevated)
  02_dashboard.png   - executive dashboard close-up
  03_budget.png      - budget command center + charts
  04_guests.png      - guest CRM + RSVP dashboard
  05_vendors.png     - vendor CRM / comparison
  06_seating.png     - seating planner
  07_timeline.png    - 18-month master timeline
  08_vision.png      - vision board image tiles
  09_weddingday.png  - wedding day run of show
  10_mobile.png      - mobile Google Sheets preview

Run: python3 build_marketing.py
"""
from __future__ import annotations

import math
import os
from PIL import Image, ImageDraw, ImageFilter, ImageFont

# ---------------------------------------------------------------------------
# Brand tokens (luxury edition)
# ---------------------------------------------------------------------------
PRIMARY = (27, 79, 72)
PRIMARY_DK = (15, 50, 45)
ACCENT = (147, 115, 86)
GOLD = (180, 145, 90)
GOLD_LT = (201, 168, 106)
SURFACE = (229, 211, 186)
HIGHLIGHT = (117, 230, 193)
BG = (251, 248, 242)        # ivory background
WHITE = (255, 255, 255)
TEXT = (51, 51, 51)
TEXT_MUTED = (130, 125, 115)
DANGER = (201, 76, 76)
SOFT_BG = (250, 247, 241)
BLUSH = (243, 228, 221)
DOT = (224, 216, 202)
WARN_BG = (251, 240, 226)
MINT_BG = (227, 248, 239)
RED_BG = (251, 230, 230)

SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"
SERIF_BI = "/usr/share/fonts/truetype/freefont/FreeSerifBoldItalic.ttf"
SERIF_I = "/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf"

SIZE = 2000


def fs(size, bold=True):
    return ImageFont.truetype(SANS_B if bold else SANS_R, size)


def fserif(size):
    return ImageFont.truetype(SERIF_B, size)


def fserif_i(size):
    return ImageFont.truetype(SERIF_BI, size)


# ---------------------------------------------------------------------------
# Primitives
# ---------------------------------------------------------------------------
def dotted_bg(canvas, dot_color=DOT, spacing=46, radius=4):
    d = ImageDraw.Draw(canvas)
    for y in range(spacing // 2, canvas.height, spacing):
        for x in range(spacing // 2, canvas.width, spacing):
            d.ellipse((x - radius, y - radius, x + radius, y + radius), fill=dot_color)


def drop_shadow(canvas, box, radius, blur=24, alpha=70):
    layer = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).rounded_rectangle(box, radius=radius, fill=(27, 79, 72, alpha))
    canvas.alpha_composite(layer.filter(ImageFilter.GaussianBlur(blur)), (0, 18))


def tsize(draw, text, font):
    b = draw.textbbox((0, 0), text, font=font)
    return b[2] - b[0], b[3] - b[1]


def tcenter(draw, xy, text, font, fill, anchor="mm"):
    draw.text(xy, text, font=font, fill=fill, anchor=anchor)


def pill(canvas, cx, cy, text, font, fg=WHITE, bg=PRIMARY,
         pad_x=60, pad_y=26, star=False, letter_spacing=False):
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    label = f"★  {text}" if star else text
    tw, th = tsize(d, label, font)
    w, h = tw + pad_x * 2, th + pad_y * 2
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(box, radius=h // 2, fill=bg)
    # subtle gold inner border
    d.rounded_rectangle((box[0]+4, box[1]+4, box[2]-4, box[3]-4),
                        radius=(h-8)//2, outline=GOLD_LT, width=2)
    d.text((cx, cy), label, font=font, fill=fg, anchor="mm")
    canvas.alpha_composite(overlay)


def gold_divider(canvas, cx, cy, width=400):
    """A thin gold rule with a center diamond — luxury separator."""
    d = ImageDraw.Draw(canvas)
    d.line((cx - width // 2, cy, cx - 30, cy), fill=GOLD, width=3)
    d.line((cx + 30, cy, cx + width // 2, cy), fill=GOLD, width=3)
    d.polygon([(cx, cy - 12), (cx + 16, cy), (cx, cy + 12), (cx - 16, cy)], fill=GOLD)


def monogram(canvas, cx, cy, r=70):
    """Gold monogram crest on a green roundel."""
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=PRIMARY, outline=GOLD_LT, width=4)
    d.ellipse((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10),
              outline=GOLD_LT, width=2)
    d.text((cx, cy), "A&J", font=fserif(int(r * 0.7)), fill=GOLD_LT, anchor="mm")
    canvas.alpha_composite(overlay)


def kpi_callout(canvas, cx, cy, label, value, width=720, height=280, value_color=None):
    box = (cx - width // 2, cy - height // 2, cx + width // 2, cy + height // 2)
    drop_shadow(canvas, box, radius=24, blur=30, alpha=85)
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    d.rounded_rectangle(box, radius=24, fill=WHITE, outline=GOLD_LT, width=3)
    d.text((cx, cy - 72), label, font=fs(44), fill=ACCENT, anchor="mm")
    d.text((cx, cy + 48), value, font=fserif(96), fill=value_color or PRIMARY, anchor="mm")
    canvas.alpha_composite(overlay)


# ---------------------------------------------------------------------------
# Faux WCC dashboard inside a laptop / panel
# ---------------------------------------------------------------------------
def draw_dashboard_in_screen(canvas, screen):
    sx0, sy0, sx1, sy1 = screen
    sw = sx1 - sx0
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)

    head_h = 84
    d.rounded_rectangle((sx0, sy0, sx1, sy0 + head_h), radius=8, fill=PRIMARY)
    d.rectangle((sx0, sy0 + 30, sx1, sy0 + head_h), fill=PRIMARY)
    d.text((sx0 + 28, sy0 + head_h // 2), "WEDDING COMMAND CENTER",
           font=fs(26), fill=WHITE, anchor="lm")
    d.text((sx1 - 28, sy0 + head_h // 2), "Alexandra & James  ·  June 12, 2027",
           font=fserif_i(24), fill=GOLD_LT, anchor="rm")
    # gold divider under header
    d.rectangle((sx0, sy0 + head_h, sx1, sy0 + head_h + 4), fill=GOLD_LT)

    # KPI grid 5 x 2
    gx = sx0 + 22
    gy = sy0 + head_h + 22
    gap = 14
    kw = (sw - 22 * 2 - gap * 4) // 5
    kh = 120
    kpis = [
        ("DAYS UNTIL", "318", PRIMARY),
        ("BUDGET LEFT", "$13.6K", PRIMARY),
        ("SPENT", "$48.5K", ACCENT),
        ("PLANNING", "62%", PRIMARY),
        ("RSVP", "74%", PRIMARY),
        ("VENDORS", "9", PRIMARY),
        ("OUTSTANDING", "$24.5K", ACCENT),
        ("GUESTS", "118", PRIMARY),
        ("TABLES", "12/16", PRIMARY),
        ("DUE THIS WK", "4", DANGER),
    ]
    for i, (lab, val, col) in enumerate(kpis):
        row, ci = divmod(i, 5)
        x0 = gx + ci * (kw + gap)
        y0 = gy + row * (kh + gap)
        box = (x0, y0, x0 + kw, y0 + kh)
        d.rounded_rectangle(box, radius=10, fill=WHITE, outline=(225, 218, 205), width=2)
        d.text((x0 + 12, y0 + 14), lab, font=fs(15), fill=ACCENT, anchor="lt")
        d.text((x0 + 12, y0 + 70), val, font=fserif(38), fill=col, anchor="lm")

    # Quick nav chips
    nav_y = gy + 2 * (kh + gap) + 8
    nav_h = 34
    nav = ["Budget", "Guests", "Vendors", "Timeline", "Seating",
           "Vision", "Payments", "Contracts", "Checklist", "Day"]
    cw = (sw - 22 * 2 - 6 * (len(nav) - 1)) / len(nav)
    for i, name in enumerate(nav):
        x0 = gx + i * (cw + 6)
        d.rounded_rectangle((x0, nav_y, x0 + cw, nav_y + nav_h), radius=17, fill=PRIMARY)
        d.text((x0 + cw / 2, nav_y + nav_h / 2), name, font=fs(15), fill=WHITE, anchor="mm")

    # Two charts
    ch_y0 = nav_y + nav_h + 20
    ch_h = sy1 - ch_y0 - 18
    chart_w = (sw - 22 * 2 - gap) // 2

    def donut_card(box, title, segs, center_top, center_bot, legend):
        """Donut + legend that scales to fill the card and centers as a block."""
        d.rounded_rectangle(box, radius=10, fill=WHITE, outline=(225, 218, 205), width=2)
        d.text((box[0] + 16, box[1] + 14), title, font=fs(20), fill=ACCENT, anchor="lt")
        title_h = 50
        avail_h = (box[3] - box[1]) - title_h
        card_w = box[2] - box[0]
        # radius adapts to card; capped so legend always fits to the right
        rr = int(max(60, min(avail_h * 0.40, card_w * 0.20)))
        dcx = box[0] + 34 + rr
        dcy = box[1] + title_h + avail_h // 2
        ang = -90
        for pct, col in segs:
            s = pct * 3.6
            d.pieslice((dcx - rr, dcy - rr, dcx + rr, dcy + rr), ang, ang + s, fill=col)
            ang += s
        hole = int(rr * 0.46)
        d.ellipse((dcx - hole, dcy - hole, dcx + hole, dcy + hole), fill=WHITE)
        d.text((dcx, dcy - hole * 0.28), center_top, font=fserif(int(rr * 0.30)),
               fill=PRIMARY, anchor="mm")
        d.text((dcx, dcy + hole * 0.42), center_bot, font=fs(int(rr * 0.16)),
               fill=TEXT_MUTED, anchor="mm")
        # legend — vertically centered block to the right of the donut
        line_h = max(34, int(avail_h / (len(legend) + 1)))
        sw_box = 22
        lx = dcx + rr + 40
        block_h = line_h * len(legend)
        ly = dcy - block_h // 2 + (line_h - sw_box) // 2
        for i, (lab, col) in enumerate(legend):
            yy = ly + i * line_h
            d.rounded_rectangle((lx, yy, lx + sw_box, yy + sw_box), radius=4, fill=col)
            d.text((lx + sw_box + 14, yy + sw_box // 2), lab, font=fs(20), fill=TEXT, anchor="lm")

    dbox = (gx, ch_y0, gx + chart_w, ch_y0 + ch_h)
    donut_card(
        dbox, "BUDGET BREAKDOWN",
        [(27, PRIMARY), (20, ACCENT), (16, HIGHLIGHT), (12, SURFACE),
         (10, GOLD), (15, TEXT_MUTED)],
        "$48K", "SPENT",
        [("Venue 27%", PRIMARY), ("Catering 20%", ACCENT),
         ("Photo 16%", HIGHLIGHT), ("Florals 12%", SURFACE),
         ("Music 10%", GOLD), ("Other 15%", TEXT_MUTED)])

    bbox = (gx + chart_w + gap, ch_y0, gx + 2 * chart_w + gap, ch_y0 + ch_h)
    donut_card(
        bbox, "RSVP PROGRESS",
        [(74, HIGHLIGHT), (12, DANGER), (14, SURFACE)],
        "74%", "REPLIED",
        [("Accepted 104", HIGHLIGHT), ("Declined 16", DANGER),
         ("Pending 20", SURFACE)])

    canvas.alpha_composite(overlay)


def draw_laptop(canvas, cx, cy, w=1500, h=900):
    overlay = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    sbox = (cx - w // 2, cy - h // 2 + 30, cx + w // 2, cy + h // 2 + 80)
    shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle(sbox, radius=40, fill=(27, 79, 72, 110))
    canvas.alpha_composite(shadow.filter(ImageFilter.GaussianBlur(50)))
    body = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(body, radius=28, fill=(38, 38, 42), outline=(20, 20, 25), width=4)
    bezel = 22
    screen = (body[0] + bezel, body[1] + bezel, body[2] - bezel, body[3] - bezel)
    d.rounded_rectangle(screen, radius=8, fill=BG)
    d.rounded_rectangle((cx - w // 2 - 120, cy + h // 2, cx + w // 2 + 120, cy + h // 2 + 24),
                        radius=12, fill=(58, 58, 62))
    d.rounded_rectangle((cx - 80, cy + h // 2 + 4, cx + 80, cy + h // 2 + 14),
                        radius=6, fill=(33, 33, 38))
    canvas.alpha_composite(overlay)
    draw_dashboard_in_screen(canvas, screen)


# ---------------------------------------------------------------------------
# Shared page scaffold for the supporting (non-hero) images
# ---------------------------------------------------------------------------
def page_scaffold(img, badge, title, subtitle):
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 175, badge, font=fs(42), pad_x=58, pad_y=24)
    tcenter(d, (SIZE // 2, 310), title, fserif(74), PRIMARY)
    gold_divider(img, SIZE // 2, 380, width=520)
    tcenter(d, (SIZE // 2, 430), subtitle, fs(32, bold=False), TEXT_MUTED)


def panel(img, box, fill_color=WHITE, title=None):
    drop_shadow(img, box, radius=24, blur=38, alpha=105)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    od.rounded_rectangle(box, radius=24, fill=fill_color, outline=GOLD_LT, width=3)
    img.alpha_composite(overlay)
    if title:
        d = ImageDraw.Draw(img)
        d.text((box[0] + 34, box[1] + 26), title, font=fs(30), fill=ACCENT, anchor="lt")


def draw_table(img, inner, headers, rows, status_col=None, status_map=None,
               col_aligns=None, header_font=None, row_font=None):
    """Generic branded table renderer. rows is list of (cells..., bg_or_None)."""
    od_layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(od_layer)
    iw = inner[2] - inner[0]
    ih = inner[3] - inner[1]
    cols = len(headers)
    col_w = iw / cols
    hdr_h = 70
    hf = header_font or fs(22)
    rf = row_font or fs(22)
    od.rounded_rectangle((inner[0], inner[1], inner[2], inner[1] + hdr_h), radius=12, fill=PRIMARY)
    od.rectangle((inner[0], inner[1] + 24, inner[2], inner[1] + hdr_h), fill=PRIMARY)
    for i, h in enumerate(headers):
        cx = inner[0] + col_w * (i + 0.5)
        od.text((cx, inner[1] + hdr_h // 2), h, font=hf, fill=WHITE, anchor="mm")
    row_h = (ih - hdr_h - 10) / len(rows)
    for ri, row in enumerate(rows):
        *cells, bg = row
        y0 = inner[1] + hdr_h + ri * row_h
        if bg is None and ri % 2 == 1:
            bg = (244, 236, 222)
        if bg:
            od.rectangle((inner[0] + 3, y0, inner[2] - 3, y0 + row_h), fill=bg)
        for ci, val in enumerate(cells):
            cx = inner[0] + col_w * (ci + 0.5)
            align = (col_aligns or {}).get(ci, "mm")
            color = PRIMARY if ci == 0 else TEXT
            if status_col is not None and ci == status_col and status_map:
                tag = status_map.get(str(val))
                if tag:
                    bg_c, fg_c = tag
                    pw = min(col_w - 16, 150)
                    od.rounded_rectangle((cx - pw/2, y0 + row_h/2 - 17,
                                          cx + pw/2, y0 + row_h/2 + 17),
                                         radius=16, fill=bg_c)
                    od.text((cx, y0 + row_h / 2), str(val), font=fs(18), fill=fg_c, anchor="mm")
                    continue
            if align == "lm":
                od.text((inner[0] + col_w * ci + 18, y0 + row_h / 2), str(val),
                        font=rf, fill=color, anchor="lm")
            else:
                od.text((cx, y0 + row_h / 2), str(val), font=rf, fill=color, anchor="mm")
    img.alpha_composite(od_layer)


# ===========================================================================
# Image 1: HERO (luxury elevated)
# ===========================================================================
def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)

    # Monogram crest at top
    monogram(img, SIZE // 2, 175, r=72)

    pill(img, SIZE // 2, 330, "ULTIMATE 32-SHEET WEDDING SYSTEM",
         font=fs(42), pad_x=64, pad_y=24)

    # Serif wordmark
    tcenter(d, (SIZE // 2, 510), "WEDDING", fserif(184), PRIMARY)
    tcenter(d, (SIZE // 2, 690), "COMMAND CENTER", fserif(128), PRIMARY)

    gold_divider(img, SIZE // 2, 800, width=620)
    tcenter(d, (SIZE // 2, 858), "BUDGET   ·   GUESTS   ·   VENDORS   ·   THE BIG DAY",
            fs(40), ACCENT)

    # Laptop
    draw_laptop(img, SIZE // 2, 1340, w=1660, h=950)

    # Floating KPI callouts
    kpi_callout(img, 360, 1115, "DAYS TO GO", "318",
                width=660, height=250, value_color=ACCENT)
    kpi_callout(img, SIZE - 360, 1115, "READINESS", "82%",
                width=660, height=250)

    # CTA
    pill(img, SIZE // 2, SIZE - 150, "GET THE COMPLETE 32-SHEET SYSTEM",
         font=fs(52), pad_x=72, pad_y=32, star=True)

    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 2: DASHBOARD CLOSEUP
# ===========================================================================
def render_dashboard(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "EXECUTIVE DASHBOARD",
                  "Every Detail. One Elegant Screen.",
                  "10 live KPIs · budget · RSVP · readiness — your whole wedding at a glance")
    box = (130, 510, SIZE - 130, SIZE - 190)
    panel(img, box, fill_color=BG)
    inner = (box[0] + 22, box[1] + 22, box[2] - 22, box[3] - 22)
    draw_dashboard_in_screen(img, inner)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 3: BUDGET
# ===========================================================================
def render_budget(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "BUDGET COMMAND CENTER",
                  "Track Every Dollar, Beautifully.",
                  "21 categories · deposits · paid · remaining — with live variance")

    # KPI strip
    kpis = [("TOTAL BUDGET", "$45,000"), ("ESTIMATED", "$48,480"),
            ("PAID", "$31,400"), ("REMAINING", "$13,600")]
    kw, kh, gap = 380, 175, 30
    start_x = (SIZE - (4 * kw + 3 * gap)) // 2
    ky = 510
    for i, (lab, val) in enumerate(kpis):
        x = start_x + i * (kw + gap)
        b = (x, ky, x + kw, ky + kh)
        drop_shadow(img, b, radius=20, blur=24, alpha=80)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        od.rounded_rectangle(b, radius=20, fill=WHITE, outline=GOLD_LT, width=2)
        od.text((x + kw // 2, ky + 52), lab, font=fs(26), fill=ACCENT, anchor="mm")
        od.text((x + kw // 2, ky + 118), val, font=fserif(58), fill=PRIMARY, anchor="mm")
        img.alpha_composite(ov)

    # Donut card
    db = (130, 760, 950, SIZE - 180)
    panel(img, db, title="CATEGORY SPENDING")
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    dcx = (db[0] + db[2]) // 2 - 110
    dcy = (db[1] + db[3]) // 2 + 35
    rr = 215
    segs = [(27, PRIMARY, "Venue 27%"), (20, ACCENT, "Catering 20%"),
            (16, HIGHLIGHT, "Photo+Video 16%"), (8, SURFACE, "Florals 8%"),
            (7, GOLD, "Music 7%"), (6, (180, 90, 90), "Attire 6%"),
            (16, TEXT_MUTED, "Other 16%")]
    ang = -90
    for pct, col, _ in segs:
        s = pct * 3.6
        od.pieslice((dcx-rr, dcy-rr, dcx+rr, dcy+rr), ang, ang+s, fill=col)
        ang += s
    od.ellipse((dcx-88, dcy-88, dcx+88, dcy+88), fill=WHITE)
    od.text((dcx, dcy-14), "$48.5K", font=fserif(34), fill=PRIMARY, anchor="mm")
    od.text((dcx, dcy+28), "ESTIMATED", font=fs(18), fill=TEXT_MUTED, anchor="mm")
    lx = dcx + rr + 38
    for i, (pct, col, lab) in enumerate(segs):
        yy = dcy - rr + 18 + i * 46
        od.rounded_rectangle((lx, yy, lx+26, yy+26), radius=5, fill=col)
        od.text((lx+42, yy+13), lab, font=fs(22), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)

    # Bar card
    bb = (1010, 760, SIZE - 130, SIZE - 180)
    panel(img, bb, title="BUDGET vs ACTUAL")
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    cats = ["Venue", "Cater", "Photo", "Video", "Floral", "Music", "Decor", "Bar"]
    plan = [12000, 9000, 4500, 3000, 3500, 1500, 2500, 3500]
    actual = [12000, 4500, 4500, 3000, 3300, 1500, 2100, 0]
    bx0, by0 = bb[0] + 50, bb[1] + 110
    bx1, by1 = bb[2] - 40, bb[3] - 90
    mx = max(plan)
    bw = (bx1 - bx0) / (len(cats) * 2.4)
    for i, (p, a) in enumerate(zip(plan, actual)):
        x = bx0 + i * (bw * 2.4) + 18
        hp = (p / mx) * (by1 - by0)
        ha = (a / mx) * (by1 - by0)
        od.rounded_rectangle((x, by1 - hp, x + bw, by1), radius=6, fill=SURFACE)
        od.rounded_rectangle((x + bw + 8, by1 - ha, x + 2*bw + 8, by1), radius=6, fill=PRIMARY)
        od.text((x + bw + 4, by1 + 20), cats[i], font=fs(18), fill=TEXT_MUTED, anchor="mt")
    od.rounded_rectangle((bb[0]+34, bb[3]-52, bb[0]+54, bb[3]-32), radius=4, fill=SURFACE)
    od.text((bb[0]+64, bb[3]-42), "Budget", font=fs(20), fill=TEXT, anchor="lm")
    od.rounded_rectangle((bb[0]+190, bb[3]-52, bb[0]+210, bb[3]-32), radius=4, fill=PRIMARY)
    od.text((bb[0]+220, bb[3]-42), "Actual", font=fs(20), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 4: GUEST CRM + RSVP
# ===========================================================================
def render_guests(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "GUEST CRM + RSVP",
                  "Your Entire Guest List, Handled.",
                  "200-row database · live RSVP counts · meal totals · seating-ready")

    # RSVP KPI cards
    kpis = [("INVITED", "140", PRIMARY), ("ACCEPTED", "104", PRIMARY),
            ("DECLINED", "16", ACCENT), ("PENDING", "20", DANGER)]
    kw, kh, gap = 380, 160, 30
    start_x = (SIZE - (4 * kw + 3 * gap)) // 2
    ky = 500
    for i, (lab, val, col) in enumerate(kpis):
        x = start_x + i * (kw + gap)
        b = (x, ky, x + kw, ky + kh)
        drop_shadow(img, b, radius=20, blur=22, alpha=78)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        od.rounded_rectangle(b, radius=20, fill=WHITE, outline=GOLD_LT, width=2)
        od.text((x + kw // 2, ky + 48), lab, font=fs(26), fill=ACCENT, anchor="mm")
        od.text((x + kw // 2, ky + 110), val, font=fserif(58), fill=col, anchor="mm")
        img.alpha_composite(ov)

    # Guest table
    gb = (130, 700, SIZE - 130, SIZE - 470)
    panel(img, gb, title="GUEST DATABASE")
    inner = (gb[0] + 28, gb[1] + 80, gb[2] - 28, gb[3] - 28)
    smap = {"Accepted": (MINT_BG, PRIMARY), "Declined": (RED_BG, DANGER),
            "Pending": (WARN_BG, ACCENT)}
    rows = [
        ("Linda & Robert Bennett", "Bennett", "2", "Accepted", "Beef", "Yes", None),
        ("Daniel Carter", "Carter", "1", "Accepted", "Chicken", "No", None),
        ("Emily Bennett", "Bennett", "1", "Accepted", "Vegetarian", "Yes", None),
        ("The Hartleys", "Hartley", "4", "Pending", "—", "No", None),
        ("Aiden Brooks", "Brooks", "0", "Declined", "—", "No", None),
        ("The Patels", "Patel", "3", "Accepted", "Vegan", "No", None),
        ("Grandma Rose", "Bennett", "1", "Accepted", "Fish", "Yes", None),
    ]
    draw_table(img, inner,
               ["GUEST", "FAMILY", "SEATS", "RSVP", "MEAL", "GIFT"],
               rows, status_col=3, status_map=smap,
               col_aligns={0: "lm", 1: "lm"}, header_font=fs(20), row_font=fs(20))

    # Meal counts strip
    mb = (130, SIZE - 440, SIZE - 130, SIZE - 150)
    panel(img, mb, title="MEAL COUNTS  (accepted guests)")
    meals = [("Beef", 38, PRIMARY), ("Chicken", 29, ACCENT), ("Fish", 18, HIGHLIGHT),
             ("Vegetarian", 12, SURFACE), ("Vegan", 5, GOLD), ("Kids", 8, (180, 90, 90))]
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    bx0, by1 = mb[0] + 60, mb[3] - 70
    by0 = mb[1] + 90
    mx = max(m[1] for m in meals)
    slot = (mb[2] - mb[0] - 120) / len(meals)
    for i, (lab, val, col) in enumerate(meals):
        x = bx0 + i * slot
        h = (val / mx) * (by1 - by0)
        od.rounded_rectangle((x, by1 - h, x + slot * 0.5, by1), radius=6, fill=col)
        od.text((x + slot * 0.25, by1 - h - 24), str(val), font=fs(24), fill=PRIMARY, anchor="mm")
        od.text((x + slot * 0.25, by1 + 22), lab, font=fs(20), fill=TEXT_MUTED, anchor="mt")
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 5: VENDORS
# ===========================================================================
def render_vendors(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "VENDOR CRM + COMPARISON",
                  "Compare, Book, and Track Every Vendor.",
                  "Quotes · deposits · contracts · star ratings — auto-ranked by score")

    # Vendor CRM table
    vb = (130, 500, SIZE - 130, 1230)
    panel(img, vb, title="VENDOR CRM")
    inner = (vb[0] + 28, vb[1] + 80, vb[2] - 28, vb[3] - 28)
    smap = {"Signed": (MINT_BG, PRIMARY), "Pending": (WARN_BG, ACCENT)}
    rows = [
        ("Venue", "The Hartwell Estate", "$12,000", "Signed", "★★★★★", None),
        ("Photographer", "Lumiere Photography", "$4,500", "Signed", "★★★★★", None),
        ("Florist", "Petal & Stem", "$3,300", "Pending", "★★★★★", None),
        ("Caterer", "Grand Catering Group", "$9,000", "Signed", "★★★★★", None),
        ("DJ", "Skyline Entertainment", "$2,500", "Signed", "★★★★☆", None),
        ("Cake", "Sweet Layers", "$850", "Signed", "★★★★☆", None),
        ("Planner", "Lane & Co. Events", "$5,000", "Signed", "★★★★★", None),
    ]
    draw_table(img, inner,
               ["TYPE", "VENDOR", "QUOTE", "CONTRACT", "RATING"],
               rows, status_col=3, status_map=smap,
               col_aligns={1: "lm"}, header_font=fs(20), row_font=fs(21))

    # Comparison card
    cb = (130, 1270, SIZE - 130, SIZE - 150)
    panel(img, cb, title="VENDOR COMPARISON  ·  auto-ranked")
    inner = (cb[0] + 28, cb[1] + 80, cb[2] - 28, cb[3] - 28)
    rows = [
        ("Lumiere Photography", "$4,500", "4.9★", "95", "1", MINT_BG),
        ("Aperture Co.", "$5,200", "4.8★", "80", "2", None),
        ("Bright Frame Studio", "$3,800", "4.6★", "84", "3", None),
    ]
    draw_table(img, inner,
               ["PHOTOGRAPHER", "PRICE", "REVIEWS", "SCORE", "RANK"],
               rows, col_aligns={0: "lm"}, header_font=fs(20), row_font=fs(22))
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 6: SEATING
# ===========================================================================
def render_seating(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "SEATING PLANNER",
                  "Tables That Arrange Themselves.",
                  "Live capacity warnings · VIP highlighting · auto seat counts")

    # Visual table layout — round tables
    vb = (130, 500, SIZE - 130, 1300)
    panel(img, vb, title="RECEPTION FLOOR PLAN")
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    tables = [
        ("1", "HEAD", 8, 10, "vip"), ("2", "Bennett", 9, 10, "ok"),
        ("3", "Carter", 10, 10, "full"), ("4", "College", 8, 10, "ok"),
        ("5", "Work", 7, 10, "ok"), ("6", "Families", 10, 10, "full"),
        ("7", "Friends", 9, 10, "ok"), ("8", "Plus-ones", 6, 10, "ok"),
    ]
    cols = 4
    cell_w = (vb[2] - vb[0] - 80) / cols
    cell_h = 320
    for i, (num, label, seated, cap, state) in enumerate(tables):
        r, c = divmod(i, cols)
        cx = vb[0] + 40 + c * cell_w + cell_w / 2
        cy = vb[1] + 130 + r * cell_h + 110
        rr = 95
        if state == "vip":
            ring = GOLD
            fillc = (246, 239, 224)
        elif state == "full":
            ring = HIGHLIGHT
            fillc = MINT_BG
        else:
            ring = PRIMARY
            fillc = WHITE
        od.ellipse((cx-rr, cy-rr, cx+rr, cy+rr), fill=fillc, outline=ring, width=6)
        # seats as small dots around
        for s in range(cap):
            a = math.radians(s / cap * 360 - 90)
            sx = cx + math.cos(a) * (rr + 20)
            sy = cy + math.sin(a) * (rr + 20)
            seat_col = PRIMARY if s < seated else (210, 205, 195)
            od.ellipse((sx-9, sy-9, sx+9, sy+9), fill=seat_col)
        od.text((cx, cy - 18), f"Table {num}", font=fs(24), fill=PRIMARY, anchor="mm")
        od.text((cx, cy + 12), label, font=fs(20), fill=TEXT_MUTED, anchor="mm")
        od.text((cx, cy + 42), f"{seated}/{cap}", font=fserif(28), fill=ring if state!="ok" else PRIMARY, anchor="mm")
        if state == "vip":
            od.text((cx, cy + 78), "★ VIP", font=fs(18), fill=GOLD, anchor="mm")
    img.alpha_composite(ov)

    # Capacity summary
    sb = (130, 1340, SIZE - 130, SIZE - 150)
    panel(img, sb, title="CAPACITY SUMMARY")
    stats = [("Total Seated", "67"), ("Total Capacity", "80"),
             ("Seats Open", "13"), ("Accepted Guests", "104")]
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    slot = (sb[2] - sb[0]) / len(stats)
    for i, (lab, val) in enumerate(stats):
        cx = sb[0] + slot * (i + 0.5)
        od.text((cx, sb[1] + 130), val, font=fserif(64), fill=PRIMARY, anchor="mm")
        od.text((cx, sb[1] + 200), lab, font=fs(24), fill=ACCENT, anchor="mm")
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 7: TIMELINE
# ===========================================================================
def render_timeline(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "MASTER TIMELINE",
                  "18 Months, Perfectly Paced.",
                  "Auto-anchored to your date — every milestone, on schedule")

    box = (150, 510, SIZE - 150, SIZE - 170)
    panel(img, box)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    # Vertical timeline spine
    spine_x = box[0] + 230
    od.line((spine_x, box[1] + 60, spine_x, box[3] - 60), fill=GOLD, width=5)

    phases = [
        ("18 MONTHS", "Set date, budget & vision", "done"),
        ("12 MONTHS", "Book venue, planner, photographer", "done"),
        ("9 MONTHS", "Order dress · book caterer & florist", "done"),
        ("6 MONTHS", "Save-the-dates · entertainment", "done"),
        ("3 MONTHS", "Invitations · menu · wedding bands", "active"),
        ("1 MONTH", "Mail invites · final fitting", "todo"),
        ("2 WEEKS", "Seating chart · final headcount", "todo"),
        ("WEDDING WEEK", "Rehearsal · pick up attire", "todo"),
        ("WEDDING DAY", "Execute run of show", "todo"),
        ("POST", "Thank-you cards · review vendors", "todo"),
    ]
    n = len(phases)
    span = (box[3] - 60) - (box[1] + 60)
    for i, (phase, task, state) in enumerate(phases):
        y = box[1] + 60 + span * i / (n - 1)
        if state == "done":
            dot_c = HIGHLIGHT; ring = PRIMARY
        elif state == "active":
            dot_c = GOLD; ring = ACCENT
        else:
            dot_c = WHITE; ring = (200, 195, 185)
        od.ellipse((spine_x - 22, y - 22, spine_x + 22, y + 22), fill=dot_c, outline=ring, width=5)
        if state == "done":
            od.text((spine_x, y), "✓", font=fs(24), fill=PRIMARY, anchor="mm")
        # Phase label left of spine
        od.text((spine_x - 50, y), phase, font=fs(26), fill=PRIMARY, anchor="rm")
        # Task right of spine
        od.text((spine_x + 50, y), task, font=fs(26, bold=False), fill=TEXT, anchor="lm")
        # status chip
        if state == "active":
            od.rounded_rectangle((box[2] - 220, y - 20, box[2] - 40, y + 20),
                                 radius=20, fill=WARN_BG)
            od.text((box[2] - 130, y), "IN PROGRESS", font=fs(18), fill=ACCENT, anchor="mm")
        elif state == "done":
            od.rounded_rectangle((box[2] - 180, y - 20, box[2] - 40, y + 20),
                                 radius=20, fill=MINT_BG)
            od.text((box[2] - 110, y), "COMPLETE", font=fs(18), fill=PRIMARY, anchor="mm")
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 8: VISION BOARD
# ===========================================================================
def render_vision(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "VISION BOARD",
                  "Pin Your Inspiration in One Place.",
                  "16 image tiles — drag & drop your dream venue, dress, florals & more")

    box = (130, 510, SIZE - 130, SIZE - 170)
    panel(img, box, fill_color=(251, 248, 242))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    tiles = ["Venue", "Dress", "Florals", "Cake", "Hair", "Makeup",
             "Decor", "Tablescapes", "Stationery", "Palette",
             "Bouquets", "Poses", "Reception", "Lighting", "Favors", "Travel"]
    cols = 4
    pad = 34
    iw = box[2] - box[0] - pad * 2
    ih = box[3] - box[1] - pad * 2
    tile_w = (iw - (cols - 1) * 24) / cols
    rows_n = 4
    tile_h = (ih - (rows_n - 1) * 24) / rows_n
    icons = ["⛪", "👰", "🌸", "🎂", "💇", "💄", "🕯", "🍽",
             "✉", "🎨", "💐", "📸", "🥂", "💡", "🎁", "🌴"]
    for i, label in enumerate(tiles):
        r, c = divmod(i, cols)
        x0 = box[0] + pad + c * (tile_w + 24)
        y0 = box[1] + pad + r * (tile_h + 24)
        # blush tile with dashed gold border
        od.rounded_rectangle((x0, y0, x0 + tile_w, y0 + tile_h), radius=14, fill=BLUSH)
        # dashed border
        dash = 14
        for xx in range(int(x0), int(x0 + tile_w), dash * 2):
            od.line((xx, y0, min(xx + dash, x0 + tile_w), y0), fill=GOLD_LT, width=3)
            od.line((xx, y0 + tile_h, min(xx + dash, x0 + tile_w), y0 + tile_h), fill=GOLD_LT, width=3)
        for yy in range(int(y0), int(y0 + tile_h), dash * 2):
            od.line((x0, yy, x0, min(yy + dash, y0 + tile_h)), fill=GOLD_LT, width=3)
            od.line((x0 + tile_w, yy, x0 + tile_w, min(yy + dash, y0 + tile_h)), fill=GOLD_LT, width=3)
        od.text((x0 + tile_w / 2, y0 + tile_h / 2 - 30), icons[i], font=fs(56), fill=ACCENT, anchor="mm")
        od.text((x0 + tile_w / 2, y0 + tile_h / 2 + 36), label, font=fs(26), fill=PRIMARY, anchor="mm")
        od.text((x0 + tile_w / 2, y0 + tile_h - 26), "⬆ drop image", font=fs(16, bold=False),
                fill=TEXT_MUTED, anchor="mm")
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 9: WEDDING DAY
# ===========================================================================
def render_weddingday(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "WEDDING DAY COMMAND CENTER",
                  "Hour-by-Hour, Nothing Forgotten.",
                  "The single run-of-show for you, your planner & every vendor")

    box = (150, 510, SIZE - 150, SIZE - 170)
    panel(img, box)
    inner = (box[0] + 30, box[1] + 40, box[2] - 30, box[3] - 40)
    schedule = [
        ("8:00 AM", "Hair & makeup begins", "Radiant Beauty", False),
        ("1:00 PM", "First look", "Couple + Photo", False),
        ("2:30 PM", "Ceremony setup complete", "Coordinator", False),
        ("4:30 PM", "CEREMONY BEGINS", "Officiant", True),
        ("5:00 PM", "Cocktail hour", "DJ / Bar", False),
        ("6:00 PM", "Reception entrance + first dance", "DJ", False),
        ("6:30 PM", "Dinner service (plated)", "Grand Catering", False),
        ("7:30 PM", "Toasts & speeches", "MOH / Best Man", False),
        ("8:00 PM", "Cake cutting", "Couple", False),
        ("10:30 PM", "Sparkler send-off", "Coordinator", True),
    ]
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    row_h = (inner[3] - inner[1]) / len(schedule)
    for i, (time, event, lead, highlight) in enumerate(schedule):
        y0 = inner[1] + i * row_h
        if highlight:
            od.rounded_rectangle((inner[0], y0 + 6, inner[2], y0 + row_h - 6),
                                 radius=12, fill=MINT_BG)
        # time block
        od.rounded_rectangle((inner[0] + 8, y0 + 14, inner[0] + 230, y0 + row_h - 14),
                             radius=10, fill=PRIMARY if not highlight else GOLD)
        od.text((inner[0] + 119, y0 + row_h / 2), time, font=fserif(34),
                fill=WHITE, anchor="mm")
        # event
        ef = fs(34) if highlight else fs(30, bold=True)
        od.text((inner[0] + 270, y0 + row_h / 2 - 16), event, font=ef,
                fill=PRIMARY, anchor="lm")
        od.text((inner[0] + 270, y0 + row_h / 2 + 22), lead, font=fs(22, bold=False),
                fill=TEXT_MUTED, anchor="lm")
        # divider
        if i < len(schedule) - 1:
            od.line((inner[0] + 270, y0 + row_h, inner[2] - 20, y0 + row_h),
                    fill=(230, 224, 212), width=1)
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
# Image 10: MOBILE
# ===========================================================================
def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "WORKS EVERYWHERE",
                  "Excel  ·  Google Sheets  ·  Mobile",
                  "Plan at your desk, check details from anywhere")

    px, py = SIZE // 2, 1300
    pw, ph = 660, 1230
    phone = (px - pw // 2, py - ph // 2, px + pw // 2, py + ph // 2)
    drop_shadow(img, phone, radius=64, blur=50, alpha=120)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle(phone, radius=64, fill=(24, 24, 28))
    bezel = 22
    screen = (phone[0] + bezel, phone[1] + bezel + 30, phone[2] - bezel, phone[3] - bezel - 30)
    od.rounded_rectangle(screen, radius=44, fill=BG)
    od.rounded_rectangle((px - 95, phone[1] + 16, px + 95, phone[1] + 50), radius=18, fill=(14, 14, 18))
    sx0, sy0, sx1, sy1 = screen

    # header with monogram
    od.rounded_rectangle((sx0, sy0, sx1, sy0 + 110), radius=44, fill=PRIMARY)
    od.rectangle((sx0, sy0 + 60, sx1, sy0 + 110), fill=PRIMARY)
    od.rectangle((sx0, sy0 + 106, sx1, sy0 + 110), fill=GOLD_LT)
    od.text(((sx0 + sx1) // 2, sy0 + 46), "A & J", font=fserif(40), fill=GOLD_LT, anchor="mm")
    od.text(((sx0 + sx1) // 2, sy0 + 86), "Wedding Command Center", font=fs(20), fill=WHITE, anchor="mm")
    img.alpha_composite(ov)

    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    y = sy0 + 150
    cards = [("DAYS UNTIL WEDDING", "318", ACCENT),
             ("READINESS SCORE", "82%", PRIMARY),
             ("BUDGET REMAINING", "$13,600", PRIMARY),
             ("RSVP COMPLETE", "74%", PRIMARY)]
    for lab, val, col in cards:
        cb = (sx0 + 30, y, sx1 - 30, y + 135)
        od.rounded_rectangle(cb, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.text((cb[0] + 26, y + 30), lab, font=fs(22), fill=ACCENT, anchor="lt")
        od.text((cb[0] + 26, y + 92), val, font=fserif(46), fill=col, anchor="lm")
        y += 152

    od.text((sx0 + 40, y + 16), "NEXT MILESTONES", font=fs(22), fill=ACCENT, anchor="lt")
    y += 52
    items = [("Mail invitations", "+10d", WARN_BG),
             ("Final dress fitting", "+45d", SURFACE),
             ("Confirm headcount", "+40d", SURFACE)]
    for task, days, dot in items:
        od.ellipse((sx0 + 40, y + 12, sx0 + 64, y + 36), fill=dot)
        od.text((sx0 + 80, y + 24), task, font=fs(24), fill=TEXT, anchor="lm")
        od.text((sx1 - 50, y + 24), days, font=fs(22), fill=PRIMARY, anchor="rm")
        y += 64
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ---------------------------------------------------------------------------
# Build all
# ---------------------------------------------------------------------------
def main():
    out_dir = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("01_hero.png",        render_hero),
        ("02_dashboard.png",   render_dashboard),
        ("03_budget.png",      render_budget),
        ("04_guests.png",      render_guests),
        ("05_vendors.png",     render_vendors),
        ("06_seating.png",     render_seating),
        ("07_timeline.png",    render_timeline),
        ("08_vision.png",      render_vision),
        ("09_weddingday.png",  render_weddingday),
        ("10_mobile.png",      render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()
