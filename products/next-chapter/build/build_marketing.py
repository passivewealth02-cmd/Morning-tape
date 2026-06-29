"""Premium marketing image set for Next Chapter™ (6 images, 2000x2000).

  01_hero.png       - feature-forward main thumbnail
  02_dashboard.png  - life dashboard close-up
  03_tasks.png      - master task manager
  04_finances.png   - financial snapshot / net worth
  05_analytics.png  - progress analytics
  06_mobile.png     - mobile preview

Run: python3 build_marketing.py
"""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFilter, ImageFont

PRIMARY = (27, 79, 72)
ACCENT = (147, 115, 86)
GOLD = (180, 145, 90)
GOLD_LT = (201, 168, 106)
SURFACE = (229, 211, 186)
HIGHLIGHT = (117, 230, 193)
BG = (251, 248, 242)
WHITE = (255, 255, 255)
TEXT = (51, 51, 51)
TEXT_MUTED = (130, 125, 115)
DANGER = (201, 76, 76)
MINT_BG = (227, 248, 239)
WARN_BG = (251, 240, 226)
DOT = (224, 216, 202)
SIZE = 2000

SANS_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
SANS_R = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
SERIF_B = "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"


def fs(s, bold=True):
    return ImageFont.truetype(SANS_B if bold else SANS_R, s)


def fserif(s):
    return ImageFont.truetype(SERIF_B, s)


def dotted_bg(c, spacing=46, r=4):
    d = ImageDraw.Draw(c)
    for y in range(spacing // 2, c.height, spacing):
        for x in range(spacing // 2, c.width, spacing):
            d.ellipse((x - r, y - r, x + r, y + r), fill=DOT)


def shadow(c, box, radius, blur=24, alpha=70):
    layer = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).rounded_rectangle(box, radius=radius, fill=(27, 79, 72, alpha))
    c.alpha_composite(layer.filter(ImageFilter.GaussianBlur(blur)), (0, 18))


def tsize(d, t, f):
    b = d.textbbox((0, 0), t, font=f)
    return b[2] - b[0], b[3] - b[1]


def tc(d, xy, t, f, fill, anchor="mm"):
    d.text(xy, t, font=f, fill=fill, anchor=anchor)


def pill(c, cx, cy, text, font, pad_x=60, pad_y=26, star=False, bg=PRIMARY, fg=WHITE):
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    label = f"★  {text}" if star else text
    tw, th = tsize(d, label, font)
    w, h = tw + pad_x * 2, th + pad_y * 2
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(box, radius=h // 2, fill=bg)
    d.rounded_rectangle((box[0] + 4, box[1] + 4, box[2] - 4, box[3] - 4),
                        radius=(h - 8) // 2, outline=GOLD_LT, width=2)
    d.text((cx, cy), label, font=font, fill=fg, anchor="mm")
    c.alpha_composite(ov)


def gold_divider(c, cx, cy, width=560):
    d = ImageDraw.Draw(c)
    d.line((cx - width // 2, cy, cx - 30, cy), fill=GOLD, width=3)
    d.line((cx + 30, cy, cx + width // 2, cy), fill=GOLD, width=3)
    d.polygon([(cx, cy - 12), (cx + 16, cy), (cx, cy + 12), (cx - 16, cy)], fill=GOLD)


def compass_crest(c, cx, cy, r=58):
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=PRIMARY, outline=GOLD_LT, width=4)
    d.ellipse((cx - r + 10, cy - r + 10, cx + r - 10, cy + r - 10), outline=GOLD_LT, width=2)
    L = r * 0.55
    s = r * 0.18
    d.polygon([(cx, cy - L), (cx + s, cy), (cx, cy + L), (cx - s, cy)], fill=HIGHLIGHT)
    d.polygon([(cx - L, cy), (cx, cy - s), (cx + L, cy), (cx, cy + s)], fill=GOLD_LT)
    d.ellipse((cx - 7, cy - 7, cx + 7, cy + 7), fill=WHITE)
    c.alpha_composite(ov)


def benefit_badge(c, cx, cy, big, small, w=440, h=170, big_color=GOLD_LT):
    box = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    shadow(c, box, 22, 26, 90)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle(box, radius=22, fill=PRIMARY, outline=GOLD_LT, width=3)
    d.text((cx, cy - h * 0.20), big, font=fserif(54), fill=big_color, anchor="mm")
    d.text((cx, cy + h * 0.26), small, font=fs(24), fill=WHITE, anchor="mm")
    c.alpha_composite(ov)


def feature_card(c, x, y, w, h, title, sub):
    box = (x, y, x + w, y + h)
    shadow(c, box, 16, 18, 60)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    d.rounded_rectangle(box, radius=16, fill=WHITE, outline=GOLD_LT, width=2)
    cyc = y + h // 2
    bx = x + 44
    d.ellipse((bx - 30, cyc - 30, bx + 30, cyc + 30), fill=HIGHLIGHT, outline=PRIMARY, width=3)
    d.text((bx, cyc - 1), "✓", font=fs(34), fill=PRIMARY, anchor="mm")
    d.text((x + 92, cyc - 22), title, font=fs(29), fill=PRIMARY, anchor="lm")
    d.text((x + 92, cyc + 24), sub, font=fs(22, bold=False), fill=TEXT_MUTED, anchor="lm")
    c.alpha_composite(ov)


def page_scaffold(img, badge, title, subtitle):
    d = ImageDraw.Draw(img)
    pill(img, SIZE // 2, 175, badge, font=fs(42), pad_x=58, pad_y=24)
    tc(d, (SIZE // 2, 310), title, fserif(70), PRIMARY)
    gold_divider(img, SIZE // 2, 380, width=520)
    tc(d, (SIZE // 2, 430), subtitle, fs(31, bold=False), TEXT_MUTED)


def panel(img, box, fill_color=WHITE, title=None):
    shadow(img, box, 24, 38, 105)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle(box, radius=24, fill=fill_color, outline=GOLD_LT, width=3)
    img.alpha_composite(ov)
    if title:
        ImageDraw.Draw(img).text((box[0] + 34, box[1] + 26), title, font=fs(30), fill=ACCENT, anchor="lt")


def draw_table(img, inner, headers, rows, status_map=None, col_aligns=None,
               header_font=None, row_font=None):
    layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(layer)
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
            if status_map is not None and str(val) in status_map:
                bgc, fgc = status_map[str(val)]
                pw = min(col_w - 16, 180)
                od.rounded_rectangle((cx - pw / 2, y0 + row_h / 2 - 17, cx + pw / 2, y0 + row_h / 2 + 17),
                                     radius=16, fill=bgc)
                od.text((cx, y0 + row_h / 2), str(val), font=fs(18), fill=fgc, anchor="mm")
                continue
            if align == "lm":
                od.text((inner[0] + col_w * ci + 18, y0 + row_h / 2), str(val), font=rf, fill=color, anchor="lm")
            else:
                od.text((cx, y0 + row_h / 2), str(val), font=rf, fill=color, anchor="mm")
    img.alpha_composite(layer)


def dashboard_screen(c, screen, charts_note=True):
    sx0, sy0, sx1, sy1 = screen
    sw = sx1 - sx0
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    head = 84
    d.rounded_rectangle((sx0, sy0, sx1, sy0 + head), radius=8, fill=PRIMARY)
    d.rectangle((sx0, sy0 + 30, sx1, sy0 + head), fill=PRIMARY)
    d.text((sx0 + 28, sy0 + head // 2), "NEXT CHAPTER", font=fs(26), fill=WHITE, anchor="lm")
    d.text((sx1 - 28, sy0 + head // 2), "Life dashboard", font=fs(22), fill=GOLD_LT, anchor="rm")
    d.rectangle((sx0, sy0 + head, sx1, sy0 + head + 4), fill=GOLD_LT)
    gx, gy, gap = sx0 + 22, sy0 + head + 22, 14
    kw = (sw - 22 * 2 - gap * 3) // 4
    kh = 120
    kpis = [
        ("DAYS IN PROCESS", "95", PRIMARY),
        ("TASKS DONE", "5", PRIMARY),
        ("DOCS COLLECTED", "9", PRIMARY),
        ("UPCOMING APPTS", "6", ACCENT),
        ("MONTHLY BUDGET", "$4,200", PRIMARY),
        ("SAVINGS", "$20.9K", PRIMARY),
        ("PARENTING SET", "100%", PRIMARY),
        ("OVERALL", "48%", PRIMARY),
    ]
    for i, (lab, val, col) in enumerate(kpis):
        r, ci = divmod(i, 4)
        x0 = gx + ci * (kw + gap)
        y0 = gy + r * (kh + gap)
        d.rounded_rectangle((x0, y0, x0 + kw, y0 + kh), radius=10, fill=WHITE, outline=(225, 218, 205), width=2)
        d.text((x0 + 14, y0 + 16), lab, font=fs(14), fill=ACCENT, anchor="lt")
        d.text((x0 + 14, y0 + 70), val, font=fserif(34), fill=col, anchor="lm")
    nav_y = gy + 2 * (kh + gap) + 8
    nav_h = 34
    nav = ["Documents", "Finances", "Parenting", "Budget", "Goals", "Contacts"]
    cw = (sw - 22 * 2 - 6 * (len(nav) - 1)) / len(nav)
    for i, name in enumerate(nav):
        x0 = gx + i * (cw + 6)
        d.rounded_rectangle((x0, nav_y, x0 + cw, nav_y + nav_h), radius=17, fill=PRIMARY)
        d.text((x0 + cw / 2, nav_y + nav_h / 2), name, font=fs(15), fill=WHITE, anchor="mm")
    cy0 = nav_y + nav_h + 22
    if charts_note and (sy1 - cy0) > 120:
        cw2 = (sw - 22 * 2 - 14) / 2
        donuts = [
            ("BUDGET BY CATEGORY", [(28, PRIMARY), (18, ACCENT), (16, HIGHLIGHT), (12, SURFACE), (26, TEXT_MUTED)]),
            ("ASSET ALLOCATION", [(33, PRIMARY), (33, ACCENT), (12, HIGHLIGHT), (10, SURFACE), (12, (180, 90, 90))]),
        ]
        for k, (label, segs) in enumerate(donuts):
            bx = gx + k * (cw2 + 14)
            box = (bx, cy0, bx + cw2, sy1 - 18)
            d.rounded_rectangle(box, radius=10, fill=WHITE, outline=(225, 218, 205), width=2)
            d.text((box[0] + 16, box[1] + 14), label, font=fs(18), fill=ACCENT, anchor="lt")
            dcx = (box[0] + box[2]) // 2
            dcy = (box[1] + box[3]) // 2 + 18
            rr = min((box[3] - box[1]) * 0.30, cw2 * 0.20)
            ang = -90
            for pct, col in segs:
                s = pct * 3.6
                d.pieslice((dcx - rr, dcy - rr, dcx + rr, dcy + rr), ang, ang + s, fill=col)
                ang += s
            hole = rr * 0.5
            d.ellipse((dcx - hole, dcy - hole, dcx + hole, dcy + hole), fill=WHITE)
    else:
        d.text((sx0 + sw / 2, cy0 + 6),
               "Documents · Finances · Budget · Parenting · Goals · Analytics — all in one file",
               font=fs(18, bold=False), fill=TEXT_MUTED, anchor="mt")
    c.alpha_composite(ov)


def laptop(c, cx, cy, w, h, charts_note=True):
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    sbox = (cx - w // 2, cy - h // 2 + 30, cx + w // 2, cy + h // 2 + 80)
    sh = Image.new("RGBA", c.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle(sbox, radius=40, fill=(27, 79, 72, 110))
    c.alpha_composite(sh.filter(ImageFilter.GaussianBlur(50)))
    body = (cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2)
    d.rounded_rectangle(body, radius=26, fill=(38, 38, 42), outline=(20, 20, 25), width=4)
    bez = 22
    screen = (body[0] + bez, body[1] + bez, body[2] - bez, body[3] - bez)
    d.rounded_rectangle(screen, radius=8, fill=BG)
    d.rounded_rectangle((cx - w // 2 - 120, cy + h // 2, cx + w // 2 + 120, cy + h // 2 + 24), radius=12, fill=(58, 58, 62))
    d.rounded_rectangle((cx - 80, cy + h // 2 + 4, cx + 80, cy + h // 2 + 14), radius=6, fill=(33, 33, 38))
    c.alpha_composite(ov)
    dashboard_screen(c, screen, charts_note=charts_note)


def render_hero(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    d = ImageDraw.Draw(img)
    compass_crest(img, SIZE // 2, 150, r=58)
    pill(img, SIZE // 2, 285, "DIVORCE ORGANIZATION & LIFE REBUILD SYSTEM", font=fs(32), pad_x=52, pad_y=22)
    tc(d, (SIZE // 2, 432), "NEXT CHAPTER", fserif(146), PRIMARY)
    gold_divider(img, SIZE // 2, 540, width=560)
    tc(d, (SIZE // 2, 600), "Organize documents, finances & schedules — and move forward with calm",
       fs(28, bold=False), TEXT_MUTED)
    laptop(img, SIZE // 2, 1000, w=1360, h=560, charts_note=False)
    benefit_badge(img, 300, 730, "20", "ORGANIZED TABS")
    benefit_badge(img, SIZE - 300, 730, "2-in-1", "EXCEL + GOOGLE SHEETS")
    tc(d, (SIZE // 2, 1360), "EVERYTHING IN ONE COMMAND CENTER", fs(34), ACCENT)
    features = [
        ("Document Vault", "know what you have"),
        ("Task Manager", "nothing slips"),
        ("Financial Snapshot", "net-worth picture"),
        ("Monthly Budget", "plan vs actual"),
        ("Parenting Organizer", "calm logistics"),
        ("Appointments", "auto countdowns"),
        ("Goal Planner", "rebuild forward"),
        ("Life Rebuild", "your new routines"),
    ]
    cols = 4
    margin = 120
    gx, gy = 24, 24
    cw = (SIZE - 2 * margin - (cols - 1) * gx) // cols
    ch = 150
    top = 1430
    for i, (t, s) in enumerate(features):
        r, cc = divmod(i, cols)
        feature_card(img, margin + cc * (cw + gx), top + r * (ch + gy), cw, ch, t, s)
    pill(img, SIZE // 2, SIZE - 120,
         "20 SHEETS · INSTANT DOWNLOAD · EXCEL + GOOGLE SHEETS",
         font=fs(40), pad_x=64, pad_y=30, star=True)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_dashboard(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "LIFE DASHBOARD",
                  "Your Whole Situation, Organized.",
                  "8 clear KPIs + budget & asset charts — calm clarity in one screen")
    box = (130, 510, SIZE - 130, SIZE - 190)
    panel(img, box, fill_color=BG)
    inner = (box[0] + 22, box[1] + 22, box[2] - 22, box[3] - 22)
    dashboard_screen(img, inner, charts_note=True)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_tasks(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "MASTER TASK MANAGER",
                  "Nothing Slips Through the Cracks.",
                  "Every to-do, by category & priority — completion tracks to your dashboard")
    box = (120, 520, SIZE - 120, SIZE - 200)
    panel(img, box, title="YOUR TASKS")
    inner = (box[0] + 28, box[1] + 80, box[2] - 28, box[3] - 28)
    smap = {"Complete": (MINT_BG, PRIMARY), "In Progress": (WARN_BG, ACCENT),
            "Not Started": ((240, 236, 228), TEXT_MUTED)}
    rows = [
        ("Legal", "Consult / retain attorney", "High", "Complete", None),
        ("Documents", "Gather 3 yrs tax returns", "High", "Complete", None),
        ("Financial", "List all accounts", "High", "Complete", None),
        ("Financial", "Open individual account", "High", "In Progress", None),
        ("Parenting", "Draft parenting schedule", "High", "In Progress", None),
        ("Property", "Home contents inventory", "Medium", "Not Started", None),
        ("Insurance", "Review health coverage", "Medium", "Not Started", None),
        ("Legal", "Prepare financial affidavit", "High", "Not Started", None),
    ]
    draw_table(img, inner, ["CATEGORY", "TASK", "PRIORITY", "STATUS"],
               rows, status_map=smap, col_aligns={1: "lm"}, header_font=fs(22), row_font=fs(24))
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_finances(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "FINANCIAL SNAPSHOT",
                  "See Your Whole Financial Picture.",
                  "Assets, debts & net worth — gathered and organized in one place")
    kpis = [("TOTAL ASSETS", "$329K"), ("TOTAL DEBTS", "$242K"),
            ("NET WORTH", "$87K"), ("SAVINGS", "$20.9K")]
    kw, kh, gap = 380, 175, 30
    sx = (SIZE - (4 * kw + 3 * gap)) // 2
    ky = 510
    for i, (lab, val) in enumerate(kpis):
        x = sx + i * (kw + gap)
        b = (x, ky, x + kw, ky + kh)
        shadow(img, b, 20, 24, 80)
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        od.rounded_rectangle(b, radius=20, fill=WHITE, outline=GOLD_LT, width=2)
        od.text((x + kw // 2, ky + 52), lab, font=fs(26), fill=ACCENT, anchor="mm")
        od.text((x + kw // 2, ky + 120), val, font=fserif(56), fill=PRIMARY, anchor="mm")
        img.alpha_composite(ov)
    db = (130, 760, 1000, SIZE - 180)
    panel(img, db, title="ASSET ALLOCATION")
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    dcx = (db[0] + db[2]) // 2 - 110
    dcy = (db[1] + db[3]) // 2 + 30
    rr = 210
    segs = [(33, PRIMARY, "Property 33%"), (33, ACCENT, "Retirement 33%"),
            (12, HIGHLIGHT, "Investments 12%"), (10, SURFACE, "Cash 10%"),
            (12, (180, 90, 90), "Vehicles 12%")]
    ang = -90
    for pct, col, _ in segs:
        s = pct * 3.6
        od.pieslice((dcx - rr, dcy - rr, dcx + rr, dcy + rr), ang, ang + s, fill=col)
        ang += s
    od.ellipse((dcx - 86, dcy - 86, dcx + 86, dcy + 86), fill=WHITE)
    od.text((dcx, dcy - 12), "$329K", font=fserif(34), fill=PRIMARY, anchor="mm")
    od.text((dcx, dcy + 28), "ASSETS", font=fs(18), fill=TEXT_MUTED, anchor="mm")
    lx = dcx + rr + 38
    for i, (pct, col, lab) in enumerate(segs):
        yy = dcy - rr + 30 + i * 56
        od.rounded_rectangle((lx, yy, lx + 26, yy + 26), radius=5, fill=col)
        od.text((lx + 42, yy + 13), lab, font=fs(23), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    bb = (1060, 760, SIZE - 130, SIZE - 180)
    panel(img, bb, title="ASSETS vs DEBTS")
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    bx0, by1 = bb[0] + 120, bb[3] - 90
    by0 = bb[1] + 130
    maxv = 329
    for i, (lab, val, col) in enumerate([("Assets", 329, PRIMARY), ("Debts", 242, ACCENT)]):
        x = bx0 + i * 280
        h = (val / maxv) * (by1 - by0)
        od.rounded_rectangle((x, by1 - h, x + 170, by1), radius=10, fill=col)
        od.text((x + 85, by1 - h - 34), f"${val}K", font=fserif(40), fill=PRIMARY, anchor="mm")
        od.text((x + 85, by1 + 30), lab, font=fs(26), fill=TEXT_MUTED, anchor="mt")
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_analytics(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "PROGRESS ANALYTICS",
                  "Watch Yourself Move Forward.",
                  "Task, document, budget & goal progress — proof of how far you've come")
    cx, cy = 560, 1050
    rr = 270
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.ellipse((cx - rr, cy - rr, cx + rr, cy + rr), fill=WHITE, outline=(230, 224, 212), width=3)
    od.pieslice((cx - rr, cy - rr, cx + rr, cy + rr), -90, -90 + 0.48 * 360, fill=HIGHLIGHT)
    od.ellipse((cx - rr + 60, cy - rr + 60, cx + rr - 60, cy + rr - 60), fill=WHITE)
    od.text((cx, cy - 20), "48%", font=fserif(120), fill=PRIMARY, anchor="mm")
    od.text((cx, cy + 80), "OVERALL PROGRESS", font=fs(26), fill=ACCENT, anchor="mm")
    img.alpha_composite(ov)
    bars = [("Task Completion", 0.42), ("Documents Collected", 0.60),
            ("Budget On Track", 0.78), ("Goal Progress", 0.38)]
    bx = 1080
    bw = SIZE - 130 - bx
    y = 820
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    for lab, pct in bars:
        od.text((bx, y), lab, font=fs(28), fill=PRIMARY, anchor="lt")
        od.rounded_rectangle((bx, y + 48, bx + bw, y + 92), radius=22, fill=(238, 232, 222))
        od.rounded_rectangle((bx, y + 48, bx + bw * pct, y + 92), radius=22, fill=HIGHLIGHT)
        od.text((bx + bw - 10, y + 8), f"{int(pct*100)}%", font=fs(26), fill=ACCENT, anchor="rt")
        y += 150
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_mobile(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    dotted_bg(img)
    page_scaffold(img, "WORKS EVERYWHERE",
                  "Excel · Google Sheets · Mobile",
                  "Update from anywhere — your command center is always with you")
    px, py = SIZE // 2, 1300
    pw, ph = 640, 1230
    phone = (px - pw // 2, py - ph // 2, px + pw // 2, py + ph // 2)
    shadow(img, phone, 64, 50, 120)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    od.rounded_rectangle(phone, radius=64, fill=(24, 24, 28))
    bez = 22
    screen = (phone[0] + bez, phone[1] + bez + 30, phone[2] - bez, phone[3] - bez - 30)
    od.rounded_rectangle(screen, radius=44, fill=BG)
    od.rounded_rectangle((px - 95, phone[1] + 16, px + 95, phone[1] + 50), radius=18, fill=(14, 14, 18))
    sx0, sy0, sx1, sy1 = screen
    od.rounded_rectangle((sx0, sy0, sx1, sy0 + 110), radius=44, fill=PRIMARY)
    od.rectangle((sx0, sy0 + 60, sx1, sy0 + 110), fill=PRIMARY)
    od.rectangle((sx0, sy0 + 106, sx1, sy0 + 110), fill=GOLD_LT)
    od.text(((sx0 + sx1) // 2, sy0 + 56), "Next Chapter", font=fserif(38), fill=GOLD_LT, anchor="mm")
    img.alpha_composite(ov)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    y = sy0 + 150
    cards = [("DAYS IN PROCESS", "95", PRIMARY), ("OVERALL PROGRESS", "48%", PRIMARY),
             ("TASKS DONE", "5 / 12", ACCENT), ("DOCS COLLECTED", "9 / 15", PRIMARY)]
    for lab, val, col in cards:
        cb = (sx0 + 30, y, sx1 - 30, y + 135)
        od.rounded_rectangle(cb, radius=18, fill=WHITE, outline=GOLD_LT, width=2)
        od.text((cb[0] + 26, y + 30), lab, font=fs(22), fill=ACCENT, anchor="lt")
        od.text((cb[0] + 26, y + 92), val, font=fserif(46), fill=col, anchor="lm")
        y += 152
    od.text((sx0 + 40, y + 16), "TODAY'S FOCUS", font=fs(22), fill=ACCENT, anchor="lt")
    y += 52
    for lab, state in [("Open individual account", True), ("Draft parenting plan", True),
                       ("Upload mortgage statement", False), ("Call financial advisor", False)]:
        col = HIGHLIGHT if state else BG
        od.ellipse((sx0 + 40, y + 6, sx0 + 78, y + 44), fill=col, outline=PRIMARY, width=3)
        if state:
            od.text((sx0 + 59, y + 24), "✓", font=fs(24), fill=PRIMARY, anchor="mm")
        od.text((sx0 + 96, y + 24), lab, font=fs(23), fill=TEXT, anchor="lm")
        y += 64
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def main():
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("01_hero.png", render_hero),
        ("02_dashboard.png", render_dashboard),
        ("03_tasks.png", render_tasks),
        ("04_finances.png", render_finances),
        ("05_analytics.png", render_analytics),
        ("06_mobile.png", render_mobile),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} images to {out_dir}")


if __name__ == "__main__":
    main()
