"""Detailed marketing image set for Restaurant Command Center™ (4 images, 2000x2000).

Deeper, benefit-driven marketing beyond the app screenshots — the kind of
detailed listing images that convert browsers into buyers:

  07_features.png   - feature spotlights (what it does + why it matters)
  08_compare.png    - "basic spreadsheet" vs Restaurant Command Center™
  09_howitworks.png - up & running in 4 steps + compatibility
  10_value.png      - what's included, who it's for & the guarantee

Reuses the branded helpers from build_marketing.py.
Run: python3 build_marketing_detail.py
"""
from __future__ import annotations
import os
from PIL import Image, ImageDraw, ImageFilter

from build_marketing import (
    SIZE, PRIMARY, PRIMARY_DK, PRIMARY_LT, ACCENT, GOLD, GOLD_LT, GOLD_HI, SURFACE,
    HIGHLIGHT, BG, WHITE, TEXT, TEXT_MUTED, DANGER, MINT_BG, WARN_BG, RED_BG, GRID,
    ROW_ALT, fs, fserif, premium_bg, pill, wordmark, gold_divider, chef_crest,
    shadow, grad_round, radial_glow, tc, donut, legend, hbars, fit_font, vgradient,
)


def wrap(d, text, font, max_w):
    words = text.split(); lines = []; cur = ""
    for w in words:
        t = (cur + " " + w).strip()
        if d.textlength(t, font=font) <= max_w:
            cur = t
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def card(img, box, radius=22, blur=16, dy=10, outline=(232, 224, 208)):
    shadow(img, box, radius, blur, 44, dy)
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle(box, radius=radius, fill=WHITE, outline=outline, width=2)
    od.rounded_rectangle((box[0], box[1] + 8, box[0] + 7, box[3] - 8), radius=3, fill=GOLD_LT)
    img.alpha_composite(ov)


def icon_badge(d, cx, cy, r, emoji_lines):
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=PRIMARY)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), outline=GOLD_HI, width=3)


def mini_donut(img, cx, cy, r, pct, label, sub):
    d = ImageDraw.Draw(img)
    donut(d, cx, cy, r, [(pct, PRIMARY), (100 - pct, SURFACE)], label, sub)


def header_band(img, tagpill, title, subtitle, band_h=380, crest=True):
    premium_bg(img, band_h=band_h)
    d = ImageDraw.Draw(img)
    if crest:
        chef_crest(img, SIZE // 2, 110, r=48)
        pill(img, SIZE // 2, 214, tagpill, font=fs(24), pad_x=40, pad_y=18)
        tc(d, (SIZE // 2, 300), title, fserif(56), WHITE)
        gold_divider(img, SIZE // 2, 360, width=480)
    else:
        pill(img, SIZE // 2, 120, tagpill, font=fs(34), pad_x=52, pad_y=22)
        tc(d, (SIZE // 2, 238), title, fserif(58), WHITE)
        gold_divider(img, SIZE // 2, 308, width=480)
    if subtitle:
        yy = 400 if crest else 352
        tc(d, (SIZE // 2, yy), subtitle, fs(23, bold=False), (226, 214, 190))


# ===========================================================================
def render_features(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "BUILT TO PROTECT YOUR MARGINS", "Where the Money Is Won",
                "Four engines that turn a busy kitchen into a profitable business", band_h=430)
    d = ImageDraw.Draw(img)
    feats = [
        ("$", "Prime Cost Engine", "donut", (65, "65%", "prime"),
         "Food + labor is 60-65% of every restaurant dollar. Track it live against your target — the single number that predicts profit."),
        ("%", "Menu Engineering", "table", None,
         "See food cost % and contribution margin on every dish. Instantly spot the plates to reprice, re-portion or cut."),
        ("■", "Inventory & Par", "stock", None,
         "Every item valued on-hand vs par, with automatic low-stock flags — so you reorder before you 86 a dish, not after."),
        ("★", "Operations Health", "bars", None,
         "One 0-100% score blends food, labor & prime cost, margin, guest ratings and stock levels — your whole operation at a glance."),
    ]
    margin = 84; gap = 28
    cw = (SIZE - 2 * margin - gap) // 2
    ch = 560
    top = 500
    for i, (emoji, title, kind, dat, body) in enumerate(feats):
        r, c = divmod(i, 2)
        x = margin + c * (cw + gap); y = top + r * (ch + gap)
        card(img, (x, y, x + cw, y + ch))
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        # header
        bx, by = x + 66, y + 66
        od.ellipse((bx - 40, by - 40, bx + 40, by + 40), fill=PRIMARY, outline=GOLD_HI, width=3)
        if kind == "stock":
            od.rounded_rectangle((bx - 22, by - 14, bx + 22, by + 20), radius=4, outline=GOLD_HI, width=4)
            od.line((bx - 22, by, bx + 22, by), fill=GOLD_HI, width=3)
            od.line((bx, by - 14, bx, by), fill=GOLD_HI, width=3)
        elif kind == "bars":
            import math as _m
            pts = []
            for k in range(10):
                a = -_m.pi / 2 + k * _m.pi / 5
                rr = 26 if k % 2 == 0 else 11
                pts.append((bx + _m.cos(a) * rr, by + _m.sin(a) * rr))
            od.polygon(pts, fill=GOLD_HI)
        else:
            od.text((bx, by - 3), emoji, font=fserif(40), fill=GOLD_HI, anchor="mm")
        od.text((x + 124, by - 16), title, font=fs(30), fill=PRIMARY, anchor="lm")
        od.text((x + 124, by + 22), "AUTOMATIC", font=fs(15), fill=ACCENT, anchor="lm")
        # body copy
        for li, line in enumerate(wrap(od, body, fs(20, bold=False), cw - 80)):
            od.text((x + 40, y + 148 + li * 32), line, font=fs(20, bold=False), fill=TEXT, anchor="lt")
        img.alpha_composite(ov)
        # mini visual
        vy = y + ch - 180
        if kind == "donut":
            mini_donut(img, x + cw // 2, vy + 30, 96, dat[0], dat[1], dat[2])
        elif kind == "table":
            od = ImageDraw.Draw(img)
            rows = [("Ribeye", "40%", "$27.60"), ("Chicken", "27%", "$20.40"), ("Risotto", "25%", "$19.60")]
            tx0 = x + 44; tx1 = x + cw - 44; ty = vy - 30
            grad_round(img, (tx0, ty, tx1, ty + 34), 6, PRIMARY_LT, PRIMARY_DK)
            for h, fx in zip(["DISH", "FC %", "MARGIN"], [0.0, 0.5, 0.78]):
                od.text((tx0 + 14 + (tx1 - tx0) * fx, ty + 17), h, font=fs(15), fill=WHITE, anchor="lm")
            for ri, (a, b, cc) in enumerate(rows):
                ry = ty + 34 + ri * 42
                if ri % 2:
                    od.rectangle((tx0, ry, tx1, ry + 42), fill=ROW_ALT)
                od.text((tx0 + 14, ry + 21), a, font=fs(19), fill=PRIMARY, anchor="lm")
                od.text((tx0 + 14 + (tx1 - tx0) * 0.5, ry + 21), b, font=fs(18, bold=False), fill=TEXT, anchor="lm")
                od.text((tx0 + 14 + (tx1 - tx0) * 0.78, ry + 21), cc, font=fs(18), fill=PRIMARY, anchor="lm")
        elif kind == "stock":
            od = ImageDraw.Draw(img)
            items = [("Ribeye (dry-aged)", "Low", RED_BG, DANGER), ("Mixed Greens", "Low", RED_BG, DANGER),
                     ("House Red Wine", "OK", MINT_BG, PRIMARY), ("Frozen Fries", "Low", RED_BG, DANGER)]
            tx0 = x + 44; tx1 = x + cw - 44; ty = vy - 30
            for ri, (nm, st, bg, fg) in enumerate(items):
                ry = ty + ri * 44
                od.rounded_rectangle((tx0, ry, tx1, ry + 38), radius=8, fill=(246, 241, 232))
                od.text((tx0 + 16, ry + 19), nm, font=fs(18, bold=False), fill=TEXT, anchor="lm")
                od.rounded_rectangle((tx1 - 96, ry + 6, tx1 - 12, ry + 32), radius=13, fill=bg)
                od.text((tx1 - 54, ry + 19), st, font=fs(15), fill=fg, anchor="mm")
        elif kind == "bars":
            hbars(img, ImageDraw.Draw(img), (x + 44, y + 300, x + cw - 44, y + ch - 30),
                  [("Food cost", 0.93, "93%"), ("Labor cost", 0.98, "98%"),
                   ("Prime cost", 0.92, "92%"), ("Guest rating", 0.92, "92%")])
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
def render_compare(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "WHY IT'S DIFFERENT", "More Than a Spreadsheet",
                "Everything a basic template leaves out — built in", band_h=430)
    rows = [
        ("Tracks daily sales & covers", True, True),
        ("Food cost % on every menu item", False, True),
        ("Live prime cost (food + labor)", False, True),
        ("Automatic P&L & net margin", False, True),
        ("Inventory value + low-stock alerts", False, True),
        ("Par levels & reorder guide", False, True),
        ("Labor cost by role vs sales", False, True),
        ("Vendor payments & AP due dates", False, True),
        ("Compliance & temp-log tracking", False, True),
        ("Reviews, waste & training", False, True),
        ("One Operations Health Score", False, True),
    ]
    margin = 110; top = 500
    tw = SIZE - 2 * margin
    c1 = margin + tw * 0.60; c2 = margin + tw * 0.83
    d = ImageDraw.Draw(img)
    # header row
    hh = 74
    grad_round(img, (margin, top, SIZE - margin, top + hh), 12, PRIMARY_LT, PRIMARY_DK)
    d.text((margin + 24, top + hh // 2), "FEATURE", font=fs(22), fill=WHITE, anchor="lm")
    d.text((c1, top + hh // 2), "Basic\nTemplate", font=fs(18), fill=(210, 210, 205), anchor="mm")
    d.text((c2, top + hh // 2), "Command\nCenter™", font=fs(18), fill=GOLD_HI, anchor="mm")
    rh = 108
    for i, (label, a, b) in enumerate(rows):
        ry = top + hh + i * rh
        if i % 2:
            d.rounded_rectangle((margin, ry, SIZE - margin, ry + rh), radius=8, fill=ROW_ALT)
        d.text((margin + 24, ry + rh // 2), label, font=fs(23, bold=False), fill=TEXT, anchor="lm")
        # basic: X or check
        if a:
            d.ellipse((c1 - 22, ry + rh // 2 - 22, c1 + 22, ry + rh // 2 + 22), outline=(180, 175, 165), width=3)
            d.line((c1 - 10, ry + rh // 2, c1 - 2, ry + rh // 2 + 10), fill=(150, 145, 135), width=4)
            d.line((c1 - 2, ry + rh // 2 + 10, c1 + 12, ry + rh // 2 - 8), fill=(150, 145, 135), width=4)
        else:
            d.ellipse((c1 - 22, ry + rh // 2 - 22, c1 + 22, ry + rh // 2 + 22), fill=RED_BG)
            d.line((c1 - 9, ry + rh // 2 - 9, c1 + 9, ry + rh // 2 + 9), fill=DANGER, width=4)
            d.line((c1 - 9, ry + rh // 2 + 9, c1 + 9, ry + rh // 2 - 9), fill=DANGER, width=4)
        # command center: check on mint
        d.ellipse((c2 - 24, ry + rh // 2 - 24, c2 + 24, ry + rh // 2 + 24), fill=PRIMARY)
        d.line((c2 - 11, ry + rh // 2 + 1, c2 - 3, ry + rh // 2 + 11), fill=HIGHLIGHT, width=5)
        d.line((c2 - 3, ry + rh // 2 + 11, c2 + 13, ry + rh // 2 - 10), fill=HIGHLIGHT, width=5)
    # bottom column highlight frame
    fy0 = top; fy1 = top + hh + len(rows) * rh
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle((c2 - 90, fy0 - 8, c2 + 90, fy1 + 8), radius=16, outline=GOLD_LT, width=4)
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
def render_howitworks(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "SET UP IN MINUTES", "Up & Running in 4 Steps",
                "No accounting degree required — if you can type, you can run it", band_h=430)
    steps = [
        ("1", "Download & open", "Buy once and instantly download your Excel file, plus a 1-click link to copy the Google Sheets version."),
        ("2", "Set your targets", "In Settings, enter your restaurant name, seats and cost targets (food, labor, prime, margin). One time only."),
        ("3", "Cost your menu & load inventory", "Enter each dish's price & food cost, then your inventory on-hand and par levels. Percentages and flags calculate themselves."),
        ("4", "Log sales & watch the dashboard", "Add daily sales and labor — prime cost, net profit and the Operations Health Score update automatically, every day."),
    ]
    margin = 100; top = 500
    sh = 300; gap = 24
    w = SIZE - 2 * margin
    d = ImageDraw.Draw(img)
    for i, (num, title, body) in enumerate(steps):
        y = top + i * (sh + gap)
        card(img, (margin, y, SIZE - margin, y + sh))
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        # number circle
        cx, cy = margin + 130, y + sh // 2
        od.ellipse((cx - 70, cy - 70, cx + 70, cy + 70), fill=PRIMARY, outline=GOLD_HI, width=4)
        od.text((cx, cy - 4), num, font=fserif(72), fill=GOLD_HI, anchor="mm")
        # text
        tx = margin + 270
        od.text((tx, cy - 66), title, font=fs(34), fill=PRIMARY, anchor="lt")
        for li, line in enumerate(wrap(od, body, fs(22, bold=False), SIZE - margin - tx - 50)):
            od.text((tx, cy - 14 + li * 34), line, font=fs(22, bold=False), fill=TEXT, anchor="lt")
        img.alpha_composite(ov)
        if i < len(steps) - 1:
            ad = ImageDraw.Draw(img)
            ax = margin + 110
            ad.line((ax, y + sh + 2, ax, y + sh + gap - 2), fill=GOLD_LT, width=4)
    img.convert("RGB").save(out, "PNG", optimize=True)


# ===========================================================================
def render_value(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "EVERYTHING YOU GET", "One File, Your Whole Restaurant",
                "Instant download · lifetime access · free updates", band_h=430)
    d = ImageDraw.Draw(img)
    margin = 84
    # left: what's included
    lx0, lx1 = margin, SIZE // 2 - 14
    card(img, (lx0, 500, lx1, 1300))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((lx0 + 40, 540), "WHAT'S INCLUDED", font=fs(26), fill=ACCENT, anchor="lt")
    includes = [
        "22-tab Excel workbook (.xlsx)", "Google Sheets edition (1-click copy)",
        "12-KPI operations dashboard", "Menu & recipe costing engine",
        "Live P&L / prime cost tracker", "Inventory + auto low-stock alerts",
        "Labor, sales & waste trackers", "Reviews, compliance & AP tools",
        "Start-Here quick-start guide (PDF)", "Free lifetime updates",
    ]
    for i, item in enumerate(includes):
        yy = 600 + i * 64
        od.ellipse((lx0 + 44, yy, lx0 + 76, yy + 32), fill=PRIMARY)
        od.line((lx0 + 52, yy + 16, lx0 + 59, yy + 24), fill=HIGHLIGHT, width=4)
        od.line((lx0 + 59, yy + 24, lx0 + 70, yy + 9), fill=HIGHLIGHT, width=4)
        od.text((lx0 + 96, yy + 16), item, font=fs(22, bold=False), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    # right top: who it's for
    rx0, rx1 = SIZE // 2 + 14, SIZE - margin
    card(img, (rx0, 500, rx1, 890))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((rx0 + 40, 540), "WHO IT'S FOR", font=fs(26), fill=ACCENT, anchor="lt")
    who = ["Full-service restaurants", "Cafés & coffee shops", "Bars & pubs",
           "Food trucks & ghost kitchens", "Caterers", "Multi-unit & franchise GMs"]
    for i, w in enumerate(who):
        r, c = divmod(i, 1)
        yy = 604 + i * 44
        od.text((rx0 + 44, yy), "•", font=fs(24), fill=GOLD_LT, anchor="lm")
        od.text((rx0 + 72, yy), w, font=fs(22, bold=False), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    # right bottom: compatibility + guarantee
    card(img, (rx0, 916, rx1, 1300))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((rx0 + 40, 956), "WORKS WITH", font=fs(26), fill=ACCENT, anchor="lt")
    for i, t in enumerate(["✓ Microsoft Excel (2019 / 2021 / 365)", "✓ Google Sheets (free account)",
                            "✓ Windows, Mac, tablet & mobile"]):
        od.text((rx0 + 44, 1020 + i * 44), t, font=fs(21, bold=False), fill=TEXT, anchor="lm")
    od.rounded_rectangle((rx0 + 40, 1160, rx1 - 40, 1264), radius=14, fill=MINT_BG)
    od.text((rx0 + 64, 1188), "★  Not working right?", font=fs(21), fill=PRIMARY, anchor="lm")
    od.text((rx0 + 64, 1228), "Message me — I help within 24 hours.", font=fs(19, bold=False), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    # bottom banner: value line
    pill(img, SIZE // 2, 1420, "REPLACES $1,000s IN SOFTWARE & CONSULTANTS", font=fs(30), pad_x=48, pad_y=22,
         grad=(PRIMARY_LT, PRIMARY_DK))
    d = ImageDraw.Draw(img)
    tc(d, (SIZE // 2, 1530), "A single point of control for the numbers that decide if you make money.",
       fs(24, bold=False), TEXT_MUTED)
    # closing crest + wordmark
    chef_crest(img, SIZE // 2, 1700, r=54)
    wordmark(img, SIZE // 2, 1850, "RESTAURANT COMMAND CENTER", 60, max_w=1500)
    img.convert("RGB").save(out, "PNG", optimize=True)


def main():
    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing")
    os.makedirs(out_dir, exist_ok=True)
    targets = [
        ("07_features.png", render_features),
        ("08_compare.png", render_compare),
        ("09_howitworks.png", render_howitworks),
        ("10_value.png", render_value),
    ]
    for name, fn in targets:
        fn(os.path.join(out_dir, name))
        print(f"  ✓ {name}")
    print(f"Wrote {len(targets)} detailed images to {out_dir}")


if __name__ == "__main__":
    main()
