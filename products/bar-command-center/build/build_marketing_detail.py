"""Detailed marketing image set for Bar & Pub Command Center™ (07-10)."""
from __future__ import annotations
import os
from PIL import Image, ImageDraw

from build_marketing import (
    SIZE, PRIMARY, PRIMARY_DK, PRIMARY_LT, ACCENT, GOLD, GOLD_LT, GOLD_HI, SURFACE,
    HIGHLIGHT, BG, WHITE, TEXT, TEXT_MUTED, DANGER, MINT_BG, WARN_BG, RED_BG, GRID,
    ROW_ALT, fs, fserif, premium_bg, pill, wordmark, gold_divider, book_crest,
    shadow, grad_round, radial_glow, tc, donut, legend, hbars, fit_font, vgradient,
)

PRINT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "marketing", "print")


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


def mini_donut(img, cx, cy, r, pct, label, sub):
    d = ImageDraw.Draw(img)
    donut(d, cx, cy, r, [(pct, PRIMARY), (100 - pct, SURFACE)], label, sub)


def header_band(img, tagpill, title, subtitle, band_h=430):
    premium_bg(img, band_h=band_h)
    d = ImageDraw.Draw(img)
    book_crest(img, SIZE // 2, 110, r=48)
    pill(img, SIZE // 2, 214, tagpill, font=fs(24), pad_x=40, pad_y=18)
    tc(d, (SIZE // 2, 300), title, fserif(54), WHITE)
    gold_divider(img, SIZE // 2, 360, width=480)
    if subtitle:
        tc(d, (SIZE // 2, 400), subtitle, fs(23, bold=False), (226, 214, 190))


def render_features(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "BUILT FOR BARS", "Where a Bar Makes Money",
                "Four ways it turns every pour into real profit")
    feats = [
        ("$", "Live Bar Score", "donut", (90, "90%", "healthy"),
         "Pour cost, margins, inventory variance, happy-hour margin & gross margin roll into one number — so you know your bar's financial health at a glance."),
        ("%", "Cost by the Pour", "table", None,
         "Cost each drink to the pour — spirit, mixer & garnish — for a true pour cost and pour-cost %. The engine behind a menu that actually profits."),
        ("box", "Keg & Draft Profit", "stock", None,
         "Every keg tracked to a cost per pint and a profit per keg — draft is your highest-margin pour, and this proves it barrel by barrel."),
        ("star", "12 Printable Pages", "print", None,
         "A matching print-ready pack — pour-cost card, keg & draft log, inventory variance sheet, liquor count sheet & an open/close checklist."),
    ]
    margin = 84; gap = 28
    cw = (SIZE - 2 * margin - gap) // 2
    ch = 560
    top = 500
    for i, (badge, title, kind, dat, body) in enumerate(feats):
        r, c = divmod(i, 2)
        x = margin + c * (cw + gap); y = top + r * (ch + gap)
        card(img, (x, y, x + cw, y + ch))
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        bx, by = x + 66, y + 66
        od.ellipse((bx - 40, by - 40, bx + 40, by + 40), fill=PRIMARY, outline=GOLD_HI, width=3)
        if kind == "table":
            od.rounded_rectangle((bx - 22, by - 16, bx + 22, by + 16), radius=5, outline=GOLD_HI, width=4)
            od.line((bx - 22, by - 2, bx + 22, by - 2), fill=GOLD_HI, width=3)
        elif kind == "stock":
            od.rounded_rectangle((bx - 22, by - 14, bx + 22, by + 20), radius=4, outline=GOLD_HI, width=4)
            od.line((bx - 22, by, bx + 22, by), fill=GOLD_HI, width=3)
        elif kind == "print":
            od.rounded_rectangle((bx - 16, by - 20, bx + 16, by + 20), radius=4, fill=GOLD_HI)
            for k in range(3):
                od.line((bx - 8, by - 8 + k * 8, bx + 8, by - 8 + k * 8), fill=PRIMARY, width=2)
        else:
            od.text((bx, by - 3), badge, font=fserif(40), fill=GOLD_HI, anchor="mm")
        od.text((x + 124, by - 16), title, font=fs(25), fill=PRIMARY, anchor="lm")
        od.text((x + 124, by + 22), "IN THE SYSTEM", font=fs(15), fill=ACCENT, anchor="lm")
        for li, line in enumerate(wrap(od, body, fs(20, bold=False), cw - 80)):
            od.text((x + 40, y + 148 + li * 32), line, font=fs(20, bold=False), fill=TEXT, anchor="lt")
        img.alpha_composite(ov)
        vy = y + ch - 180
        if kind == "donut":
            mini_donut(img, x + cw // 2, vy + 30, 96, dat[0], dat[1], dat[2])
        elif kind == "table":
            od = ImageDraw.Draw(img)
            rows = [("Pour cost", "$2.20"), ("Menu price", "$13.00"), ("Pour-cost %", "17%")]
            tx0 = x + 44; tx1 = x + cw - 44; ty = vy - 30
            grad_round(img, (tx0, ty, tx1, ty + 34), 6, PRIMARY_LT, PRIMARY_DK)
            for h, fx in zip(["OLD FASHIONED", "VALUE"], [0.0, 0.62]):
                od.text((tx0 + 14 + (tx1 - tx0) * fx, ty + 17), h, font=fs(15), fill=WHITE, anchor="lm")
            for ri, (a, b) in enumerate(rows):
                ry = ty + 34 + ri * 42
                if ri % 2:
                    od.rectangle((tx0, ry, tx1, ry + 42), fill=ROW_ALT)
                od.text((tx0 + 14, ry + 21), a, font=fs(17), fill=PRIMARY, anchor="lm")
                od.text((tx0 + 14 + (tx1 - tx0) * 0.62, ry + 21), b, font=fs(16, bold=False), fill=TEXT, anchor="lm")
        elif kind == "stock":
            od = ImageDraw.Draw(img)
            items = [("Weekly drink sales", "$16,940", MINT_BG, PRIMARY), ("Avg profit / keg", "$571", MINT_BG, PRIMARY),
                     ("Happy-hour margin", "79%", SURFACE, (110, 88, 58)), ("Weekly units poured", "1,870", MINT_BG, PRIMARY)]
            tx0 = x + 44; tx1 = x + cw - 44; ty = vy - 30
            for ri, (nm, st, bg, fg) in enumerate(items):
                ry = ty + ri * 44
                od.rounded_rectangle((tx0, ry, tx1, ry + 38), radius=8, fill=(246, 241, 232))
                od.text((tx0 + 16, ry + 19), nm, font=fs(18, bold=False), fill=TEXT, anchor="lm")
                od.rounded_rectangle((tx1 - 120, ry + 6, tx1 - 12, ry + 32), radius=13, fill=bg)
                od.text((tx1 - 66, ry + 19), st, font=fs(15), fill=fg, anchor="mm")
        elif kind == "print":
            tw = cw - 88; pw = (tw - 4 * 16) / 5; ph = int(pw * 11 / 8.5)
            px0 = x + 44; py0 = vy - 40
            for k in range(5):
                tp = os.path.join(PRINT_DIR, f"page_{k+1:02d}.png")
                px = px0 + k * (pw + 16)
                if os.path.exists(tp):
                    thumb = Image.open(tp).convert("RGB").resize((int(pw), int(ph)))
                    img.paste(thumb, (int(px), int(py0)))
                    ImageDraw.Draw(img).rounded_rectangle((px, py0, px + pw, py0 + ph), radius=4, outline=(210, 203, 190), width=2)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_compare(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "WHY IT'S DIFFERENT", "More Than a Sales Log",
                "Everything a basic bar sheet leaves out — built in")
    rows = [
        ("A weekly sales log", True, True),
        ("Pour-cost engine per drink", False, True),
        ("Pour-cost % on every drink", False, True),
        ("Keg & draft profit per keg", False, True),
        ("Inventory variance / shrinkage", False, True),
        ("Liquor count & par tracker", False, True),
        ("Happy-hour margin planner", False, True),
        ("Waste & spill log", False, True),
        ("Ordering, cash & events", False, True),
        ("Live Bar Score dashboard", False, True),
        ("Google Sheets AND printable PDF", False, True),
    ]
    margin = 110; top = 500
    tw = SIZE - 2 * margin
    c1 = margin + tw * 0.60; c2 = margin + tw * 0.83
    d = ImageDraw.Draw(img)
    hh = 74
    grad_round(img, (margin, top, SIZE - margin, top + hh), 12, PRIMARY_LT, PRIMARY_DK)
    d.text((margin + 24, top + hh // 2), "FEATURE", font=fs(22), fill=WHITE, anchor="lm")
    d.text((c1, top + hh // 2), "A Basic\nSales Log", font=fs(18), fill=(210, 210, 205), anchor="mm")
    d.text((c2, top + hh // 2), "Command\nCenter™", font=fs(18), fill=GOLD_HI, anchor="mm")
    rh = 108
    for i, (label, a, b) in enumerate(rows):
        ry = top + hh + i * rh
        if i % 2:
            d.rounded_rectangle((margin, ry, SIZE - margin, ry + rh), radius=8, fill=ROW_ALT)
        d.text((margin + 24, ry + rh // 2), label, font=fs(23, bold=False), fill=TEXT, anchor="lm")
        if a:
            d.ellipse((c1 - 22, ry + rh // 2 - 22, c1 + 22, ry + rh // 2 + 22), outline=(180, 175, 165), width=3)
            d.line((c1 - 10, ry + rh // 2, c1 - 2, ry + rh // 2 + 10), fill=(150, 145, 135), width=4)
            d.line((c1 - 2, ry + rh // 2 + 10, c1 + 12, ry + rh // 2 - 8), fill=(150, 145, 135), width=4)
        else:
            d.ellipse((c1 - 22, ry + rh // 2 - 22, c1 + 22, ry + rh // 2 + 22), fill=RED_BG)
            d.line((c1 - 9, ry + rh // 2 - 9, c1 + 9, ry + rh // 2 + 9), fill=DANGER, width=4)
            d.line((c1 - 9, ry + rh // 2 + 9, c1 + 9, ry + rh // 2 - 9), fill=DANGER, width=4)
        d.ellipse((c2 - 24, ry + rh // 2 - 24, c2 + 24, ry + rh // 2 + 24), fill=PRIMARY)
        d.line((c2 - 11, ry + rh // 2 + 1, c2 - 3, ry + rh // 2 + 11), fill=HIGHLIGHT, width=5)
        d.line((c2 - 3, ry + rh // 2 + 11, c2 + 13, ry + rh // 2 - 10), fill=HIGHLIGHT, width=5)
    fy0 = top; fy1 = top + hh + len(rows) * rh
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.rounded_rectangle((c2 - 90, fy0 - 8, c2 + 90, fy1 + 8), radius=16, outline=GOLD_LT, width=4)
    img.alpha_composite(ov)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_howitworks(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "SET UP IN AN AFTERNOON", "Price Your Bar in 4 Steps",
                "No spreadsheet skills needed — if you can type, you can run it")
    steps = [
        ("1", "Cost your pours", "Buy once and instantly download the Google Sheets system plus the 12-page printable pack, then cost each drink to the pour — spirit, mixer & garnish."),
        ("2", "Build your menu", "See pour-cost % and margin on every drink, and profit per keg on every draft — price for a healthy pour instead of guessing."),
        ("3", "Kill the shrinkage", "Count liquor against par, log waste & spill, and run inventory variance — the gap between theoretical and actual is profit walking out."),
        ("4", "Run the bar", "Plan happy hour, track sales, cash, ordering & events, and watch your Bar Score climb."),
    ]
    margin = 100; top = 500
    sh = 300; gap = 24
    for i, (num, title, body) in enumerate(steps):
        y = top + i * (sh + gap)
        card(img, (margin, y, SIZE - margin, y + sh))
        ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
        cx, cy = margin + 130, y + sh // 2
        od.ellipse((cx - 70, cy - 70, cx + 70, cy + 70), fill=PRIMARY, outline=GOLD_HI, width=4)
        od.text((cx, cy - 4), num, font=fserif(72), fill=GOLD_HI, anchor="mm")
        tx = margin + 270
        od.text((tx, cy - 66), title, font=fs(34), fill=PRIMARY, anchor="lt")
        for li, line in enumerate(wrap(od, body, fs(22, bold=False), SIZE - margin - tx - 50)):
            od.text((tx, cy - 14 + li * 34), line, font=fs(22, bold=False), fill=TEXT, anchor="lt")
        img.alpha_composite(ov)
        if i < len(steps) - 1:
            ImageDraw.Draw(img).line((margin + 130, y + sh + 2, margin + 130, y + sh + gap - 2), fill=GOLD_LT, width=4)
    img.convert("RGB").save(out, "PNG", optimize=True)


def render_value(out):
    img = Image.new("RGBA", (SIZE, SIZE), BG + (255,))
    header_band(img, "EVERYTHING YOU GET", "One System, Whole Bar",
                "Instant download · lifetime access · free updates")
    d = ImageDraw.Draw(img)
    margin = 84
    lx0, lx1 = margin, SIZE // 2 - 14
    card(img, (lx0, 500, lx1, 1300))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((lx0 + 40, 540), "WHAT'S INCLUDED", font=fs(26), fill=ACCENT, anchor="lt")
    includes = [
        "14-tab Google Sheets system", "Excel (.xlsx) edition",
        "12-page printable PDF pack", "Live dashboard + Bar Score",
        "Pour-cost-per-drink engine", "Keg & draft profit tracker",
        "Inventory variance & liquor count", "Happy-hour margin planner",
        "Waste, sales, cash & events", "Start-Here guide + free updates",
    ]
    for i, item in enumerate(includes):
        yy = 600 + i * 64
        od.ellipse((lx0 + 44, yy, lx0 + 76, yy + 32), fill=PRIMARY)
        od.line((lx0 + 52, yy + 16, lx0 + 59, yy + 24), fill=HIGHLIGHT, width=4)
        od.line((lx0 + 59, yy + 24, lx0 + 70, yy + 9), fill=HIGHLIGHT, width=4)
        od.text((lx0 + 96, yy + 16), item, font=fs(22, bold=False), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    rx0, rx1 = SIZE // 2 + 14, SIZE - margin
    card(img, (rx0, 500, rx1, 890))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((rx0 + 40, 540), "WHO IT'S FOR", font=fs(26), fill=ACCENT, anchor="lt")
    who = ["Bars, pubs & taprooms", "Cocktail & wine bars", "Breweries & beer gardens",
           "Restaurant bar programs", "Nightclubs & lounges", "Anyone pouring for profit"]
    for i, w in enumerate(who):
        yy = 604 + i * 44
        od.text((rx0 + 44, yy), "•", font=fs(24), fill=GOLD_LT, anchor="lm")
        od.text((rx0 + 72, yy), w, font=fs(20, bold=False), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    card(img, (rx0, 916, rx1, 1300))
    ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); od = ImageDraw.Draw(ov)
    od.text((rx0 + 40, 956), "WORKS WITH", font=fs(26), fill=ACCENT, anchor="lt")
    for i, t in enumerate(["✓ Google Sheets (free account)", "✓ Microsoft Excel (2019 / 2021 / 365)",
                            "✓ Print at home — US Letter PDF"]):
        od.text((rx0 + 44, 1020 + i * 44), t, font=fs(21, bold=False), fill=TEXT, anchor="lm")
    od.rounded_rectangle((rx0 + 40, 1160, rx1 - 40, 1264), radius=14, fill=MINT_BG)
    od.text((rx0 + 64, 1188), "★  Stuck on setup?", font=fs(21), fill=PRIMARY, anchor="lm")
    od.text((rx0 + 64, 1228), "Message me — I help within 24 hours.", font=fs(19, bold=False), fill=TEXT, anchor="lm")
    img.alpha_composite(ov)
    pill(img, SIZE // 2, 1420, "COST THE POUR · POUR MORE PROFIT", font=fs(27), pad_x=44, pad_y=22,
         grad=(PRIMARY_LT, PRIMARY_DK))
    d = ImageDraw.Draw(img)
    tc(d, (SIZE // 2, 1530), "Cost every pour, kill your shrinkage, and keep more of every dollar over the bar.",
       fs(23, bold=False), TEXT_MUTED)
    book_crest(img, SIZE // 2, 1700, r=54)
    wordmark(img, SIZE // 2, 1850, "BAR & PUB COMMAND CENTER", 52, max_w=1200)
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
