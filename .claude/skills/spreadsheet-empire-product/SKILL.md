---
name: spreadsheet-empire-product
description: >
  Build a complete premium "Command Center" digital product for Etsy — a
  multi-tab Google Sheets/Excel workbook (and optionally a printable PDF pack),
  plus 10 marketing images and 4 docs, all in one forest-green + gold brand.
  Use this whenever the user asks to "make/create/do a [niche] one", "new
  product", a planner/tracker/dashboard/command center, or a spreadsheet product
  for any audience (gig drivers, families, hobbies, small business, etc.).
---

# Spreadsheet Empire — Product Builder

You build **premium, production-ready digital products** for high-ticket Etsy
sales under ONE consistent brand: a forest-green + gold "Command Center" identity.
Every product is a multi-tab **Google Sheets / Excel workbook** with a live
dashboard, delivered with **10 marketing images** and **4 docs**. Some products
are **dual-format** (workbook **+** a print-ready PDF pack).

This skill makes every product come out **identical in quality and brand** no
matter the niche. Follow it exactly.

## Hard rules (never break these)

1. **Never deploy / publish / host anything.** No Vercel, no web hosting. The
   user downloads files from chat only. (Committing/pushing to the designated git
   branch IS allowed and expected.)
2. **Verify every KPI number** shown in marketing against the actual workbook
   data with a Python replica before you draw it. Marketing numbers must match
   the sheet exactly. If they don't, fix the data or the copy.
3. **Image 3 must be a distinct sheet**, never a repeat of the hero dashboard.
4. **Do not put the model identifier** in commits, docs, code comments, or any
   pushed/shipped artifact — chat replies only.
5. **Custom (non-`build_log`) tables must start in column B** when the sheet uses
   a width-2 margin column A (`set_widths(ws,[2, ...])`). Writing content into
   column A crushes it. Pass `start_col=2` and write cells to columns 2+.
6. **Fonts are DejaVu.** They render `$ % ★ ✓ • ▲ ►` but NOT color emoji or `■`.
   Use vector-drawn icons for anything DejaVu can't render (see BRAND.md).
7. **One crest per product** — a unique vector emblem that nods to the niche
   (steering wheel for rideshare, carrot for Instacart, backpack for back-to-
   school, open book for homeschool…) in the shared green+gold frame.

## The 8-step workflow (do in order)

1. **Design** the product: pick a name (`<Niche> Command Center™`), a sample
   persona, ~18 tabs (Dashboard + ~16 domain tabs + Settings), and the **12
   dashboard KPIs**. Decide single-format vs dual-format (add a 12-page PDF when
   the niche loves printables — families, school, home).
2. **Lock the numbers first.** Write the sample data as explicit Python constants
   and run a replica to compute every KPI + a blended "health/readiness/on-track"
   score. Only proceed once the numbers are clean and quotable.
3. **`build_xlsx.py`** — copy `reference/helpers_xlsx.py` verbatim, then add
   domain dropdown lists, one builder function per tab, cross-linked named ranges
   (interdependent totals reference each other by name), and a Dashboard with 12
   `kpi_card`s + a health table (`HealthRange`) + charts. Build & load-check it.
4. **Verify** with openpyxl: 19 sheets, all named ranges resolve, custom tables
   start in column B. Re-run the KPI replica against the constants.
5. **`build_marketing.py`** — copy `reference/helpers_marketing.py`, add the
   crest, `TABS`, `FILE_LABEL`, `KPIS`, `content_dashboard()`, `content_*()` for
   the 3 sheet-spread images, and 6 `render_*()`: 01 hero, 02 inside (all tabs),
   03/04/05 distinct sheets, 06 mobile (or printables showcase if dual-format).
6. **`build_marketing_detail.py`** — import from `build_marketing`; render 07
   features (4 spotlights), 08 compare ("basic X vs Command Center"), 09
   how-it-works (4 steps), 10 value (included / who / guarantee). 10 images total.
7. **(Dual-format only) `build_pdf.py`** — copy `reference/helpers_pdf.py`; one
   function per printable page; save a multi-page US-Letter PDF + page PNGs
   (reuse those PNGs as the 06 printables-showcase image).
8. **Docs → commit → deliver.** Write `README.md`, `GOOGLE_SHEETS.md`,
   `BUILD_INSTRUCTIONS.md`, `ETSY_LISTING.md` (see `reference/ETSY_TEMPLATE.md`).
   Remove `__pycache__`, `git add`, commit to the designated branch, push.
   Deliver via chat: images in batches of 3 (`display:"render"`), then the
   `.xlsx`/PDF/listing (`display:"attach"`).

## Read these before building

- `reference/WORKFLOW.md` — the detailed process, KPI-verification discipline,
  the dashboard/health-score pattern, and delivery steps.
- `reference/BRAND.md` — palette, fonts, glyph safety, and the crest recipe.
- `reference/helpers_xlsx.py`, `reference/helpers_marketing.py`,
  `reference/helpers_pdf.py` — copy these verbatim; they are the shared engine.
- `reference/ETSY_TEMPLATE.md` — the listing structure (title/tags/description).

Design content is the only thing that changes per product. The helpers, brand,
workflow and quality bar never change.
