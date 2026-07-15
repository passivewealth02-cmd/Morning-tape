# Installing this skill on another Claude account / machine

This folder is a self-contained **Claude Code skill**. Copy the whole
`spreadsheet-empire-product/` folder into a skills directory, then Claude will
load it automatically when you ask it to build a product.

## Option A — for everything you do (recommended)

Put it in your personal skills folder so it works in every project:

```bash
mkdir -p ~/.claude/skills
cp -r spreadsheet-empire-product ~/.claude/skills/
```

## Option B — for one repo only

Put it in the project you build products in:

```bash
mkdir -p <your-repo>/.claude/skills
cp -r spreadsheet-empire-product <your-repo>/.claude/skills/
```

## Using it

Just talk normally — "make a Command Center for dog groomers", "new product for
wedding planners", "do a budget spreadsheet for nurses". Claude will pick up the
skill (its `description` triggers on product/planner/tracker/dashboard requests),
read `SKILL.md`, follow the 8-step workflow, copy the helper blocks from
`reference/`, and produce the same brand every time.

You can also invoke it explicitly: `/spreadsheet-empire-product`.

## What's inside

- `SKILL.md` — the orchestrator (rules + 8-step workflow).
- `reference/WORKFLOW.md` — detailed process + KPI-verification discipline.
- `reference/BRAND.md` — palette, fonts, glyph safety, crest recipe.
- `reference/helpers_xlsx.py` / `helpers_marketing.py` / `helpers_pdf.py` — the
  shared engine; copy these verbatim into each new product's build files.
- `reference/ETSY_TEMPLATE.md` — the listing structure.
- `reference/example_homeschool/` — a complete, runnable worked example
  (dual-format: workbook + printable PDF).

## Requirements on the machine

- Python 3.10+ with `openpyxl` and `pillow` (`pip install openpyxl pillow`).
- The DejaVu fonts (`fonts-dejavu` — preinstalled on most Linux; the paths in the
  helpers are `/usr/share/fonts/truetype/dejavu/`).
- Only if you want the "shared to a git branch" workflow: a repo + push access.
  The product files themselves are just local files you download.
