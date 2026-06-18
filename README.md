# DS Setup Repo

A **git-versioned, queryable catalog** of every product we set up — built from our
DS setup sheets. One product = one structured JSON file. Git gives us full history
(who changed what, when, and why), a review gate so bad data can't silently overwrite
good data, and a clean way for Claude Code to pull, transform, and export the data.

## What's in here

```
products/<brand-slug>/DS<number>.json   The catalog. One file per product.
index/catalog.csv                        Auto-generated flat "spreadsheet view".
schema/product.schema.json               The 81-field product schema.
scripts/                                 import / validate / export / build_index.
docs/GUIDE_FOR_BRUCE.md                  Plain-English how-to.
.claude/commands/                        Slash commands that run the safe workflow.
CLAUDE.md                                Rules Claude follows when working here.
```

## Quick start

1. **Clone & open in Claude Code**
   ```
   git clone https://github.com/jmaitner/DS_Setup_Repo
   cd DS_Setup_Repo
   pip install openpyxl
   claude
   ```
2. **Ask Claude** things like:
   - "Import this setup sheet" (drag the .xlsx path in)
   - "What's the wholesale cost and case qty for all First4 Figures products?"
   - "Pull every image URL for the Bladez brand"
   - "Export WMS + Shopify files for DS12345 and DS12346"
   - "DS12345's MSRP is wrong, it should be 29.99"

Claude will use the slash commands and scripts below, always on a branch + PR.

## The rules (short version)

- **Never edit `main` directly** — branch → PR → review → merge.
- **Never blindly overwrite** conflicting data — a human decides which value is right.
- **Validate before pushing** (`python scripts/validate.py`); CI blocks bad merges.
- **Rebuild the index** after changes (`python scripts/build_index.py`).

See [`CLAUDE.md`](CLAUDE.md) for the full rules and [`docs/GUIDE_FOR_BRUCE.md`](docs/GUIDE_FOR_BRUCE.md)
for a step-by-step guide.

## Scripts

| Command | What it does |
|---|---|
| `python scripts/import_setup_sheet.py "<file.xlsx>"` | Setup sheet → product JSONs |
| `python scripts/validate.py` | Check the catalog (the safety gate) |
| `python scripts/build_index.py` | Regenerate `index/catalog.csv` |
| `python scripts/export.py --brand <slug> --supplier <id> --out ./out` | JSON → WMS/channel files |
