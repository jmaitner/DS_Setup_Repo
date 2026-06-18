# Vendor Intake Workflow (the ongoing cadence)

This is how a setup sheet (and its images/docs) becomes catalog data. Run it with the
`/onboard-vendor` slash command, which does these steps for you.

```
  Setup sheet (.xlsx)
        │
        ▼
  [1] import_setup_sheet.py ──► products/<vendor>/DS#####.json   (one file per product)
        │                         • vendor auto-detected from the file path
        │                         • prices cleaned, Excel errors dropped
        ▼
  [2] validate.py  ──►  ERRORS block;  WARNINGS = data-quality issues for human review
        │              (missing/odd UPCs, dup UPCs, missing cost/MSRP, no images)
        ▼
  [3] CONFIRM with Jackson  ◄── do not invent barcodes/prices; he decides on conflicts
        │
        ▼
  [4] Images/docs uploaded to Cloudinary (~/.claude/Cloudinary/)  ──► URL spreadsheet
        │
        ▼
  [5] merge_urls.py  ──►  images.urls  /  compliance.documents   (by DS#, non-destructive)
        │
        ▼
  [6] build_index.py  ──►  index/catalog.csv refreshed
        │
        ▼
  [7] branch ──► commit ──► Pull Request ──► review ──► merge to main
```

## Why two phases (sheet first, images later)
Jackson's described cadence: hand over the **setup sheet** first (we import + verify the
data), then later grant access to the **image folder** (we upload to Cloudinary and merge
the URLs). The catalog supports a product existing with data but no images yet — the
`# Images = 0` shows up in `index/catalog.csv` and as a validation warning, so nothing
gets lost; we just fill images when they arrive.

## Data-quality issues are surfaced, never silently "fixed"
The validator flags problems but the **human decides**. Examples seen in the initial seed:
- Some First 4 Figures rows have edition text (e.g. "Zelda (Standard edition)") in the
  barcode column instead of a UPC → flagged, left as-is for Jackson to correct at source.
- Duplicate UPCs across variants → flagged for review.
- `'x'` placeholders in cost columns → normalized to empty (no numeric meaning).

## Updating existing data later
Use `/update-product`. Conflicting values are shown old-vs-new with their source; the
human chooses. Re-importing a sheet is non-destructive by default (existing files are
skipped); `--update` overwrites and the PR diff is the review.
