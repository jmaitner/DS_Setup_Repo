---
description: Full intake cadence for a new/updated setup sheet — import, verify, images to Cloudinary, merge, PR
argument-hint: <path to setup sheet .xlsx>  (image folder optional, ask later)
---

Run the standard vendor intake cadence for: $ARGUMENTS

This is the repeatable flow for when Jackson hands over a setup sheet (and later the
images). Do each step, confirming with the user at the checkpoints.

### Phase 1 — Import & verify (do now)
1. `git checkout -b intake/<vendor>-<today>`.
2. `python scripts/import_setup_sheet.py "<sheet path>"` (vendor auto-detected from the
   `Priority Vendors/<Vendor>` path; pass `--vendor` if the path doesn't contain that).
3. `python scripts/validate.py` and `python scripts/build_index.py`.
4. **Report to the user and get confirmation:** products created/updated/skipped, and
   every WARNING grouped by type — especially: non-numeric/missing UPCs, duplicate UPCs,
   missing wholesale_cost/MSRP, products with no images. These are data-quality issues
   for a human to confirm, not for you to silently fix. Do NOT invent barcodes or prices.

### Phase 2 — Images & docs (when the user grants the image/doc folder)
5. Ask the user for the path to the image folder (and any testing-doc folder).
6. Upload to Cloudinary using the existing Cloudinary tooling in `~/.claude/Cloudinary/`
   (creds in `~/.claude/Cloudinary/.env`). Match images to products by DS# / vendor item #
   / name per the Cloudinary image-upload skill. Produce a URL spreadsheet in that skill's
   standard format (DS# + Image URL columns; or Doc Type/Doc URL columns for testing docs).
7. Merge the resulting URLs into the catalog:
   `python scripts/merge_urls.py "<that url file>.xlsx"`  (auto-detects images vs docs;
   non-destructive — reports CONFLICTS instead of overwriting; add `--overwrite` only with
   the user's OK).
8. `python scripts/validate.py` and `python scripts/build_index.py` again.

### Phase 3 — Land it
9. Show the user the final diff / summary.
10. `git add -A && git commit -m "Intake <Vendor>: <N> products (+images/docs)"`.
11. `git push -u origin HEAD && gh pr create --fill`; share the PR link for review.

Never commit to `main` directly. Never overwrite conflicting data without explicit OK.
