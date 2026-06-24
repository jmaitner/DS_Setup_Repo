# DS Setup Repo — Rules for Claude

This repository is the **single source of truth for DS product setup data**. Every
product we receive a setup sheet for lives here as one JSON file, versioned in git.
You (Claude) help Jackson and Bruce query, add, update, and export this data **safely**.

Read this whole file before touching anything.

---

## The golden rules (safety first)

1. **Never edit data directly on `main`.** Every change — adding products, fixing a
   field, importing a sheet — happens on a **branch**, then a **Pull Request**. `main`
   is protected.
2. **Never blindly overwrite.** If data already exists and the new info disagrees,
   **stop and show the diff** to the user. We often receive *wrong* vendor info; the
   human decides which version is correct. Do not assume the newest value wins.
3. **Always validate before pushing.** Run `python scripts/validate.py`. If it reports
   ERRORS, fix them before opening a PR. The GitHub Action will block the merge anyway.
4. **Always rebuild the index after data changes.** Run `python scripts/build_index.py`
   so `index/catalog.csv` stays in sync, and commit it with your change.
5. **One product = one file.** `products/<vendor-slug>/DS<number>.json`. Folders are keyed
   on **vendor** (who we buy from — Headstart, First 4 Figures, etc.), NOT the marketing
   `brand` field. The DS number is the primary key. Filename and folder must match the
   data (validate.py enforces this).
6. **Explain what you're about to do before doing it**, especially for Bruce. Prefer the
   slash commands in `.claude/commands/` — they encode the safe workflow step by step.

---

## What the data looks like

Each product JSON follows `schema/product.schema.json` (81 fields from the DS Only tab),
grouped as: top-level identity (`ds_number`, `vendor`, `brand`, `vendor_item_number`,
`product_name`), then `identity`, `pricing`, `sourcing`, `content`, `attributes`,
`images` (Cloudinary URLs), `dimensions` (`item` / `package` / `master_case`),
`compliance`, and `_meta`. Note `vendor` (who we buy from) drives the folder; `brand` is
the per-product marketing brand and can differ.

Images are **Cloudinary URLs**, never binary files in the repo.

**Three layers, joined by DS#.** (1) Product *data* in `products/` (above). (2) Channel
*status* — lifecycle + per-channel state (live/error/not_listed/…), case#, listing ids,
issues — in `status/DS#####.json` (see `docs/STATUS_LAYER.md`). (3) *Cases* — support cases &
vendor/compliance issues, each attributed to items/brands with tags — in `cases/<id>.json`
(see `docs/CASES.md`). They change for different reasons; never mix them. Channels: Amazon,
Walmart 1P, Walmart 3P, TikTok, Best Buy, NocNoc, Shopify (Face2FaceFun), Toys R Us, eBay,
Target Plus.

---

## The scripts (all under `scripts/`)

| Script | Purpose |
|---|---|
| `import_setup_sheet.py "<file.xlsx>"` | Sheet → product JSONs. Reads **Move Forward Items** tab if present (the curated carry-forward list), else **DS Only**; `--sheet` to force. Non-destructive (skips existing unless `--update`). Vendor auto-detected from path. |
| `merge_urls.py "<urls.xlsx>"` | Fold image-URL or compliance-doc files into products by DS#. Non-destructive (reports conflicts). |
| `validate.py` | The safety gate. Errors block merge; warnings are reported. `--strict` to fail on warnings. |
| `export.py` | Product JSON → WMS + channel upload files. Reuses `ds_automation.py`. |
| `gen_amazon.py --vendor <name>` | Amazon flat-file from catalog. Generic field names by default; `--template <Amazon Toys template.xlsx>` fills Amazon's own category template. Skips discontinued; adds category recommendation. |
| `build_index.py` | Regenerate `index/catalog.csv` and `index/status.csv`. |
| `set_supplier.py <vendor-slug> <id>` | Stamp Apprise supplier ID onto a vendor's products (sheets with blank B3). `--list` shows current IDs. |
| `init_status.py` | Create a `status/DS#####.json` for every product missing one (all channels `not_listed`). |
| `update_status.py <DS#> ...` | Edit lifecycle / a channel's state, case#, listing id, issue. |
| `ingest_status.py "<report.xlsx>" --channel <ch>` | Bulk-update one channel's status from a channel export/error report (match by DS#/SKU/id). |
| `pull_goflow.py` | Sync status from GoFlow (the hub) for all connected channels in one sweep (match by `product.item_number`=DS#). Needs `GOFLOW_KEY` env. Runs 2×/week via `.github/workflows/sync-status.yml`. |
| `build_dashboard.py` | Regenerate `site/index.html` — a self-contained visual dashboard (catalog + status + **cases** + detail). Open in a browser or host `site/`. Rerun after data/status/case changes. |
| `case.py` | Create/update/attribute cases in the **cases layer** (`cases/<id>.json`). Upsert by id; `--ds`/`--brand`/`--tags`/`--status`; `--stamp` writes the case# onto linked items' channel. `--list` shows all. See `docs/CASES.md`. |
| `ds_automation.py` | Vendored transformer (the 5 WMS + 7 channel generators). Single source of the column map. |
| `ds_schema.py` | Conversion between sheet rows, product JSON, and the transformer's dict. |

---

## Standard workflows

For the full vendor intake cadence (sheet → verify → images to Cloudinary → merge → PR),
use `/onboard-vendor` — see `docs/INTAKE_WORKFLOW.md`.

### Add / import a setup sheet
```
git checkout -b import/<brand>-<date>
python scripts/import_setup_sheet.py "C:\path\to\Brand 2026 Setup.xlsx"
python scripts/validate.py
python scripts/build_index.py
git add -A && git commit -m "Import <Brand> setup sheet"
git push -u origin HEAD
gh pr create --fill
```

### Update a field (e.g. wrong price)
Edit the product JSON. Then **show the user the git diff** and confirm before committing.
```
git checkout -b fix/<ds-number>-<what>
# edit products/<brand>/DS<number>.json
python scripts/validate.py && python scripts/build_index.py
git add -A && git commit -m "Fix <field> for DS<number>: <old> -> <new> (reason)"
git push -u origin HEAD && gh pr create --fill
```

### Export upload files
```
python scripts/export.py --brand <slug> --supplier <id> --category Toys --out ./out
# add --channels --shopify-type "..." for the 7 channel files
```
`out/` is git-ignored — exports are disposable artifacts, not catalog data.

### Query / pull images / build reports
Read `index/catalog.csv` for a fast overview, or read individual product JSONs for
detail. For images, pull the `images.urls` for the requested DS#s / brand.

---

## When something is wrong

- **Conflicting data:** show both values + their source (`_meta.source_file`, git blame)
  and ask the human which is correct. Never pick silently.
- **Validation errors you can't resolve:** report them; don't push.
- **A product that should be deleted:** confirm with the user first, then remove the file
  on a branch + PR (git keeps the history).

Keep `docs/GUIDE_FOR_BRUCE.md` and the slash commands updated as the workflow evolves.
