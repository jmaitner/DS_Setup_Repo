---
description: Generate WMS and/or sales-channel upload files from the catalog
argument-hint: <brand or DS numbers, and which channels>
---

The user wants upload files for: $ARGUMENTS

1. Figure out the selection: a `--brand <slug>`, a list of `--ds <numbers>`, or `--all`.
2. Determine the Apprise `--supplier` id. Prefer the one stored in the products'
   `_meta.supplier_id`; if missing or inconsistent, ask the user.
3. Run the export to the git-ignored `out/` folder:
   ```
   python scripts/export.py --brand <slug> --supplier <id> --category Toys --out ./out
   ```
   - The 5 WMS files (new_product, supplier, uom, purchase_cost, goflow) are always made.
   - Add `--channels` for the 7 channel files. When using channels, ask for the category
     values the user needs: `--tiktok-cat`, `--bestbuy-cat`, `--target-cat`,
     `--shopify-type`.
   - For the Apprise New Product file, ask whether to skip items already in the system
     (`--min-ds <first new DS#>`).
4. Read the export's validation output. If "VALIDATION FAILED" appears, surface it.
5. Tell the user exactly which files were written to `out/` and what each is for.

Exports are disposable artifacts — never commit them.
