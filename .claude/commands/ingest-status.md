---
description: Bulk-update channel status from a channel export / error report spreadsheet
argument-hint: <path to the channel report .xlsx> (tell me which channel)
---

The user wants to ingest a channel report into the status layer: $ARGUMENTS

1. Ask which channel this report is for if not stated (amazon, walmart_1p, walmart_3p,
   tiktok, bestbuy, nocnoc, shopify_f2f, toysrus, ebay, target_plus).
2. INSPECT the spreadsheet's header row first (read the first sheet's columns) so you can
   map them correctly. Identify:
   - the key column: DS Number, or a SKU/seller-SKU (= vendor item number), or a listing id
   - which column holds the listing state / error message / case number / listing id (if any)
3. Create a branch: `git checkout -b status/ingest-<channel>-<today>`.
4. Run `scripts/ingest_status.py` with the right flags, e.g.:
   ```
   python scripts/ingest_status.py "report.xlsx" --channel walmart_1p \
       --match-type sku --match-col "Seller SKU" --set-state live --id-col "Item ID"
   ```
   or for an error report:
   ```
   python scripts/ingest_status.py "errors.xlsx" --channel amazon \
       --match-col "DS Number" --set-state error --issue-col "Error" --case-col "Case ID"
   ```
   Use `--dry-run` first to preview matches.
5. Report matched / updated / unmatched counts. Unmatched keys mean rows with no product in
   the catalog — surface them; don't guess.
6. `validate.py`, `build_index.py`, show the diff, commit, push, PR.

Never commit to `main` directly. Only the named channel's fields are touched.
