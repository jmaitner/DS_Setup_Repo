---
description: Query the catalog and/or build a custom spreadsheet or report
argument-hint: <what you want to know / build>
---

The user wants: $ARGUMENTS

This is READ-ONLY unless they explicitly ask for a file. No branch or PR needed for pure
lookups.

1. For a quick overview, read `index/catalog.csv` (Brand, DS#, UPC, name, costs, MSRP,
   case qty, image count, compliance docs flag).
2. For field-level detail, read the individual `products/<brand-slug>/DS<number>.json`
   files that match.
3. Answer concisely. If the user asked for a deliverable (Excel/CSV report, pricing list,
   profitability inputs, catalog by brand), build it with the xlsx skill and save it to
   `out/` (which is git-ignored — reports are artifacts, not catalog data).
4. If a lookup reveals missing or clearly wrong data, mention it and offer to fix it via
   `/update-product` (which goes through the review gate).

Do not modify any product JSON in this command.
