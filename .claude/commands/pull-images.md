---
description: Pull product image URLs (and optionally download them) for a brand or DS list
argument-hint: <brand or DS numbers>
---

The user wants images for: $ARGUMENTS

1. Identify the matching products (by brand slug or DS numbers) from
   `products/<brand-slug>/DS<number>.json`.
2. Collect each product's `images.urls` (Cloudinary URLs; the first URL is primary).
3. Present them clearly — grouped by DS number — as a list, or build a CSV/Excel in `out/`
   with columns: DS#, Product Name, UPC, Image 1..N. Match whatever format the user asked
   for.
4. If the user wants the actual image FILES (not just URLs), download each URL into
   `out/images/<DS number>/` named by position (1.jpg, 2.jpg, ...). Note in your summary
   that these are local copies; the catalog still references the Cloudinary URLs.
5. Flag any product with zero image URLs so the user knows what's missing.

Read-only on the catalog — do not modify product JSON.
