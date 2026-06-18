#!/usr/bin/env python3
"""
export.py  —  Turn catalog product JSON back into upload files.

Reuses the generator functions in the vendored ds_automation.py (the exact
same code that runs your manual setup-sheet workflow), but driven
non-interactively from the catalog and command-line flags.

WMS files (always generated):
  new_product, supplier, uom, purchase_cost, goflow

Channel files (with --channels):
  walmart, tiktok, bestbuy, toysrus, target, ebay, shopify

Examples:
  # All BLADEZ products -> WMS files in ./out
  python scripts/export.py --brand bladez --supplier 1514 --category Toys --out ./out

  # Specific DS numbers, include channel files
  python scripts/export.py --ds 12345 12346 --supplier 1514 --category Toys \\
      --channels --shopify-type "Action Figure" --out ./out

  # Everything in the catalog
  python scripts/export.py --all --supplier 1514 --category Toys --out ./out
"""

import argparse
import glob
import json
import os
import sys

# ds_automation prints unicode (✓ → ) — keep Windows' cp1252 console from crashing.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import ds_automation as A
from ds_schema import brand_slug, product_to_automation_dict

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRODUCTS_DIR = os.path.join(REPO_ROOT, "products")


def load_products(brand=None, ds_list=None, take_all=False):
    files = sorted(glob.glob(os.path.join(PRODUCTS_DIR, "**", "*.json"), recursive=True))
    out = []
    ds_set = set(str(d).strip() for d in ds_list) if ds_list else None
    slug = brand_slug(brand) if brand else None
    for path in files:
        with open(path, encoding="utf-8") as f:
            p = json.load(f)
        if slug and brand_slug(p.get("brand")) != slug:
            continue
        if ds_set and str(p.get("ds_number")).strip() not in ds_set:
            continue
        out.append(p)
    if not (brand or ds_list or take_all):
        return []
    return out


def resolve_supplier(products, cli_supplier):
    if cli_supplier:
        return cli_supplier
    ids = {(p.get("_meta") or {}).get("supplier_id")
           for p in products if (p.get("_meta") or {}).get("supplier_id")}
    if len(ids) == 1:
        return ids.pop()
    return None


def main():
    ap = argparse.ArgumentParser(description="Export catalog products to upload files.")
    sel = ap.add_mutually_exclusive_group(required=True)
    sel.add_argument("--brand", help="export all products for this brand")
    sel.add_argument("--ds", nargs="+", help="export these DS number(s)")
    sel.add_argument("--all", action="store_true", help="export the entire catalog")

    ap.add_argument("--supplier", help="Apprise supplier id (else read from product _meta)")
    ap.add_argument("--category", default="", help="WMS product category, e.g. Toys")
    ap.add_argument("--lead-time", type=int, default=A.DEFAULT_LEAD_TIME)
    ap.add_argument("--moq", type=int, default=A.DEFAULT_MOQ)
    ap.add_argument("--min-ds", type=int, default=0,
                    help="New Product file: only include DS# >= this (skip existing)")
    ap.add_argument("--out", default=os.path.join(REPO_ROOT, "out"), help="output folder")
    ap.add_argument("--base", default="catalog_export", help="output filename prefix")

    ap.add_argument("--channels", action="store_true", help="also generate channel files")
    ap.add_argument("--tiktok-cat", default="")
    ap.add_argument("--bestbuy-cat", default="")
    ap.add_argument("--target-cat", default="")
    ap.add_argument("--shopify-type", default="")
    args = ap.parse_args()

    raw = load_products(brand=args.brand, ds_list=args.ds, take_all=args.all)
    if not raw:
        sys.exit("No matching products found.")

    supplier = resolve_supplier(raw, args.supplier)
    if supplier is None:
        sys.exit("ERROR: supplier id not set and not found in product _meta. Pass --supplier.")

    products = [product_to_automation_dict(p) for p in raw]
    os.makedirs(args.out, exist_ok=True)
    base = args.base
    j = lambda name: os.path.join(args.out, name)

    print(f"Exporting {len(products)} product(s)  supplier={supplier}  -> {args.out}\n")

    A.generate_new_product(products, supplier, args.category,
                           j(f"{base}_new_product.xlsx"), min_ds_number=args.min_ds)
    A.generate_supplier(products, supplier, args.lead_time, args.moq,
                        j(f"{base}_supplier.xlsx"))
    A.generate_uom(products, j(f"{base}_uom.xlsx"))
    A.generate_purchase_cost(products, supplier, j(f"{base}_purchase_cost.xlsx"))
    A.generate_goflow_csv(products, j(f"{base}_goflow.csv"))
    A.validate_new_product(products, j(f"{base}_new_product.xlsx"), min_ds_number=args.min_ds)

    if args.channels:
        print("\nGenerating channel files...")
        A.generate_walmart_xlsx(products, supplier, j(f"{base}_walmart.xlsx"))
        A.generate_tiktok_xlsx(products, args.tiktok_cat, j(f"{base}_tiktok.xlsx"))
        A.generate_bestbuy_xlsx(products, args.bestbuy_cat, j(f"{base}_bestbuy.xlsx"))
        A.generate_toysrus_xlsx(products, j(f"{base}_toysrus.xlsx"))
        A.generate_target_csv(products, args.target_cat, j(f"{base}_target.csv"))
        A.generate_ebay_csv(products, j(f"{base}_ebay.csv"))
        A.generate_shopify_xlsx(products, args.shopify_type, j(f"{base}_shopify.xlsx"))

    print(f"\nDone. Files in {args.out}")


if __name__ == "__main__":
    main()
