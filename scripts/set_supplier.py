#!/usr/bin/env python3
"""
set_supplier.py  —  Stamp the Apprise supplier ID onto a vendor's products.

Several setup sheets ship with a blank B3 (supplier number), so those products have
_meta.supplier_id = null and `export.py` needs --supplier each time. Run this once per
vendor to record the real ID; exports then pick it up automatically.

Usage:
  python scripts/set_supplier.py <vendor-slug> <supplier_id>
  python scripts/set_supplier.py canal-toys 1530
  python scripts/set_supplier.py --list            # show current supplier IDs per vendor
"""

import glob
import json
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRODUCTS_DIR = os.path.join(REPO_ROOT, "products")


def list_suppliers():
    by_vendor = {}
    for path in glob.glob(os.path.join(PRODUCTS_DIR, "*", "*.json")):
        slug = os.path.basename(os.path.dirname(path))
        with open(path, encoding="utf-8") as f:
            p = json.load(f)
        sid = (p.get("_meta") or {}).get("supplier_id")
        by_vendor.setdefault(slug, set()).add(sid)
    for slug in sorted(by_vendor):
        ids = sorted(str(s) for s in by_vendor[slug])
        print(f"  {slug:24s} supplier_id(s): {', '.join(ids)}")


def main():
    args = sys.argv[1:]
    if not args or args[0] == "--list":
        list_suppliers()
        return
    if len(args) != 2:
        sys.exit("Usage: python scripts/set_supplier.py <vendor-slug> <supplier_id>  (or --list)")

    slug, supplier = args[0], args[1]
    try:
        supplier = int(supplier)
    except ValueError:
        pass  # keep as string if not purely numeric

    vendor_dir = os.path.join(PRODUCTS_DIR, slug)
    if not os.path.isdir(vendor_dir):
        sys.exit(f"ERROR: no products/{slug}/ folder. Run --list to see vendor slugs.")

    n = 0
    for path in glob.glob(os.path.join(vendor_dir, "*.json")):
        with open(path, encoding="utf-8") as f:
            p = json.load(f)
        p.setdefault("_meta", {})["supplier_id"] = supplier
        with open(path, "w", encoding="utf-8") as f:
            json.dump(p, f, indent=2, ensure_ascii=False, default=str)
            f.write("\n")
        n += 1

    print(f"Set supplier_id={supplier} on {n} product(s) in products/{slug}/")
    print("Next: python scripts/validate.py && python scripts/build_index.py")


if __name__ == "__main__":
    main()
