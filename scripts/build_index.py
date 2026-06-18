#!/usr/bin/env python3
"""
build_index.py  —  Regenerate index/catalog.csv, the flat "spreadsheet view".

Walks products/**/*.json and writes one row per product with the fields you
scan most often.  Run this after importing or editing products so the index
stays in sync (the PR check will remind you if it's stale).

Usage:
  python scripts/build_index.py
"""

import csv
import glob
import json
import os
import sys

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRODUCTS_DIR = os.path.join(REPO_ROOT, "products")
INDEX_PATH = os.path.join(REPO_ROOT, "index", "catalog.csv")

COLUMNS = [
    "Vendor", "Brand", "DS#", "UPC", "Product Name", "Vendor Item#",
    "Wholesale Cost", "Drop Ship Cost", "MSRP", "Case Qty",
    "# Images", "Has Compliance Docs", "Supplier ID", "Source File", "File Path",
]


def has_docs(comp):
    doc_fields = ("compliance_cert", "doc", "sds", "sds_url", "cpsia",
                  "test_reports", "cpc", "letter_of_compliance")
    return "Yes" if any(comp.get(f) for f in doc_fields) else "No"


def main():
    files = sorted(glob.glob(os.path.join(PRODUCTS_DIR, "**", "*.json"), recursive=True))
    rows = []
    for path in files:
        with open(path, encoding="utf-8") as f:
            p = json.load(f)
        pricing = p.get("pricing") or {}
        ident = p.get("identity") or {}
        mc = (p.get("dimensions") or {}).get("master_case", {}) or {}
        comp = p.get("compliance") or {}
        meta = p.get("_meta") or {}
        rows.append([
            p.get("vendor", ""),
            p.get("brand", ""),
            p.get("ds_number", ""),
            ident.get("upc", ""),
            p.get("product_name", ""),
            p.get("vendor_item_number", ""),
            pricing.get("wholesale_cost", ""),
            pricing.get("drop_ship_cost", ""),
            pricing.get("msrp", ""),
            mc.get("case_qty", ""),
            len((p.get("images") or {}).get("urls") or []),
            has_docs(comp),
            meta.get("supplier_id", ""),
            meta.get("source_file", ""),
            os.path.relpath(path, REPO_ROOT).replace("\\", "/"),
        ])

    rows.sort(key=lambda r: (str(r[0]).lower(), str(r[1])))
    os.makedirs(os.path.dirname(INDEX_PATH), exist_ok=True)
    with open(INDEX_PATH, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(COLUMNS)
        w.writerows([["" if v is None else v for v in r] for r in rows])

    print(f"Wrote {len(rows)} row(s) -> {os.path.relpath(INDEX_PATH, REPO_ROOT)}")


if __name__ == "__main__":
    main()
