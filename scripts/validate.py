#!/usr/bin/env python3
"""
validate.py  —  The safety gate.

Loads every products/**/*.json and checks it for hard ERRORS (which block a
merge) and soft WARNINGS (which don't, but get reported).  This is what the
GitHub Action runs on every Pull Request, and what you should run locally
before pushing.

ERRORS (exit code 1):
  - file is not valid JSON
  - missing ds_number / brand / product_name / UPC
  - duplicate ds_number across the catalog
  - file is in the wrong place (folder != brand slug, or name != DS<number>.json)
  - a price / numeric field that is not a number

WARNINGS (exit code 0, but printed):
  - duplicate UPC
  - UPC not 12-14 digits
  - missing wholesale_cost / msrp / case_qty
  - no image URLs, or an image URL that isn't http(s)

Usage:
  python scripts/validate.py
  python scripts/validate.py --strict     # treat warnings as errors too
"""

import argparse
import glob
import json
import os
import sys

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ds_schema import brand_slug

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRODUCTS_DIR = os.path.join(REPO_ROOT, "products")
SCHEMA_PATH = os.path.join(REPO_ROOT, "schema", "product.schema.json")

NUMERIC_PRICE_FIELDS = ("drop_ship_cost", "wholesale_cost",
                        "ds_cost_domestic", "import_cost", "msrp")


def load_schema_validator():
    """Return a jsonschema validator for product.schema.json, or None if unavailable."""
    try:
        from jsonschema import Draft7Validator
    except ImportError:
        print("NOTE: jsonschema not installed - skipping schema checks "
              "(run: pip install -r requirements.txt). Rule checks still run.\n")
        return None
    try:
        with open(SCHEMA_PATH, encoding="utf-8") as f:
            schema = json.load(f)
        return Draft7Validator(schema)
    except Exception as e:
        print(f"NOTE: could not load schema ({e}) - skipping schema checks.\n")
        return None


def is_number(v):
    if v is None:
        return True  # blank is allowed; "missing" is a separate check
    if isinstance(v, (int, float)):
        return True
    try:
        float(str(v).replace("$", "").replace(",", "").strip())
        return True
    except (TypeError, ValueError):
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--strict", action="store_true",
                    help="treat warnings as errors (exit 1 if any)")
    args = ap.parse_args()

    files = sorted(glob.glob(os.path.join(PRODUCTS_DIR, "**", "*.json"), recursive=True))
    errors, warnings = [], []
    ds_seen, upc_seen = {}, {}
    validator = load_schema_validator()

    if not files:
        print("No product files found under products/. Nothing to validate.")
        return

    for path in files:
        rel = os.path.relpath(path, REPO_ROOT)
        try:
            with open(path, encoding="utf-8") as f:
                p = json.load(f)
        except Exception as e:
            errors.append(f"{rel}: invalid JSON - {e}")
            continue

        # ── schema conformance (structure / types / unknown fields) ──
        if validator is not None:
            for err in validator.iter_errors(p):
                loc = ".".join(str(x) for x in err.absolute_path) or "(root)"
                errors.append(f"{rel}: schema [{loc}] {err.message}")

        ds = str(p.get("ds_number") or "").strip()
        vendor = p.get("vendor")
        name = p.get("product_name")
        upc = str((p.get("identity") or {}).get("upc") or "").strip()

        # ── required fields ──
        if not ds:
            errors.append(f"{rel}: missing ds_number")
        if not vendor:
            errors.append(f"{rel}: missing vendor")
        if not p.get("brand"):
            warnings.append(f"{rel}: missing brand")
        if not name:
            errors.append(f"{rel}: missing product_name")
        if not upc:
            errors.append(f"{rel}: missing identity.upc")

        # ── location / naming (folder keyed on vendor) ──
        if ds and vendor:
            expected_dir = os.path.join(PRODUCTS_DIR, brand_slug(vendor))
            expected_name = f"DS{ds}.json"
            if os.path.dirname(path) != expected_dir:
                errors.append(f"{rel}: wrong folder - vendor '{vendor}' should live in "
                              f"products/{brand_slug(vendor)}/")
            if os.path.basename(path) != expected_name:
                errors.append(f"{rel}: filename should be {expected_name}")

        # ── duplicate ds_number ──
        if ds:
            if ds in ds_seen:
                errors.append(f"{rel}: duplicate ds_number {ds} (also in {ds_seen[ds]})")
            else:
                ds_seen[ds] = rel

        # ── numeric prices ──
        pricing = p.get("pricing") or {}
        for fld in NUMERIC_PRICE_FIELDS:
            if not is_number(pricing.get(fld)):
                errors.append(f"{rel}: pricing.{fld} is not a number ({pricing.get(fld)!r})")

        # ── WARNINGS ──
        if upc:
            if upc in upc_seen:
                warnings.append(f"{rel}: duplicate UPC {upc} (also in {upc_seen[upc]})")
            else:
                upc_seen[upc] = rel
            digits = upc.replace("-", "").replace(" ", "")
            if not (digits.isdigit() and 12 <= len(digits) <= 14):
                warnings.append(f"{rel}: UPC {upc!r} is not 12-14 digits")

        if pricing.get("wholesale_cost") in (None, ""):
            warnings.append(f"{rel}: missing pricing.wholesale_cost")
        if pricing.get("msrp") in (None, ""):
            warnings.append(f"{rel}: missing pricing.msrp")
        if (p.get("dimensions") or {}).get("master_case", {}).get("case_qty") in (None, ""):
            warnings.append(f"{rel}: missing dimensions.master_case.case_qty")

        urls = (p.get("images") or {}).get("urls") or []
        if not urls:
            warnings.append(f"{rel}: no image URLs")
        for u in urls:
            if not str(u).lower().startswith(("http://", "https://")):
                warnings.append(f"{rel}: image URL not http(s): {u!r}")

    # ── status layer (status/DS#####.json) ──
    status_files = sorted(glob.glob(os.path.join(REPO_ROOT, "status", "*.json")))
    status_validator = None
    status_schema_path = os.path.join(REPO_ROOT, "schema", "status.schema.json")
    if status_files and validator is not None and os.path.exists(status_schema_path):
        try:
            from jsonschema import Draft7Validator
            with open(status_schema_path, encoding="utf-8") as f:
                status_validator = Draft7Validator(json.load(f))
        except Exception as e:
            print(f"NOTE: could not load status schema ({e}).\n")

    for path in status_files:
        rel = os.path.relpath(path, REPO_ROOT)
        try:
            with open(path, encoding="utf-8") as f:
                s = json.load(f)
        except Exception as e:
            errors.append(f"{rel}: invalid JSON - {e}")
            continue
        if status_validator is not None:
            for err in status_validator.iter_errors(s):
                loc = ".".join(str(x) for x in err.absolute_path) or "(root)"
                errors.append(f"{rel}: status schema [{loc}] {err.message}")
        sds = str(s.get("ds_number") or "").strip()
        exp_name = f"DS{sds}.json"
        if sds and os.path.basename(path) != exp_name:
            errors.append(f"{rel}: filename should be {exp_name}")
        if sds and ds_seen and sds not in ds_seen:
            errors.append(f"{rel}: status for DS{sds} but no such product in catalog")

    # ── report ──
    print(f"Validated {len(files)} product file(s)"
          + (f" and {len(status_files)} status file(s)" if status_files else "") + ".\n")
    if warnings:
        print(f"WARNINGS ({len(warnings)}):")
        for w in warnings:
            print(f"  ! {w}")
        print()
    if errors:
        print(f"ERRORS ({len(errors)}):")
        for e in errors:
            print(f"  X {e}")
        print(f"\nVALIDATION FAILED - {len(errors)} error(s). Fix before merging.")
        sys.exit(1)

    if args.strict and warnings:
        print(f"STRICT MODE: {len(warnings)} warning(s) treated as failure.")
        sys.exit(1)

    print("VALIDATION PASSED." + (f"  ({len(warnings)} warning(s))" if warnings else ""))


if __name__ == "__main__":
    main()
