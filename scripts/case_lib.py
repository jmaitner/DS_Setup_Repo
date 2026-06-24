#!/usr/bin/env python3
"""
case_lib.py  —  Shared helpers for the cases / compliance layer.

Cases live in cases/<id>.json, one per issue, attributed to items (DS#) and/or a brand.
Separate from products/ (what an item is) and status/ (where it's listed) — a case is an
ISSUE thread. Cross-links: case.linked_ds <-> item, case.case_number <-> status channel.
"""

import glob
import json
import os
import re

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CASES_DIR = os.path.join(REPO_ROOT, "cases")
PRODUCTS_DIR = os.path.join(REPO_ROOT, "products")

STATUSES = ["needs_review", "open", "pending", "resolved", "closed"]


def slug_id(source, case_number=None, title=None):
    """Stable case id: '<source>-<case#>' when a case number exists, else '<source>-<title-slug>'."""
    src = re.sub(r"[^a-z0-9]+", "-", (source or "other").lower()).strip("-")
    if case_number:
        return f"{src}-{str(case_number).strip()}"
    t = re.sub(r"[^a-z0-9]+", "-", (title or "case").lower()).strip("-")[:50]
    return f"{src}-{t}"


def case_path(cid):
    return os.path.join(CASES_DIR, f"{cid}.json")


def load_case(cid):
    p = case_path(cid)
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save_case(doc):
    os.makedirs(CASES_DIR, exist_ok=True)
    with open(case_path(doc["id"]), "w", encoding="utf-8") as f:
        json.dump(doc, f, indent=2, ensure_ascii=False, default=str)
        f.write("\n")


def all_cases():
    out = []
    for p in sorted(glob.glob(os.path.join(CASES_DIR, "*.json"))):
        with open(p, encoding="utf-8") as f:
            out.append(json.load(f))
    return out


def new_case(cid, source, title, today=None):
    return {
        "id": cid, "source": source, "case_number": None, "title": title,
        "description": None, "status": "needs_review", "tags": [], "brand": None,
        "linked_ds": [], "email_link": None, "email_id": None,
        "opened": today, "updated": today, "owner": None, "notes": None,
        "_meta": {},
    }


def catalog_ds_numbers():
    out = set()
    for p in glob.glob(os.path.join(PRODUCTS_DIR, "**", "*.json"), recursive=True):
        with open(p, encoding="utf-8") as f:
            out.add(str(json.load(f).get("ds_number")).strip())
    return out
