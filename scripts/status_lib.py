#!/usr/bin/env python3
"""
status_lib.py  —  Shared definitions for the channel-status layer.

The status layer lives in `status/DS#####.json`, keyed by DS number, SEPARATE from
the product data in `products/`. DS# is the join key. Status changes constantly
(listings go live, errors appear, items get discontinued); product data rarely does.

Single source of truth for: the channel list, the allowed states/lifecycles, and the
shape of a status document.
"""

import json
import os

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATUS_DIR = os.path.join(REPO_ROOT, "status")
PRODUCTS_DIR = os.path.join(REPO_ROOT, "products")

# The 10 sales channels we set products up on. Key = stable id; value = display label.
CHANNELS = {
    "amazon":      "Amazon",
    "walmart_1p":  "Walmart 1P",
    "walmart_3p":  "Walmart 3P",
    "tiktok":      "TikTok Shop",
    "bestbuy":     "Best Buy",
    "nocnoc":      "NocNoc",
    "shopify_f2f": "Shopify (Face2FaceFun)",
    "toysrus":     "Toys R Us",
    "ebay":        "eBay",
    "target_plus": "Target Plus",
}

# Per-channel listing state.
CHANNEL_STATES = [
    "not_listed",         # not set up here (default)
    "planned",            # intend to list
    "setup_in_progress",  # being built
    "pending",            # submitted, awaiting channel approval
    "live",               # active and selling
    "error",              # rejected / suppressed-with-error / needs fixing
    "suppressed",         # live listing currently suppressed
    "discontinued",       # was live, taken down
]

# Overall product lifecycle (independent of any single channel).
LIFECYCLES = ["active", "planned", "on_hold", "discontinued", "dropped"]


def channel_default():
    return {
        "state": "not_listed",
        "id": None,            # ASIN / item id / listing id on that channel
        "url": None,           # direct link to the listing (from GoFlow store_page_url)
        "case_number": None,   # Amazon/Walmart support case #
        "issue": None,         # short description of current error/blocker
        "notes": None,
        "updated": None,       # ISO date this channel entry last changed
    }


def new_status_doc(ds_number, today=None):
    return {
        "ds_number": str(ds_number).strip(),
        "lifecycle": "active",
        "lifecycle_note": None,
        "channels": {ch: channel_default() for ch in CHANNELS},
        "updated": today,
    }


def status_path(ds_number):
    return os.path.join(STATUS_DIR, f"DS{str(ds_number).strip()}.json")


def load_status(ds_number):
    p = status_path(ds_number)
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save_status(doc):
    os.makedirs(STATUS_DIR, exist_ok=True)
    with open(status_path(doc["ds_number"]), "w", encoding="utf-8") as f:
        json.dump(doc, f, indent=2, ensure_ascii=False, default=str)
        f.write("\n")


def catalog_ds_numbers():
    """Set of DS numbers that exist in the product catalog."""
    import glob
    out = set()
    for path in glob.glob(os.path.join(PRODUCTS_DIR, "**", "*.json"), recursive=True):
        with open(path, encoding="utf-8") as f:
            out.add(str(json.load(f).get("ds_number")).strip())
    return out
