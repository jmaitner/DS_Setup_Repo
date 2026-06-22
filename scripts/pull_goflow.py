#!/usr/bin/env python3
"""
pull_goflow.py  —  Sync the channel-status layer from GoFlow (the hub).

GoFlow's GET /v1/listings reports, per listing: product.item_number (= our DS#),
store.channel (which channel), status (active/in_review/inactive/unknown),
store_provided_id (the channel's listing id), and store_page_url. One paginated sweep
covers every channel you've connected in GoFlow, so this single puller can drive most/all
of the status layer.

Reads GOFLOW_KEY (bearer token) and GOFLOW_BASE_URL from the environment. Updates
status/DS#####.json in place, touching ONLY GoFlow-derived fields (state / id / url /
updated) and preserving anything set by hand (case_number / issue / notes).

Usage (env: GOFLOW_KEY required):
  GOFLOW_BASE_URL=https://distributionsolutions.goflow.com python scripts/pull_goflow.py
  python scripts/pull_goflow.py --stores-only      # just list connected stores (auth check)
  python scripts/pull_goflow.py --limit 200        # stop after N listings (testing)
  python scripts/pull_goflow.py --dry-run          # report, write nothing
"""

import argparse
import json
import os
import sys
import urllib.parse
import urllib.request
from datetime import date

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from status_lib import (CHANNELS, new_status_doc, load_status, save_status,
                        catalog_ds_numbers)

BASE = (os.environ.get("GOFLOW_BASE_URL") or "https://distributionsolutions.goflow.com").rstrip("/")
KEY = os.environ.get("GOFLOW_KEY", "")

# GoFlow store.channel  ->  our channel key. Only the channels we track; others ignored.
CHANNEL_MAP = {
    "amazon_marketplace_usa": "amazon",
    "amazon_vendor_usa": "amazon",
    "walmart_marketplace_usa": "walmart_3p",
    "walmart_vendor_usa": "walmart_1p",
    "tik_tok": "tiktok",
    "best_buy": "bestbuy",
    "noc_noc": "nocnoc",
    "shopify": "shopify_f2f",
    "toys_r_us_usa": "toysrus",
    "ebay": "ebay",
    "target_plus_marketplace": "target_plus",
}

# GoFlow listing.status  ->  our channel state. 'unknown' => leave existing state alone.
STATE_MAP = {"active": "live", "in_review": "pending", "inactive": "not_listed"}
# precedence so a 'live' listing wins over a 'pending'/'not_listed' for the same DS#+channel
RANK = {"live": 3, "pending": 2, "not_listed": 1}


def api_get(url):
    if not url.startswith("http"):
        url = BASE + ("" if url.startswith("/") else "/") + url
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {KEY}",
        "Accept": "application/json",
    })
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "replace")[:300]
        sys.exit(f"ERROR: GoFlow {e.code} on {url}\n  {body}\n"
                 f"  (check GOFLOW_KEY / that auth is a Bearer token)")
    except Exception as e:
        sys.exit(f"ERROR: request failed on {url}: {e}")


def get_stores():
    stores = {}  # store id -> {channel, our, name, status}
    data = api_get("/v1/stores")
    for s in data.get("data", []):
        gch = s.get("channel")
        stores[s.get("id")] = {
            "channel": gch, "our": CHANNEL_MAP.get(gch),
            "name": s.get("name"), "status": s.get("status"),
        }
    return stores


def iter_listings(limit=None):
    url, n = "/v1/listings", 0
    while url:
        data = api_get(url)
        for row in data.get("data", []):
            yield row
            n += 1
            if limit and n >= limit:
                return
        url = data.get("next")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--stores-only", action="store_true")
    ap.add_argument("--limit", type=int)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    if not KEY:
        sys.exit("ERROR: GOFLOW_KEY not set in environment.")
    print(f"GoFlow base: {BASE}")

    stores = get_stores()
    tracked = {sid: s for sid, s in stores.items() if s["our"]}
    print(f"\nConnected stores: {len(stores)}  (tracked channels: {len(tracked)})")
    for sid, s in sorted(stores.items(), key=lambda kv: str(kv[1]['channel'])):
        mark = f"-> {s['our']}" if s["our"] else "(not tracked)"
        print(f"  [{s['status']}] {s['channel']}  {mark}")
    present = sorted({s["our"] for s in tracked.values()})
    missing = [c for c in CHANNELS if c not in present]
    print(f"\nOur channels covered by GoFlow: {present}")
    print(f"Our channels NOT in GoFlow (need another source): {missing}")
    if args.stores_only:
        return

    catalog = catalog_ds_numbers()
    today = date.today().isoformat()
    # best[(ds, our_channel)] = (rank, state, id, url)
    best = {}
    seen_listings = matched = unmatched = 0

    for li in iter_listings(limit=args.limit):
        seen_listings += 1
        ds = str((li.get("product") or {}).get("item_number") or "").strip()
        store = li.get("store") or {}
        our = CHANNEL_MAP.get(store.get("channel"))
        if not our:
            continue
        gstatus = li.get("status")
        state = STATE_MAP.get(gstatus)
        if not state:  # 'unknown' or unmapped -> don't touch state
            continue
        if not ds or ds not in catalog:
            unmatched += 1
            continue
        matched += 1
        rank = RANK.get(state, 0)
        key = (ds, our)
        if key not in best or rank > best[key][0]:
            best[key] = (rank, state, li.get("store_provided_id"), li.get("store_page_url"))

    # apply to status files
    by_ds = {}
    for (ds, our), (_, state, lid, url) in best.items():
        by_ds.setdefault(ds, {})[our] = (state, lid, url)

    changed = 0
    per_channel = {c: {} for c in CHANNELS}
    for ds, chans in by_ds.items():
        doc = load_status(ds) or new_status_doc(ds, today)
        touched = False
        for our, (state, lid, url) in chans.items():
            ch = doc["channels"].setdefault(our, {})
            if ch.get("state") != state or ch.get("id") != lid or ch.get("url") != url:
                ch["state"], ch["id"], ch["url"] = state, lid, url
                ch["updated"] = today
                touched = True
            per_channel[our][state] = per_channel[our].get(state, 0) + 1
        if touched:
            doc["updated"] = today
            changed += 1
            if not args.dry_run:
                save_status(doc)

    print(f"\nListings seen: {seen_listings}   matched to catalog: {matched}   "
          f"unmatched: {unmatched}")
    print(f"Status docs {'that would change' if args.dry_run else 'updated'}: {changed}")
    print("Per-channel states set:")
    for c in CHANNELS:
        if per_channel[c]:
            print(f"  {CHANNELS[c]:24} " + ", ".join(f"{k}={v}" for k, v in per_channel[c].items()))
    if not args.dry_run:
        print("\nNext: python scripts/validate.py && python scripts/build_index.py && python scripts/build_dashboard.py")


if __name__ == "__main__":
    main()
