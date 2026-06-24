#!/usr/bin/env python3
"""
case.py  —  Create / update / attribute cases (the compliance & issues layer).

Upsert by id (create if new, else update only the fields you pass). Attribute to items
(--ds) and/or a brand (--brand), tag them, set status, and optionally stamp the linked
items' channel status with the case number.

Examples:
  # create a case from an Amazon support case, link items, tag it
  python scripts/case.py --source amazon --case-number 20917568511 \\
      --title "High return-rate flag" --status open --tags returns,listing --ds 5530713

  # attribute an existing case to more items + mark resolved
  python scripts/case.py --id amazon-20917568511 --ds 5530714,5530715 --status resolved

  # brand-level vendor/compliance issue
  python scripts/case.py --source compliance --title "Target WERCs registration (Canal)" \\
      --brand "Canal Toys" --status resolved --tags compliance,werCs,target

  python scripts/case.py --list
"""

import argparse
import os
import sys
from datetime import date

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from case_lib import (STATUSES, slug_id, load_case, save_case, new_case, all_cases)
from status_lib import CHANNELS, load_status, save_status


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--id")
    ap.add_argument("--source")
    ap.add_argument("--case-number", dest="case_number")
    ap.add_argument("--title")
    ap.add_argument("--desc", dest="description")
    ap.add_argument("--status", choices=STATUSES)
    ap.add_argument("--tags", help="comma-separated (replaces existing)")
    ap.add_argument("--add-tags", help="comma-separated (appends)")
    ap.add_argument("--brand")
    ap.add_argument("--ds", help="comma-separated DS#s to add to linked_ds")
    ap.add_argument("--email-link", dest="email_link")
    ap.add_argument("--email-id", dest="email_id")
    ap.add_argument("--notes")
    ap.add_argument("--stamp", action="store_true",
                    help="also set the case_number on linked items' channel (source must be a channel)")
    ap.add_argument("--stamp-state", choices=list(__import__("status_lib").CHANNEL_STATES), default="error")
    args = ap.parse_args()

    if args.list:
        for c in sorted(all_cases(), key=lambda c: (c.get("status"), c.get("source"))):
            print(f"  [{c['status']:11}] {c['id']:34} {','.join(c.get('tags') or []):24} "
                  f"ds={len(c.get('linked_ds') or [])}  {c['title'][:50]}")
        return

    today = date.today().isoformat()
    cid = args.id or slug_id(args.source, args.case_number, args.title)
    doc = load_case(cid)
    if not doc:
        if not args.source or not args.title:
            sys.exit("New case needs at least --source and --title.")
        doc = new_case(cid, args.source, args.title, today)

    for fld in ("source", "case_number", "title", "description", "status", "brand",
                "email_link", "email_id", "notes"):
        v = getattr(args, fld)
        if v is not None:
            doc[fld] = v
    if args.tags is not None:
        doc["tags"] = [t.strip() for t in args.tags.split(",") if t.strip()]
    if args.add_tags:
        doc["tags"] = sorted(set((doc.get("tags") or []) + [t.strip() for t in args.add_tags.split(",") if t.strip()]))
    if args.ds:
        add = [d.strip() for d in args.ds.split(",") if d.strip()]
        doc["linked_ds"] = sorted(set((doc.get("linked_ds") or []) + add))
    doc["updated"] = today
    save_case(doc)
    print(f"Saved case {cid}  status={doc['status']}  items={len(doc.get('linked_ds') or [])}  tags={doc.get('tags')}")

    if args.stamp and doc["source"] in CHANNELS:
        n = 0
        for ds in doc.get("linked_ds") or []:
            st = load_status(ds)
            if not st:
                continue
            ch = st["channels"].setdefault(doc["source"], {})
            if doc.get("case_number"):
                ch["case_number"] = doc["case_number"]
            ch["state"] = args.stamp_state
            ch["issue"] = doc["title"]
            ch["updated"] = today
            st["updated"] = today
            save_status(st)
            n += 1
        print(f"  stamped {doc['source']} case#/issue on {n} item(s)")


if __name__ == "__main__":
    main()
