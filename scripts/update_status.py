#!/usr/bin/env python3
"""
update_status.py  —  Edit lifecycle and/or a channel's status for one or more DS#.

Examples:
  # Mark items discontinued
  python scripts/update_status.py 5531025 5531026 --lifecycle discontinued \\
      --lifecycle-note "Vendor EOL Jan 2026"

  # Set a channel's listing state + case number
  python scripts/update_status.py 5531540 --channel walmart_1p --state error \\
      --case CASE-12345 --issue "Main image rejected"

  # Mark a listing live with its id
  python scripts/update_status.py 5531525 --channel amazon --state live --id B0ABC123

Creates the status file if missing. Stamps 'updated' to today.
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
from status_lib import (CHANNELS, CHANNEL_STATES, LIFECYCLES,
                        new_status_doc, load_status, save_status)


def main():
    ap = argparse.ArgumentParser(description="Update channel/lifecycle status for DS#(s).")
    ap.add_argument("ds", nargs="+", help="one or more DS numbers")
    ap.add_argument("--lifecycle", choices=LIFECYCLES)
    ap.add_argument("--lifecycle-note")
    ap.add_argument("--channel", choices=list(CHANNELS))
    ap.add_argument("--state", choices=CHANNEL_STATES)
    ap.add_argument("--id", dest="listing_id")
    ap.add_argument("--url", dest="listing_url")
    ap.add_argument("--case", dest="case_number")
    ap.add_argument("--issue")
    ap.add_argument("--notes")
    args = ap.parse_args()

    if not args.lifecycle and not args.channel:
        sys.exit("Nothing to do: pass --lifecycle and/or --channel (with field flags).")
    if args.channel and not any([args.state, args.listing_id, args.listing_url,
                                 args.case_number, args.issue, args.notes]):
        sys.exit("--channel given but no field to set (--state/--id/--url/--case/--issue/--notes).")

    today = date.today().isoformat()
    for ds in args.ds:
        ds = str(ds).strip().lstrip("DSds")
        doc = load_status(ds) or new_status_doc(ds, today)

        if args.lifecycle:
            doc["lifecycle"] = args.lifecycle
        if args.lifecycle_note is not None:
            doc["lifecycle_note"] = args.lifecycle_note

        if args.channel:
            ch = doc["channels"].setdefault(args.channel, {})
            if args.state:        ch["state"] = args.state
            if args.listing_id:   ch["id"] = args.listing_id
            if args.listing_url:  ch["url"] = args.listing_url
            if args.case_number:  ch["case_number"] = args.case_number
            if args.issue:        ch["issue"] = args.issue
            if args.notes:        ch["notes"] = args.notes
            ch["updated"] = today

        doc["updated"] = today
        save_status(doc)
        summary = []
        if args.lifecycle: summary.append(f"lifecycle={args.lifecycle}")
        if args.channel: summary.append(f"{args.channel}={doc['channels'][args.channel].get('state')}")
        print(f"  DS{ds}: {', '.join(summary) or 'updated'}")

    print("\nNext: python scripts/validate.py && python scripts/build_index.py")


if __name__ == "__main__":
    main()
