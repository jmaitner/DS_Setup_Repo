#!/usr/bin/env python3
"""
init_status.py  —  Create a status/DS#####.json for every catalog product that
doesn't have one yet (lifecycle=active, all 10 channels = not_listed).

Non-destructive: never touches an existing status file. Run after importing new
products so each has a status record to fill in.

Usage:
  python scripts/init_status.py
  python scripts/init_status.py --dry-run
"""

import sys

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from status_lib import new_status_doc, status_path, save_status, catalog_ds_numbers


def main():
    dry = "--dry-run" in sys.argv
    # date is passed in to avoid Date.now()-style nondeterminism concerns; use today
    from datetime import date
    today = date.today().isoformat()

    created, existing = 0, 0
    for ds in sorted(catalog_ds_numbers()):
        if os.path.exists(status_path(ds)):
            existing += 1
            continue
        if dry:
            print(f"  ~ would create status for DS{ds}")
        else:
            save_status(new_status_doc(ds, today))
        created += 1

    print(f"\nStatus files {'to create' if dry else 'created'}: {created}   "
          f"already existed: {existing}")
    if not dry:
        print("Next: python scripts/validate.py && python scripts/build_index.py")


if __name__ == "__main__":
    main()
