---
description: Update an item's lifecycle or a channel's setup status (branch + PR)
argument-hint: <DS#(s) and what changed, e.g. "5531025 discontinued" or "5531540 walmart 1p live">
---

The user wants to update channel/lifecycle status for: $ARGUMENTS

Status lives in `status/DS#####.json` (separate from product data, joined by DS#). Follow
the safe workflow:

1. Identify the DS number(s) and what's changing (lifecycle, or a specific channel's state/
   case number/listing id/issue/notes). The 10 channels: amazon, walmart_1p, walmart_3p,
   tiktok, bestbuy, nocnoc, shopify_f2f, toysrus, ebay, target_plus.
2. Create a branch: `git checkout -b status/<ds-or-batch>-<what>`.
3. Apply with `scripts/update_status.py`, e.g.:
   - `python scripts/update_status.py 5531025 --lifecycle discontinued --lifecycle-note "..."`
   - `python scripts/update_status.py 5531540 --channel walmart_1p --state error --case CASE-123 --issue "..."`
   - `python scripts/update_status.py 5531525 --channel amazon --state live --id B0ABC123`
4. `python scripts/validate.py` then `python scripts/build_index.py`.
5. Show the user the git diff of the status file(s); confirm before committing.
6. Commit, push, open a PR, share the link.

Never commit to `main` directly.
