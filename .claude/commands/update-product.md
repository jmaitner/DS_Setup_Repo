---
description: Safely change a field on one or more products (branch + diff review + PR)
argument-hint: <DS number(s) and what to change>
---

The user wants to update: $ARGUMENTS

This is the sensitive path — we protect against blindly overwriting good data with wrong
data. Follow exactly:

1. Locate the product file(s): `products/<brand-slug>/DS<number>.json`.
2. Show the user the CURRENT value of the field(s) they want to change.
3. **If the new value conflicts with existing data**, do NOT just overwrite. Show both the
   old and proposed new value, note the source (`_meta.source_file`) and what git blame
   says, and ask the user to confirm which is correct. The human decides.
4. Create a branch: `git checkout -b fix/DS<number>-<short-what>`.
5. Make the edit. Update `_meta.last_verified` to today's date.
6. Run `python scripts/validate.py` then `python scripts/build_index.py`.
7. Show the user the `git diff` and get explicit confirmation before committing.
8. Commit with a message that records the reason and the old→new value:
   `git commit -m "Fix <field> DS<number>: <old> -> <new> (<reason>)"`.
9. Push and open a PR: `git push -u origin HEAD && gh pr create --fill`. Share the link.

Never commit directly to `main`. Never pick a value silently when data conflicts.
