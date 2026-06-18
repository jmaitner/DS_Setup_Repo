---
description: Import a DS setup sheet (.xlsx) into the catalog on a branch + PR
argument-hint: <path-to-setup-sheet.xlsx>
---

Import the setup sheet at: $ARGUMENTS

Follow the safe workflow exactly:

1. Confirm the file exists and is the 2026 template (has a "DS Only" sheet).
2. Create a branch: `git checkout -b import/<brand>-<today>` (derive brand from the sheet).
3. Run `python scripts/import_setup_sheet.py "$ARGUMENTS"`.
   - It is non-destructive by default (skips products that already exist). If it reports
     skips and the user actually wants to update them, re-run with `--update` ONLY after
     showing the user which products would change.
4. Run `python scripts/validate.py`. If there are ERRORS, fix them and re-validate. Do not
   continue past errors.
5. Run `python scripts/build_index.py`.
6. Show the user a summary: how many products created/updated/skipped, and any warnings.
7. Commit: `git add -A && git commit -m "Import <Brand> setup sheet"`.
8. Push and open a PR: `git push -u origin HEAD && gh pr create --fill`.
9. Give the user the PR link and tell them it needs review before merging.

Never commit directly to `main`.
