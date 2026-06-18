# Guide for Bruce — Using the DS Setup Repo Safely

This repo holds all our product setup data. You don't need to be a programmer to use
it — **just talk to Claude Code in plain English.** This guide explains what's safe,
what to avoid, and the words to say.

---

## The one thing to remember

**You can't break anything permanently.** Every change is tracked by git and reviewed
before it becomes official. Nothing you do touches the "real" catalog (`main`) until
someone approves it. So explore freely.

---

## One-time setup

1. Install: **Git**, **Python** (with `pip install openpyxl`), and **GitHub CLI** (`gh`).
2. Get access to the repo from Jackson (he'll add you on GitHub).
3. In a terminal:
   ```
   git clone https://github.com/jmaitner/DS_Setup_Repo
   cd DS_Setup_Repo
   gh auth login
   claude
   ```
4. Now just type what you want. Claude knows the rules (they live in `CLAUDE.md`).

---

## Things you can ask Claude to do

| You want to… | Say something like… |
|---|---|
| Add a new setup sheet | "Import this setup sheet: `C:\...\Bladez 2026 Setup.xlsx`" |
| Look something up | "What's the cost and case qty for all Head Start products?" |
| Get images | "Give me every image URL for First4 Figures" |
| Fix a wrong value | "DS12345's MSRP should be 29.99, not 24.99" |
| Make upload files | "Export WMS and Shopify files for the Bladez brand" |
| Build a spreadsheet | "Make me an Excel of all products under $10 wholesale with their MSRP" |

Claude will always tell you what it's about to do **before** it does it. Read that, and
say "go ahead" if it looks right.

---

## How a change becomes official (the safe loop)

1. Claude makes the change on a **branch** (a private copy — `main` is untouched).
2. Claude runs **validation** to catch mistakes (missing UPC, bad price, duplicates).
3. Claude opens a **Pull Request** — a request to merge your change.
4. **Jackson (or you) reviews it** on GitHub: you see exactly what changed, side by side.
5. Approve → it merges into `main` and is now official. Or request changes / close it.

If validation finds an error, Claude will stop and tell you. Don't push past it — fix
it or ask Jackson.

---

## Golden rules (the same ones Claude follows)

1. **Never edit `main` directly.** Always a branch + Pull Request. (Claude handles this.)
2. **Don't overwrite when info conflicts.** We get wrong vendor data all the time. If the
   new value disagrees with what's there, Claude will show you both and ask which is
   right — **ask Jackson if you're unsure.** Don't guess.
3. **Validate before pushing.** Claude runs `validate.py` for you.
4. **When in doubt, ask** — in chat to Claude, or message Jackson. Nothing is urgent
   enough to skip the review step.

---

## If you get stuck

- "I don't understand what this PR is changing" → ask Claude: "explain this diff in plain
  English."
- "Validation failed" → ask Claude: "what's the error and how do I fix it?"
- Still stuck → send Jackson the PR link.
