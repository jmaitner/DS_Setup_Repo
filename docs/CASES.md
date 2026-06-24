# The Cases / Compliance Layer

Cases are issues that need tracking — marketplace support cases (Amazon/Walmart case #s),
vendor/compliance threads (e.g. a Target WERCs registration handled with Canal), pricing
reviews, safety-doc requests, etc. Each lives in `cases/<id>.json`, **attributed to items
(DS#) and/or a brand**, with tags and a status. It's a third layer alongside product `data`
and channel `status`, joined by DS#.

```
cases/<id>.json   id · source · case_number · title · description · status · tags
                  · brand · linked_ds[] · email_link · email_id · opened/updated · notes
```
- **status:** `needs_review` (auto-found, awaiting your confirmation) · `open` · `pending` · `resolved` · `closed`
- **id:** the channel case # when there is one (`amazon-20917568511`, `walmart-3p-15089649`), else a `source-title` slug.
- A case can cover many items; an item can have many cases. `case_number` cross-links to a
  product's `status` channel.

## How cases get created
1. **Automatically** — the weekday `channel-case-morning-digest` routine scans Outlook broadly
   for issue/case emails, creates **draft cases (`needs_review`)** with suggested tags + linked
   items/brand, rebuilds, and digests them for you.
2. **By hand (via Claude)** — tell Claude "open a case for …" or "attach case X to DS#… / the
   Canal brand, tag it compliance, mark resolved", and it runs `scripts/case.py`.

## Working with cases
```
python scripts/case.py --list
python scripts/case.py --source amazon --case-number 20917568511 \
    --title "High return-rate flag" --status open --tags returns,listing --ds 5530713
python scripts/case.py --id amazon-20917568511 --status resolved          # update
python scripts/case.py --id walmart-3p-15089649 --ds 5531001,5531002 --stamp  # also stamp items' channel case#
```
The dashboard's **Cases** view lists everything (open first), filterable, with tags, linked
items, and a link to the source email. Edits follow the normal branch → validate → PR flow;
`validate.py` checks each case against `schema/case.schema.json` and warns on linked DS#s
that aren't in the catalog.
