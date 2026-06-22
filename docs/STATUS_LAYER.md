# The Status Layer — channel setup status, separate from product data

Product **data** (what an item is) and channel **status** (where it's listed and how it's
doing) change for different reasons, so they live in separate places joined by **DS#**:

```
products/<vendor>/DS#####.json   ← product data (name, price, images, compliance)
status/DS#####.json              ← lifecycle + per-channel status  ← THIS LAYER
index/status.csv                 ← auto-generated flat view (a column per channel)
```

## What a status file holds

- **lifecycle**: `active` · `planned` · `on_hold` · `discontinued` · `dropped` (+ a note)
- **channels** — one entry per sales channel, each with:
  - **state**: `not_listed` · `planned` · `setup_in_progress` · `pending` · `live` · `error` · `suppressed` · `discontinued`
  - **id** (ASIN / item id / listing id), **case_number** (Amazon/Walmart support case),
    **issue** (current error/blocker), **notes**, **updated** (date)

The 10 channels: **Amazon, Walmart 1P, Walmart 3P, TikTok Shop, Best Buy, NocNoc,
Shopify (Face2FaceFun), Toys R Us, eBay, Target Plus.**

## How to use it

| Task | How |
|---|---|
| New products got imported | `python scripts/init_status.py` (creates a status record for each, all channels `not_listed`) |
| Mark an item discontinued | `python scripts/update_status.py <DS#> --lifecycle discontinued --lifecycle-note "..."` |
| Record a channel listing / error | `python scripts/update_status.py <DS#> --channel amazon --state live --id B0XXX` |
| Bulk-update from a channel export/error report | `python scripts/ingest_status.py "report.xlsx" --channel walmart_1p --match-type sku --match-col "Seller SKU" --set-state live` |
| See everything at a glance | open `index/status.csv` |

`ingest_status.py` matches report rows to products by **DS#**, **SKU** (vendor item number),
or listing **id**, and updates just that one channel — so you can drop in a Walmart error
report, an Amazon "active listings" export, etc., and it folds the results into the clean
status DB, reporting any rows it couldn't match. Claude will inspect the report's columns
and pick the right `--*-col` flags for you.

## Automated sync from GoFlow (the hub)

GoFlow is the single source for most channels. `scripts/pull_goflow.py` calls GoFlow's
`GET /v1/listings`, matches each listing to our catalog by `product.item_number` (= DS#),
maps `store.channel` → our channel and GoFlow `status` (active/in_review/inactive) →
our state (live/pending/not_listed), and records the listing id + URL. It updates only
GoFlow-derived fields (state/id/url) and preserves hand-set fields (case_number/issue/notes).

`.github/workflows/sync-status.yml` runs it **Tue + Fri** (and on-demand via "Run workflow"),
then rebuilds the index + dashboard and opens a PR with the changes. Needs repo secret
`GOFLOW_KEY`. First run: trigger manually and check the logs — it prints which channels are
connected in GoFlow (covered by the hub) vs. which need another source.

## Rules
Same as everything else: edits go on a **branch + PR**, `validate.py` checks status files
against `schema/status.schema.json` (and that every status DS# has a real product), and
`build_index.py` refreshes `index/status.csv`. Use `/update-status` and `/ingest-status`.
