# ARC — build tracker

Live page: https://cybersecdigestdaily.github.io/arc/
Spreadsheet (rebuilt on every check): https://cybersecdigestdaily.github.io/arc/arc-build.xlsx
Live CSV (for Google Sheets `=IMPORTDATA(...)`): https://cybersecdigestdaily.github.io/arc/prices.csv

Tracks UK prices and stock for every part of an all-black AM5 living-room gaming PC in a Jonsbo D33 (13 lines incl. Windows), keeps price history, moves parts to cheaper vetted UK shops and suggests compatible swaps.

## How it works

| Layer | What | When |
|---|---|---|
| Price/stock checker | `checker.py` reads `build.json`, pulls each part's full UK offer list from PriceSpy UK (shop, price, stock, condition, delivery) plus the shops' own pages (schema.org data or Shopify product data), keeps vetted UK shops only, ranks offers by item price + delivery to Northern Ireland, and writes `docs/data.json` and `docs/history.json` | GitHub Actions, about 4× a day, and on every change to `build.json` |
| Spreadsheet | `sheet.py` turns the same results into `docs/arc-build.xlsx` (build, prices at every shop, every offer, price history, swaps and upgrades, fit check) and `docs/prices.csv`. Both are published with the page, not committed | Every checker run |
| Amazon check | Amazon blocks automated checks from the cloud, so a Claude scheduled task on Don's computer reads every Amazon Prime listing in the built-in browser (delivery set to an NI postcode), records verified prices under retailer "Amazon Prime" and adds cheaper Prime finds, including PriceSpy's cheaper Amazon marketplace listings (`parts[].amazon_check`) | Daily, about 07:21 UK, when the computer is on |
| Daily sweep | A Claude scheduled task applies button presses from the page (emailed), verifies prices the checker can't read live (Scan, Watercooling UK and CCL block bots; Overclockers UK's PriceSpy feed lags its site), vets new shops' NI delivery, finds cheaper UK offers, proposes swaps that pass the fit check with equal or better spec, fixes dead links, and commits `build.json` | Daily, about 07:46 UK |
| Page | `docs/index.html` (static), published to the `gh-pages` branch by the Action | Always |

## Rules

- UK shops only, and they must deliver to Northern Ireland. New stock only: marketplace sellers (eBay, OnBuy, non-Prime Amazon sellers…), grey-market key sellers, used/refurbished stock and unvetted shops are ignored. Amazon counts when sold by Amazon (PriceSpy shop "Amazon.co.uk") or dispatched by Amazon (Prime) with free NI delivery (retailer "Amazon Prime", verified on the item page).
- Totals and picks come from the cheapest basket: item prices plus one NI delivery per shop, with free-delivery thresholds (`free_over`) applied, so a cheap item rides along on an order that's already paying delivery. Each part also keeps its cheapest offer on its own (item + single-order NI delivery) for alerts and price history.
- A retailer marked `"pricespy_feed": false` is never taken from PriceSpy (its feed lags); its own pages and the sweep's verified prices count instead.
- The same part at a cheaper vetted shop moves automatically. A different model is only ever a suggestion until approved, and must pass the D33 fit check (`build.json` → `fit_limits`), match or beat the current part's spec, and save at least £10 (colour swaps excepted).
- Bought parts stop being checked. The sweep switches itself off once everything is bought.

## Common edits

- **Bought something:** press **Bought it** on the page (opens a pre-filled email; the next sweep applies it), or set `"bought": true`, `"paid"`, `"paid_retailer"` and `"bought_on"` on the part in `build.json`.
- **Swap ideas:** **Approve** / **Not for me** on the page, or set the suggestion's `"status"` to `"approved"` / `"rejected"`.
- **Vet a shop:** add it under `retailers` using the exact PriceSpy shop name, with `ni` (`verified` / `assumed` / `unverified` / `no`) and `ni_first` (single-parcel NI delivery cost).
- Saving `build.json` re-runs the checker (and rebuilds the spreadsheet) within a couple of minutes.

## Files

- `build.json` — the build: parts, PriceSpy ids and match rules, known shop offers, fit limits, vetted shops, swap ideas, upgrades, sweep-verified prices, log.
- `checker.py` — the price/stock checker. `sheet.py` — the spreadsheet/CSV export. Python standard library only.
- `docs/index.html` — the page. `docs/data.json`, `docs/history.json` — written by the checker; don't hand-edit.

Public repo: no names, addresses or other personal details anywhere.
