# ARC — build tracker

Live page: https://cybersecdigestdaily.github.io/arc/

Tracks UK prices and stock for every part of an AM5 small-form-factor gaming build in a Jonsbo Z20, keeps price history, moves parts to cheaper vetted UK shops and suggests cheaper compatible swaps.

## How it works

| Layer | What | When |
|---|---|---|
| Price/stock checker | `checker.py` reads `build.json`, pulls each part's full UK offer list from PriceSpy UK (shop, price, stock, condition, delivery) plus direct shop checks, keeps vetted UK shops only, ranks offers by item price + delivery to Northern Ireland, and writes `docs/data.json` and `docs/history.json` | GitHub Actions, about 4× a day, and on every change to `build.json` |
| Daily sweep | A Claude scheduled task applies button presses from the page (emailed), verifies prices the checker can't read live (Overclockers UK, Scan, Watercooling UK and CCL block bots), vets new shops' NI delivery, finds cheaper UK offers, proposes swaps that pass the fit check with equal or better spec, fixes dead links, and commits `build.json` | Daily, about 07:46 UK |
| Page | `docs/index.html` (static), published to the `gh-pages` branch by the Action | Always |

## Rules

- UK shops only, and they must deliver to Northern Ireland. New stock only: marketplace sellers (eBay, OnBuy, Amazon Marketplace…), used/refurbished stock and unvetted shops are ignored.
- The headline price per part is the cheapest in-stock vetted offer by landed cost: item + NI delivery for a single-item order. Buying several parts from one shop together usually saves on delivery.
- The same part at a cheaper vetted shop moves automatically. A different model is only ever a suggestion until approved, and must pass the Z20 fit check (`build.json` → `fit_limits`), match or beat the current part's spec, and save at least £10.
- Bought parts stop being checked. The sweep switches itself off once everything is bought.

## Common edits

- **Bought something:** press **Bought it** on the page (opens a pre-filled email; the next sweep applies it), or set `"bought": true`, `"paid"`, `"paid_retailer"` and `"bought_on"` on the part in `build.json`.
- **Swap ideas:** **Approve** / **Not for me** on the page, or set the suggestion's `"status"` to `"approved"` / `"rejected"`.
- **Vet a shop:** add it under `retailers` using the exact PriceSpy shop name, with `ni` (`verified` / `assumed` / `unverified` / `no`) and `ni_first` (single-parcel NI delivery cost).
- Saving `build.json` re-runs the checker within a couple of minutes.

## Files

- `build.json` — the build: parts, PriceSpy ids and match rules, known shop offers, fit limits, vetted shops, swap ideas, sweep-verified prices, log.
- `checker.py` — the price/stock checker (Python standard library only).
- `docs/index.html` — the page. `docs/data.json`, `docs/history.json` — written by the checker; don't hand-edit.

Public repo: no names, addresses or other personal details anywhere.
