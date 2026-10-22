#!/usr/bin/env python3
"""ARC price checker.

Reads build.json, checks live UK prices and stock for every part that isn't bought yet, and writes
docs/data.json (what the page shows) and docs/history.json (one price point per part per UK day).

Sources, in order of trust:
  1. PriceSpy UK product pages: every UK shop's price, stock, condition and GB delivery for a product.
  2. Direct retailer checks for offers marked "shopify" (product .js endpoint) or "jsonld" (schema.org data).
  3. "manual" prices verified by Claude's daily sweep (build.json -> manual), used while fresh.

Only shops listed in build.json -> retailers can be picked as the best offer. Marketplaces, used stock and
shops in blocked_shops are ignored; any other shop is reported as "unvetted" for the daily sweep to vet.
The best offer is the cheapest in-stock one by landed cost: item price + delivery to Northern Ireland
for a single-item order.

Runs on GitHub Actions (~4x a day and on every change to build.json). Standard library only.
Also rebuilds docs/arc-build.xlsx and docs/prices.csv (see sheet.py) on every run.
"""

import gzip
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

ROOT = os.path.dirname(os.path.abspath(__file__))
BUILD = os.path.join(ROOT, "build.json")
DOCS = os.path.join(ROOT, "docs")
DATA = os.path.join(DOCS, "data.json")
HIST = os.path.join(DOCS, "history.json")
XLSX = os.path.join(DOCS, "arc-build.xlsx")
CSV = os.path.join(DOCS, "prices.csv")

UK = ZoneInfo("Europe/London")
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36")
DELAY = float(os.environ.get("ARC_DELAY", "2"))
HIST_DAYS = 180
EVENTS_KEEP = 80
OFFERS_SHOWN = 10
IN_STOCK = {"instock", "limitedavailability", "limitedstock", "instoreonly"}
OUT_STOCK = {"outofstock", "soldout", "discontinued", "preorder", "backorder", "incoming"}


# ---------------------------------------------------------------- helpers

def now_utc():
    return datetime.now(timezone.utc)


def iso(dt):
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def uk_today():
    return now_utc().astimezone(UK).date()


def load(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return default


def save(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
        f.write("\n")
    os.replace(tmp, path)


def money(x):
    return None if x is None else round(float(x) + 1e-9, 2)


def gbp(x):
    return "£{:,.2f}".format(x) if x is not None else "—"


def dig(d, *path):
    """Safe nested get. PriceSpy's flight data uses "$..." strings as references/undefined: treat as None."""
    for k in path:
        if isinstance(d, dict):
            d = d.get(k)
        elif isinstance(d, list) and isinstance(k, int) and -len(d) <= k < len(d):
            d = d[k]
        else:
            return None
    if isinstance(d, str) and d.startswith("$"):
        return None
    return d


class Blocked(Exception):
    pass


def fetch(url, accept="text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"):
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": accept,
        "Accept-Language": "en-GB,en;q=0.9",
        "Accept-Encoding": "gzip",
    })
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            body = r.read()
            if (r.headers.get("Content-Encoding") or "").lower() == "gzip":
                body = gzip.decompress(body)
    except urllib.error.HTTPError as e:
        if e.code in (401, 403, 429, 503):
            raise Blocked("HTTP %d" % e.code)
        raise
    text = body.decode("utf-8", "ignore")
    if re.search(r"<title>\s*(Just a moment|Attention Required|Access denied)", text[:4000], re.I):
        raise Blocked("bot check")
    return text


def stock_flag(status):
    s = (status or "").replace(" ", "").replace("_", "").lower()
    if s in IN_STOCK:
        return True
    if s in OUT_STOCK:
        return False
    return None


def matches(name, rule):
    if not rule:
        return True
    name = name or ""
    for pat in rule.get("must", []):
        if not re.search(pat, name, re.I):
            return False
    for pat in rule.get("not", []):
        if re.search(pat, name, re.I):
            return False
    return True


# ---------------------------------------------------------------- sources

def pricespy(pid):
    """Return (offers, title). Raises Blocked or ValueError (markup changed)."""
    html = fetch("https://pricespy.co.uk/product.php?p=%s" % pid)
    title = re.search(r"<title>([^<]*)", html)
    title = title.group(1).strip() if title else ""
    chunks = []
    for m in re.finditer(r'self\.__next_f\.push\(\[1,"((?:[^"\\]|\\.)*)"\]\)', html, re.S):
        try:
            chunks.append(json.loads('"' + m.group(1) + '"'))
        except ValueError:
            pass
    flight = "".join(chunks)
    dec = json.JSONDecoder()
    raw, seen = [], set()
    for m in re.finditer(r'\{"index":\d+,"shopId":', flight):
        try:
            o, _ = dec.raw_decode(flight, m.start())
        except ValueError:
            continue
        key = o.get("id") or (o.get("shopId"), o.get("shopOfferId"))
        if key in seen:
            continue
        seen.add(key)
        raw.append(o)
    if not raw:
        count = re.search(r'"offerCount"\s*:\s*(\d+)', html)
        if count and int(count.group(1)) > 0:
            raise ValueError("offer list not found (page markup may have changed)")
    offers = []
    for o in raw:
        offers.append({
            "shop": dig(o, "shop", "name") or str(o.get("shopId")),
            "sku": dig(o, "shopOfferId"),
            "name": dig(o, "name") or "",
            "price": money(dig(o, "price", "amount")),
            "was": money(dig(o, "originalPrice", "amount")),
            "stock": dig(o, "stockStatus") or "Unknown",
            "condition": dig(o, "condition") or "Unknown",
            "ship": money(dig(o, "shipping", "nodes", 0, "price", "amount")),
            "url": dig(o, "clickoutUrl"),
            "source": "pricespy",
        })
    return offers, title


def shopify(url):
    u = url.split("?")[0].rstrip("/")
    if not u.endswith(".js"):
        u += ".js"
    d = json.loads(fetch(u, accept="application/json"))
    price = d.get("price")
    if isinstance(price, int):
        price = price / 100.0  # Shopify .js prices are integer pence
    avail = d.get("available")
    return {"price": money(price), "stock": "InStock" if avail else "OutOfStock", "name": d.get("title") or ""}


def jsonld(url):
    html = fetch(url)
    for block in re.findall(r'<script[^>]*application/ld\+json[^>]*>(.*?)</script>', html, re.S | re.I):
        try:
            d = json.loads(block.strip())
        except ValueError:
            continue
        items = d if isinstance(d, list) else d.get("@graph", [d]) if isinstance(d, dict) else []
        for it in items:
            if not isinstance(it, dict) or "Product" not in str(it.get("@type")):
                continue
            offers = it.get("offers")
            offers = offers[0] if isinstance(offers, list) and offers else offers
            if not isinstance(offers, dict):
                continue
            price = offers.get("price") or offers.get("lowPrice")
            avail = str(offers.get("availability") or "").rsplit("/", 1)[-1]
            if price is not None:
                return {"price": money(price), "stock": avail or "Unknown", "name": it.get("name") or ""}
    m = re.search(r'product:price:amount"\s+content="([\d.]+)"', html)
    if m:
        a = re.search(r'product:availability"\s+content="([^"]+)"', html)
        return {"price": money(m.group(1)), "stock": a.group(1) if a else "Unknown", "name": ""}
    raise ValueError("no price found")


# ---------------------------------------------------------------- pricing logic

def delivery(shop, listed_ship, cfg):
    """(cost, basis, note) for a single-item order to Northern Ireland."""
    r = cfg["retailers"].get(shop, {})
    ni = r.get("ni", "unverified")
    est = float(cfg["settings"].get("ni_default_estimate", 15))
    if ni == "no":
        return None, "no", "Doesn't deliver to NI"
    if r.get("ni_first") is not None:
        return float(r["ni_first"]), ni, r.get("note") or ""
    if ni == "assumed":
        return float(listed_ship or 0), ni, r.get("note") or "Delivery cost at checkout"
    if ni == "verified":
        return est, "estimate", r.get("note") or "Delivery shown at checkout (estimate used)"
    cost = float(listed_ship) if listed_ship else est
    return cost, "unverified", "NI delivery not confirmed yet (estimate used)"


def shop_status(shop, cfg):
    if shop in cfg["retailers"]:
        return "allowed"
    low = (shop or "").lower()
    if any(low == b.lower() for b in cfg.get("blocked_shops", [])):
        return "blocked"
    return "unvetted"


def display(shop, cfg):
    return cfg["retailers"].get(shop, {}).get("display") or shop


def build_offers(item, cfg, today, prev_item, run):
    """Collect and rank offers for a part or a suggestion. Returns a dict of results."""
    rule = item.get("match")
    known = {(o.get("retailer"), (o.get("sku") or "").lower()): o for o in item.get("offers", []) if o.get("sku")}
    known_url = {}
    for o in item.get("offers", []):
        if o.get("url") and "/search?" not in o["url"]:
            known_url.setdefault(o.get("retailer"), o["url"])
    offers, unvetted, notes = [], [], []
    excluded = {"marketplace": 0, "used": 0, "other_variant": 0, "untrusted_feed": 0}
    status = "no_source"
    checked = None

    # 1) PriceSpy
    pid = item.get("pricespy")
    if pid:
        try:
            raw, _title = pricespy(pid)
            run["pricespy_ok"] += 1
            status, checked = "live", iso(now_utc())
            for o in raw:
                sku_hit = (o["shop"], (o["sku"] or "").lower()) in known
                if not sku_hit and not matches(o["name"], rule):
                    excluded["other_variant"] += 1
                    continue
                if (o["condition"] or "").lower() not in ("new", "unknown"):
                    excluded["used"] += 1
                    continue
                if cfg["retailers"].get(o["shop"], {}).get("pricespy_feed") is False:
                    excluded["untrusted_feed"] += 1  # this shop's PriceSpy prices lag its site; use direct/sweep prices
                    continue
                st = shop_status(o["shop"], cfg)
                if st == "blocked":
                    excluded["marketplace"] += 1
                    continue
                if st == "unvetted":
                    unvetted.append({"shop": o["shop"], "price": o["price"], "stock": o["stock"], "name": o["name"][:90]})
                    continue
                if sku_hit or o["shop"] in known_url:
                    o["url"] = known_url.get(o["shop"]) or o["url"]
                offers.append(o)
        except Blocked as e:
            run["pricespy_blocked"] += 1
            status = "blocked"
            notes.append("PriceSpy refused this run (%s); showing the last good prices" % e)
        except Exception as e:  # noqa: BLE001
            run["errors"].append("%s: PriceSpy %s: %s" % (item.get("id"), pid, e))
            status = "error"
            notes.append("PriceSpy page couldn't be read (%s)" % e)
        time.sleep(DELAY)

    # Carry forward the last good PriceSpy offers when this run couldn't read them
    if status in ("blocked", "error") and prev_item:
        for o in prev_item.get("offers", []):
            if o.get("source") == "pricespy":
                c = dict(o)
                c["stale"] = True
                offers.append(c)
        checked = prev_item.get("checked")

    # 2) Direct retailer checks
    live_shops = {o["shop"] for o in offers if not o.get("stale")}
    for o in item.get("offers", []):
        kind = o.get("check")
        if kind not in ("shopify", "jsonld") or o.get("retailer") in live_shops:
            continue
        try:
            r = shopify(o["url"]) if kind == "shopify" else jsonld(o["url"])
            offers.append({"shop": o["retailer"], "sku": o.get("sku"), "name": r.get("name") or item.get("name"),
                           "price": r["price"], "was": None, "stock": r["stock"], "condition": "New",
                           "ship": None, "url": o["url"], "source": "direct"})
            live_shops.add(o["retailer"])
            if status in ("no_source", "blocked", "error"):
                status = "live" if status == "no_source" else status
            checked = checked or iso(now_utc())
        except Exception as e:  # noqa: BLE001
            notes.append("%s check failed (%s)" % (display(o["retailer"], cfg), e))
        time.sleep(DELAY / 2)

    # 3) Prices verified by the daily sweep
    fresh_days = int(cfg["settings"].get("manual_fresh_days", 3))
    for m in cfg.get("manual", {}).get(item.get("id"), []):
        if m.get("retailer") in live_shops or m.get("price") is None:
            continue
        try:
            age = (today - date.fromisoformat(m.get("checked_on"))).days
        except (TypeError, ValueError):
            age = 999
        offers.append({"shop": m["retailer"], "sku": m.get("sku"), "name": item.get("name"),
                       "price": money(m["price"]), "was": None,
                       "stock": "InStock" if m.get("in_stock") else "OutOfStock",
                       "condition": "New", "ship": None, "url": m.get("url") or known_url.get(m["retailer"]),
                       "source": "verified", "checked_on": m.get("checked_on"),
                       "stale": age > fresh_days, "note": m.get("note")})

    # Rank: landed cost (item + NI delivery for a single-item order)
    for o in offers:
        cost, basis, note = delivery(o["shop"], o.get("ship"), cfg)
        o["delivery"] = money(cost)
        o["ni"] = basis
        o["delivery_note"] = note
        o["landed"] = money(o["price"] + cost) if (o.get("price") is not None and cost is not None) else None
        o["in_stock"] = stock_flag(o.get("stock"))
        o["display"] = display(o["shop"], cfg)

    def pool(stock_value, allow_stale):
        return [o for o in offers if o["landed"] is not None and o["in_stock"] is stock_value
                and o["ni"] != "no" and (allow_stale or not o.get("stale"))]

    best = None
    for stock_value, allow_stale in ((True, False), (True, True), (None, False), (None, True), (False, False), (False, True)):
        p = pool(stock_value, allow_stale)
        if p:
            best = min(p, key=lambda o: (o["landed"], o["price"]))
            break

    offers.sort(key=lambda o: (0 if o["in_stock"] else 1 if o["in_stock"] is None else 2,
                               o["landed"] if o["landed"] is not None else 1e9))
    # One row per shop in the visible list (keep each shop's best)
    shown, seen_shops = [], set()
    for o in offers:
        if o["shop"] in seen_shops:
            continue
        seen_shops.add(o["shop"])
        shown.append(o)
    cheaper_unvetted = sorted(
        [u for u in unvetted if u["price"] is not None and stock_flag(u["stock"]) is not False
         and (best is None or u["price"] < best["price"])],
        key=lambda u: u["price"])[:5]
    return {
        "status": status if offers or status != "no_source" else "no_source",
        "checked": checked,
        "best": slim(best),
        "offers": [slim(o) for o in shown[:OFFERS_SHOWN]],
        "unvetted": cheaper_unvetted,
        "excluded": excluded,
        "notes": notes,
    }


def slim(o):
    if not o:
        return None
    keys = ("shop", "display", "sku", "name", "price", "was", "delivery", "landed", "ni", "delivery_note",
            "stock", "in_stock", "source", "url", "stale", "checked_on", "note")
    return {k: o.get(k) for k in keys if o.get(k) not in (None, "", False) or k in ("in_stock",)}


# ---------------------------------------------------------------- fit check

def fit_checks(specs, lim):
    """specs maps part id -> spec dict; lim is build.json fit_limits. Returns a list of checks."""
    cooler = specs.get("cooler", {})
    ram = specs.get("ram", {})
    gpu = specs.get("gpu", {})
    psu = specs.get("psu", {})
    board = specs.get("board", {})
    bottom = specs.get(lim.get("bottom_fan_part", "fans"), {})
    out = []

    def add(check, value, limit, unit, ok, part, kind="max", note=None):
        c = {"check": check, "value": value, "limit": limit, "unit": unit, "ok": ok, "part": part, "kind": kind}
        if note:
            c["note"] = note
        out.append(c)

    def total(key):
        return sum(v.get(key, 0) for v in specs.values() if isinstance(v.get(key, 0), (int, float)))

    if "radiator_mm" in cooler and lim.get("radiator_mm"):
        add("Top radiator size", cooler["radiator_mm"], " / ".join(str(x) for x in lim["radiator_mm"]), "mm",
            cooler["radiator_mm"] in lim["radiator_mm"], "cooler", "in", "AIO radiator in the top mount")
    elif "height_mm" in cooler and lim.get("cooler_height_mm"):
        add("CPU cooler height", cooler["height_mm"], lim["cooler_height_mm"], "mm",
            cooler["height_mm"] <= lim["cooler_height_mm"], "cooler")
    if "height_mm" in ram and lim.get("ram_height_mm"):
        clearance = min(lim["ram_height_mm"], cooler.get("ram_clearance_mm", lim["ram_height_mm"]))
        add("RAM height", ram["height_mm"], clearance, "mm", ram["height_mm"] <= clearance, "ram")
    if "length_mm" in gpu and lim.get("gpu_length_mm"):
        add("GPU length", gpu["length_mm"], lim["gpu_length_mm"], "mm", gpu["length_mm"] <= lim["gpu_length_mm"], "gpu")
    if "thickness_mm" in gpu and lim.get("gpu_plus_bottom_fan_mm"):
        t = gpu["thickness_mm"] + bottom.get("thickness_mm", 0)
        add("GPU + bottom fan thickness", t, lim["gpu_plus_bottom_fan_mm"], "mm", t <= lim["gpu_plus_bottom_fan_mm"], "gpu",
            note="%smm card + %smm bottom fans" % (gpu["thickness_mm"], bottom.get("thickness_mm", 0)))
    if "length_mm" in psu and lim.get("psu_length_mm"):
        add("PSU length", psu["length_mm"], lim["psu_length_mm"], "mm", psu["length_mm"] <= lim["psu_length_mm"], "psu")
    if "watts" in psu and lim.get("psu_min_w"):
        add("PSU wattage", psu["watts"], lim["psu_min_w"], "W", psu["watts"] >= lim["psu_min_w"], "psu", "min")
        if lim.get("est_load_w"):
            pct = round(lim["est_load_w"] / psu["watts"] * 100)
            add("Full-load draw vs PSU", "%sW (%d%%)" % (lim["est_load_w"], pct), "80%", "", pct <= 80, "psu", "in",
                "Estimated whole-system draw at full load")
    fans = total("case_fans")
    if fans and lim.get("case_fans_max"):
        add("Case fans", fans, lim["case_fans_max"], "", fans <= lim["case_fans_max"], "fans")
    if "form" in board and lim.get("motherboard"):
        add("Motherboard size", board["form"], " / ".join(lim["motherboard"]), "", board["form"] in lim["motherboard"], "board", "in")
    usb = total("usb2_header")
    if usb and "usb2_headers" in board:
        add("USB 2.0 headers (screens)", usb, board["usb2_headers"], "", usb <= board["usb2_headers"], "board",
            note="The cooler's LCD and the front screen each take one")
    groups = total("argb_groups")
    if groups and "argb_headers" in board:
        extra = sum(max(0, v.get("argb_outputs", 1) - 1) for v in specs.values() if "argb_outputs" in v)
        cap = board["argb_headers"] + extra
        add("ARGB lighting groups", groups, cap, "", groups <= cap, "board",
            note="%d board headers%s" % (board["argb_headers"], " + %d from the splitter" % extra if extra else ""))
    fh = total("fan_headers_needed")
    if fh and "fan_headers" in board:
        add("Fan/pump headers", fh, board["fan_headers"], "", fh <= board["fan_headers"], "board")
    return out


# ---------------------------------------------------------------- history + events

def upsert_history(series, day, res):
    best = res.get("best")
    if not best:
        return
    entry = [day, best["price"], best["landed"], 1 if best.get("in_stock") else 0, best.get("display")]
    if series and series[-1][0] == day:
        old = series[-1]
        better = (entry[3] and not old[3]) or (entry[3] == old[3] and entry[2] is not None
                                                and (old[2] is None or entry[2] < old[2]))
        if better:
            series[-1] = entry
    else:
        series.append(entry)
    cutoff = (date.fromisoformat(day) - timedelta(days=HIST_DAYS)).isoformat()
    while series and series[0][0] < cutoff:
        series.pop(0)


def lows(series, today_s):
    past = [e for e in series if e[0] != today_s and e[3] and e[1] is not None]
    if not past:
        return None
    low = min(past, key=lambda e: e[1])
    return {"price": low[1], "date": low[0], "retailer": low[4], "days": len({e[0] for e in series})}


def yesterday_price(series, today_s):
    """Best price on the previous UK day, only if we have a point for exactly that day."""
    y = (date.fromisoformat(today_s) - timedelta(days=1)).isoformat()
    past = [e for e in series if e[0] == y and e[1] is not None]
    return past[-1][1] if past else None


def detect_events(part, res, prev, cfg, ts):
    s = cfg["settings"]
    ev = []
    best = res.get("best")
    pbest = (prev or {}).get("best")
    name = part["slot"]
    if not best or best.get("stale"):
        return ev
    was_in = bool(pbest and pbest.get("in_stock"))
    now_in = bool(best.get("in_stock"))
    if pbest and was_in and not now_in:
        ev.append({"ts": ts, "level": "alert", "type": "soldout", "part": part["id"],
                   "text": "%s: no vetted UK shop has it in stock right now" % name})
    elif pbest and now_in and not was_in:
        ev.append({"ts": ts, "level": "alert", "type": "restock", "part": part["id"],
                   "text": "%s back in stock: %s at %s" % (name, gbp(best["price"]), best["display"])})
    if pbest and now_in and was_in and best.get("landed") is not None and pbest.get("landed") is not None:
        drop = pbest["landed"] - best["landed"]
        pct = drop / pbest["landed"] * 100 if pbest["landed"] else 0
        big = drop >= s["alert_drop_gbp"] or (pct >= s["alert_drop_pct"] and drop >= s["alert_drop_min_gbp"])
        if big:
            via = "" if best["display"] == pbest.get("display") else " (was %s at %s)" % (gbp(pbest["price"]), pbest.get("display"))
            ev.append({"ts": ts, "level": "alert", "type": "drop", "part": part["id"],
                       "text": "%s down %s: %s at %s%s" % (name, gbp(drop), gbp(best["price"]), best["display"], via)})
        elif -drop >= s["alert_drop_gbp"]:
            ev.append({"ts": ts, "level": "info", "type": "rise", "part": part["id"],
                       "text": "%s up %s to %s at %s" % (name, gbp(-drop), gbp(best["price"]), best["display"])})
    return ev


# ---------------------------------------------------------------- main

def main():
    cfg = load(BUILD, None)
    if not cfg:
        sys.exit("build.json missing or invalid")
    cfg["settings"].setdefault("ni_default_estimate", 15)
    prev = load(DATA, {})
    hist = load(HIST, {"_about": "One point per part per UK day: [date, price, landed, in_stock, shop]", "parts": {}, "suggestions": {}})
    hist.setdefault("parts", {})
    hist.setdefault("suggestions", {})
    start = now_utc()
    ts = iso(start)
    today = uk_today()
    today_s = today.isoformat()
    run = {"started": ts, "source": os.environ.get("ARC_SOURCE", "github-actions" if os.environ.get("GITHUB_ACTIONS") else "local"),
           "pricespy_ok": 0, "pricespy_blocked": 0, "errors": []}
    prev_parts = {p["id"]: p for p in prev.get("parts", [])}
    prev_sugg = {sgg["id"]: sgg for sgg in prev.get("suggestions", [])}
    prev_upg = {u["id"]: u for u in prev.get("upgrades", [])}
    base_date = cfg.get("baseline", {}).get("date")

    parts_out, events, specs = [], [], {}
    for part in cfg["parts"]:
        pid = part["id"]
        specs[pid] = part.get("spec", {})
        series = hist["parts"].setdefault(pid, [])
        b = part.get("baseline") or {}
        if not series and b.get("price") is not None and base_date and base_date < today_s:
            cost, _, _ = delivery(b.get("retailer"), None, cfg)
            series.append([base_date, b["price"], money(b["price"] + (cost or 0)), 1, display(b.get("retailer"), cfg)])
        out = {k: part.get(k) for k in ("id", "slot", "name", "bought", "paid", "paid_retailer", "bought_on", "spec", "pricespy", "note")}
        out["baseline"] = dict(b, date=base_date) if b else None
        out["phase1"] = pid in cfg["settings"].get("phase1", {}).get("parts", [])
        if part.get("bought"):
            out.update({"status": "bought", "best": None, "offers": [], "history": series[-90:]})
            parts_out.append(out)
            continue
        res = build_offers(part, cfg, today, prev_parts.get(pid), run)
        upsert_history(series, today_s, res)
        out.update(res)
        best = res.get("best")
        y = yesterday_price(series, today_s)
        out["change"] = {
            "vs_yesterday": money(best["price"] - y) if best and y is not None else None,
            "vs_baseline": money(best["price"] - b["price"]) if best and b.get("price") is not None else None,
            "vs_prev_run": money(best["landed"] - prev_parts[pid]["best"]["landed"])
            if best and prev_parts.get(pid, {}).get("best") and best.get("landed") is not None
            and prev_parts[pid]["best"].get("landed") is not None else None,
        }
        lo = lows(series, today_s)
        out["low"] = lo
        sig = []
        if best:
            if not best.get("in_stock"):
                sig.append("soldout" if best.get("in_stock") is False else "stock_unknown")
            if b.get("price") is not None and best["price"] <= b["price"] * 0.95:
                sig.append("below_sheet")
            if lo and lo["days"] >= cfg["settings"]["low_min_days"] and best.get("in_stock"):
                if best["price"] < lo["price"]:
                    sig.append("new_low")
                elif best["price"] <= lo["price"] * (1 + cfg["settings"]["near_low_pct"] / 100):
                    sig.append("near_low")
            if best.get("stale"):
                sig.append("not_live")
            if best.get("ni") in ("unverified", "estimate"):
                sig.append("ni_check")
        else:
            sig.append("no_price")
        out["signals"] = sig
        out["history"] = series[-90:]
        events += detect_events(part, res, prev_parts.get(pid), cfg, ts)
        parts_out.append(out)

    # Suggestions (only pending ones are checked live)
    by_id = {p["id"]: p for p in parts_out}
    sugg_out = []
    for sg in cfg.get("suggestions", []):
        o = {k: sg.get(k) for k in ("id", "part", "name", "spec", "why", "status", "added", "decided", "pricespy")}
        if sg.get("status") == "pending":
            res = build_offers(sg, cfg, today, prev_sugg.get(sg["id"]), run)
            series = hist["suggestions"].setdefault(sg["id"], [])
            upsert_history(series, today_s, res)
            o.update({k: res[k] for k in ("status_live", "best", "offers", "checked", "notes") if k in res})
            o["live"] = res["status"]
            cur = by_id.get(sg["part"], {}).get("best")
            o["saving"] = money(cur["landed"] - res["best"]["landed"]) if cur and res.get("best") and cur.get("landed") is not None and res["best"].get("landed") is not None else None
            alt_specs = dict(specs)
            alt_specs[sg["part"]] = dict(specs.get(sg["part"], {}), **(sg.get("spec") or {}))
            checks = fit_checks(alt_specs, cfg["fit_limits"])
            o["fit_ok"] = all(c["ok"] for c in checks)
            o["fit_problems"] = [c["check"] for c in checks if not c["ok"]]
            prev_o = prev_sugg.get(sg["id"])
            th = cfg["settings"]["suggest_min_saving_gbp"]
            if o["saving"] is not None and o["saving"] >= th and (res.get("best") or {}).get("in_stock") \
                    and not (prev_o and (prev_o.get("saving") or 0) >= th):
                events.append({"ts": ts, "level": "alert", "type": "swap", "part": sg["part"],
                               "text": "Swap idea: %s saves %s (%s at %s)" % (sg["name"], gbp(o["saving"]),
                                                                             gbp(res["best"]["price"]), res["best"]["display"])})
        sugg_out.append(o)

    # Upgrades / alternatives: tracked for reference only (no events, no approval flow)
    upg_out = []
    for u in cfg.get("upgrades", []):
        o = {k: u.get(k) for k in ("id", "part", "name", "spec", "why", "pricespy")}
        res = build_offers(u, cfg, today, prev_upg.get(u["id"]), run)
        o.update({"best": res.get("best"), "offers": res.get("offers", [])[:5], "checked": res.get("checked"), "live": res["status"]})
        cur = by_id.get(u["part"], {})
        cur_price = cur.get("paid") if cur.get("bought") else (cur.get("best") or {}).get("price")
        o["extra"] = money(res["best"]["price"] - cur_price) if res.get("best") and cur_price is not None else None
        alt_specs = dict(specs)
        alt_specs[u["part"]] = dict(specs.get(u["part"], {}), **(u.get("spec") or {}))
        checks = fit_checks(alt_specs, cfg["fit_limits"])
        o["fit_ok"] = all(c["ok"] for c in checks)
        o["fit_problems"] = [c["check"] for c in checks if not c["ok"]]
        upg_out.append(o)

    # Totals
    unbought = [p for p in parts_out if not p.get("bought")]
    paid = money(sum((p.get("paid") or 0) for p in parts_out if p.get("bought")))
    remaining = money(sum(p["best"]["price"] for p in unbought if p.get("best")))
    missing = [p["slot"] for p in unbought if not p.get("best")]
    shops = {}
    for p in unbought:
        bst = p.get("best")
        if bst:
            shops.setdefault(bst["shop"], []).append(bst)
    deliv, deliv_lines = 0.0, []
    for shop, items in shops.items():
        r = cfg["retailers"].get(shop, {})
        first = items[0].get("delivery") or 0
        extra = r.get("ni_extra")
        cost = first  # one order per shop; checkout decides parcels
        deliv += cost
        deliv_lines.append({"shop": display(shop, cfg), "parts": len(items), "cost": money(cost),
                            "basis": items[0].get("ni"), "extra_parcel": extra, "note": r.get("order_note")})
    phase = cfg["settings"].get("phase1", {})
    p1 = [p for p in parts_out if p["id"] in phase.get("parts", [])]
    p1_cost = money(sum((p.get("paid") if p.get("bought") else (p.get("best") or {}).get("price") or 0) for p in p1))
    totals = {
        "parts_total": len(parts_out), "parts_left": len(unbought), "paid": paid, "remaining": remaining,
        "delivery": money(deliv), "delivery_lines": deliv_lines,
        "build_total": money((paid or 0) + (remaining or 0)),
        "baseline_parts": cfg.get("baseline", {}).get("parts_total"),
        "delta_vs_baseline": money((paid or 0) + (remaining or 0) - cfg.get("baseline", {}).get("parts_total", 0))
        if cfg.get("baseline", {}).get("parts_total") else None,
        "missing": missing,
        "budget": cfg["settings"].get("budget"),
        "left_after_delivery": money(cfg["settings"]["budget"] - (paid or 0) - (remaining or 0) - deliv)
        if cfg["settings"].get("budget") else None,
    }
    phase1 = dict(phase, cost=p1_cost, left=money(phase.get("budget", 0) - (p1_cost or 0)),
                  items=[{"id": p["id"], "slot": p["slot"], "price": p.get("paid") if p.get("bought") else (p.get("best") or {}).get("price"),
                          "bought": p.get("bought")} for p in p1]) if phase else None

    fit = fit_checks(specs, cfg["fit_limits"])
    old_events = prev.get("events", [])
    all_events = (events + old_events)[:EVENTS_KEEP]

    run["finished"] = iso(now_utc())
    data = {
        "generated": ts,
        "today": today_s,
        "title": cfg.get("title"),
        "title_long": cfg.get("title_long"),
        "subtitle": cfg.get("subtitle"),
        "site": cfg.get("site"),
        "updated": cfg.get("updated"),
        "last_sweep": cfg.get("last_sweep"),
        "baseline": cfg.get("baseline"),
        "settings": cfg.get("settings"),
        "fit_limits": cfg.get("fit_limits"),
        "retailers": {k: {kk: v.get(kk) for kk in ("display", "ni", "ni_first", "ni_extra", "note", "checked_on")}
                      for k, v in cfg.get("retailers", {}).items()},
        "run": run,
        "totals": totals,
        "phase1": phase1,
        "parts": parts_out,
        "suggestions": sugg_out,
        "upgrades": upg_out,
        "fit": fit,
        "events": all_events,
        "log": cfg.get("log", [])[-40:],
    }
    # Spreadsheet + CSV (published with the page). Never let an export problem stop the price data.
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("arc_sheet", os.path.join(ROOT, "sheet.py"))
        sheet = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(sheet)
        sheet.export(data, hist, cfg, XLSX, CSV, cfg.get("site") or "")
        data["exports"] = {"xlsx": "arc-build.xlsx", "csv": "prices.csv", "built": iso(now_utc())}
    except Exception as e:  # noqa: BLE001
        run["errors"].append("spreadsheet export: %s" % e)
    save(DATA, data)
    save(HIST, hist)
    print("ARC check %s: %d PriceSpy pages ok, %d blocked, %d errors, %d events. Build %s (%s vs sheet)." % (
        ts, run["pricespy_ok"], run["pricespy_blocked"], len(run["errors"]), len(events),
        gbp(totals["build_total"]), gbp(totals["delta_vs_baseline"])))
    for p in parts_out:
        b = p.get("best")
        print("  %-11s %-10s %s" % (p["slot"], gbp(b["price"]) if b else "—",
                                    "%s · %s · %s" % (b.get("display"), b.get("stock"), b.get("source")) if b else p.get("status")))
    for e in events:
        print("  EVENT", e["level"], e["text"])
    for err in run["errors"]:
        print("  ERROR", err)


if __name__ == "__main__":
    main()
