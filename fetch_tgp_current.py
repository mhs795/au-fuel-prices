"""Scrape AIP's live terminal gate price tables, to carry TGP past the workbook.

Same job as fetch_gas_current.py does for gas: the master file is monthly-ish and the
live source is the only thing that reaches today.

AIP's August 2026 move to WordPress broke the workbook - www.aip.com.au now serves one
ending 26 June - but the app that actually produces the numbers is still running on
api.aip.com.au and still publishing every weekday. The new site's own pages query it
through a broken SQL view and render nothing, so this reads the app directly.

  fetch_tgp_current.py <out.csv>

It publishes a rolling five business days, so the out CSV is a cache: it is merged, not
overwritten, and grows a day per run. Nothing here is irreplaceable the way the Gas
Supply Hub cache is - the workbook carries the same numbers once AIP fixes it, and the
workbook wins wherever it reaches - but a daily build keeps the series unbroken in the
meantime.

Two known limits, both deliberate:
  * Seven capitals only. AIP's volume-weighted National average is published in the
    workbook and nowhere else, so national is left blank for these days rather than
    recomputed here with weights that are not AIP's.
  * http only - api.aip.com.au serves no TLS. These are published prices, and every
    value is bounds-checked below before it is written.
"""
import csv, os, re, sys, urllib.request
from datetime import datetime

URL = "http://api.aip.com.au/public/tgpTables"
UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
OUT = sys.argv[1] if len(sys.argv) > 1 else "work/petrol/tgp_current.csv"
CITIES = ["Sydney", "Melbourne", "Brisbane", "Adelaide", "Perth", "Darwin", "Hobart"]
# Cents per litre, GST inclusive. Wide enough to survive any plausible market, tight
# enough to catch a parse that has picked up a year, a percentage or a stray index.
LO, HI = 50.0, 500.0


def text(fragment):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", fragment)).strip()


def parse(html):
    """-> {(iso date, fuel, city): price}. Reads the ULP table then the diesel one."""
    out = {}
    for fuel, heading in (("petrol", "Petrol (ULP"), ("diesel", "Diesel (")):
        start = html.find(heading)
        if start < 0:
            raise SystemExit(f"{URL}: no '{heading}' heading - the page has changed shape")
        m = re.search(r"(?s)<table.*?</table>", html[start:])
        if not m:
            raise SystemExit(f"{URL}: no table under '{heading}'")
        rows = re.findall(r"(?s)<tr.*?</tr>", m.group(0))

        dates = []
        for cell in re.findall(r"(?s)<th.*?</th>", rows[0]):
            t = text(cell)
            # "Wednesday 12 August 2026" - the weekday is part of the cell, not a column
            hit = re.search(r"(\d{1,2} [A-Za-z]+ \d{4})", t)
            if hit:
                dates.append(datetime.strptime(hit.group(1), "%d %B %Y").date().isoformat())
        if len(dates) < 2:
            raise SystemExit(f"{URL}: read {len(dates)} dates from the {fuel} header")

        seen = set()
        for row in rows[1:]:
            cells = [text(c) for c in re.findall(r"(?s)<t[hd].*?</t[hd]>", row)]
            if not cells:
                continue
            city = next((c for c in CITIES if c.lower() == cells[0].lower()), None)
            if not city:
                continue
            seen.add(city)
            for d, cell in zip(dates, cells[1:]):
                try:
                    price = float(cell)
                except ValueError:
                    continue    # a genuinely blank day, not a parse failure
                if not (LO <= price <= HI):
                    raise SystemExit(f"{URL}: {fuel} {city} {d} reads {price} c/L, outside "
                                     f"[{LO}, {HI}] - refusing to write it")
                out[(d, fuel, city.lower())] = price
        missing = [c for c in CITIES if c not in seen]
        if missing:
            raise SystemExit(f"{URL}: {fuel} table is missing {missing}")
    return out


def merge(path, fresh):
    """Fold the rolling window into the cache. Fresh values win on the days they cover."""
    cols = [f"TGP_{f}_{c.lower()}" for f in ("petrol", "diesel") for c in CITIES]
    rows = {}
    if os.path.exists(path):
        with open(path, newline="") as fh:
            for r in csv.DictReader(fh):
                d = r.get("date", "")[:10]
                if d:
                    rows[d] = {c: r[c] for c in cols if r.get(c) not in (None, "")}
    added = 0
    for (d, fuel, city), price in fresh.items():
        if d not in rows:
            rows[d] = {}
            added += 1
        rows[d][f"TGP_{fuel}_{city}"] = price

    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["date"] + cols)
        for d in sorted(rows):
            w.writerow([d] + [rows[d].get(c, "") for c in cols])
    os.replace(tmp, path)
    return sorted(rows), added


req = urllib.request.Request(URL, headers={"User-Agent": UA})
fresh = parse(urllib.request.urlopen(req, timeout=120).read().decode("utf-8", errors="replace"))
dates, added = merge(OUT, fresh)
window = sorted({d for d, _, _ in fresh})
print(f"wrote {OUT} days {len(dates)} {dates[0]} -> {dates[-1]} "
      f"(live window {window[0]} -> {window[-1]}, {added} new)", flush=True)
