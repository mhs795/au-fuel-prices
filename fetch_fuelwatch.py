"""Download FuelWatch WA monthly retail CSVs and aggregate to daily averages.

Each monthly file is reduced to a small per-month aggregate cached on disk. Closed
months never change, so only the current and previous month are re-fetched on a
rerun - the first build downloads ~1.3 GB, later ones a few MB. The raw station-level
files are streamed and discarded; only the aggregates are kept.
"""
import csv, io, json, os, sys, urllib.request
from collections import defaultdict
from datetime import datetime, date

UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
BASE = "https://www.fuelwatch.wa.gov.au/api/report/monthly-retail-prices"
OUT = sys.argv[1] if len(sys.argv) > 1 else "fw_daily.csv"
CACHE = os.path.join(os.path.dirname(os.path.abspath(OUT)), "fw_cache")
START_YEAR = 2015
FIELDS = ["date", "product", "scope", "mean_price_cpl", "n_observations"]


def fetch(url, tries=4):
    for a in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
            return urllib.request.urlopen(req, timeout=300).read()
        except Exception:
            if a == tries - 1:
                raise
            import time; time.sleep(5 * (a + 1))


def month_key(name):
    """FuelWatchRetail-MM-YYYY.csv -> (YYYY, MM)"""
    parts = name.replace(".csv", "").split("-")
    return parts[2], parts[1]


def aggregate(raw):
    agg = defaultdict(lambda: [0.0, 0])
    rdr = csv.DictReader(io.StringIO(raw.decode("utf-8-sig", errors="replace")))
    for row in rdr:
        try:
            price = float(row["PRODUCT_PRICE"])
        except (TypeError, ValueError):
            continue
        if price <= 0:
            continue
        try:
            d = datetime.strptime(row["PUBLISH_DATE"].strip(), "%d/%m/%Y").date().isoformat()
        except (ValueError, AttributeError):
            continue
        prod = (row["PRODUCT_DESCRIPTION"] or "").strip()
        scopes = ["WA"] + (["Metro"] if (row["REGION_DESCRIPTION"] or "").strip() == "Metro" else [])
        for scope in scopes:
            e = agg[(d, prod, scope)]
            e[0] += price
            e[1] += 1
    return agg


os.makedirs(CACHE, exist_ok=True)
listing = json.loads(fetch(f"{BASE}?dateFrom={START_YEAR}-01-01&dateTo=2030-12-31"))
files = sorted((f for f in listing if int(month_key(f["fileName"])[0]) >= START_YEAR),
               key=lambda f: month_key(f["fileName"]))

today = date.today()
cur = (f"{today.year:04d}", f"{today.month:02d}")
pm = date(today.year, today.month, 1).toordinal() - 1
prev_d = date.fromordinal(pm)
prev = (f"{prev_d.year:04d}", f"{prev_d.month:02d}")

fetched = reused = 0
for f in files:
    key = month_key(f["fileName"])
    cp = os.path.join(CACHE, f"{key[0]}-{key[1]}.csv")
    # closed months are final; only the live and just-closed months can still change
    if os.path.exists(cp) and os.path.getsize(cp) > 0 and key not in (cur, prev):
        reused += 1
        continue
    agg = aggregate(fetch(f["url"]))
    tmp = cp + ".tmp"
    with open(tmp, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(FIELDS)
        for (d, prod, scope), (s, c) in sorted(agg.items()):
            w.writerow([d, prod, scope, round(s / c, 3), c])
    os.replace(tmp, cp)
    fetched += 1
    print(f"  fetched {f['fileName']} ({len(agg)} daily/product rows)", flush=True)

print(f"{fetched} month(s) downloaded, {reused} reused from cache", flush=True)

rows = []
for name in sorted(os.listdir(CACHE)):
    if name.endswith(".csv"):
        with open(os.path.join(CACHE, name), newline="") as fh:
            rows.extend(list(csv.DictReader(fh)))
rows.sort(key=lambda r: (r["date"], r["product"], r["scope"]))
with open(OUT, "w", newline="") as fh:
    w = csv.DictWriter(fh, FIELDS)
    w.writeheader()
    w.writerows(rows)
print("wrote", OUT, len(rows), "rows", flush=True)
