"""Download and cache AEMO Gas Supply Hub daily transaction summaries from nemweb.

  fetch_gsh.py <cache_dir> <out.csv>

Each published file carries a rolling window of the last 30 gas days, and nemweb keeps
roughly the last 95 daily files, so the deepest history reachable in one run is about four
months - there is no master file going back to the hub's 2014 start, which is why this
series was a documented gap for so long. The cache is therefore the point: every file ever
downloaded is kept, and the output is the union of all of them, so the series grows one day
per run and reaches back further than any single fetch could. Never clear the cache to
"force a refresh" - published files are immutable, and what would be lost is the only copy
of the days that have already rolled off nemweb.

Where two cached files cover the same gas date, the later publication wins: files are
processed in filename order, which sorts by gas date then by publication ID.

Only Wallumbilla and South East Queensland are written out. They are the hub's commodity
gas trading locations; the other product locations in the same file (MSP, MAP, SWQP, RBP,
EGP, QGP, CUL, WIL, UGS NNT, WAL Compression) are pipeline capacity, compression and
storage products, priced per GJ of service rather than per GJ of gas, and would be a
different quantity charted on the same axis. The raw files are cached whole, so adding a
location later needs no refetching.
"""
import csv, io, os, re, sys, time, zipfile, urllib.request, urllib.error
from collections import defaultdict

CACHE, OUT = sys.argv[1], sys.argv[2]
BASE = "https://nemweb.com.au/Reports/Current/GSH/GSH_Historical_Trans_Summary/"
UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
# (file's product location, column stem in the output)
HUBS = [("WAL", "GSH_WAL"), ("SEQ", "GSH_SEQ")]


def fetch(url, tries=4):
    err = ""
    for a in range(tries):
        if a:
            time.sleep(3 * a)
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            return urllib.request.urlopen(req, timeout=90).read()
        except Exception as e:
            err = f"{type(e).__name__} {e}"
    raise SystemExit(f"ERROR: could not fetch {url} after {tries} tries: {err}")


os.makedirs(CACHE, exist_ok=True)
listing = fetch(BASE).decode("utf-8", "replace")
names = sorted(set(re.findall(r"PUBLIC_HISTORICALTRANSACTIONSUMMARY_\d+_\d+\.zip", listing)))
if not names:
    raise SystemExit(f"ERROR: no GSH files found in the listing at {BASE}. The report "
                     f"directory has probably moved; the cache still holds every file "
                     f"downloaded so far, so do not delete it.")

new = 0
for n in names:
    p = os.path.join(CACHE, n)
    if os.path.exists(p) and os.path.getsize(p) > 200:
        continue                      # published files never change
    body = fetch(BASE + n)
    with open(p + ".tmp", "wb") as fh:
        fh.write(body)
    os.replace(p + ".tmp", p)
    new += 1
cached = sorted(f for f in os.listdir(CACHE) if f.endswith(".zip"))
print(f"  {new} new file(s) downloaded, {len(cached) - new} already cached, "
      f"{len(names)} currently listed on nemweb", flush=True)

# date -> location -> (price, quantity), later files overwriting earlier ones
day = defaultdict(dict)
for n in cached:
    with zipfile.ZipFile(os.path.join(CACHE, n)) as z:
        inner = [m for m in z.namelist() if m.upper().endswith(".CSV")]
        if len(inner) != 1:
            raise SystemExit(f"ERROR: {n} holds {len(inner)} CSVs, expected exactly 1")
        text = z.read(inner[0]).decode("utf-8", "replace")
    hdr = None
    for row in csv.reader(io.StringIO(text)):
        if not row:
            continue
        if row[0] == "I":
            hdr = row[4:]
        elif row[0] == "D":
            if hdr is None:
                raise SystemExit(f"ERROR: {n} has a data row before its header row")
            r = dict(zip(hdr, row[4:]))
            d = r["GAS_DATE"][:10].replace("/", "-")
            day[d][r["PRODUCT_LOCATION"]] = (r["VOLUME_WEIGHTED_AVERAGE_PRICE"],
                                             r["TOTAL_QUANTITY"])

cols = []
for _, stem in HUBS:
    cols += [f"{stem}_price", f"{stem}_qty_GJ"]

traded = defaultdict(int)
with open(OUT, "w", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(["date"] + cols)
    for d in sorted(day):
        line = [d]
        for loc, _ in HUBS:
            price, qty = day[d].get(loc, ("", ""))
            q = float(qty) if qty not in ("", None) else None
            # A zero-quantity row repeats the last traded price rather than reporting a
            # new one. Writing it would put a price on a day nothing changed hands, and
            # would drag every monthly average towards whatever last traded, so the
            # price is left blank and only the quantity - a real zero - is recorded.
            line += ["" if not q else round(float(price), 4),
                     "" if q is None else int(q)]
            if q:
                traded[loc] += 1
        w.writerow(line)

dates = sorted(day)
print(f"wrote {OUT} days {len(dates)} {dates[0]} -> {dates[-1]}"
      + "".join(f", {loc} traded on {traded[loc]}" for loc, _ in HUBS), flush=True)
