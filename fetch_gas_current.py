"""Bring the DWGM and STTM price series up to the last completed gas day.

  fetch_gas_current.py <cache_dir> <out.csv>

The two AEMO master workbooks the gas series is built from are republished monthly, so
they end at the close of the last completed month and the gas tab used to run up to six
weeks behind everything else. The same clearing prices are on nemweb every night, in the
MIBB reports:

  VicGas int310  - Victorian schedule prices, five intervals per gas day, rolling ~1 year
  STTM   int651  - ex-ante market price per hub, in Day01-Day31 + CurrentDay bundles
  STTM   int657  - ex-post imbalance price per hub, same bundles

These are not a different or lesser dataset. Checked against the master workbooks over
every overlapping day available on 16 Aug 2026, they are identical to the last published
decimal: 365/365 days exact for the 6am DWGM price and for the mean of the five schedules,
57/57 hub-days exact for STTM ex-ante and 60/60 for ex-post, max absolute difference
0.000000 throughout. build_gas.py still lets the master workbook win wherever both cover a
day, so these values only ever extend the series past the workbook's end - which keeps that
equality true by construction and leaves any later AEMO revision free to take effect.

The STTM bundles are named by day of month and recycled every month, so the cache is keyed
by the publication timestamp inside each member filename instead. Unlike the Gas Supply Hub
cache this one is not load-bearing: the master workbook catches up each month, so a lost
cache costs nothing that the next monthly refresh does not restore.

The ex-ante price is set a day ahead, so int651 always carries tomorrow's gas date. The
series stops at the last completed gas day rather than running into the future.
"""
import csv, datetime, io, os, re, sys, time, zipfile, urllib.request
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor

CACHE, OUT = sys.argv[1], sys.argv[2]
UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
VICGAS = "https://nemweb.com.au/REPORTS/CURRENT/VicGas/int310_v4_price_and_withdrawals_1.csv"
STTM = "https://nemweb.com.au/REPORTS/CURRENT/STTM/"
WANT = ("int651", "int657")          # ex-ante price, ex-post imbalance price
HUBS = ("SYD", "ADL", "BRI")
TODAY = datetime.date.today()


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


def gas_date(s):
    """MIBB stamps dates as '16 Aug 2026'."""
    return datetime.datetime.strptime(s.strip(), "%d %b %Y").date()


os.makedirs(CACHE, exist_ok=True)

# --- Victoria: one file, a rolling year, no cache needed ---------------------
vic_raw = fetch(VICGAS).decode("utf-8", "replace")
vic = defaultdict(dict)
for r in csv.DictReader(io.StringIO(vic_raw)):
    if not r.get("price_value"):
        continue
    vic[gas_date(r["gas_date"])][int(r["schedule_interval"])] = float(r["price_value"])

# --- STTM: 31 day-of-month bundles plus today's ------------------------------
def grab(name):
    got = 0
    try:
        body = fetch(STTM + name, tries=2)
    except SystemExit:
        return 0                      # a day-of-month bundle that does not exist yet
    with zipfile.ZipFile(io.BytesIO(body)) as z:
        for m in z.namelist():
            if not any(m.startswith(w) for w in WANT):
                continue
            p = os.path.join(CACHE, os.path.basename(m))
            if os.path.exists(p):
                continue              # keyed by publication timestamp, so never restated
            with open(p + ".tmp", "wb") as fh:
                fh.write(z.read(m))
            os.replace(p + ".tmp", p)
            got += 1
    return got

names = ["CurrentDay.zip"] + [f"Day{d:02d}.zip" for d in range(1, 32)]
with ThreadPoolExecutor(6) as ex:
    new = sum(ex.map(grab, names))
cached = sorted(f for f in os.listdir(CACHE) if f.endswith(".csv"))
print(f"  VicGas int310 {len(vic)} gas days; STTM {new} new report(s), "
      f"{len(cached)} cached", flush=True)

exante, expost = defaultdict(dict), defaultdict(dict)
for f in cached:
    with open(os.path.join(CACHE, f), newline="") as fh:
        for r in csv.DictReader(fh):
            hub = r.get("hub_identifier")
            if hub not in HUBS or not r.get("gas_date"):
                continue
            if f.startswith("int651") and r.get("ex_ante_market_price"):
                exante[gas_date(r["gas_date"])][hub] = float(r["ex_ante_market_price"])
            elif f.startswith("int657") and r.get("ex_post_imbalance_price"):
                expost[gas_date(r["gas_date"])][hub] = float(r["ex_post_imbalance_price"])

cols = ["VIC_DWGM_6am", "VIC_DWGM_schedule_mean", "SYD_exante", "SYD_expost",
        "ADL_exante", "ADL_expost", "BRI_exante", "BRI_expost", "VIC_DWGM_n_schedules"]
# Tomorrow's ex-ante price is real and already set, but this is a record of gas days that
# have happened, so the series stops at the last completed one.
dates = sorted(d for d in set(vic) | set(exante) | set(expost) if d < TODAY)
if not dates:
    raise SystemExit("ERROR: no MIBB gas days before today; the report paths have "
                     "probably moved.")

with open(OUT, "w", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(["date"] + cols)
    for d in dates:
        v, xa, xp = vic.get(d, {}), exante.get(d, {}), expost.get(d, {})
        sched = sorted(v.values()) and list(v.values())
        row = [d.isoformat(),
               v.get(1, ""),
               round(sum(v.values()) / len(v), 4) if v else ""]
        for hub in HUBS:
            row += [xa.get(hub, ""), xp.get(hub, "")]
        row.append(len(v))
        w.writerow(row)
print(f"wrote {OUT} days {len(dates)} {dates[0]} -> {dates[-1]}", flush=True)
