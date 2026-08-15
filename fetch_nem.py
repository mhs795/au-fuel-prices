"""Download AEMO 'aggregated price and demand' monthly files for every NEM region.

History runs from the start of the NEM in December 1998. Regions did not all exist for
the whole period - Tasmania joined in May 2005 and the Snowy region was abolished after
June 2008 - so a 404 is recorded as "this region did not exist that month" rather than
treated as an error.
"""
import os, sys, time, datetime, itertools, urllib.request, urllib.error
from concurrent.futures import ThreadPoolExecutor

OUT = sys.argv[1]
UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
REGIONS = ["NSW1", "QLD1", "SA1", "SNOWY1", "TAS1", "VIC1"]
FIRST = (1998, 12)
# months outside these bounds are known not to exist and are not requested at all
LIMITS = {"TAS1": ((2005, 5), (9999, 12)), "SNOWY1": ((1998, 12), (2008, 6))}

today = datetime.date.today()
months = [(y, m) for y in range(FIRST[0], today.year + 1) for m in range(1, 13)
          if FIRST <= (y, m) <= (today.year, today.month)]
cur = (today.year, today.month)
prev_d = datetime.date(today.year, today.month, 1) - datetime.timedelta(days=1)
prev = (prev_d.year, prev_d.month)


def get(task):
    reg, (y, m) = task
    lo, hi = LIMITS.get(reg, (FIRST, (9999, 12)))
    if not (lo <= (y, m) <= hi):
        return None
    ym = f"{y}{m:02d}"
    p = os.path.join(OUT, f"{ym}_{reg}.csv")
    # closed months are final; only the live and just-closed months can still change
    if (y, m) not in (cur, prev) and os.path.exists(p) and os.path.getsize(p) > 200:
        return None
    url = f"https://www.aemo.com.au/aemo/data/nem/priceanddemand/PRICE_AND_DEMAND_{ym}_{reg}.csv"
    err = ""
    for a in range(4):
        if a:
            time.sleep(3 * a)
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            body = urllib.request.urlopen(req, timeout=90).read()
            if body.startswith(b"REGION"):
                with open(p + ".tmp", "wb") as fh:
                    fh.write(body)
                os.replace(p + ".tmp", p)
                return None
            return f"BADCONTENT {ym} {reg}"
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None          # region genuinely absent that month
            err = f"HTTP {e.code}"
        except Exception as e:
            err = f"{type(e).__name__} {e}"
    return f"FAIL {ym} {reg} {err}"


os.makedirs(OUT, exist_ok=True)
tasks = list(itertools.product(REGIONS, months))
with ThreadPoolExecutor(6) as ex:
    bad = [r for r in ex.map(get, tasks) if r]
have = len([f for f in os.listdir(OUT) if f.endswith(".csv")])
print(f"{len(tasks)} candidates, {have} files on disk, {len(bad)} failed", flush=True)
for r in bad[:40]:
    print(" ", r, flush=True)
if bad:
    raise SystemExit(1)
