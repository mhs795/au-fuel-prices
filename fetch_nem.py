"""Download AEMO 'aggregated price and demand' monthly files for every NEM region.

History runs from the start of the NEM in December 1998. Regions did not all exist for
the whole period - Tasmania joined in May 2005 and the Snowy region was abolished after
June 2008 - so LIMITS skips those month-region pairs without requesting them. Any 404
that survives that filter is reported as a missing file, not swallowed.

A cached closed month is reused only if it actually runs to the end of its month, so a
month captured while it was still live cannot freeze into the series as a part-month.
"""
import calendar, os, sys, time, datetime, itertools, urllib.request, urllib.error
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


def month_complete(p, y, m):
    """Does a cached month actually run to the end of that month?

    A month first downloaded while it was still the live month is partial. Once it is no
    longer the current or previous month it would otherwise be accepted forever on the
    strength of existing and being non-empty, freezing a part-month into the series with
    nothing downstream able to notice.
    """
    last_day = calendar.monthrange(y, m)[1]
    try:
        with open(p, newline="") as fh:
            rows = fh.read().rstrip("\n").split("\n")
        if len(rows) < 2:
            return False
        stamp = rows[-1].split(",")[1].strip().strip('"')
        seen = datetime.datetime.strptime(stamp[:10], "%Y/%m/%d").date()
    except Exception:
        return False
    # the final interval of a month is stamped 00:00 on the 1st of the next month
    return seen >= datetime.date(y, m, last_day)


def get(task):
    reg, (y, m) = task
    lo, hi = LIMITS.get(reg, (FIRST, (9999, 12)))
    if not (lo <= (y, m) <= hi):
        return None
    ym = f"{y}{m:02d}"
    p = os.path.join(OUT, f"{ym}_{reg}.csv")
    # closed months are final; only the live and just-closed months can still change
    if ((y, m) not in (cur, prev) and os.path.exists(p) and os.path.getsize(p) > 200
            and month_complete(p, y, m)):
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
                # LIMITS already skips the only two legitimate absences (TAS1 before
                # 2005-05, SNOWY1 after 2008-06) without requesting them, so a 404 that
                # reaches here is unexpected - treating it as "did not exist" would punch
                # a silent month-long hole in one region.
                return f"MISSING {ym} {reg} (404 inside the region's active window)"
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
