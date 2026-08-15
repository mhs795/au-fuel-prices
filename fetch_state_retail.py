"""Build daily retail pump price averages for NSW and Queensland.

Both states publish price-CHANGE events, not daily snapshots, so a daily average
requires carrying each site's last posted price forward until it changes. Sites that
have not reported for STALE_DAYS are dropped so closed sites stop dragging the mean.

Monthly aggregates are cached. Because the forward-fill carries across month
boundaries, each month also caches its end-of-month price state, so a rerun only needs
the previous month's state plus the months that can still change.
"""
import calendar, csv, io, json, os, re, sys, urllib.request
from collections import defaultdict
from datetime import datetime, date, timedelta

UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
OUT = sys.argv[1] if len(sys.argv) > 1 else "state_daily.csv"
CACHE = os.path.join(os.path.dirname(os.path.abspath(OUT)), "state_cache")
STALE_DAYS = 30
FIELDS = ["date", "state", "product", "mean_price_cpl", "n_sites"]

NSW_PKG = "https://data.nsw.gov.au/data/api/3/action/package_show?id=a97a46fc-2bdd-4b90-ac7f-0cb1e8d7ac3b"
QLD_SEARCH = "https://www.data.qld.gov.au/api/3/action/package_search?q=%22Fuel+price+reporting%22&rows=30"
MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}
MONTHS.update({m.lower(): i for i, m in enumerate(calendar.month_abbr) if m})

# Source fuel labels -> the common labels used in the workbook
NSW_FUEL = {"U91": "ULP", "E10": "E10", "P95": "PULP95", "P98": "PULP98",
            "DL": "Diesel", "PDL": "Diesel_premium", "LPG": "LPG"}


def qld_fuel(name):
    n = (name or "").strip().lower()
    if "lpg" in n: return "LPG"
    if "diesel" in n:
        return "Diesel_premium" if ("premium" in n or "pdl" in n) else "Diesel"
    if "e10" in n: return "E10"
    if "98" in n: return "PULP98"
    if "95" in n: return "PULP95"
    if "unleaded" in n or n == "ulp": return "ULP"
    return None      # E85, OPAL and anything else are deliberately excluded


def fetch(url, tries=4):
    for a in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
            return urllib.request.urlopen(req, timeout=600).read()
        except Exception:
            if a == tries - 1:
                raise
            import time; time.sleep(5 * (a + 1))


def parse_month(name):
    m = re.search(r"(" + "|".join(sorted(MONTHS, key=len, reverse=True)) + r")\w*\s+(\d{4})",
                  (name or "").lower())
    if m:
        return int(m.group(2)), MONTHS[m.group(1)]
    m = re.search(r"(\d{4})[-_](\d{2})", name or "")
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(1)), int(m.group(2))
    return None


def discover():
    """-> {(state, year, month): url}"""
    out = {}
    pkg = json.loads(fetch(NSW_PKG))["result"]
    for r in pkg["resources"]:
        fmt = (r.get("format") or "").lower()
        if not (fmt.startswith("xlsx") or fmt.startswith("excel") or fmt == "csv"):
            continue
        ym = parse_month(r.get("name"))
        if ym:
            out[("NSW", ym[0], ym[1])] = r["url"]
    for p in json.loads(fetch(QLD_SEARCH))["result"]["results"]:
        for r in p.get("resources", []):
            if (r.get("format") or "").upper() != "CSV":
                continue
            ym = parse_month(r.get("name"))
            if ym:
                out[("QLD", ym[0], ym[1])] = r["url"]
    return out


def _nsw_rows(raw):
    """NSW publishes some months as xlsx and some as plain CSV; both share columns."""
    if raw[:2] == b"PK":
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(values_only=True):
            yield row
        wb.close()
    else:
        for row in csv.reader(io.StringIO(raw.decode("utf-8-sig", errors="replace"))):
            yield row


def read_nsw(raw):
    rows = _nsw_rows(raw)
    # some months put a title row above the header, so find the real header row
    hdr = None
    for _ in range(10):
        try:
            cand = [str(h or "").strip() for h in next(rows)]
        except StopIteration:
            break
        if "ServiceStationName" in cand and "Price" in cand:
            hdr = cand
            break
    if hdr is None:
        raise ValueError("no header row found in first 10 rows")
    ix = {h: i for i, h in enumerate(hdr)}
    # the column has been called both FuelCode and FuelType across the archive
    fuel_col = "FuelCode" if "FuelCode" in ix else ("FuelType" if "FuelType" in ix else None)
    need = ("ServiceStationName", "Address", "PriceUpdatedDate", "Price")
    if fuel_col is None or not all(k in ix for k in need):
        raise ValueError(f"unexpected NSW columns: {hdr}")
    for r in rows:
        try:
            price = float(r[ix["Price"]])
        except (TypeError, ValueError):
            continue
        if not (20 <= price <= 600):
            continue
        rawfuel = str(r[ix[fuel_col]] or "").strip()
        # codes (U91/DL/...) in most files, plain names in some of the older ones
        fuel = NSW_FUEL.get(rawfuel.upper()) or qld_fuel(rawfuel)
        if not fuel:
            continue
        ts = r[ix["PriceUpdatedDate"]]
        if hasattr(ts, "date"):
            d = ts.date()
        else:
            s = str(ts)[:19]
            try:
                d = datetime.strptime(s, "%Y-%m-%d %H:%M:%S").date()
            except ValueError:
                try:
                    d = datetime.strptime(s[:10], "%d/%m/%Y").date()
                except ValueError:
                    continue
        site = f"{r[ix['ServiceStationName']]}|{r[ix['Address']]}"
        yield d, site, fuel, price


def read_qld(raw):
    txt = raw.decode("utf-8-sig", errors="replace")
    for r in csv.DictReader(io.StringIO(txt)):
        fuel = qld_fuel(r.get("Fuel_Type"))
        if not fuel:
            continue
        try:
            price = float(r["Price"]) / 10.0      # published in tenths of a cent
        except (TypeError, ValueError, KeyError):
            continue
        if not (20 <= price <= 600):
            continue
        raw_ts = (r.get("TransactionDateutc") or "").strip()
        d = None
        for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
            try:
                d = (datetime.strptime(raw_ts, fmt) + timedelta(hours=10)).date()  # UTC -> AEST
                break
            except ValueError:
                continue
        if d is None:
            continue
        yield d, str(r.get("SiteId") or r.get("Site_Name")), fuel, price


def month_days(y, m):
    n = calendar.monthrange(y, m)[1]
    return [date(y, m, d) for d in range(1, n + 1)]


def process_month(state, y, m, url, carry):
    """carry: {site|fuel: [iso_date, price]} entering the month. Returns (rows, carry_out)."""
    raw = fetch(url)
    reader = read_nsw if state == "NSW" else read_qld
    events = defaultdict(list)          # day -> list of (key, fuel, price, day)
    for d, site, fuel, price in reader(raw):
        events[d].append((f"{site}\x1f{fuel}", fuel, price, d))

    cur = {k: (date.fromisoformat(v[0]), v[1]) for k, v in carry.items()}
    rows = []
    for day in month_days(y, m):
        for key, fuel, price, ed in events.get(day, []):
            cur[key] = (ed, price)
        cutoff = day - timedelta(days=STALE_DAYS)
        acc = defaultdict(lambda: [0.0, 0])
        for key, (upd, price) in cur.items():
            if upd < cutoff or upd > day:
                continue
            a = acc[key.split("\x1f")[1]]
            a[0] += price
            a[1] += 1
        for fuel, (s, c) in sorted(acc.items()):
            if c >= 5:                  # too few sites reporting to be a meaningful average
                rows.append([day.isoformat(), state, fuel, round(s / c, 3), c])
    # drop anything already stale at month end so the carried state cannot grow forever
    end = month_days(y, m)[-1]
    carry_out = {k: [u.isoformat(), p] for k, (u, p) in cur.items()
                 if u >= end - timedelta(days=STALE_DAYS)}
    return rows, carry_out


def main():
    os.makedirs(CACHE, exist_ok=True)
    found = discover()
    keys = sorted(found)
    print(f"{len(keys)} state-months discovered "
          f"(NSW {sum(1 for k in keys if k[0]=='NSW')}, QLD {sum(1 for k in keys if k[0]=='QLD')})",
          flush=True)

    today = date.today()
    live = {(today.year, today.month)}
    pm = date(today.year, today.month, 1) - timedelta(days=1)
    live.add((pm.year, pm.month))

    fetched = reused = 0
    for state in ("NSW", "QLD"):
        months = [k for k in keys if k[0] == state]
        carry = {}
        for (_, y, m) in months:
            agg_p = os.path.join(CACHE, f"{state}-{y}-{m:02d}.csv")
            st_p = os.path.join(CACHE, f"{state}-{y}-{m:02d}.state.json")
            if (y, m) not in live and os.path.exists(agg_p) and os.path.exists(st_p):
                carry = json.load(open(st_p))
                reused += 1
                continue
            try:
                rows, carry = process_month(state, y, m, found[(state, y, m)], carry)
            except Exception as e:
                print(f"  !! {state} {y}-{m:02d} failed: {type(e).__name__}: {e}", flush=True)
                continue
            with open(agg_p + ".tmp", "w", newline="") as fh:
                w = csv.writer(fh); w.writerow(FIELDS); w.writerows(rows)
            os.replace(agg_p + ".tmp", agg_p)
            json.dump(carry, open(st_p + ".tmp", "w"))
            os.replace(st_p + ".tmp", st_p)
            fetched += 1
            print(f"  {state} {y}-{m:02d}: {len(rows)} daily rows, {len(carry)} sites carried",
                  flush=True)

    print(f"{fetched} month(s) processed, {reused} reused from cache", flush=True)
    out = []
    for name in sorted(os.listdir(CACHE)):
        if name.endswith(".csv"):
            with open(os.path.join(CACHE, name), newline="") as fh:
                out.extend(list(csv.DictReader(fh)))
    out.sort(key=lambda r: (r["date"], r["state"], r["product"]))
    with open(OUT, "w", newline="") as fh:
        w = csv.DictWriter(fh, FIELDS); w.writeheader(); w.writerows(out)
    print("wrote", OUT, len(out), "rows", flush=True)


if __name__ == "__main__":
    main()
