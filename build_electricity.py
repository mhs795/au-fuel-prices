"""Aggregate AEMO NEM 'aggregated price and demand' files to daily, monthly, quarterly
and annual regional series.

Each interval's SETTLEMENTDATE is the interval-ENDING timestamp, so one second is
subtracted before taking the calendar date. That puts the 00:00 interval (which ends at
midnight) on the day it actually belongs to, and is agnostic to the 30-minute /
5-minute settlement change of 1 October 2021.

Monthly, quarterly and annual demand-weighted prices are accumulated from the raw
intervals, not averaged from the daily figures - an average of daily averages is not a
demand-weighted average and would quietly misprice volatile months.

  build_electricity.py <src_dir> <daily.csv> <monthly.csv> <quarterly.csv> <annual.csv>
"""
import csv, glob, os, sys
from collections import defaultdict
from datetime import datetime, timedelta

# Older AEMO files (pre-2004) stamp the time without seconds.
TS_FORMATS = ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M")


def parse_ts(v):
    if not v:
        return None
    t = str(v).strip().strip('"')
    for fmt in TS_FORMATS:
        try:
            return datetime.strptime(t, fmt)
        except ValueError:
            continue
    return None


SRC, DAILY, MONTHLY, QUARTERLY, ANNUAL = sys.argv[1:6]


def quarter_key(iso):
    return f"{iso[:4]}-Q{(int(iso[5:7]) - 1) // 3 + 1}"


def new():
    return {"pq": 0.0, "q": 0.0, "p": 0.0, "n": 0,
            "min": float("inf"), "max": float("-inf"), "dsum": 0.0, "days": set()}


acc = {"daily": defaultdict(new), "monthly": defaultdict(new),
       "quarterly": defaultdict(new), "annual": defaultdict(new)}
unparsed = 0
files = sorted(glob.glob(os.path.join(SRC, "*.csv")))
print(f"{len(files)} files", flush=True)
for fp in files:
    with open(fp, newline="") as fh:
        for row in csv.DictReader(fh):
            ts = parse_ts(row.get("SETTLEMENTDATE"))
            if ts is None:
                unparsed += 1
                continue
            try:
                rrp = float(row["RRP"])
                dem = float(row["TOTALDEMAND"])
            except (ValueError, KeyError, TypeError):
                unparsed += 1
                continue
            d = (ts - timedelta(seconds=1)).date()
            reg = row["REGION"]
            iso = d.isoformat()
            for freq, key in (("daily", iso), ("monthly", iso[:7]),
                              ("quarterly", quarter_key(iso)), ("annual", iso[:4])):
                a = acc[freq][(reg, key)]
                a["pq"] += rrp * dem
                a["q"] += dem
                a["p"] += rrp
                a["n"] += 1
                a["dsum"] += dem
                a["days"].add(iso)
                if rrp < a["min"]: a["min"] = rrp
                if rrp > a["max"]: a["max"] = rrp

if unparsed:
    raise SystemExit(f"ERROR: {unparsed} rows could not be parsed - refusing to write a "
                     f"silently incomplete series. Check SETTLEMENTDATE/RRP formats.")

regions = sorted({k[0] for k in acc["daily"]})
print("regions", regions, flush=True)


def write(path, freq, label):
    keys = sorted({k[1] for k in acc[freq]})
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh)
        hdr = [label]
        for r in regions:
            hdr += [f"{r}_price_dwa", f"{r}_price_mean", f"{r}_price_min",
                    f"{r}_price_max", f"{r}_demand_mean_mw", f"{r}_intervals"]
        if freq != "daily":
            hdr.append("n_days")
        w.writerow(hdr)
        for k in keys:
            row = [k]
            days = 0
            for r in regions:
                a = acc[freq].get((r, k))
                if not a or a["n"] == 0:
                    row += ["", "", "", "", "", ""]
                    continue
                days = max(days, len(a["days"]))
                row += [round(a["pq"] / a["q"], 4) if a["q"] else "",
                        round(a["p"] / a["n"], 4),
                        round(a["min"], 4), round(a["max"], 4),
                        round(a["dsum"] / a["n"], 2), a["n"]]
            if freq != "daily":
                row.append(days)
            w.writerow(row)
    print(f"wrote {path} {len(keys)} rows {keys[0]} -> {keys[-1]}", flush=True)


write(DAILY, "daily", "date")
write(MONTHLY, "monthly", "month")
write(QUARTERLY, "quarterly", "quarter")
write(ANNUAL, "annual", "year")
