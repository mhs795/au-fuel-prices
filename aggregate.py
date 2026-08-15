"""Roll a daily series up to monthly, quarterly and annual averages.

Every price column becomes a simple mean of the daily values present in the period -
days with no value are skipped, not treated as zero. Count columns are handled on their
own terms: interval and schedule counts are summed, so is traded volume - a month's
gas is the sum of its days, not their average - and site counts are averaged. An n_days
column records how many days each period actually rests on, so partial periods (the
current month, the current quarter, the current year, a month with a source gap) are
visible rather than silently understated.

Quarters are calendar quarters, labelled YYYY-Qn, to match the annual tabs.

  aggregate.py <daily.csv> <monthly.csv> <quarterly.csv> <annual.csv>
"""
import csv, sys
from collections import defaultdict

SUM_SUFFIXES = ("_intervals", "_n_schedules", "_qty_GJ")
SITE_SUFFIXES = ("_sites",)


def month_key(iso):
    return iso[:7]


def quarter_key(iso):
    return f"{iso[:4]}-Q{(int(iso[5:7]) - 1) // 3 + 1}"


def year_key(iso):
    return iso[:4]


def load(path):
    with open(path, newline="") as fh:
        r = csv.reader(fh)
        return next(r), list(r)


def roll(header, rows, keyfn, label):
    cols = header[1:]
    acc = defaultdict(lambda: defaultdict(lambda: [0.0, 0]))
    days = defaultdict(set)
    for row in rows:
        key = keyfn(row[0])
        days[key].add(row[0])
        for c, v in zip(cols, row[1:]):
            if v in ("", None):
                continue
            try:
                x = float(v)
            except ValueError:
                continue
            e = acc[key][c]
            e[0] += x
            e[1] += 1
    out = []
    for key in sorted(acc):
        line = [key]
        for c in cols:
            s, n = acc[key].get(c, [0.0, 0])
            if n == 0:
                line.append("")
            elif c.endswith(SUM_SUFFIXES):
                line.append(int(round(s)))
            elif c.endswith(SITE_SUFFIXES):
                line.append(int(round(s / n)))
            else:
                line.append(round(s / n, 4))
        line.append(len(days[key]))
        out.append(line)
    return [label] + cols + ["n_days"], out


def write(path, header, rows):
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(rows)
    print(f"wrote {path} {len(rows)} rows {rows[0][0]} -> {rows[-1][0]}", flush=True)


if __name__ == "__main__":
    daily, monthly, quarterly, annual = sys.argv[1:5]
    hdr, rows = load(daily)
    write(monthly, *roll(hdr, rows, month_key, "month"))
    write(quarterly, *roll(hdr, rows, quarter_key, "quarter"))
    write(annual, *roll(hdr, rows, year_key, "year"))
