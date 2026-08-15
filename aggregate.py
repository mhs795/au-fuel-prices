"""Roll a daily series up to monthly and annual averages.

Every price column becomes a simple mean of the daily values present in the period -
days with no value are skipped, not treated as zero. Count columns are handled on their
own terms: interval and schedule counts are summed, site counts are averaged. An n_days
column records how many days each period actually rests on, so partial periods (the
current month, the current year, a month with a source gap) are visible rather than
silently understated.

  aggregate.py <daily.csv> <monthly.csv> <annual.csv>
"""
import csv, sys
from collections import defaultdict

SUM_SUFFIXES = ("_intervals", "_n_schedules")
SITE_SUFFIXES = ("_sites",)


def load(path):
    with open(path, newline="") as fh:
        r = csv.reader(fh)
        return next(r), list(r)


def roll(header, rows, width, label):
    cols = header[1:]
    acc = defaultdict(lambda: defaultdict(lambda: [0.0, 0]))
    days = defaultdict(set)
    for row in rows:
        key = row[0][:width]
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
    daily, monthly, annual = sys.argv[1:4]
    hdr, rows = load(daily)
    write(monthly, *roll(hdr, rows, 7, "month"))
    write(annual, *roll(hdr, rows, 4, "year"))
