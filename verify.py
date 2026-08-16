"""Assert the invariants this workbook's own notes claim, and fail the build if any break.

Every check here exists because something actually went wrong. The pipeline's failure mode
is not a crash - it is publishing numbers that were never measured, which looks exactly
like data until someone reads the site counts. In particular:

  * NSW retail days 1-9 of each month were silently forward-filled for 50 months, because
    a date parser dropped every unpadded day and the fill carried the previous month's
    prices over the hole. The tell was NSW_sites repeating EXACTLY across nine days -
    a live feed never repeats a site count. -> check_no_stalled_runs
  * Queensland lost six whole months to a spaced column header and an ISO timestamp, and
    the same fill papered over them. -> check_no_stalled_runs, check_site_minimum
  * Calendar 2021 electricity mixed 30-minute and 5-minute intervals without weighting
    them by duration, so the annual did not reconcile to its own quarters. -> check_annual_reconciles_to_quarters

  verify.py <work_dir>

Exits non-zero on any failure, so `set -e` in update.sh stops the publish.
"""
import csv, sys, os
from datetime import date, timedelta

WORK = sys.argv[1] if len(sys.argv) > 1 else "work"
failures = []
notes = []


def load(name):
    p = os.path.join(WORK, name)
    if not os.path.exists(p):
        return None, []
    with open(p, newline="") as fh:
        r = csv.reader(fh)
        return next(r), [row for row in r if row and row[0].strip()]


def fail(check, msg):
    failures.append(f"{check}: {msg}")


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- date hygiene
def check_dates(name):
    hdr, rows = load(name)
    if not hdr:
        return
    seen, prev = set(), None
    for row in rows:
        d = row[0][:10]
        if d in seen:
            fail("dates", f"{name}: duplicate date {d}")
        seen.add(d)
        if prev and d <= prev:
            fail("dates", f"{name}: {d} is not after {prev}")
        prev = d
    # daily tabs that claim continuous coverage must not skip calendar days
    if name in ("elec_daily.csv", "gas_daily.csv", "retail_daily.csv") and rows:
        a = date.fromisoformat(rows[0][0][:10])
        b = date.fromisoformat(rows[-1][0][:10])
        missing = (b - a).days + 1 - len(rows)
        if missing:
            fail("dates", f"{name}: {missing} calendar day(s) missing between {a} and {b}")


# --------------------------------------------- the forward-fill / stalled feed
def check_no_stalled_runs(name, groups, limit=5):
    """A run of days where EVERY column of a group repeats, site counts included.

    Prices can genuinely sit still - WA LPG held one value for 29 days and that was real,
    because its site counts moved underneath. What cannot happen naturally is the whole
    group freezing, count and all: that is a source gap being carried forward.
    """
    hdr, rows = load(name)
    if not hdr:
        return
    for g in groups:
        idx = [i for i, c in enumerate(hdr) if c.startswith(g + "_")]
        if not idx:
            continue
        run, start, prev = 0, None, None
        for row in rows:
            cur = tuple(row[i] if i < len(row) else "" for i in idx)
            if all(v == "" for v in cur):
                run, prev = 0, None
                continue
            if prev is not None and cur == prev:
                run += 1
                if run == 1:
                    start = row[0][:10]
                if run >= limit:
                    fail("stalled", f"{name}: {g} unchanged (prices AND counts) for "
                                    f"{run + 1} days ending {row[0][:10]} (from {start}) "
                                    f"- forward-fill over a source gap?")
                    break
            else:
                run, prev = 0, cur


# ------------------------------------------------------------- retail specific
def check_site_minimum(minimum=5):
    hdr, rows = load("retail_daily.csv")
    if not hdr:
        return
    for i, c in enumerate(hdr):
        if not c.endswith("_sites"):
            continue
        base = c[:-6]
        try:
            pi = hdr.index(base)
        except ValueError:
            continue
        for row in rows:
            n = num(row[i]) if i < len(row) else None
            price = row[pi] if pi < len(row) else ""
            if price.strip() and n is not None and n < minimum:
                fail("sites", f"retail: {base} written on {row[0][:10]} with only "
                              f"{int(n)} site(s) - below the documented minimum of {minimum}")
                return


# ---------------------------------------------------------------- TGP specific
def check_tgp_weekdays():
    hdr, rows = load("tgp_daily.csv")
    if not hdr:
        return
    for row in rows:
        if date.fromisoformat(row[0][:10]).weekday() >= 5:
            fail("tgp", f"weekend row {row[0][:10]} - TGP is collated on business days only")
            return


def check_national_within_range():
    hdr, rows = load("tgp_daily.csv")
    if not hdr:
        return
    for fuel in ("petrol", "diesel"):
        ni = hdr.index(f"TGP_{fuel}_national") if f"TGP_{fuel}_national" in hdr else None
        caps = [i for i, c in enumerate(hdr)
                if c.startswith(f"TGP_{fuel}_") and not c.endswith("_national")]
        if ni is None or not caps:
            continue
        for row in rows:
            nat = num(row[ni])
            vals = [num(row[i]) for i in caps if i < len(row) and num(row[i]) is not None]
            if nat is None or not vals:
                continue
            if not (min(vals) - 1e-6 <= nat <= max(vals) + 1e-6):
                fail("tgp", f"{fuel} national {nat} outside the capital range "
                            f"[{min(vals)}, {max(vals)}] on {row[0][:10]}")
                return


# -------------------------------------------------------- electricity specific
def check_price_bounds():
    hdr, rows = load("elec_daily.csv")
    if not hdr:
        return
    regions = sorted({c.split("_")[0] for c in hdr if c.endswith("_price_dwa")})
    for r in regions:
        try:
            lo, dwa, mean, hi = (hdr.index(f"{r}_price_{k}")
                                 for k in ("min", "dwa", "mean", "max"))
        except ValueError:
            continue
        for row in rows:
            a, b, c, d = (num(row[lo]), num(row[dwa]), num(row[mean]), num(row[hi]))
            if None in (a, b, c, d):
                continue
            if not (a <= b <= d and a <= c <= d):
                fail("elec", f"{r} on {row[0][:10]}: min {a} / dwa {b} / mean {c} / max {d} "
                             f"are not ordered")
                return


def quarter_hours(label, intervals):
    """Hours of trading time a quarter's intervals represent.

    Five-minute settlement began 1 Oct 2021, exactly on a quarter boundary, so every
    quarter is homogeneous: 30-minute intervals up to 2021-Q3, 5-minute from 2021-Q4.
    """
    y, q = int(str(label)[:4]), int(str(label)[-1])
    return intervals * (0.5 if (y, q) < (2021, 4) else 1.0 / 12.0)


def check_annual_reconciles_to_quarters():
    """An annual demand-weighted price must equal its quarters weighted by ENERGY.

    A bare min/max bound is too loose to catch this: the unweighted 2021 accumulation put
    VIC1's annual at 42.26 against quarters of 27.16 / 78.56 / 65.45 / 33.33, which is
    still inside that span. What it cannot survive is the actual identity - annual_dwa =
    sum(q_dwa * q_demand * q_hours) / sum(q_demand * q_hours) - because counting a
    five-minute interval as equal to a thirty-minute one gives Q4 six times its true
    weight. Tolerance is generous; the 2021 error was 24%.
    """
    ah, arows = load("elec_annual.csv")
    qh, qrows = load("elec_quarterly.csv")
    if not ah or not qh:
        return
    regions = sorted({c.split("_")[0] for c in ah if c.endswith("_price_dwa")})
    for r in regions:
        cols = [f"{r}_price_dwa", f"{r}_demand_mean_mw", f"{r}_intervals"]
        if any(c not in ah or c not in qh for c in cols):
            continue
        api, aqi = ah.index(cols[0]), qh.index(cols[0])
        qdi, qii = qh.index(cols[1]), qh.index(cols[2])
        for row in arows:
            year = str(row[0])[:4]
            ann = num(row[api])
            if ann is None:
                continue
            wsum = psum = 0.0
            for q in qrows:
                if not str(q[0]).startswith(year):
                    continue
                dwa, dem, iv = num(q[aqi]), num(q[qdi]), num(q[qii])
                if None in (dwa, dem, iv):
                    continue
                w = dem * quarter_hours(q[0], iv)
                wsum += w
                psum += dwa * w
            if wsum <= 0:
                continue
            want = psum / wsum
            if abs(want - ann) > max(0.5, abs(want) * 0.02):
                fail("elec", f"{r} {year}: annual dwa {ann} but its quarters reconcile to "
                             f"{want:.2f} - intervals not weighted by duration?")
                return


# ---------------------------------------------------------------- gas specific
def check_gsh_no_trade_rule():
    """A hub price is written only on days something traded; a no-trade day is blank
    with a real zero quantity. Both halves have to hold or the rule is not being applied."""
    hdr, rows = load("gas_daily.csv")
    if not hdr:
        return
    for hub in ("WAL", "SEQ"):
        pc, qc = f"GSH_{hub}_price", f"GSH_{hub}_qty_GJ"
        if pc not in hdr or qc not in hdr:
            continue
        pi, qi = hdr.index(pc), hdr.index(qc)
        for row in rows:
            if pi >= len(row) or qi >= len(row):
                continue
            p, q = row[pi].strip(), num(row[qi])
            if q is None:
                continue
            if not p and q > 0:
                fail("gas", f"{hub} {row[0][:10]}: no price but {q} GJ traded")
                return
            if p and q == 0:
                fail("gas", f"{hub} {row[0][:10]}: price {p} on zero traded volume")
                return


def check_no_negative_prices():
    for name, cols in (("gas_daily.csv", None), ("tgp_daily.csv", None),
                       ("retail_daily.csv", None)):
        hdr, rows = load(name)
        if not hdr:
            continue
        for i, c in enumerate(hdr):
            if i == 0 or c.endswith("_sites") or c.endswith("_qty_GJ") \
               or c.endswith("_n_schedules"):
                continue
            for row in rows:
                v = num(row[i]) if i < len(row) else None
                if v is not None and v < 0:
                    fail("range", f"{name}: negative value {v} in {c} on {row[0][:10]}")
                    return


# --------------------------------------------------------------- rollup sanity
def check_n_days():
    for daily, rolled in (("elec_daily.csv", "elec_monthly.csv"),
                          ("gas_daily.csv", "gas_monthly.csv"),
                          ("tgp_daily.csv", "tgp_monthly.csv"),
                          ("retail_daily.csv", "retail_monthly.csv")):
        dh, drows = load(daily)
        rh, rrows = load(rolled)
        if not dh or not rh or "n_days" not in rh:
            continue
        counts = {}
        for row in drows:
            counts[row[0][:7]] = counts.get(row[0][:7], 0) + 1
        ni = rh.index("n_days")
        for row in rrows:
            key = str(row[0])[:7]
            want, got = counts.get(key), num(row[ni])
            if want is not None and got is not None and int(got) != want:
                fail("rollup", f"{rolled}: {key} n_days={int(got)} but the daily tab has "
                               f"{want} rows")
                return


CHECKS = [
    ("date hygiene", lambda: [check_dates(n) for n in
                              ("elec_daily.csv", "gas_daily.csv", "tgp_daily.csv",
                               "retail_daily.csv")]),
    ("no stalled feeds (retail)", lambda: check_no_stalled_runs(
        "retail_daily.csv", ["NSW", "QLD", "WA", "Perth"])),
    ("no stalled feeds (TGP)", lambda: check_no_stalled_runs(
        "tgp_daily.csv", ["TGP"], limit=8)),
    ("retail site minimum", check_site_minimum),
    ("TGP business days only", check_tgp_weekdays),
    ("TGP national within capitals", check_national_within_range),
    ("electricity price bounds", check_price_bounds),
    ("electricity annual reconciles", check_annual_reconciles_to_quarters),
    ("gas hub no-trade rule", check_gsh_no_trade_rule),
    ("no negative prices", check_no_negative_prices),
    ("rollup n_days", check_n_days),
]

print("verifying the published invariants", flush=True)
for label, fn in CHECKS:
    before = len(failures)
    fn()
    print(f"  {'FAIL' if len(failures) > before else 'ok  '}  {label}", flush=True)

for n in notes:
    print(f"  note: {n}", flush=True)

if failures:
    print(f"\n{len(failures)} invariant(s) broken:", flush=True)
    for f in failures:
        print(f"  - {f}", flush=True)
    raise SystemExit(1)
print("all invariants hold", flush=True)
