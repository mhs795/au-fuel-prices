"""Build the two petrol series: wholesale terminal gate prices, and retail pump prices.

Writes two CSVs because they are different things on different footings - TGP is a
clean national weekday series, retail is a patchier state-by-state one.

  build_petrol.py <AIP.xlsx> <tgp_out.csv> [<fw_daily.csv> <state_daily.csv> <retail_out.csv>]
"""
import csv, os, sys
from collections import defaultdict
from datetime import datetime
import openpyxl

CITIES = ["Sydney", "Melbourne", "Brisbane", "Adelaide", "Perth", "Darwin", "Hobart", "National"]
PRODUCTS = ["ULP", "E10", "PULP95", "PULP98", "Diesel", "LPG"]
# NT and Darwin come from MyFuel NT, which stopped publishing after November 2024; the
# columns are kept because blanks read as missing everywhere in this workbook.
SCOPES = ["NSW", "QLD", "WA", "Perth", "NT", "Darwin"]
MIN_SITES = 5   # too few sites reporting to be a meaningful average
# FuelWatch's own product labels -> the common labels shared with NSW and QLD
WA_FUEL = {"ULP": "ULP", "PULP": "PULP95", "98 RON": "PULP98", "Diesel": "Diesel",
           "Brand Diesel": "Diesel_premium", "LPG": "LPG", "E10": "E10"}


def as_date(v):
    if v is None:
        return None
    if hasattr(v, "date"):
        return v.date().isoformat()
    t = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(t[:11], fmt).date().isoformat()
        except ValueError:
            continue
    return None


def check_aip_header(row, sheet):
    """The city columns are read positionally, so verify AIP has not reordered them.

    Without this a column reorder upstream would silently relabel entire cities - Perth
    prices published as Sydney - and the per-cell `except: pass` below guarantees it
    would never raise.
    """
    got = [str(v or "").strip() for v in row[1:len(CITIES) + 1]]
    for want, have in zip(CITIES, got):
        key = want.lower()
        if key == "national":
            if "national" not in have.lower():
                raise SystemExit(f"AIP sheet '{sheet}': expected a National column, got {got}")
        elif key not in have.lower():
            raise SystemExit(f"AIP sheet '{sheet}': expected {CITIES}, got {got}")


def build_tgp(aip_path, out_path):
    tgp = defaultdict(dict)
    wb = openpyxl.load_workbook(aip_path, read_only=True, data_only=True)
    for fuel, sheet in (("petrol", "Petrol TGP"), ("diesel", "Diesel TGP")):
        all_rows = list(wb[sheet].iter_rows(values_only=True))
        check_aip_header(all_rows[0], sheet)
        for row in all_rows[1:]:
            d = as_date(row[0])
            if d is None:
                continue
            for i, city in enumerate(CITIES, start=1):
                try:
                    tgp[d][f"TGP_{fuel}_{city.lower()}"] = float(row[i])
                except (TypeError, ValueError, IndexError):
                    pass
    wb.close()
    cols = [f"TGP_{f}_{c.lower()}" for f in ("petrol", "diesel") for c in CITIES]
    # Precedence: the AIP workbook wins on every date it covers, then the live feed,
    # then the archive. So a day is only ever taken from a lesser source when the
    # better one does not reach it, and AIP's own revisions always come through.
    src = os.path.dirname(aip_path) or "."
    live = fill_gaps(os.path.join(src, "tgp_current.csv"), tgp, cols)
    archive_path = os.path.join(src, "tgp_archive.csv")
    filled = fill_gaps(archive_path, tgp, cols)
    write_archive(archive_path, tgp, cols)
    dates = sorted(d for d in tgp if d >= "2004-01-01")   # the full span of the AIP file
    with open(out_path, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["date"] + cols)
        for d in dates:
            w.writerow([d] + [tgp[d].get(c, "") for c in cols])
    extra = ", ".join(f"{n} from the {label}" for n, label in
                      ((live, "live feed"), (filled, "archive")) if n)
    print(f"wrote {out_path} days {len(dates)} {dates[0]} -> {dates[-1]}"
          + (f" ({extra})" if extra else ""), flush=True)


def fill_gaps(path, tgp, cols):
    """Fill dates the workbook does not reach, from a CSV in tgp_daily.csv's own shape.

    Only ever fills - a date the workbook covers is left exactly as the workbook has it.
    """
    filled = 0
    if not os.path.exists(path):
        return filled
    with open(path, newline="") as fh:
        for row in csv.DictReader(fh):
            d = row.get("date", "")[:10]
            if not d or d in tgp:
                continue
            vals = {c: float(row[c]) for c in cols if row.get(c) not in (None, "")}
            if vals:
                tgp[d] = vals
                filled += 1
    return filled


def write_archive(path, tgp, cols):
    """Keep days AIP has published in the past but does not publish today.

    The AIP workbook is the whole history in one file, so it is normally its own
    archive - until it goes backwards. AIP's August 2026 move to WordPress republished
    a workbook ending 26 June while the series had already run to 14 August, and every
    build reads the workbook fresh, so the shorter file would have silently truncated
    the series by seven weeks on the next run. This is the same problem as the Gas
    Supply Hub cache: data with no upstream to re-fetch it from, so it is tracked in
    git rather than left on one disk.

    The workbook always wins on any date it covers - the archive only fills dates it
    does not reach - so AIP revisions still come straight from AIP.
    """
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["date"] + cols)
        for d in sorted(tgp):
            w.writerow([d] + [tgp[d].get(c, "") for c in cols])
    os.replace(tmp, path)


def build_retail(fw_path, state_paths, out_path):
    price = defaultdict(dict)     # date -> {scope_product: value}
    sites = defaultdict(lambda: defaultdict(int))

    with open(fw_path, newline="") as fh:
        for r in csv.DictReader(fh):
            prod = WA_FUEL.get(r["product"])
            if prod not in PRODUCTS:
                continue
            n = int(r["n_observations"])
            # The same five-site floor NSW and QLD apply. Without it WA published days
            # resting on a single station - Perth_E10 sat at 116.400 for seven days in
            # June 2017 off one site, while Perth_sites read ~390.
            if n < MIN_SITES:
                continue
            scope = "Perth" if r["scope"] == "Metro" else "WA"
            price[r["date"]][f"{scope}_{prod}"] = float(r["mean_price_cpl"])
            sites[r["date"]][f"{scope}_{prod}"] = n

    # Every state feed writes the same five columns, so NSW/QLD and NT read identically;
    # the "state" column carries the scope, which is why NT contributes both NT and Darwin.
    for state_path in state_paths:
        if not os.path.exists(state_path):
            continue
        with open(state_path, newline="") as fh:
            for r in csv.DictReader(fh):
                if r["product"] not in PRODUCTS:
                    continue
                scope = r["state"]
                if scope not in SCOPES:
                    continue
                price[r["date"]][f"{scope}_{r['product']}"] = float(r["mean_price_cpl"])
                sites[r["date"]][f"{scope}_{r['product']}"] = int(r["n_sites"])

    cols = [f"{s}_{p}" for s in SCOPES for p in PRODUCTS]
    # Per product, not per scope. A single count per scope was the max across that
    # scope's products, so it described the widest series and overstated every other one:
    # WA_sites read 940 on 2026-08-14 while WA_LPG rested on 39 sites.
    site_cols = [f"{s}_{p}_sites" for s in SCOPES for p in PRODUCTS]
    dates = sorted(d for d in price if d >= "2015-01-01")
    with open(out_path, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["date"] + cols + site_cols)
        for d in dates:
            w.writerow([d] + [price[d].get(c, "") for c in cols]
                       + [sites[d].get(c, "") for c in cols])
    print(f"wrote {out_path} days {len(dates)} {dates[0]} -> {dates[-1]}", flush=True)


if __name__ == "__main__":
    build_tgp(sys.argv[1], sys.argv[2])
    if len(sys.argv) > 3:
        # argv[6:] are further state-shaped CSVs (NT), appended after the retail output
        build_retail(sys.argv[3], [sys.argv[4]] + sys.argv[6:], sys.argv[5])
