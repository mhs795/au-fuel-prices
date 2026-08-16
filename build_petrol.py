"""Build the two petrol series: wholesale terminal gate prices, and retail pump prices.

Writes two CSVs because they are different things on different footings - TGP is a
clean national weekday series, retail is a patchier state-by-state one.

  build_petrol.py <AIP.xlsx> <tgp_out.csv> [<fw_daily.csv> <state_daily.csv> <retail_out.csv>]
"""
import csv, sys
from collections import defaultdict
from datetime import datetime
import openpyxl

CITIES = ["Sydney", "Melbourne", "Brisbane", "Adelaide", "Perth", "Darwin", "Hobart", "National"]
PRODUCTS = ["ULP", "E10", "PULP95", "PULP98", "Diesel", "LPG"]
SCOPES = ["NSW", "QLD", "WA", "Perth"]
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
    dates = sorted(d for d in tgp if d >= "2004-01-01")   # the full span of the AIP file
    with open(out_path, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["date"] + cols)
        for d in dates:
            w.writerow([d] + [tgp[d].get(c, "") for c in cols])
    print(f"wrote {out_path} days {len(dates)} {dates[0]} -> {dates[-1]}", flush=True)


def build_retail(fw_path, state_path, out_path):
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

    with open(state_path, newline="") as fh:
        for r in csv.DictReader(fh):
            if r["product"] not in PRODUCTS:
                continue
            scope = r["state"]
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
        build_retail(sys.argv[3], sys.argv[4], sys.argv[5])
