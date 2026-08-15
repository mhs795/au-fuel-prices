"""Build a daily east-coast wholesale gas price series ($/GJ) from two AEMO
master workbooks: the Victorian DWGM prices-and-demand file and the STTM
price-and-withdrawals file.
"""
import csv, sys
from datetime import datetime
from collections import defaultdict
import openpyxl

DWGM, STTM, OUT = sys.argv[1], sys.argv[2], sys.argv[3]

def as_date(v):
    """Gas_Date arrives as a datetime in some rows and a DD/MM/YYYY string in others."""
    if v is None:
        return None
    if hasattr(v, "date"):
        return v.date().isoformat()
    t = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(t[:10], fmt).date().isoformat()
        except ValueError:
            continue
    raise ValueError(f"unparsed gas date: {v!r}")

# --- Victoria DWGM: five scheduling intervals per gas day (06,10,14,18,22) ---
wb = openpyxl.load_workbook(DWGM, read_only=True, data_only=True)
vic = defaultdict(dict)
for gd, hour, price in list(wb["Prices"].iter_rows(values_only=True))[1:]:
    d = as_date(gd)
    if d is None or price is None:
        continue
    try:
        vic[d][int(hour)] = float(price)
    except (TypeError, ValueError):
        continue
wb.close()

# --- STTM hubs: one ex-ante and one ex-post price per gas day per hub ---
wb = openpyxl.load_workbook(STTM, read_only=True, data_only=True)
sttm = defaultdict(dict)
for hub, sheet in (("SYD", "SYD price and withdrawals"),
                   ("ADL", "ADL price and withdrawals"),
                   ("BRI", "BRI price and withdrawals")):
    for row in list(wb[sheet].iter_rows(values_only=True))[1:]:
        d = as_date(row[0])
        if d is None:
            continue
        for label, val in ((f"{hub}_exante", row[1]), (f"{hub}_expost", row[2])):
            try:
                sttm[d][label] = float(val)
            except (TypeError, ValueError):
                pass
wb.close()

dates = sorted(d for d in set(vic) | set(sttm) if d >= "2007-02-01")
cols = ["VIC_DWGM_6am", "VIC_DWGM_schedule_mean",
        "SYD_exante", "SYD_expost", "ADL_exante", "ADL_expost",
        "BRI_exante", "BRI_expost"]

with open(OUT, "w", newline="") as fh:
    w = csv.writer(fh)
    w.writerow(["date"] + cols + ["VIC_DWGM_n_schedules"])
    for d in dates:
        v = vic.get(d, {})
        sched = [p for p in v.values() if p is not None]
        six = v.get(6, "")
        mean = round(sum(sched) / len(sched), 4) if sched else ""
        s = sttm.get(d, {})
        w.writerow([d, six, mean] + [s.get(c, "") for c in cols[2:]] + [len(sched)])
print("wrote", OUT, "days", len(dates), dates[0], "->", dates[-1], flush=True)
