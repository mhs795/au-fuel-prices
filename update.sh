#!/usr/bin/env bash
# Rebuild the AU daily energy & fuel price workbook from source. Safe to re-run.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORK="${FUEL_WORK:-$HERE/work}"
DEST="${FUEL_DEST:-$HOME/GoogleDrive/WORK/information library/data}"
OUT="$DEST/AU daily energy and fuel prices.xlsx"
UA='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36'

mkdir -p "$WORK/nem" "$WORK/gas" "$WORK/petrol"

# One build at a time. Two runs share $WORK and the per-month caches, and they collide on
# the same <month>.state.json.tmp - one renames it, the other dies on a missing file, and
# the cached retail chain is left half-rebuilt.
exec 9>"$WORK/.lock"
if ! flock -n 9; then
  echo "another build is already running (lock: $WORK/.lock) - not starting a second" >&2
  exit 1
fi

# Fail before the downloads, not after them: this is 5-10 minutes of fetching and the
# destination is usually an rclone mount that may simply not be up.
if [ ! -d "$DEST" ]; then
  echo "destination directory missing: $DEST" >&2
  echo "is the Google Drive mount up?  mount | grep GoogleDrive" >&2
  exit 1
fi
if ! touch "$DEST/.write-test" 2>/dev/null; then
  echo "destination is not writable: $DEST" >&2
  exit 1
fi
rm -f "$DEST/.write-test"

echo "== 1/5 AEMO NEM monthly price files (incremental)"
python3 "$HERE/fetch_nem.py" "$WORK/nem"

echo "== 2/5 AEMO gas master workbooks and Gas Supply Hub"
curl -fsSL -A "$UA" -o "$WORK/gas/dwgm-prices-and-demand.xlsx" \
  "https://www.aemo.com.au/-/media/files/gas/dwgm/dwgm-prices-and-demand.xlsx?rev=948bdc7238b54023afede846391eb2d4&sc_lang=en"
curl -fsSL -A "$UA" -o "$WORK/gas/sttm-price-and-withdrawals.xlsx" \
  "https://www.aemo.com.au/-/media/files/gas/sttm/data/sttm-price-and-withdrawals.xlsx?rev=f16b91ed263c4e06953189dde1a7e758&sc_lang=en"
# The hub only publishes a rolling 30-day window and nemweb keeps ~95 daily files, so the
# cache under work/gas/gsh_cache is the series: it grows a day per run and holds the only
# copy of everything that has already rolled off. Never delete it.
python3 "$HERE/fetch_gsh.py" "$WORK/gas/gsh_cache" "$WORK/gsh_daily.csv"
# The two master workbooks above are monthly, so on their own the gas series ends at the
# close of the last completed month. The same clearing prices are on nemweb nightly; this
# only ever extends the series past the workbooks' end, so revisions still come from them.
python3 "$HERE/fetch_gas_current.py" "$WORK/gas/mibb_cache" "$WORK/gas_current.csv"

echo "== 3/5 AIP terminal gate prices"
python3 "$HERE/fetch_aip.py" "$WORK/petrol/AIP_TGP.xlsx"

# FUEL_NO_RETAIL=1 skips retail entirely and omits the retail tab from the workbook.
if [ "${FUEL_NO_RETAIL:-0}" = "1" ]; then
  echo "== 4/5 retail SKIPPED (no-retail build)"
else
  echo "== 4/5 retail pump prices (WA FuelWatch, NSW FuelCheck, QLD)"
  python3 "$HERE/fetch_fuelwatch.py" "$WORK/petrol/fw_daily.csv"
  python3 "$HERE/fetch_state_retail.py" "$WORK/petrol/state_daily.csv"
fi

echo "== 5/5 building series and workbook"
python3 "$HERE/build_electricity.py" "$WORK/nem" \
  "$WORK/elec_daily.csv" "$WORK/elec_monthly.csv" "$WORK/elec_quarterly.csv" \
  "$WORK/elec_annual.csv"
python3 "$HERE/build_gas.py" "$WORK/gas/dwgm-prices-and-demand.xlsx" \
  "$WORK/gas/sttm-price-and-withdrawals.xlsx" "$WORK/gas_daily.csv" "$WORK/gsh_daily.csv" \
  "$WORK/gas_current.csv"
python3 "$HERE/aggregate.py" "$WORK/gas_daily.csv" "$WORK/gas_monthly.csv" \
  "$WORK/gas_quarterly.csv" "$WORK/gas_annual.csv"

if [ "${FUEL_NO_RETAIL:-0}" = "1" ]; then
  python3 "$HERE/build_petrol.py" "$WORK/petrol/AIP_TGP.xlsx" "$WORK/tgp_daily.csv"
else
  python3 "$HERE/build_petrol.py" "$WORK/petrol/AIP_TGP.xlsx" "$WORK/tgp_daily.csv" \
    "$WORK/petrol/fw_daily.csv" "$WORK/petrol/state_daily.csv" "$WORK/retail_daily.csv"
  python3 "$HERE/aggregate.py" "$WORK/retail_daily.csv" "$WORK/retail_monthly.csv" \
    "$WORK/retail_quarterly.csv" "$WORK/retail_annual.csv"
fi
python3 "$HERE/aggregate.py" "$WORK/tgp_daily.csv" "$WORK/tgp_monthly.csv" \
  "$WORK/tgp_quarterly.csv" "$WORK/tgp_annual.csv"

if [ -f "$OUT" ]; then cp -p "$OUT" "$WORK/previous_workbook.xlsx"; fi
if [ "${FUEL_NO_RETAIL:-0}" = "1" ]; then
  python3 "$HERE/build_workbook.py" "$WORK" "$WORK/AU daily energy and fuel prices.xlsx" --no-retail
else
  python3 "$HERE/build_workbook.py" "$WORK" "$WORK/AU daily energy and fuel prices.xlsx"
fi
# Publish atomically. $WORK and $DEST are different filesystems (the destination is an
# rclone mount), so a plain mv degrades to a copy written straight to the final path - an
# interruption mid-copy leaves a truncated workbook where the published one used to be.
# Copy to a temp file ON the destination first, verify it, then rename within that
# filesystem, which is atomic.
BUILT="$WORK/AU daily energy and fuel prices.xlsx"
TMP="$DEST/.AU daily energy and fuel prices.xlsx.tmp-$$"
trap 'rm -f "$TMP"' EXIT
cp "$BUILT" "$TMP"
sync "$TMP" 2>/dev/null || true
built_size=$(stat -c%s "$BUILT")
copied_size=$(stat -c%s "$TMP")
if [ "$built_size" != "$copied_size" ]; then
  echo "publish aborted: copied $copied_size bytes, expected $built_size - $OUT left unchanged" >&2
  exit 1
fi
mv "$TMP" "$OUT"
trap - EXIT
rm -f "$BUILT"
echo "done -> $OUT"
