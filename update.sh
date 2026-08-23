#!/usr/bin/env bash
# Rebuild the AU daily energy & fuel price workbook from source. Safe to re-run.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORK="${FUEL_WORK:-$HERE/work}"
DEST="${FUEL_DEST:-$HOME/GoogleDrive/WORK/information library/data}"
OUT="$DEST/AU daily energy and fuel prices.xlsx"
LOCAL_OUT="$HERE/AU daily energy and fuel prices.xlsx"
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

backup_cache() {  # <path relative to $HERE> <label> <what is tracked, for the quiet line>
  # Keep caches with no upstream backed up OFF this machine. Once a day rolls off the
  # publisher's window it exists nowhere else, so leaving it on one disk is not safe.
  # git is the backup - the files are small - but only if something actually commits
  # them, so the build does it rather than relying on anyone remembering.
  local path="$1" label="$2" quiet="$3"
  git -C "$HERE" rev-parse --git-dir >/dev/null 2>&1 || return 0
  if [ -z "$(git -C "$HERE" status --porcelain -- "$path")" ]; then
    echo "   $label: nothing new, $quiet"
    return 0
  fi
  local n
  n=$(git -C "$HERE" status --porcelain -- "$path" | wc -l)
  git -C "$HERE" add -- "$path"
  git -C "$HERE" commit -q -m "$label: $n change(s) to $(date +%Y-%m-%d)" -- "$path"
  # Push is best-effort: no network should not fail a build, but an unpushed cache is
  # only as safe as this disk, so say so loudly.
  if git -C "$HERE" push -q origin HEAD 2>/dev/null; then
    echo "   $label: $n change(s) committed and pushed"
  else
    echo "   !! $label: $n change(s) committed but NOT PUSHED -" >&2
    echo "   !! this data exists only on this disk until you push" >&2
  fi
}

backup_cache work/gas/gsh_cache "gas supply hub cache" \
  "$(ls "$WORK/gas/gsh_cache" 2>/dev/null | wc -l) file(s) tracked"

echo "== 3/5 AIP terminal gate prices"
python3 "$HERE/fetch_aip.py" "$WORK/petrol/AIP_TGP.xlsx"
# The workbook above is the whole history but it is only as current as AIP's website,
# which since the August 2026 rebuild ends 26 June. api.aip.com.au is the app that
# actually produces the numbers and is still publishing every weekday, so this carries
# the series to yesterday. Same arrangement as the gas MIBB fetch above: master file for
# history and revisions, live feed for the tail.
python3 "$HERE/fetch_tgp_current.py" "$WORK/petrol/tgp_current.csv"

# FUEL_NO_RETAIL=1 skips retail entirely and omits the retail tab from the workbook.
if [ "${FUEL_NO_RETAIL:-0}" = "1" ]; then
  echo "== 4/5 retail SKIPPED (no-retail build)"
else
  echo "== 4/5 retail pump prices (WA FuelWatch, NSW FuelCheck, QLD, NT MyFuel)"
  python3 "$HERE/fetch_fuelwatch.py" "$WORK/petrol/fw_daily.csv"
  python3 "$HERE/fetch_state_retail.py" "$WORK/petrol/state_daily.csv"
  # MyFuel NT stopped publishing after November 2024, so this is a closed archive: every
  # workbook is cached and reused, and a run normally fetches nothing.
  python3 "$HERE/fetch_nt_fuel.py" "$WORK/petrol/nt_daily.csv"
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
    "$WORK/petrol/fw_daily.csv" "$WORK/petrol/state_daily.csv" "$WORK/retail_daily.csv" \
    "$WORK/petrol/nt_daily.csv"
  python3 "$HERE/aggregate.py" "$WORK/retail_daily.csv" "$WORK/retail_monthly.csv" \
    "$WORK/retail_quarterly.csv" "$WORK/retail_annual.csv"
fi
python3 "$HERE/aggregate.py" "$WORK/tgp_daily.csv" "$WORK/tgp_monthly.csv" \
  "$WORK/tgp_quarterly.csv" "$WORK/tgp_annual.csv"

# The TGP archive has the same problem as the hub cache while AIP's workbook is stuck at
# 26 June 2026: the live tables publish five business days and then drop them, so a day
# missed by every run of this build is gone. Back it up the same way.
backup_cache work/petrol/tgp_archive.csv "terminal gate price archive" \
  "$(( $(wc -l < "$WORK/petrol/tgp_archive.csv" 2>/dev/null || echo 1) - 1 )) day(s) tracked"

# Check the invariants the workbook's own notes tab claims, BEFORE building and
# publishing. A broken invariant here means a source changed shape and the series is
# carrying numbers nobody measured.
python3 "$HERE/verify.py" "$WORK"

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
# Two copies on purpose: the Drive one to share and read, a local one that does not
# depend on the mount being up. The local copy IS the build output, so it is moved
# rather than re-copied.
mv "$BUILT" "$LOCAL_OUT"
echo "done -> $OUT"
echo "     -> $LOCAL_OUT"
