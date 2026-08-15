#!/usr/bin/env bash
# Rebuild the AU daily energy & fuel price workbook from source. Safe to re-run.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WORK="${FUEL_WORK:-$HERE/work}"
DEST="${FUEL_DEST:-$HOME/GoogleDrive/WORK/information library/data}"
OUT="$DEST/AU daily energy and fuel prices.xlsx"
UA='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36'

mkdir -p "$WORK/nem" "$WORK/gas" "$WORK/petrol"

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
  "$WORK/gas/sttm-price-and-withdrawals.xlsx" "$WORK/gas_daily.csv" "$WORK/gsh_daily.csv"
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
mv "$WORK/AU daily energy and fuel prices.xlsx" "$OUT"
echo "done -> $OUT"
