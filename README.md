# AU daily energy & fuel prices

Builds one Excel workbook of Australian **electricity, gas and petrol prices** at daily,
monthly and annual frequency, entirely from published source files.

Output: `~/GoogleDrive/WORK/information library/data/AU daily energy and fuel prices.xlsx`

## Usage

```bash
PRICES              # rebuild everything, including retail pump prices
PRICES NO RETAIL    # skip retail; the workbook is built without those tabs
```

`PRICES` is a wrapper on `~/.local/bin/` that calls `update.sh`. A rerun takes about
5–10 minutes: closed months are cached, so only the current and previous month are
re-fetched. The first build downloads several GB.

## Coverage

| Series | From | Frequency of source |
|---|---|---|
| Electricity — 6 NEM regions, $/MWh | Dec 1998 | continuous |
| Gas — VIC DWGM + SYD/ADL/BRI STTM, $/GJ | Feb 2007 | monthly refresh |
| Petrol terminal gate prices — 7 capitals + national, c/L | Jan 2004 | weekdays |
| Petrol retail — NSW, QLD, WA + Perth, c/L | 2015 (WA), Aug 2016 (NSW), Feb 2019 (QLD) | monthly refresh |

## Sources

- **Electricity** — AEMO aggregated price and demand data
- **Gas** — AEMO DWGM and STTM master workbooks
- **Petrol wholesale** — Australian Institute of Petroleum terminal gate prices
- **Petrol retail** — FuelWatch WA, NSW FuelCheck (data.nsw.gov.au), QLD Fuel Price
  Reporting (data.qld.gov.au)

Full method, caveats and exact file URLs are written into the workbook's
*Sources & notes* tab by `notes.py`.

## Design notes

- **Nothing is estimated.** Every number traces to a published file. Blanks stay blank —
  no interpolation, no carry-forward, no filling.
- **Electricity monthly/annual prices are accumulated from the raw dispatch intervals**,
  not averaged from daily figures; an average of daily averages is not a demand-weighted
  average.
- **Caching is by closed month.** Historical months never change, so they are aggregated
  once and cached under `work/`. Deleting `work/` forces a full rebuild.
- **The builders fail loudly.** `build_electricity.py` aborts rather than write a series
  with unparsed rows — an earlier bug silently dropped everything before Nov 2003 because
  pre-2004 AEMO files omit seconds from the timestamp.

## Layout

| File | Role |
|---|---|
| `update.sh` | orchestrator; `PRICES` calls this |
| `fetch_nem.py` | AEMO NEM monthly price files |
| `fetch_aip.py` | AIP terminal gate price workbook (scrapes the weekly link) |
| `fetch_fuelwatch.py` | FuelWatch WA retail, cached per month |
| `fetch_state_retail.py` | NSW + QLD retail, price-change events → daily averages |
| `build_electricity.py` | NEM intervals → daily/monthly/annual |
| `build_gas.py` | AEMO gas workbooks → daily |
| `build_petrol.py` | AIP + retail sources → daily |
| `aggregate.py` | generic daily → monthly/annual averaging |
| `notes.py` | content of the Sources & notes tab |
| `build_workbook.py` | assembles and formats the workbook |

## Install

```bash
ln -sf "$PWD/PRICES" ~/.local/bin/PRICES
ln -sf "$PWD/PRICES" ~/.local/bin/prices
```

`PRICES` expects the scripts at `~/Documents/Work/au_fuel_prices/`; edit `SCRIPT` at the
top of the wrapper if they live elsewhere. Set `FUEL_DEST` to change where the workbook
is written.
