# AU daily energy & fuel prices

Builds one Excel workbook of Australian **electricity, gas and petrol prices** at daily,
monthly, quarterly and annual frequency, entirely from published source files.

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
| Gas — VIC DWGM + SYD/ADL/BRI STTM, $/GJ | Feb 2007 | monthly workbook, extended nightly |
| Gas — Gas Supply Hub, Wallumbilla + SEQ, $/GJ | Apr 2026 | nightly, 30-day window |
| Petrol terminal gate prices — 7 capitals + national, c/L | Jan 2004 | weekdays |
| Petrol retail — NSW, QLD, WA + Perth, c/L | 2015 (WA), Aug 2016 (NSW), Feb 2019 (QLD) | monthly refresh |

## Sources

- **Electricity** — AEMO aggregated price and demand data
- **Gas** — AEMO DWGM and STTM master workbooks, extended past the workbooks' end by the
  nightly nemweb MIBB reports (`int310` VicGas, `int651`/`int657` STTM)
- **Gas Supply Hub** — nemweb `GSH_Historical_Trans_Summary`, a rolling 30-day window
- **Petrol wholesale** — Australian Institute of Petroleum terminal gate prices
- **Petrol retail** — FuelWatch WA, NSW FuelCheck (data.nsw.gov.au), QLD Fuel Price
  Reporting (data.qld.gov.au)

Full method, caveats and exact file URLs are written into the workbook's
*Sources & notes* tab by `notes.py`. A *Summary charts* tab holds native Excel charts of
every series at each frequency, built by `build_charts.py` from cells on the data tabs.

## Design notes

- **Nothing is estimated.** Every number traces to a published file. Blanks stay blank —
  no interpolation, no carry-forward, no filling.
- **Electricity monthly/quarterly/annual prices are accumulated from the raw dispatch
  intervals**, not averaged from daily figures; an average of daily averages is not a
  demand-weighted average.
- **Caching is by closed month.** Historical months never change, so they are aggregated
  once and cached under `work/`. Deleting `work/` forces a full rebuild — **except
  `work/gas/gsh_cache`, which must never be deleted.** AEMO publishes only a rolling
  30-day window of Gas Supply Hub trades and nemweb keeps roughly 95 days, so that cache
  is the only copy of every day that has already rolled off. Nothing can rebuild it.
- **Interval-duration weighting.** Electricity periods are weighted by how much *time*
  each interval represents, not by interval count — trading intervals went from 30
  minutes to 5 minutes on 1 October 2021, so counting them equally overweights the
  five-minute era 6:1. Calendar 2021 is the only period that straddles the change.
- **The builders fail loudly.** Every fetcher counts the rows it could not use and aborts
  rather than let a source format change turn into prices quietly carried forward over
  the gap; a month that parses to nothing is never cached. `build_electricity.py` aborts
  rather than write a series with unparsed rows — an earlier bug silently dropped everything before Nov 2003 because
  pre-2004 AEMO files omit seconds from the timestamp.

## Layout

| File | Role |
|---|---|
| `update.sh` | orchestrator; `PRICES` calls this |
| `PRICES` | terminal wrapper installed on `~/.local/bin`; calls `update.sh` |
| `fetch_nem.py` | AEMO NEM monthly price files |
| `fetch_aip.py` | AIP terminal gate price workbook (scrapes the weekly link) |
| `fetch_fuelwatch.py` | FuelWatch WA retail, cached per month |
| `fetch_state_retail.py` | NSW + QLD retail, price-change events → daily averages |
| `build_electricity.py` | NEM intervals → daily/monthly/quarterly/annual |
| `fetch_gas_current.py` | nightly nemweb MIBB reports, to carry gas past the workbooks |
| `fetch_gsh.py` | nemweb Gas Supply Hub trades; **caches irreplaceable history** |
| `build_gas.py` | AEMO gas workbooks → daily |
| `build_petrol.py` | AIP + retail sources → daily |
| `aggregate.py` | generic daily → monthly/quarterly/annual averaging |
| `build_charts.py` | native Excel charts on the Summary charts tab |
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
