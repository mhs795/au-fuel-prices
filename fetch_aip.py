"""Download the current AIP daily Terminal Gate Price workbook.

The filename carries the publication date and changes weekly, so the link is
scraped from AIP's site rather than hard-coded.

AIP rebuilt the site on WordPress in August 2026. The old page
www.aip.com.au/historical-ulp-and-diesel-tgp-data and every
/sites/default/files/...AIP_TGP_Data_<date>.xlsx URL under it now return 404 -
that is what broke this fetch. The workbook now lives under /wp-content/uploads/.
Two ways in, because the obvious one is broken at the source: the Terminal Gate
Prices page links the workbook via /resource-api/..., which itself 404s, so scrape
the resources page that actually carries the link and fall back to the WordPress
media API if the page markup changes again.
"""
import datetime, json, os, re, sys, urllib.request

PAGES = [
    "https://aip.com.au/resources/historical-ulp-and-diesel-tgp-data/",
    "https://aip.com.au/pricing/terminal-gate-prices/",
]
MEDIA_API = ("https://aip.com.au/wp-json/wp/v2/media"
             "?search=AIP_TGP_Data&per_page=100&_fields=source_url,date")
# Matches both the current /wp-content/uploads/ layout and the retired
# /sites/default/files/ one, with or without the www host.
LINK = re.compile(r'https?://(?:www\.)?aip\.com\.au/[^"\'\s<>]*?'
                  r'AIP_TGP_Data_\d{1,2}-[A-Za-z]{3}-\d{4}\.xlsx')
UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
OUT = sys.argv[1]
# AIP collates on weekdays, so a workbook more than a fortnight old means they have
# stopped publishing, not that the week is quiet. Say so rather than silently
# freezing the series.
STALE_DAYS = 14


def get(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    return urllib.request.urlopen(req, timeout=180).read()


def published(url):
    """Sort key from the D-Mon-YYYY stamp in the filename.

    The day is not zero-padded, so a plain string sort puts 9-Sep above 14-Sep.
    Anything unparseable sorts to the bottom rather than winning by accident.
    """
    m = re.search(r'AIP_TGP_Data_(\d{1,2}-[A-Za-z]{3}-\d{4})\.xlsx', url)
    if not m:
        return (datetime.date.min, url)
    try:
        return (datetime.datetime.strptime(m.group(1), "%d-%b-%Y").date(), url)
    except ValueError:
        return (datetime.date.min, url)


def last_date(path):
    """Last date carried by the workbook's petrol sheet, or None if unreadable."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            dates = [r[0] for r in wb["Petrol TGP"].iter_rows(values_only=True)
                     if r and hasattr(r[0], "date")]
        finally:
            wb.close()
        return max(d.date() for d in dates) if dates else None
    except Exception:
        return None


links, tried = set(), []
for page in PAGES:
    try:
        links |= set(LINK.findall(get(page).decode("utf-8", errors="replace")))
    except Exception as e:
        tried.append(f"{page}: {e}")
    if links:
        break

if not links:
    try:
        for item in json.loads(get(MEDIA_API).decode("utf-8", errors="replace")):
            url = item.get("source_url", "")
            if LINK.fullmatch(url):
                links.add(url)
    except Exception as e:
        tried.append(f"{MEDIA_API}: {e}")

if not links:
    raise SystemExit("no AIP_TGP_Data link found. Tried:\n  " + "\n  ".join(
        tried or PAGES + [MEDIA_API]))

url = max(links, key=published)
stamp = published(url)[0]
print("downloading", url, flush=True)
data = get(url)

# Publish only if this is not a step backwards. AIP's WordPress migration left a
# workbook seven weeks older than the one already on disk; overwriting would have
# silently truncated the series by that much, and the build would have gone on to
# publish it. The workbook is the full history every time, so the newer file is
# strictly better and the older one has nothing to add.
# The suffix has to stay .xlsx: openpyxl refuses to open a file by any other
# extension, and a temp name it cannot read would make the check below blind.
tmp = OUT + ".tmp.xlsx"
with open(tmp, "wb") as fh:
    fh.write(data)
new, old = last_date(tmp), (last_date(OUT) if os.path.exists(OUT) else None)
if old and (new is None or new < old):
    # Unreadable counts as a step backwards, not as a pass. Treating an unknown
    # coverage as acceptable is what let a stale workbook overwrite a newer one.
    os.remove(tmp)
    coverage = new or "nothing readable"
    print(f"!! AIP is publishing {coverage}, not past the {old} workbook already in "
          f"{OUT} - keeping the existing file", file=sys.stderr)
    print(f"!! the TGP series stops at {old} until AIP publishes a newer workbook",
          file=sys.stderr)
elif new is None and old is None:
    raise SystemExit(f"downloaded {url} but cannot read a date out of it")
else:
    os.replace(tmp, OUT)
    print("wrote", OUT, "covering to", new, flush=True)

age = (datetime.date.today() - stamp).days if stamp != datetime.date.min else None
if age is not None and age > STALE_DAYS:
    print(f"!! newest AIP workbook is {stamp} ({age} days old) - check whether AIP "
          f"has stopped publishing at {PAGES[0]}", file=sys.stderr)
