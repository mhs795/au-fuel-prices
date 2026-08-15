"""Download the current AIP daily Terminal Gate Price workbook.

The filename carries the publication date and changes weekly, so the link is
scraped from the AIP historical TGP page rather than hard-coded.
"""
import datetime, re, sys, urllib.request

PAGE = "https://www.aip.com.au/historical-ulp-and-diesel-tgp-data"
UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36"
OUT = sys.argv[1]


def get(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    return urllib.request.urlopen(req, timeout=180).read()


html = get(PAGE).decode("utf-8", errors="replace")
links = re.findall(r'https://www\.aip\.com\.au/sites/default/files/[^"\']*?AIP_TGP_Data_[^"\']*?\.xlsx', html)
if not links:
    raise SystemExit("no AIP_TGP_Data link found on " + PAGE)


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


url = max(set(links), key=published)
print("downloading", url, flush=True)
open(OUT, "wb").write(get(url))
print("wrote", OUT, flush=True)
