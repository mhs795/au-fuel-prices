"""Download the current AIP daily Terminal Gate Price workbook.

The filename carries the publication date and changes weekly, so the link is
scraped from the AIP historical TGP page rather than hard-coded.
"""
import re, sys, urllib.request

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
url = sorted(set(links))[-1]
print("downloading", url, flush=True)
open(OUT, "wb").write(get(url))
print("wrote", OUT, flush=True)
