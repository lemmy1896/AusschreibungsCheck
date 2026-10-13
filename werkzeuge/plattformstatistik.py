"""Zählt, über welche Vergabeplattformen die Bekanntmachungen der letzten Tage laufen."""

import collections
import datetime as dt
import re
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ausschreibungscheck as ac  # noqa: E402

TAGE = int(sys.argv[1]) if len(sys.argv) > 1 else 14


def main():
    hosts = collections.Counter()
    beispiele = {}
    gesamt = 0
    for i in range(1, TAGE + 1):
        tag = dt.date.today() - dt.timedelta(days=i)
        url = ac.EXPORT_URL.format(day=tag.isoformat(), format="ocds.zip")
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": ac.USER_AGENT}), timeout=120) as a:
                releases = ac.releases_aus_zip(a.read())
        except Exception as fehler:  # noqa: BLE001
            print(tag, "Fehler:", fehler)
            continue
        print(tag, len(releases), "Bekanntmachungen")
        for r in releases:
            gesamt += 1
            b = ac.aus_release(r)
            gefunden = set()
            for link in b.dokument_links:
                host = re.sub(r"^https?://([^/:]+).*", r"\1", link).lower()
                host = re.sub(r"^www\.", "", host)
                gefunden.add(host)
                beispiele.setdefault(host, link)
            hosts.update(gefunden or {"(kein Unterlagen-Link)"})
    print(f"\n{gesamt} Bekanntmachungen in {TAGE} Tagen. Plattformen nach Häufigkeit:")
    for host, n in hosts.most_common(40):
        print(f"{n:6d}  {100 * n / max(gesamt, 1):5.1f} %  {host:45s} {beispiele.get(host, '')[:90]}")


if __name__ == "__main__":
    main()
