"""Prüft den echten Tagesexport von oeffentlichevergabe.de und gibt eine Strukturübersicht aus.

Läuft auf GitHub Actions (dort ist die Seite erreichbar) und hilft, Parser und
Plattformfilter an die tatsächlichen Daten anzupassen.
"""

import collections
import datetime as dt
import io
import json
import re
import sys
import urllib.request
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ausschreibungscheck as ac  # noqa: E402


def pfade_mit(obj, suchtext, pfad=""):
    if isinstance(obj, dict):
        for k, v in obj.items():
            yield from pfade_mit(v, suchtext, f"{pfad}.{k}")
    elif isinstance(obj, list):
        for v in obj:
            yield from pfade_mit(v, suchtext, f"{pfad}[]")
    elif isinstance(obj, str) and suchtext in obj.lower():
        yield pfad


def main():
    for format_ in ("ocds.zip", "eforms.zip"):
        tag = dt.date.today() - dt.timedelta(days=1)
        for _ in range(5):
            url = ac.EXPORT_URL.replace("format=ocds.zip", f"format={format_}").format(day=tag.isoformat())
            print(f"\n=== {url}")
            try:
                with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": ac.USER_AGENT}), timeout=120) as a:
                    inhalt = a.read()
                    print("HTTP", a.status, a.headers.get("Content-Type"), len(inhalt), "Bytes")
                break
            except Exception as fehler:  # noqa: BLE001
                print("Fehler:", fehler)
                tag -= dt.timedelta(days=1)
        else:
            continue

        with zipfile.ZipFile(io.BytesIO(inhalt)) as zf:
            namen = zf.namelist()
            print(f"{len(namen)} Dateien im Zip, z. B.:", namen[:5])
            if format_ != "ocds.zip":
                continue
            releases = ac.releases_aus_zip(inhalt)
        print(f"{len(releases)} Releases gelesen")
        if not releases:
            with zipfile.ZipFile(io.BytesIO(inhalt)) as zf:
                print(zf.read(namen[0])[:1500].decode("utf-8", "replace"))
            continue

        print("Schlüssel des ersten Release:", list(releases[0].keys()))
        print("Schlüssel von tender:", list((releases[0].get("tender") or {}).keys()))
        hosts = collections.Counter()
        for r in releases:
            for u in ac.alle_urls(r):
                hosts[re.sub(r"^https?://([^/]+).*", r"\1", u).lower()] += 1
        print("Häufigste Hosts in Links:")
        for h, n in hosts.most_common(40):
            print(f"  {n:6d}  {h}")

        treffer = [r for r in releases if "evergabe-online" in json.dumps(r).lower()]
        print(f"Releases mit 'evergabe-online': {len(treffer)}")
        fundorte = collections.Counter(p for r in treffer for p in pfade_mit(r, "evergabe-online"))
        for p, n in fundorte.most_common(10):
            print(f"  {n:6d}  {p}")

        profil = ac.Profil.laden(Path(__file__).resolve().parents[1] / "profil_ifs.yaml")
        alle = [ac.aus_release(r) for r in releases]
        print("Beispiel-Bekanntmachung:", alle[0])
        bewertet = [t for b in alle if (t := ac.bewerten(b, profil))]
        print(f"Profiltreffer über alle Plattformen: {len(bewertet)}, davon evergabe-online: "
              f"{sum(ac.passt_plattform(t.bekanntmachung, profil) for t in bewertet)}")
        for t in sorted(bewertet, key=lambda t: -t.punkte)[:15]:
            b = t.bekanntmachung
            print(f"  {t.punkte:3d}  {b.titel[:90]} | {b.vergabestelle[:40]} | {b.regionen[:2]} | {b.link[:80]}")


if __name__ == "__main__":
    main()
