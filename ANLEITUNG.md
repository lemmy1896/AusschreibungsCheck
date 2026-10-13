# AusschreibungsCheck

Sucht täglich neue öffentliche Ausschreibungen, die zum Suchprofil passen, und schreibt sie in eine Excel-Liste.

## Benutzung
1. Den Ordner mit `AusschreibungsCheck.exe` und `profil_ifs.yaml` an einen festen Ort legen (z. B. Dokumente\AusschreibungsCheck).
2. `AusschreibungsCheck.exe` doppelklicken.
3. Das Programm lädt alle Tage seit dem letzten Lauf (beim ersten Mal die letzten 30 Tage), bewertet die Bekanntmachungen und öffnet danach die Excel-Liste.

Windows zeigt beim ersten Start eventuell „Der Computer wurde durch Windows geschützt“. Über „Weitere Informationen“ → „Trotzdem ausführen“ startet das Programm.

## Excel-Liste
- Grün markiert und „neu“: Treffer, die in einem früheren Lauf noch nicht dabei waren.
- „Punkte“: Summe der Gewichte aller Suchbegriffe, CPV-Codes und Regionen, die gepasst haben. „Treffergründe“ zeigt, welche das waren.
- „Link“ führt zur Bekanntmachung.
- Alle Listen liegen im Unterordner `ergebnisse`.

## Suchprofil anpassen
`profil_ifs.yaml` mit einem Texteditor (z. B. Editor/Notepad) öffnen. Dort stehen Suchbegriffe mit Gewicht, CPV-Codes, Bonusregionen, Ausschlusswörter und die Mindestpunktzahl. Änderungen gelten beim nächsten Start.

## Datenquelle
Offene Tagesexporte des Bekanntmachungsservice des Bundes (oeffentlichevergabe.de, Datenservice Öffentlicher Einkauf). Gefiltert wird auf Bekanntmachungen, die auf evergabe-online.de verweisen.

## Für Entwickler
```
pip install -r requirements-dev.txt
python -m pytest -q
python ausschreibungscheck.py --von 2026-10-01      # ab einem bestimmten Tag
python ausschreibungscheck.py --zip export.zip      # lokale Exportdatei testen
```
Bei jedem Push auf `main` baut GitHub die .exe (Actions → „Windows-Programm bauen“ → Artefakt). Ein Tag `v*` erzeugt zusätzlich ein Release mit Zip-Download.
