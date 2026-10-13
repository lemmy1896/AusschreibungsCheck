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
Offene Tagesexporte des Bekanntmachungsservice des Bundes (oeffentlichevergabe.de, Datenservice Öffentlicher Einkauf). Gefiltert wird auf Bekanntmachungen, deren Vergabeunterlagen auf evergabe-online.de liegen. Die Angebotsfrist kommt aus dem eForms-Export; Bekanntmachungen mit abgelaufener Frist werden ausgelassen.

## Für Entwickler
```
pip install -r requirements-dev.txt
python -m pytest -q
python ausschreibungscheck.py --von 2026-10-01      # ab einem bestimmten Tag
python ausschreibungscheck.py --zip export.zip      # lokale Exportdatei testen
```
Bei jedem Push auf `main` baut GitHub die .exe und veröffentlicht sie als neueste Version unter Releases (https://github.com/lemmy1896/AusschreibungsCheck/releases/latest).
