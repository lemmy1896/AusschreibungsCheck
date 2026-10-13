# AusschreibungsCheck

Sucht täglich neue öffentliche Ausschreibungen, die zum Suchprofil passen, und schreibt sie in eine Excel-Liste.

## Benutzung
1. Den Ordner mit `AusschreibungsCheck.exe` und `profil_ifs.yaml` an einen festen Ort legen (z. B. Dokumente\AusschreibungsCheck).
2. `AusschreibungsCheck.exe` doppelklicken. Es öffnet sich das Programmfenster.
3. Plattformen anhaken, bei Bedarf zusätzliche Schlagwörter eintragen und „Suche starten“ klicken.
4. Das Programm lädt alle Tage seit dem letzten Lauf (beim ersten Mal die letzten 30 Tage), bewertet die Bekanntmachungen und öffnet danach die Excel-Liste.

Windows zeigt beim ersten Start eventuell „Der Computer wurde durch Windows geschützt“. Über „Weitere Informationen“ → „Trotzdem ausführen“ startet das Programm.

## Programmfenster
- **Plattformen:** e-Vergabe des Bundes (evergabe-online.de), DTVP, evergabe.de, subreport ELViS und RIB meinauftrag. „Alle Plattformen“ nimmt zusätzlich alle übrigen Bekanntmachungen des Bundes-Exports mit.
- **Zeitraum:** „Seit dem letzten Lauf“ prüft nur neue Tage. „Ab Datum“ prüft alle Ausschreibungen ab dem eingetragenen Tag (TT.MM.JJJJ), z. B. nach geänderten Schlagwörtern. Bereits geladene Tage kommen aus dem Zwischenspeicher und gehen schnell.
- **Schlagwörter:** die komplette Liste des Profils mit Punkten. Zum Hinzufügen Schlagwort und Punkte eintragen und „Hinzufügen / Punkte ändern“ klicken. Zum Ändern einen Eintrag anklicken, Punkte anpassen und denselben Knopf klicken. Zum Löschen Eintrag anklicken und „Entfernen“ (oder Entf-Taste). Änderungen landen sofort in `profil_ifs.yaml`; beim ersten Speichern wird eine Sicherung `profil_ifs.yaml.bak` angelegt.

## Excel-Liste
- Grün markiert und „neu“: Treffer, die in einem früheren Lauf noch nicht dabei waren.
- „Punkte“: Summe der Gewichte aller Suchbegriffe, CPV-Codes und Regionen, die gepasst haben. „Treffergründe“ zeigt, welche das waren.
- „Link“ führt zur Bekanntmachung.
- Alle Listen liegen im Unterordner `ergebnisse`.

## Suchprofil anpassen
Schlagwörter am einfachsten im Programmfenster bearbeiten. CPV-Codes, Bonusregionen, Ausschlusswörter, Plattformen und die Mindestpunktzahl stehen in `profil_ifs.yaml` und lassen sich mit einem Texteditor (z. B. Editor/Notepad) ändern. Änderungen gelten beim nächsten Start.

## Datenquelle
Offene Tagesexporte des Bekanntmachungsservice des Bundes (oeffentlichevergabe.de, Datenservice Öffentlicher Einkauf). Die Plattform wird am Link zu den Vergabeunterlagen erkannt; die Liste zeigt sie in der Spalte „Plattform“. Die Angebotsfrist kommt aus dem eForms-Export; Bekanntmachungen mit abgelaufener Frist werden ausgelassen.

## Für Entwickler
```
pip install -r requirements-dev.txt
python -m pytest -q
python gui.py                                       # Programmfenster
python ausschreibungscheck.py --von 2026-10-01      # ohne Fenster, ab einem bestimmten Tag
python ausschreibungscheck.py --plattform DTVP --begriff "Starkregen: 12"
python ausschreibungscheck.py --zip export.zip      # lokale Exportdatei testen
```
Bei jedem Push auf `main` baut GitHub die .exe und veröffentlicht sie als neueste Version unter Releases (https://github.com/lemmy1896/AusschreibungsCheck/releases/latest).
