# AusschreibungsCheck – Designentscheidungen

Stand: 13.10.2026

## Ziel
Ausführbares Programm (keine Website), das deutsche Ausschreibungsplattformen nach passenden Bekanntmachungen durchsucht. Zunächst nur evergabe-online.de, Suchprofil auf Basis der Referenzen von ifs Hannover (www.ifs-hannover.de).

## Entscheidungen

| # | Thema | Entscheidung | Begründung |
|---|---|---|---|
| 1 | Datenquelle | Offene Tagesexporte des Bekanntmachungsservice des Bundes (oeffentlichevergabe.de), gefiltert auf evergabe-online | evergabe-online hat keine dokumentierte API/RSS; der Bundesdienst bündelt seit 10/2023 alle Pflichtbekanntmachungen und ist später auf alle Plattformen erweiterbar |
| 2 | Sprache/Form | Python, gepackt als einzelne .exe (PyInstaller) | schnell zu bauen und zu ändern |
| 3 | Profilformat | YAML-Datei neben der .exe | ohne neue .exe änderbar, Kommentare, Gewichtung, mehrere Profile möglich |
| 4 | Ausgabe | Excel-Liste (.xlsx) pro Lauf: Relevanzwert, Titel, Vergabestelle, Ort, Frist, Link, Markierung „neu“ | filterbar, weiterleitbar, keine Mail-Zugangsdaten nötig |
| 5 | Start | Manuell per Doppelklick; lädt alle Tage seit dem letzten Lauf nach (erster Lauf: 30 Tage) und öffnet danach die Excel-Liste | einfachster Start; Zeitplan später ohne Programmänderung ergänzbar |
| 6 | Bau der .exe | GitHub Actions baut auf Windows bei jedem Push; Download als Artefakt bzw. Release | kein Zugriff auf Julius' Rechner, keine Python-Installation nötig |
| 7 | Regionen | Bundesweite Suche, Bonuspunkte für Referenzregionen | keine passende Ausschreibung soll durchrutschen |

## Offene Punkte
- Exakte Export-Endpunkte (vermutlich `/api/notice-exports?pubDay=JJJJ-MM-TT&format=ocds.zip`) und das Feld, an dem sich die Herkunftsplattform evergabe-online erkennen lässt, live prüfen. Aus der Cloud-Umgebung war die Seite nicht erreichbar.
- Suchprofil mit Julius abstimmen (CPV-Codes, Regionen, Ausschlusswörter).

## Suchprofil-Entwurf (ifs Hannover)
- Themen: Generalentwässerungsplanung/GEP, Kanalnetzberechnung, hydraulische Berechnung, Regenwasserbewirtschaftung, Entwässerungskonzept, Regenrückhaltebecken, Regenklärbecken, Retentionsbodenfilter, Versickerung, Mischwasser, Fremdwasser, Fachbeitrag WRRL/Wasserrahmenrichtlinie, Tausalzgutachten, hydrologisches Gutachten, Beckenbuch, Durchflussmessung, Abflussmessung, Hydrometrie, SüwV-kom, Probenahme, Grundwassermessung, Machbarkeitsstudie Gewässer
- CPV-Kandidaten: 71311000, 71322000, 71335000, 71313000, 90713100, 73110000
- Regionen mit Referenzen: Niedersachsen, Berlin, Hamburg, Nordrhein-Westfalen, Brandenburg
- Auftraggebertypen: Kommunen, Stadtentwässerungen/Abwasserverbände, Wasserverbände, Autobahn GmbH, Landesbehörden
