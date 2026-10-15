# CDE-Änderungsübernahme

Autor: **Nico Lehr** · Stand: 15.10.2026 · Version 0.1.0 (Prototyp, nur mit erfundenen Daten getestet)

Seitenprojekt, **getrennt von der Masterarbeit (bess-pipeline)**. Läuft lokal, ohne Internet, ohne KI-Dienst, ohne API-Key.

## Was macht das Tool?

Du hast ein **PDF mit sichtbaren Änderungen** (blau = neu, rot oder durchgestrichen = entfällt, bis hin zu ganzen
Abschnitten) und eine **bestehende CDE-Excel** (nach Sections geordnet, Textblöcke dahinter, kein einheitliches Format).
Das Tool findet die Stellen in der Excel, übernimmt **nur die inhaltlichen Änderungen** und schreibt eine **neue** Excel.
Die Original-Excel bleibt unberührt. Dazu gibt es ein Protokoll, in dem jede Änderung steht.

**Sicherheitsnetz (Bilanz):** Jede erkannte Änderung ist entweder *übernommen*, *zu prüfen* oder *nicht übernommen (mit Grund)*.
Die Wörter-Summe muss zum PDF passen. So geht nichts still verloren.

## Starten (Windows)

1. Python **3.11** installiert? `py -3.11 --version`
2. Einmalig: `scripts\setup.bat` (legt `.venv` an und installiert die Pakete). Wenn pip im Firmennetz Zertifikatsfehler meldet:
   **nicht** mit `--trusted-host` umgehen, sondern IT fragen.
3. Dashboard: `scripts\start_dashboard.bat` → http://localhost:5100/ (bess-pipeline nutzt einen anderen Port, beide können parallel laufen).
4. Ohne Dashboard: `python -m src.cli lauf --pdf aenderungen.pdf --excel cde.xlsx` (mit `--alle-vorschlaege` auch unsichere Treffer).
5. Tests: `python -m pytest -q` (immer in der `.venv`).

## Dashboard (4 Tabs)

*Neuer Lauf* (PDF + Excel ziehen) · *Prüfen & Übernehmen* (Liste, PDF-Ausschnitt, Filter, Freigabe unsicherer Vorschläge) ·
*Verlauf* (alle Läufe, Downloads) · *Farben im PDF* (welche Farben gefunden und wie gedeutet wurden – wichtig zum Kalibrieren).

## Ordner

```
config/       paths.yaml, server.yaml (Port 5100), farben.yaml, zuordnung.yaml  (alle Einstellungen, kein Code)
src/          pdf_reader, excel_model, matcher, planner, applier, protokoll, runs, vorschau, cli, api/
frontend/     Dashboard (HTML/CSS/JS, eigenes Projekt, Stil wie bess)
scripts/      setup.bat, start_dashboard.bat
tests/        pytest + fixtures (erfundene PDFs/Excel werden zur Laufzeit erzeugt)
docs/         konzept_v1.md, auftrag/ (Auftrag + Antworten), adr/ (Entscheidungen)
referenz/     bess-dashboard/  Kopien zum Nachschlagen (werden nicht geladen)
.claude/      entwicklungsregeln.md, design-guideline.md  (Kopien in docs/regeln/; scripts\regeln_nach_claude.bat kopiert sie)
data/         Läufe (runs/), append-only, wird NICHT versioniert
```

## Was funktioniert (getestet mit erfundenen Daten)

- 4 PDF-Markierungsarten: Textfarbe, Hintergrund, Durchstreichen-Linie, PDF-Annotationen.
- 3 verschiedene Excel-Aufbauten (Section/Titel/Inhalt getrennt, kombiniert "3.2 Titel", eine Zeile pro Absatz).
- Wort-, Satz-, Absatz-, ganze Section neu/gelöscht, Umnummerierung, mehrseitige Dokumente, Kopf-/Fußzeilen werden ignoriert.
- Unsichere oder abweichende Stellen werden **nicht** still geändert, sondern mit Grund aufgelistet.
- 26 automatische Tests grün.

## Bekannte Grenzen (ehrlich)

- Nur **einspaltige Text-PDFs**; Tabellen und Bilder im PDF werden nicht verglichen; **Scans** (nur Bild) werden nicht ausgewertet (Warnung).
- Teil-Formatierung (z. B. einzelne fette Wörter) in geänderten Excel-Zellen kann verloren gehen.
- openpyxl kann Bilder/Diagramme in der Excel verlieren (Warnung); Formeln und verbundene Zellen werden beim Zeilen-Einfügen/Löschen nicht angepasst (Warnung).
- Noch **keine echten Beispiele** getestet: Farb- und Zeilenabstand-Schwellen müssen mit echten PDFs kalibriert werden (Tab "Farben im PDF", `config/farben.yaml`).
- Excel-Aufbau wird per Heuristik erkannt; bei Abweichungen hilft später eine YAML-Beschreibung je Layout.

## Nächste Schritte

1. 3–5 echte Beispielpaare (PDF + Excel vorher + Excel nachher) vom anderen Team, **nach Freigabe**, nur lokal, getrennt in Entwicklungs- und Testbeispiele.
2. Messen: Bilanz + Vergleich mit der echten "nachher"-Excel; Schwellen kalibrieren.
3. Farblegende (später), Layout-Beschreibung per YAML, optional lokales LLM nur als Prüfer.

## Für Claude Code (online/Cloud)

Erst `CLAUDE.md` und `.claude/entwicklungsregeln.md` (bzw. `docs/regeln/`) lesen. **Nur erfundene Daten verwenden.** Keine echten Dokumente anfordern,
keine API-Keys einbauen, `bess-pipeline` nicht anfassen, Port 5100 behalten. Im Cloud-Container ist Python evtl. neuer als 3.11:
Code darf keine 3.12+-Syntax benötigen.
