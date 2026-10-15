# CDE-Änderungsübernahme

Autor: **Nico Lehr** · Stand: 15.10.2026 · Version 0.2.0 (nur mit erfundenen Daten getestet; Zahlen siehe `docs/status.md`)

Seitenprojekt, **getrennt von der Masterarbeit (bess-pipeline)**. Läuft lokal, ohne Internet, ohne KI-Dienst, ohne API-Key.

## Was macht das Tool?

Du hast ein **PDF mit sichtbaren Änderungen** (blau = neu, rot oder durchgestrichen = entfällt, bis hin zu ganzen
Abschnitten) und eine **bestehende CDE-Excel** (nach Sections geordnet, Textblöcke dahinter, kein einheitliches Format).
Das Tool findet die Stellen in der Excel, übernimmt **nur die inhaltlichen Änderungen** und schreibt eine **neue** Excel.
Die Original-Excel bleibt unberührt. Dazu gibt es ein Protokoll, in dem jede Änderung steht.

**Grundsatz: lieber nicht übernehmen als falsch übernehmen.** Jede erkannte Änderung ist entweder *übernommen*,
*zu prüfen* oder *nicht übernommen (mit Grund)*. Zusätzlich meldet das Tool ehrlich, was es **nicht lesen konnte**
(Scan-Seiten, Zwei-Spalten-Text, Tabellen, Bilder, unbekannte Farben): dann steht in der Bilanz **„unvollständig“**, nie ein grünes „alles ok“.

## Starten (Windows)

1. Python **3.11** installiert? `py -3.11 --version`
2. Einmalig: `scripts\setup.bat` (legt `.venv` an und installiert die Pakete). Wenn pip im Firmennetz Zertifikatsfehler meldet:
   **nicht** mit `--trusted-host` umgehen, sondern IT fragen.
3. Dashboard: `scripts\start_dashboard.bat` → http://localhost:5100/ (bess-pipeline nutzt einen anderen Port, beide können parallel laufen).
4. Tests: `python -m pytest -q` (immer in der `.venv`).

## Befehle auf der Kommandozeile (`python -m src.cli …`)

| Befehl | Wofür |
|---|---|
| `analyse --pdf A.pdf --excel B.xlsx` | nur auswerten, Bilanz anzeigen, nichts schreiben |
| `lauf --pdf A.pdf --excel B.xlsx` | auswerten und alle **sicheren** Änderungen eintragen (`--alle-vorschlaege` nimmt auch Vorschläge mit) |
| `messen --korpus entwicklung` | **Messwerkzeug**: Paare (PDF + Excel vorher + nachher) ausführen und mit „nachher“ vergleichen — Treffergenauigkeit, Abdeckung, **stille Fehler**, Fehlerarten. Ausgabe nur Zahlen, nie Dokumentinhalt |
| `farben --pdf A.pdf` | welche Farben im PDF vorkommen und wie sie gedeutet werden (nur Zahlen und Farbcodes) |
| `profil liste / vorschau / bestaetigen / analyse` | Excel-Aufbau von Hand festlegen, wenn die Automatik unsicher ist. **Neue Profile gehen nie direkt live**: erst Vorschau, dann bewusste Bestätigung |

Echte Beispiele sicher einspeisen: `docs/anleitung_echte_paare.md`. Schwellen verstehen/kalibrieren: `docs/kalibrierung.md`.
Testdaten erzeugen (erfunden): `python scripts/erzeuge_paare.py --korpus entwicklung --seeds 1-60`.

## Dashboard

*Neuer Lauf* (PDF + Excel ziehen) · *Prüfen & Übernehmen* (Liste mit Suche/Filter/Sortierung, PDF-Ausschnitt, „Dort steht jetzt“ neben
„Im PDF: alt → neu“, Hinweise und Wirkung je Änderung, Freigabe unsicherer Vorschläge, frühere Ergebnisse mit Download) ·
*Verlauf* (alle Läufe; **Archivieren statt Löschen**, Wiederherstellen) · *Farben im PDF* (zum Kalibrieren).
Bedienbar per Tastatur, hell und dunkel, Meldungen in einfachem Deutsch.

## Ordner

```
config/       paths.yaml, server.yaml (Port 5100), farben.yaml, zuordnung.yaml, layouts.yaml (Vorlage)  (alle Einstellungen, kein Code)
src/          pdf_reader, excel_model, matcher, planner, applier, protokoll, runs, messen, layouts, cli, api/
frontend/     Dashboard (HTML/CSS/JS)
scripts/      setup.bat, start_dashboard.bat, erzeuge_paare.py (erfundene Testpaare)
tests/        pytest; fixtures/ = Generator für erfundene PDFs/Excel; redteam/ = gemeine Fälle (erwartete Fehler sind als xfail markiert)
docs/         ziel_v1.md (Ziel + Kriterien), status.md (was geht / was nicht), konzept_v1.md, paar_format.md, kalibrierung.md,
              anleitung_echte_paare.md, messung_baseline.md, adr/ (Entscheidungen), auftrag/, regeln/
data/         Läufe (runs/), Paare (paare/), Messberichte (messung/), Profile (profile/) — append-only, wird NICHT versioniert
```

## Wie gut ist es?

**Kurz und ehrlich:** Auf erfundenen Entwicklungsbeispielen sind die stillen Fehler von 64 auf 7 gesunken, auf einem getrennten, ungesehenen Testkorpus sind es 23 (Ziel: 0). Das Tool warnt ehrlich, ist aber noch nicht „blind vertrauenswürdig“: automatisch übernommene Änderungen bitte stichprobenartig prüfen.

Aktuelle Messwerte und ehrliche Grenzen stehen in **`docs/status.md`**. Die Ziele und Kriterien stehen in **`docs/ziel_v1.md`**.
Alle Zahlen stammen von **erfundenen** Beispielen; sie sind kein Versprechen für echte Dokumente.

## Bekannte Grenzen (ehrlich)

- Nur **einspaltige Text-PDFs** werden gelesen. Tabellen, Bilder, Zwei-Spalten-Text und **Scans** werden **erkannt und gemeldet**, aber nicht verarbeitet (kein OCR).
- Teil-Formatierung (z. B. einzelne fette Wörter) in geänderten Excel-Zellen kann verloren gehen.
- Formeln, bedingte Formatierung, Datenprüfung, Filter, Excel-Tabellen und verbundene Zellen an der Einfügestelle werden **nicht repariert**, sondern führen zu „prüfen“ mit Warnung.
- Noch **keine echten Beispiele** getestet: Farb- und Zeilenabstand-Schwellen müssen mit echten PDFs kalibriert werden (`docs/kalibrierung.md`).
- Automatisch Übernommenes lässt sich im Dashboard noch nicht einzeln abwählen (die Schnittstelle kennt nur Freigaben für Vorschläge).

## Nächste Schritte

1. 3–5 echte Beispielpaare (PDF + Excel vorher + Excel nachher) vom anderen Team, **nach Freigabe**, nur lokal, getrennt in Entwicklungs- und Testbeispiele (`docs/anleitung_echte_paare.md`).
2. Messen (`python -m src.cli messen`), Schwellen nur an den Entwicklungsbeispielen kalibrieren, einmal mit den Testbeispielen messen.
3. Farblegende (später), Layout-Profile für Excel-Typen, die die Automatik nicht sicher erkennt, optional lokales LLM nur als Prüfer.

## Für Claude Code (online/Cloud)

Erst `CLAUDE.md` und `.claude/entwicklungsregeln.md` (bzw. `docs/regeln/`) lesen. **Nur erfundene Daten verwenden.** Keine echten Dokumente anfordern,
keine API-Keys einbauen, `bess-pipeline` nicht anfassen, Port 5100 behalten. Im Cloud-Container ist Python evtl. neuer als 3.11:
Code darf keine 3.12+-Syntax benötigen.
