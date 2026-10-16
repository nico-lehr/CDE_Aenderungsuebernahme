# Übergabe: lokal weiterarbeiten (Windows, VS Code, Ollama)

Autor: Nico Lehr · Stand: 16.10.2026 · Arbeitsbranch der Cloud-Sitzung: `claude/cool-keller-5tt3mh` (Stand ist nach `main` gemergt)

> **Ohne GitHub-Sync?** Updates per ZIP einspielen: `docs/arbeiten_ohne_github_sync.md` (Skript `scripts\update_einspielen.bat`).
> **Neue lokale Claude-Code-Sitzung:** fertiger Prompt mit allen Infos und Aufgaben in `docs/prompt_lokale_sitzung.md` (Datenschutz-Regeln beachten).

## 1. Wo steht das Projekt?
- **Etappe 1 von Auftrag 002 ist fertig** (Gegenprobe gegen das PDF, genauerer PDF-Leser, Reißverschluss-Abgleich, zweiter Farb-Leser, realistischere Testdaten). Zahlen und Grenzen: `docs/status.md`, Roadmap: `docs/plan_auftrag_002.md`.
- **Neu festgelegt, aber noch nicht eingebaut** (Etappe 2a): Farblogik im PDF (durchgestrichen = entfällt, rot/blau ohne Strich = neu) und farbige Markierung in der Excel
  (rot = neu, blau durchgestrichen = alt). Grundlage ist da: `src/markierung.py`, `docs/adr/0005-farblogik-und-excel-markierung.md`.
- **Vorbereitet für deine Beispielprojekte:** Steckbrief-Werkzeug (`steckbrief`), Fehler-Schablonen (`messen --schablonen`) und der Fragebogen für Ollama (`docs/fragebogen_ollama.md`).

## 2. Einrichten (einmalig)
1. Python **3.11** installiert? `py -3.11 --version`.
2. Im Projektordner (am besten auf **C:**, nicht auf dem langsamen Laufwerk M:): `scripts\setup.bat` (legt `.venv` an und installiert die Pakete).
   Meldet pip einen Zertifikatsfehler im Firmennetz: **nicht** mit `--trusted-host` umgehen, IT fragen.
3. Tests: `.venv\Scripts\python -m pytest -q` (Soll: alles grün; einige Tests brauchen LibreOffice und werden sonst übersprungen).
4. Optional LibreOffice (nur für die realistischeren Test-PDFs): Windows-Installation genügt, `soffice` muss im PATH sein.

## 3. Testdaten erzeugen (erfunden; liegen unter `data/`, nicht in Git)
```
.venv\Scripts\python scripts\erzeuge_paare.py --korpus entwicklung --seeds 1-60
.venv\Scripts\python scripts\erzeuge_paare.py --korpus entwicklung --seeds 101-160
.venv\Scripts\python scripts\erzeuge_paare.py --korpus ex_test --seeds 1001-1030
.venv\Scripts\python -m src.cli messen --korpus entwicklung --mit-vorschlaegen
```
Seeds ab **2000** sind für den späteren ungesehenen Test reserviert (Regel 11): nicht erzeugen, nicht ansehen.

## 4. Deine echten Beispielprojekte (nur lokal!)
Reihenfolge, sicher für sensible Daten (siehe auch `docs/anleitung_steckbrief.md` und `docs/anleitung_echte_paare.md`):
1. **Steckbrief** (nur Zahlen, Farbcodes, Formen – kein Text):
   `.venv\Scripts\python -m src.cli steckbrief --pdf A.pdf --vorher V.xlsx --nachher N.xlsx --name projekt1`
   danach prüfen: `.venv\Scripts\python -m src.cli steckbrief --pruefe data\steckbrief\001_projekt1\steckbrief.json`.
   Wähle `--name` neutral (Datei-, Ordner- und Blattnamen können echte Namen enthalten).
2. **Messzahlen** (nur Zahlen): echte Paare wie in `docs/anleitung_echte_paare.md` nach `data\paare\entwicklung\echt_01\…` legen (mit zufälliger Aufteilung in Entwicklung und Test), dann `python -m src.cli messen --korpus entwicklung --paar echt_01 --schablonen`. Die Berichte enthalten nie Dokumentinhalte, nur Zahlen und Formen. Hinweis: Die Aufteilung „test“ in der alten Anleitung heißt jetzt `test` nur für den späteren ungesehenen Korpus; der frühere Testkorpus heißt `ex_test`.
3. **Fragebogen** mit deiner lokalen Ollama beantworten lassen (`docs/fragebogen_ollama.md`); die Antwort selbst lesen, bevor du sie teilst.
4. Der Cloud-Sitzung (oder einer lokalen Claude-Sitzung) gibst du **nur** Steckbrief, Messzahlen, Schablonen und die geprüften Fragebogen-Antworten. Echte PDFs/Excel gehen nie in Git (`.gitignore`) und nie an einen Cloud-Dienst.

## 5. Wie geht es weiter (Reihenfolge)
1. Steckbriefe auswerten ⇒ ADR 0005 mit den echten Werten bestätigen (Hex-Farben, Trenner Alt/Neu, Darstellung neuer und gestrichener Zeilen).
2. **Etappe 2a** (siehe Plan): LEGENDE (PDF-Farblogik) · MARKIERUNG (Excel-Markierung im Applier + markierte Gegenprobe) · GEN v3 / NACHBAU (Generator aus dem Steckbrief).
3. Messen auf Entwicklung + Nachbau; Zwischenbericht.
4. **Etappe 2b**: Excel-Zeilen (Formeln, Füllspalten), Dashboard-Ansicht „PDF sagt / Excel hat / Ergebnis“ und Abwahl, Ollama-Zweitmeinung (`LLMClient`), Kalibrierung, ungesehener Test, Prüfer, Version 0.3.0.

## 6. Regeln für jede Sitzung (lokal oder Cloud)
- `CLAUDE.md` lesen; immer in der `.venv`; nur erfundene Daten in Code, Tests, Kommentaren, Commits; keine API-Keys; kein `git add -A`.
- Test-Beispiele nie zum Einstellen von Schwellen benutzen (Regel 11). `ex_test` (1001–1030) gilt seit Etappe 1 als **Entwicklung**; der ungesehene Test kommt aus Seeds ≥ 2001.
- Unsicheres nie automatisch (Regel 27). Stille Fehler dürfen von Schritt zu Schritt nie steigen.
- Ollama nur über `LLMClient`, `temperature 0`, nie als Entscheider (Regel 12/22).
- Vor „fertig“: `pytest -q` komplett grün, `messen` vorher/nachher vergleichen (`--vergleich`).

## 7. Offene Punkte (Stand Etappe 1)
- Formeln und Füllspalten in neuen Zeilen werden noch nicht gefüllt (alle Zellen erreicht: ~58 %).
- Grüne/unbekannte Farben und Zwei-Spalten/Tabellen/Scans: nur erkannt und gemeldet; Vorschläge dazu schwach (ex_test 48 %). Mögliche Hilfe: Ollama-Lesevorschlag (Bildmodell).
- Die 4 verbleibenden stillen Fehler der Entwicklung: abgetrennter Satzteil als eigene Zeile; Füllzelle der neuen Zeile in der falschen Zeile; Details in `docs/status.md`.
- Dashboard kann automatische Einträge noch nicht einzeln abwählen (Schnittstelle `abwahl` ist im Applier vorbereitet).
- Aufräumen: In der Cloud-Sitzung entstanden lokale Hilfs-Arbeitsordner (`.claude/worktrees/`, ignoriert); sie gehören nicht zu `main`.
