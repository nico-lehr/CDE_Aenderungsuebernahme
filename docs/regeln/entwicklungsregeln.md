# Projekt-Regeln: CDE-Änderungsübernahme

Autor: Nico Lehr · Stand: 15.10.2026
Zweck: Arbeitsgrundlage für KI-gestützte Entwicklung (Claude Code). Die Nummern entsprechen der bess-pipeline
(`bess-pipeline/.claude/entwicklungsregeln.md`), damit "Regel 16" in beiden Projekten dasselbe heißt.
Nicht übernommene bess-Regeln (13, 15, 19, 23, 24-Teile, 25) betreffen LLM-Extraktion/Wissenschaft und gelten hier nicht.

**Nach jeder Code-Änderung den Flask-Server neu starten** (Strg+C, dann `scripts\start_dashboard.bat`).

---

## Projekt-Rahmen (NEU, nur hier)

A. **Seitenprojekt, getrennt von der Masterarbeit.** Der Ordner `bess-pipeline` ist **nur lesbar**. Dort wird nichts
   bearbeitet, verschoben oder gelöscht. Arbeitsordner ist ausschließlich dieser Projektordner.
B. **Eigener Port 5100** (`config/server.yaml`), nur `127.0.0.1`. Kein CORS, keine Cookies/Sessions. Nie denselben Port,
   Datenordner oder dieselbe `.venv` wie bess nutzen. Beide Projekte müssen parallel laufen können.
C. **Immer in der `.venv` arbeiten** (Python 3.11): `py -3.11 -m venv .venv`, danach `scripts\setup.bat`.
   Neue Abhängigkeiten gegen 3.11 prüfen, dann in `requirements.txt` pinnen (UTF-8). pip-Zertifikatsfehler im Firmennetz
   **nicht** mit `--trusted-host` umgehen, sondern nachfragen.
D. **Bilanz-Prinzip:** Nichts geht still verloren. Jede im PDF erkannte Änderung steht entweder in der Excel
   (übernommen), wartet auf Freigabe (prüfen) oder steht mit Grund auf der Liste "nicht übernommen". Die Wörter-Summe
   muss zur Zahl der markierten PDF-Wörter passen. Ist das nicht so, ist das ein Fehler, kein Schönheitsfehler.
E. **Regelbasiert statt KI.** Kern ist deterministisch (pdfplumber + Textvergleich). Ein lokales LLM (Ollama) wäre höchstens
   später ein Prüfer, nie der Entscheider, und nur über das `LLMClient`-Interface (Regel 22).
F. **Lieber nicht übernehmen als falsch übernehmen.** Unsichere Stellen (Übereinstimmung < mittel) nie automatisch (Regel 27).
G. **Original-Excel bleibt unverändert.** Es wird immer auf einer Kopie gearbeitet; Ergebnis liegt in einem neuen Ordner.
H. **Cloud/Online-Arbeit nur mit erfundenen Daten.** Der Projektordner enthält keine echten Dokumente des anderen Teams.
   Echte Beispiele erst nach Freigabe von Nico, und sie bleiben lokal (siehe Regel 16/17, `.gitignore`).
I. **Schreibweise:** Dokumente, die weitergegeben werden könnten, tragen den Namen **Nico Lehr** als Autor.

---

## Architektur (Pipeline, jedes Modul eine Aufgabe)

| Modul | Aufgabe |
|---|---|
| `src/pdf_reader.py` | PDF → Wörter mit Status (bleibt/neu/entfällt) → Zeilen → Blöcke. Erkennt Farbe, Hintergrund, Durchstreichen, Annotationen |
| `src/excel_model.py` | Excel-Aufbau erkennen (Section-/Titel-/Inhaltsspalte), Kandidaten mit Wortpositionen |
| `src/matcher.py` | PDF-Block ↔ Excel-Stelle zuordnen, mit Sicherheit (hoch/mittel/niedrig) |
| `src/planner.py` | Konkrete Operationen planen (ersetzen, einfügen, löschen, Zeile einfügen) |
| `src/applier.py` | Auf Kopie anwenden, Erwartungstext prüfen, Konflikte erkennen, Nachkontrolle |
| `src/protokoll.py` | Bilanz + Änderungsprotokoll (JSON/Excel) |
| `src/runs.py` | Läufe, append-only im `data/runs/` |
| `src/api/` | Flask-API (nur Darstellung/Transport, keine Fachlogik) |
| `frontend/` | Dashboard (nur Darstellung) |
| `config/*.yaml` | Pfade, Port, Farben, Schwellen |
| `tests/` | pytest mit erfundenen, zur Laufzeit erzeugten PDFs/Excel |

---

## Regeln

### Allgemeine Arbeitsweise
1. So einfach wie möglich, kein Overengineering.
2. Bei Unklarheiten nachfragen statt raten.
3. Token-sparend arbeiten: gezielte Edits statt ganze Dateien neu schreiben.
4. Keine ungefragten Umbenennungen oder Restrukturierungen.
5. Nur an Dateien arbeiten, die für die Aufgabe relevant sind.

### Code-Struktur
6. Bestehendes wiederverwenden (auch `referenz/bess-dashboard/` und `frontend/`), bevor Neues gebaut wird.
7. Funktionen klein und einzweckig.
8. Keine Fachlogik in `frontend/` und `src/api/`; keine UI-Logik in `pdf_reader/matcher/planner/applier`.
9. Bestehende Schnittstellen (Funktionssignaturen, JSON-Schema) nicht ohne Anfrage ändern.
10. Neue Abhängigkeit nur mit kurzer Begründung.
22. LLM-Zugriff (falls je nötig) ausschließlich über ein abstraktes `LLMClient`-Interface, nie direkt gegen ein Backend.
28. **Keine hartcodierten Pfade/Adressen.** Alles aus `config/paths.yaml` / `config/server.yaml`; Pfade mit `pathlib.Path`.
29. **Läufe und Ergebnisse sind append-only.** Nichts überschreiben; jede Übernahme legt einen neuen `ergebnis_NNN/`-Ordner an.
    Keine Versionierung von Läufen über Git.
27. **Unsichere Treffer nie automatisch.** Niedrige Sicherheit → "zu prüfen" bzw. "nicht übernommen", nie stillschweigend.
24. Das JSON-Schema (`plan.json`, `ergebnis.json`) hat `schema_version` (`SCHEMA_VERSION` in `models.py`). Änderungen → Version erhöhen
    und Migration/Notiz in `docs/adr/`.

### Wissenschaftliche Sauberkeit / Tests
11. **Test-Beispiele (echte Paare von Nicos Team) nie zum Debuggen/Tunen benutzen.** Entwicklungsbeispiele und Testbeispiele
    getrennt halten; Schwellen (`config/farben.yaml`, `config/zuordnung.yaml`) nur anhand der Entwicklungsbeispiele ändern.
12. Falls je ein LLM eingesetzt wird: `temperature=0`.
14. Größere Entscheidungen als kurzes ADR in `docs/adr/`.
20. Bei Fehlern erst Ursache verstehen, keine Symptom-Patches.
21. Kleine, einzeln nachvollziehbare Commits. **Nie `git add -A` / `git add .`**, nur gezielt `git add <Datei>`.

### Daten und Datenschutz
16. **Niemals echte PDF-/Excel-Inhalte, Kundennamen oder Angebotsdaten** in Code, Kommentare, Logs, Tests oder Commit-Messages.
    Tests nutzen nur erfundene Inhalte (`tests/fixtures`). `.gitignore` ignoriert `*.pdf`, `*.xlsx` & Co. und `data/`.
17. **Keine externen Netzwerkaufrufe mit Kundendaten** ohne Freigabe. Firmeninterne Ziele (Ollama lokal/Cluster) sind ok.
    Keine API-Keys im Projekt. Claude (Assistent) bekommt echte Dokumente des anderen Teams nur nach ausdrücklicher Freigabe durch Nico.
18. Neue Bibliotheken vorher auf Internet-/Telemetrie-Zugriff zur Laufzeit prüfen.
30. Neue Datei, die Geheimes/Umgebungsspezifisches enthalten kann → vor Abschluss prüfen, ob `.gitignore` sie erfasst. Im Zweifel nicht versionieren und nachfragen.

### Ablauf
- Tests: `python -m pytest -q` (in der `.venv`). Vor "fertig" müssen alle grün sein.
- Start: `scripts\start_dashboard.bat` → http://localhost:5100/
- M: ist langsam; Code besser auf C: bearbeiten und später zurückkopieren.

## Offene Klärung (Stand 15.10.2026)
- Datenschutz/Code-Rechte für Daten des anderen Teams: **nicht geklärt** → bis dahin nur erfundene Daten, Cloud-Arbeit nur an synthetischen Beispielen.
- Echte Beispielpaare (PDF + Excel vorher + Excel nachher): kommen laut Nico in der Folgewoche.
