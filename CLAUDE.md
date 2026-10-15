## Entwicklungsregeln (Kurzfassung — Vollversion: .claude/entwicklungsregeln.md)

Projekt: **CDE-Änderungsübernahme** (Autor: Nico Lehr). Seitenprojekt, getrennt von der Masterarbeit (bess-pipeline).

ich möchte immer in der .venv Umgebung arbeiten!

Projekt-Standard: Python 3.11 (`.venv` mit `py -3.11 -m venv .venv` anlegen, dann `scripts\setup.bat`).
`requirements.txt` ist UTF-8-kodiert und mit 3.11-kompatiblen Versionen gepinnt. Neue Abhängigkeiten immer gegen 3.11 prüfen, bevor sie gepinnt werden.

- `bess-pipeline` ist **read only** — dort nie etwas ändern. Eigener Port **5100** (nie den von bess nutzen).
- Regel 11: Test-Beispiele NIE für Debugging/Schwellen-Tuning anfassen
- Regel 16: Keine echten PDF-/Excel-Inhalte/Kundendaten in Code, Tests, Kommentare, Commits (nur erfundene Daten)
- Regel 17: Keine externen API-Calls mit Kundendaten; keine API-Keys; firmeninterne Ziele (Ollama) ok
- Regel 22: LLM-Zugriff (falls nötig) nur über LLMClient-Interface
- Regel 27: Unsichere Treffer nie automatisch übernehmen
- Regel 28: Keine hartcodierten Pfade, nur config/*.yaml
- Regel 29: Läufe/Ergebnisse append-only, nie überschreiben
- Bilanz-Prinzip: nichts geht still verloren (Details in den Regeln, Punkt D)
- Nie `git add -A`; nur gezielt Dateien hinzufügen
- Cloud-Arbeit nur mit erfundenen Daten. Echte Dokumente erst nach Freigabe durch Nico.
- Tests: `python -m pytest -q` — vor "fertig" alles grün

Bei Detailfragen zu Architektur/Regeln → .claude/entwicklungsregeln.md lesen (Kopie: docs/regeln/entwicklungsregeln.md).
Bei Designaufgaben: DesignGuide findest du hier, halte dich an diese: .claude/design-guideline.md (Kopie: docs/regeln/design-guideline.md).
Falls `.claude/` fehlt: `scripts\regeln_nach_claude.bat` ausführen bzw. die Dateien aus docs/regeln/ nehmen.
Einstieg und Überblick: README.md, docs/konzept_v1.md, docs/auftrag/
