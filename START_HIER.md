# START HIER – CDE-Änderungsübernahme (Autor: Nico Lehr)

1. `README.md` lesen (Überblick, Starten, Grenzen).
2. `docs/konzept_v1.md` (Zielbild, Messung, Machbarkeitstest) und `docs/auftrag/` (Auftrag + Antworten).
3. Regeln: `CLAUDE.md` (Kurzfassung) und `docs/regeln/entwicklungsregeln.md` (Vollversion, plus `design-guideline.md`).
   **Einmalig:** `scripts\regeln_nach_claude.bat` ausführen – kopiert beide nach `.claude\`, wo Claude Code sie erwartet.
4. Einrichten: `scripts\setup.bat` → Start: `scripts\start_dashboard.bat` → http://localhost:5100/
5. Tests: `python -m pytest -q`

**Wichtig:** `bess-pipeline` ist read only. Nur erfundene Daten im Ordner. Echte Beispiele erst nach Freigabe, nur lokal.
