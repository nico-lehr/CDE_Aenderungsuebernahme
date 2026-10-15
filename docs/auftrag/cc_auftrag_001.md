# Auftrag 001 an Claude Code – CDE-Änderungsübernahme

Autor: Nico Lehr · Stand: 15.10.2026

## Rolle und Arbeitsweise
Du bist der leitende Entwickler dieses Projekts. Ich (Nico) habe keine Programmiererfahrung: Erkläre Ergebnisse in einfachen Worten
(Fachbegriffe kurz erklären), duze mich, bleibe knapp. Du arbeitest **selbstständig** und **mit vielen Subagents** (siehe unten).
Frag nur nach, wenn eine Entscheidung nicht rückgängig zu machen ist oder nur ich sie treffen kann.

## Zuerst lesen (in dieser Reihenfolge, nichts davon doppelt lesen)
1. `CLAUDE.md` → `.claude/entwicklungsregeln.md` (falls `.claude/` fehlt: `docs/regeln/`)
2. `README.md`, `docs/konzept_v1.md`, `docs/auftrag/auftrag_und_antworten.md`, `docs/adr/0001-*.md`
3. Dann gezielt den Code: `src/` (pdf_reader, excel_model, matcher, planner, applier, protokoll, runs), `tests/`, `frontend/`

## Worum geht es (Kurzfassung)
Ein PDF zeigt Änderungen: **blau = neu**, **rot oder durchgestrichen = entfällt** (Wörter, Sätze, Absätze, ganze Sections/Subsections).
Das Tool überträgt nur diese **inhaltlichen** Änderungen in eine **bestehende CDE-Excel** (nach Sections geordnet, Textblöcke dahinter,
**kein einheitliches Format** → der Aufbau muss dynamisch erkannt werden). Ergebnis: neue Excel + Protokoll; Original bleibt unverändert.
Nichts darf still verloren gehen (Bilanz-Prinzip). Lieber nicht übernehmen als falsch übernehmen. Das Tool ist für ein anderes Team gedacht.

## Stand
Prototyp v0.1 ist fertig und mit **erfundenen** Daten getestet (26 Tests grün): 4 PDF-Markierungsarten (Farbe, Hintergrund, Strich,
Annotation) × 3 Excel-Layouts, Dashboard auf Port 5100, CLI, Protokoll, Bilanz. **Noch keine echten Beispiele.**
Bekannte Grenzen: nur einspaltige Text-PDFs; Tabellen/Bilder/Scans nicht; Teilformatierung in Zellen geht verloren; Formeln/verbundene
Zellen werden bei Zeilen-Einfügen/-Löschen nicht angepasst; Schwellen (Farbe, Zeilenabstand, Absatz) nicht an echten PDFs kalibriert;
Excel-Aufbau nur per Heuristik.

## Harte Grenzen (nicht verhandelbar)
- **Nur erfundene Daten.** Keine echten Dokumente anfordern, einbauen oder in Commits/Logs zitieren (Regel 16). Keine API-Keys, keine
  externen Dienste mit Kundendaten (Regel 17). Kein LLM im Kern; höchstens später als Prüfer über `LLMClient` (Regel 22).
- `bess-pipeline` existiert hier nicht / ist read only. Port **5100** und alle Pfade nur aus `config/*.yaml` (Regel 28).
- Immer in der `.venv`; Ziel-Python **3.11**. Neue Pakete nur mit Begründung, gegen 3.11 prüfen, pip-Zertifikatsprobleme nicht mit
  `--trusted-host` umgehen. Cloud-Python kann neuer sein: keine 3.12+-Syntax.
- Läufe/Ergebnisse append-only (Regel 29). Unsichere Treffer nie automatisch (Regel 27). Schema-Änderungen → `SCHEMA_VERSION` erhöhen.
- Git: nie `git add -A` / `git add .`, nur gezielt Dateien; kleine Commits; keine PDFs/Excel committen (`.gitignore`).
- Alles, was weitergegeben werden könnte, trägt den Autor **Nico Lehr**. Bestehende Schnittstellen nicht ohne Not ändern.
- Test-Beispiele (später echt) nie zum Tunen benutzen (Regel 11): Entwicklungsdaten und Testdaten strikt trennen.

## Dein Ziel – such das beste selbst
Ich gebe dir bewusst kein starres Ziel. **Schritt 0 (vor jedem Code):** Schreibe `docs/ziel_v1.md` mit
(a) dem besten Zielbild aus Sicht eines Teams, das dem Tool vertrauen muss, (b) messbaren Erfolgskriterien (mindestens: Bilanz immer stimmig,
Anzahl stiller Fehler = 0, Treffergenauigkeit, Abdeckung ohne Nachfrage), (c) einer begründeten Prioritätenliste der Arbeitspakete
(Risiko × Nutzen), (d) Dingen, die du bewusst NICHT tust. Danach setzt du die Liste von oben nach unten um.

Kandidaten für Arbeitspakete (Vorschläge – streichen, ergänzen, umsortieren erlaubt, mit Begründung):
1. **Realismus der Testdaten:** Generator für deutlich realistischere erfundene PDFs/Excel (andere Schriftgrößen, Umbrüche, Silbentrennung, Kopf-/Fußzeilen,
   Aufzählungen, Seitenumbrüche mitten im Absatz, Tabellen/Bilder als Störer, Zwei-Spalten, Scan-Seite, unterschiedliche Rot-/Blautöne, verschiedene
   Markierungs-Mischformen). Ziel: Schwachstellen finden, bevor echte Daten kommen.
2. **Messwerkzeug:** Skript/CLI `python -m src.cli messen` – führt Paare (PDF + Excel vorher + Excel nachher) aus, vergleicht Ergebnis mit „nachher“,
   gibt Treffergenauigkeit, Abdeckung, stille Fehler und Fehlerarten aus (PDF falsch gelesen / Layout falsch erkannt / Zuordnung falsch / Plan falsch).
   Muss sofort mit echten Paaren laufen, die ich später lokal in einen ignorierten Ordner lege.
3. **Layout-Profile in YAML:** Wenn die automatische Erkennung unsicher ist, kann ein Profil pro Excel-Typ Spalten/Modell festlegen; Dashboard zeigt
   erkannten Aufbau und lässt ihn korrigieren (neue Profile nie direkt live – erst Vorschau/Prüfung).
4. **Robustheit der Zuordnung/Planung:** Stress- und Zufallstests (Property-/Fuzz-Tests), Fälle wie doppelte Absätze, umgestellte Reihenfolge, verschobene
   Sections, Umnummerierung, Aufzählungen, Fußnoten, Sonderzeichen/Silbentrennung, sehr lange Dokumente (Performance), Zellen mit Rich Text.
5. **Excel-Sicherheit:** Bilder/Diagramme/Formeln/verbundene Zellen/Filter/bedingte Formatierung nicht beschädigen oder klar warnen; Teilformatierung in
   geänderten Zellen erhalten, wo möglich.
6. **Dashboard-Qualität:** Prüfansicht (PDF-Ausschnitt neben Excel-Vorher/Nachher), schnelle Freigabe, Filter, Tastatur, Fehlermeldungen in einfachem Deutsch,
   Barrierefreiheit, Dark Mode; Design nach `.claude/design-guideline.md`.
7. **Kalibrierung vorbereiten:** Konfig-Schwellen dokumentieren, Werkzeug „Farben im PDF“ ausbauen, Anleitung für mich, wie ich echte Beispiele sicher lokal einspeise.
8. **Dokumentation:** README aktuell halten, ADRs für Entscheidungen, `docs/status.md` (was geht, was nicht, was fehlt), kurze Anleitung für das andere Team.

## Subagents (bitte aktiv nutzen)
Arbeite als Orchestrator. Zerlege in unabhängige Pakete und starte **parallele Subagents** (gern in eigenen Worktrees, damit sich nichts überschreibt).
Empfohlene Rollen:
- **Erkunder** (nur lesen): Code-Review des bestehenden Stands, Liste der Schwachstellen mit Belegen.
- **Red-Team/Tester:** versucht, das Tool mit gemeinen erfundenen Fällen zu brechen; liefert reproduzierbare Tests, repariert nichts selbst.
- **Daten-Bauer:** realistische synthetische PDFs/Excel (Paket 1).
- **Mess-Bauer:** Paket 2.
- **Kern-Entwickler (je Modul ein Agent):** pdf_reader / excel_model+matcher / planner+applier – nur im eigenen Modul ändern (Regel 5, 9).
- **Frontend-Agent:** Paket 6.
- **Prüfer (unabhängig):** liest Diffs und Tests, prüft Regeln 11/16/17/27/28/29, sucht eingeschleuste echte Inhalte oder Geheimnisse. Der Prüfer darf nicht der sein, der den Code geschrieben hat.
Regeln: Jeder Subagent bekommt ein klares Briefing (Ziel, Dateien, Grenzen, Testkommando). Du führst die Ergebnisse zusammen, löst Konflikte, lässt
`python -m pytest -q` laufen. Nichts gilt als fertig, solange Tests rot sind oder ein Prüfer Regelverstöße meldet.

## Fertig-Kriterien
- Alle Tests grün (in der `.venv`), neue Funktionen haben Tests, Bilanz stimmt in jedem Test.
- Keine Regel verletzt, keine echten Inhalte, keine Geheimnisse; `.gitignore` deckt neue Dateitypen ab (Regel 30).
- `docs/ziel_v1.md` und `docs/status.md` aktuell; README stimmt mit dem Code überein.
- Server startet auf Port 5100 mit `scripts\start_dashboard.bat`.

## Abschlussbericht an mich (kurz, einfache Worte)
1. Was ist jetzt besser (3–6 Sätze)? 2. Messwerte aus den synthetischen Tests (Tabelle). 3. Was kann das Tool weiterhin NICHT. 4. Was brauchst du von mir
(z. B. welche echten Beispiele, welche Entscheidung). 5. Wie starte ich es auf meinem Rechner.
