# Prompt für eine neue, lokal laufende Claude-Code-Sitzung

Autor: Nico Lehr · Stand: 16.10.2026
Benutzung: Öffne den Projektordner in VS Code, starte Claude Code und füge den Text zwischen den Linien als **erste Nachricht** ein.

**Wichtiger Hinweis zum Datenschutz:** Auch eine lokal gestartete Claude-Code-Sitzung schickt alles, was sie liest, an Anthropic (Cloud-Modell). Deshalb gilt für diese Sitzung dieselbe Regel wie für die Cloud-Sitzung:
**Sie darf keine echten Dokumente lesen.** Echte PDFs/Excel liest nur deine lokale Ollama. Der Prompt unten enthält die Regel und die Sperre.

---------------------------------------------------------------------------------------------------

Du bist der leitende Entwickler im Projekt **CDE-Änderungsübernahme** (Autor: Nico Lehr). Ich (Nico) habe keine Programmiererfahrung: Erkläre in einfachen Worten, duze mich, bleibe knapp, **stelle Rückfragen, wenn sie zu einem besseren Ergebnis führen**. Du arbeitest selbstständig, gern mit parallelen Subagents (eigene Git-Worktrees, klare Besitzlisten pro Datei), und führst die Ergebnisse zusammen.

## 0. Zuerst lesen (in dieser Reihenfolge)
1. `CLAUDE.md` und `docs/regeln/entwicklungsregeln.md` (Regeln 11, 16, 17, 22, 27, 28, 29, Bilanz-Prinzip).
2. `docs/uebergabe_lokal.md` (Setup, Befehle, Regeln für lokale Sitzungen) und `docs/plan_auftrag_002.md` (Roadmap, das ist deine Arbeitsgrundlage).
3. `docs/ziel_v2.md`, `docs/adr/0004-gegenprobe-schema3.md`, `docs/adr/0005-farblogik-und-excel-markierung.md`, `docs/status.md`.
4. Code gezielt: `src/pdf_reader.py`, `src/excel_model.py`, `src/matcher.py`, `src/abgleich.py`, `src/planner.py`, `src/applier.py`, `src/gegenprobe.py`, `src/markierung.py`, `src/sicht.py`, `src/messen.py`, `src/steckbrief.py`, `tests/`.

## 1. Harte Grenzen (nicht verhandelbar)
- **Datenschutz:** Echte Beispielprojekte (PDF/Excel von meinem Team) dürfen **nur meine lokale Ollama** sehen. **Du liest sie nie** – auch nicht „kurz zum Prüfen“. Gesperrte Orte: alles unter `data\paare\echt*`, `data\paare\entwicklung\echt_*`, jeder Ordner mit echten Dateien. Du arbeitest nur mit **erfundenen** Daten (Generator) und mit dem **Steckbrief** (Zahlen/Formen), den ich dir einfüge, nachdem ich ihn mit `steckbrief --pruefe` geprüft habe. Bitte mich nie, echte Inhalte einzufügen.
- Keine API-Keys im Projekt, keine externen Dienste mit Kundendaten (Regel 17). KI-Zugriff im Tool nur über `LLMClient`, `temperature 0`, nie als Entscheider (Regel 12/22).
- Immer in der `.venv` (Python 3.11, keine 3.12+-Syntax), **keine neuen Pakete** (pip im Firmennetz gesperrt; nicht mit `--trusted-host` umgehen). Pfade/Schwellen nur in `config/*.yaml` (Regel 28), Port 5100, `bess-pipeline` nicht anfassen.
- Unsicheres wird **nie automatisch** eingetragen (Regel 27). Lieber „prüfen“ als falsch. Läufe/Ergebnisse append-only (Regel 29). Schemaänderung ⇒ `SCHEMA_VERSION` erhöhen + ADR (Regel 24).
- Test-Beispiele nie zum Einstellen von Schwellen benutzen (Regel 11). Korpora: `entwicklung` (Seeds 1–60, 101–160), `ex_test` (1001–1030, jetzt offen Entwicklung). **Seeds ab 2000 sind für den späteren ungesehenen Test reserviert:** nicht erzeugen, nicht ansehen, genau einmal am Ende messen.
- Nie `git add -A` / `git add .`, nur gezielt Dateien; kleine Commits; keine PDFs/Excel in Git. Commit-Nachrichten enden mit den Zeilen, die dir das System vorgibt.
- **Stop-Regel:** Die Zahl der stillen Fehler (`messen`) darf von Merge zu Merge nie steigen. Vor „fertig“ immer `python -m pytest -q` komplett grün.

## 2. Worum geht es (Kurzfassung)
PDF mit sichtbaren Änderungen + bestehende CDE-Excel (nach Sections geordnet, Textblöcke dahinter, kein einheitliches Format) ⇒ neue Excel + Protokoll; Original bleibt unverändert; nichts geht still verloren (Bilanz).
**Zielbild (von mir festgelegt):**
- **Farben im PDF sind uneinheitlich:** Neuer Inhalt ist mal rot, mal blau. **Nur Durchstreichen** bedeutet „entfällt“ (egal welche Farbe). Es gibt keine PDFs, in denen Rot ohne Strich „entfällt“ bedeutet.
- **In der Excel müssen die Änderungen markiert werden:** neu = **rot**, alt = **blau + durchgestrichen**. Ersetzte Stelle: Alt direkt vor Neu in derselben Zelle, ein Leerzeichen dazwischen. Ganz gestrichene Section: **Zeile bleibt**, Inhalt blau durchgestrichen. Neue Section/Zeile/Absatz: rot. Gelöschter Text: blau durchgestrichen stehen lassen.
- Neue Zeilen „wie ein Mensch“: Nachbarzeile gleichen Typs als Vorlage (Formeln mit angepassten Bezügen, immer gleiche Werte übernehmen, individuelle Spalten leer + Hinweis).
- Ollama (meine 12-GB-GPU) nur als Zweitmeinung: Vorschläge/Bremse, einzeln abschaltbar, Standard aus, nur an, wenn eine Messung bei mir den Nutzen zeigt.

## 3. Stand (Ende Etappe 1, alles auf `main`)
Fertig: Gegenprobe A1/A3 + Absatz-Sperre (`src/gegenprobe.py`), Absatz-Abgleich und Reißverschluss (`src/abgleich.py`), genauerer PDF-Leser, zweiter Farb-Leser und Farb-Inventur (`src/sicht.py`), Generator v2 + LibreOffice-PDFs, Messwerkzeug mit Restarbeit/Schablonen, Steckbrief-Werkzeug (`src/steckbrief.py`), Grundlage der Excel-Markierung (`src/markierung.py`, ADR 0005, Konfig `markierung:` in `config/zuordnung.yaml`, Test-Weiche `markierung.aenderungen: keine` in `tests/conftest.py`; Tests mit `@pytest.mark.markierung_an` prüfen den neuen Standard).
Messwerte (erfundene Daten): Entwicklung 120 Paare: 4 stille Fehler, 98,1 % der Automatik richtig, Vorschläge nach Freigabe 87 %, Abdeckung (Haupttext) 87,4 %, alle Zellen 57,8 %, Lese-Lücke 2,2 %, Zweiter-Lauf-Test 120/120. `ex_test`: 95,8 % richtig, Vorschläge 48 %, Lese-Lücke 7 %.
**Noch nicht eingebaut:** neue Farblogik im PDF, Markierung in der Excel (der Applier schreibt noch „sauber ersetzt“), Formeln/Füllspalten in neuen Zeilen, Ollama-Anschluss, Dashboard-Abwahl.

## 4. Aufgaben in dieser Reihenfolge
**A. Einrichten und prüfen (zuerst, kurz):** `scripts\setup.bat`, `.venv\Scripts\python -m pytest -q` (alles grün; einige Tests brauchen LibreOffice und werden sonst übersprungen). Testkorpora erzeugen: `python scripts\erzeuge_paare.py --korpus entwicklung --seeds 1-60`, `... --seeds 101-160`, `--korpus ex_test --seeds 1001-1030`; `python -m src.cli messen --korpus entwicklung --mit-vorschlaegen` (Basiszahlen notieren).
**B. Steckbrief auswerten (sobald ich ihn einfüge):** Aus `steckbrief.json` die echten Werte entnehmen: Hex-Farben rot/blau, Trenner zwischen Alt und Neu, Darstellung neuer und gestrichener Zeilen, Legende je Dokument, Strich-Geometrie. `docs/adr/0005-…` und `config/zuordnung.yaml` (`markierung:`) daran anpassen (Annahmen bestätigen oder korrigieren, im ADR vermerken). Unklares **mich fragen**.
**C. Etappe 2a (siehe `docs/plan_auftrag_002.md`, Abschnitt 5):**
1. **LEGENDE** (`src/pdf_reader.py`, `config/farben.yaml`, `src/planner.py`): Strich ⇒ entfällt (jede Farbe); rot **oder** blau ohne Strich ⇒ neu; `gemischt_blau_strich` ist jetzt normal (aus der Sperrliste); Dokument-Legende im Ergebnis; „steht schon da“-Schutz für farbigen Text, der exakt an der Stelle steht; neue Red-Team-Fälle (rot-neu, blau-neu, gemischt im selben Dokument, schwarz/rot/blau + Strich).
2. **MARKIERUNG** (`src/applier.py`, `src/gegenprobe.py` additiv, `src/excel_model.py`): Änderungen als Rich-Text-Läufe schreiben (`src/markierung.py`: `zustand_markiert`, `schreibe`, `faerbe_zelle`, `ansichten`); Basisschrift/Zahlenformat/Füllung erhalten; **A3 markiert:** `alt_ansicht` == Original (nichts geht verloren), `neu_ansicht` == PDF-Neu; Excel-Modell matcht auf `neu_ansicht`, erkennt markierte Stellen (Idempotenz: zweiter Lauf ändert nichts); `geloeschte_section: markieren`.
3. **GEN v3 / NACHBAU** (`tests/fixtures/*`, `scripts/erzeuge_paare.py`, `src/messen.py`): Generator v3 mit Legende je Dokument (blau-neu, rot-neu, gemischt; Strich in schwarz/rot/blau) und **markierter** `nachher.xlsx` (Option `nachher_format: sauber|markiert`); Messung „Format-Treue“ je Lauf (alt/neu/normal); **Nachbau-Generator aus dem Steckbrief** (`--steckbrief`), Korpus `nachbau_*` (Entwicklung).
Danach: `messen` auf allen Entwicklungskorpora (vorher/nachher, `--vergleich`), Docs (`docs/status.md`, `docs/plan_auftrag_002.md`) aktualisieren, committen, **kurzen Zwischenbericht an mich in einfachen Worten**, dann Pause bis ich „weiter“ sage.
**D. Etappe 2b:** `src/zeilen.py` (Excel-Semantik beim Einfügen/Löschen: Formeln per `openpyxl.formula.tokenizer`/`translate`, Namen, bedingte Formatierung, Datenprüfung, Filter, Tabellen, Merges, Listen-Summen; Vorlagenzeile; Konstanten kopieren) · Dashboard („PDF sagt / Excel hat / Ergebnis“, Gründe aus `pruefcodes` in einfachem Deutsch, einzelne Automatik abwählbar via `abwahl`, Gruppen-Freigabe) · Ollama-Zweitmeinung (`src/llm.py`: `LLMClient`, `AusClient` Standard, `OllamaClient` nur `urllib` mit `ProxyHandler({})`, `FakeClient` für Tests; Funktionen: Lesevorschlag (Bildmodell), Vorschlag bei Abweichung mit Regel-Wächter, Tipp, Lektor; `messen --llm`; `docs/ollama.md`) · Kalibrierung nur an Entwicklungsdaten · **ungesehener Test** (Seeds ab 2001, Generator v3, 2 neue Familien) genau einmal messen · unabhängiger Prüfer (Regeln 11/16/17/22/27/28/29, keine echten Inhalte/Schlüssel) · Version 0.3.0.
**E. Offene kleine Punkte:** die 4 verbleibenden stillen Fehler der Entwicklung (`docs/status.md`); Aufräumen der Hilfsordner `.claude/worktrees/` (nur nach meinem Okay); `main` aktuell halten (mein Team zieht per ZIP von GitHub, ich kann nicht synchronisieren – siehe `docs/arbeiten_ohne_github_sync.md`).

## 5. Zielwerte (nur erfundene Daten)
Stille Fehler Entwicklung **0** (ex_test ≤ 1, ungesehener Test ≤ 2, Ziel 0) · Automatik richtig ≥ 99,8 % · Vorschläge nach Freigabe ≥ 85 % · Abdeckung ohne Nachfrage (Haupttext) ≥ 90 % (Standardlayout ≥ 97 %) · alle ableitbaren Zellen ≥ 85 % · Lese-Lücke ≤ 1,5 % (alle gemeldet) · Zweiter Lauf ändert nichts 100 % · Regel-27-Verstöße 0 · Original unverändert.
Sei ehrlich: Bei ca. 300 automatischen Zellen heißt „0 stille Fehler“ nur „Fehlerquote vermutlich unter etwa 1 %“. Echte PDFs sind der eigentliche Test; berichte Abweichungen, statt Messungen zu beschönigen.

## 6. Arbeitsweise
- Plane vor jeder größeren Etappe kurz (Dateien, Besitzlisten, Tests, Messung), setze um, miss, berichte.
- Kleine Commits; nach jedem Merge `pytest -q` und `messen --korpus entwicklung --mit-vorschlaegen --vergleich`.
- Red-Team-Konvention: `@pytest.mark.redteam` + `@pytest.mark.xfail(strict=True, raises=AssertionError, reason="RT-nn …; Besitzer …")`; behobene Fälle: Marker entfernen.
- Beginne jetzt mit **A** und sag mir danach in drei Sätzen, was du gesehen hast und was du als Nächstes tust.

---------------------------------------------------------------------------------------------------
