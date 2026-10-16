
# Plan Auftrag 002 – „Wie ein sorgfältiger Mensch“ (Arbeitsgrundlage für die lokale Fortsetzung)

Autor: Nico Lehr · Stand: 16.10.2026 · gehört zu `docs/ziel_v2.md`, ADR 0004 (Gegenprobe, Schema 3) und ADR 0005 (Farblogik, Excel-Markierung)

## 1. Ziel in einem Satz
Das Tool soll PDF-Änderungen so in die CDE-Excel übertragen, wie ein sorgfältiger Mensch es täte: ganze Absätze lesen, in derselben Reihenfolge durch PDF und
Excel gehen, nur Markiertes ändern, das Ergebnis **gegen das PDF** prüfen und **anhalten**, sobald etwas nicht zusammenpasst (lieber „prüfen“ als falsch).

## 2. Zielbild – was Nico gesagt hat (maßgeblich)
- **Farblogik im PDF ist nicht einheitlich:** Neuer Inhalt ist mal **rot**, mal **blau**. **Nur Durchstreichen** bedeutet „entfällt“ (egal welche Farbe).
  Es gibt keine PDFs, in denen Rot ohne Strich „entfällt“ bedeutet. Regel: durchgestrichen ⇒ entfällt; rot oder blau ohne Strich ⇒ neu.
- **In der Excel müssen Änderungen markiert werden:** **rot = neuer Inhalt**, **blau + durchgestrichen = alter Inhalt**. Ersetzte Stelle: Alt (blau, durchgestrichen)
  **direkt vor** Neu (rot) in derselben Zelle, ein Leerzeichen dazwischen. Ganz gestrichene Section: **Zeile bleibt**, Inhalt blau durchgestrichen. Neue Section/Zeile: rot.
- **Neue Zeilen „wie ein Mensch“:** Nachbarzeile gleichen Typs als Vorlage (Formeln mit angepassten Bezügen, immer gleiche Werte übernehmen, individuelle Spalten leer + Hinweis).
- **Datenschutz:** Echte Beispielprojekte darf **nur eine lokale KI** sehen. Die Cloud-Sitzung bekommt nie echte Inhalte, nur den **Steckbrief** (Zahlen/Formen) und Messzahlen.
- **Ollama (12 GB GPU):** nur als Zweitmeinung (Vorschläge/Bremse), nie als Entscheider; jede Funktion einzeln abschaltbar und nur an, wenn sie sich messbar lohnt.

## 3. Warum das Tool vorher falsch lag (Befund, belegt)
1. Nachkontrolle verglich nur mit dem eigenen Plan; „sicher“ schon bei Score 0,90. 2. Satzzeichen erbten die Farbe des Nachbarworts („Anlage. insgesamt.“).
3. Striche über ganze Zeilen, Sonderfarben, Fußnoten und gedrehte Seiten wurden still falsch gelesen. 4. Absätze über Seitenumbrüche wurden zerrissen.
5. Zuordnung unscharf, Dubletten und zitierter Alttext führten zu vielen unnötigen Rückfragen. 6. Formeln/Füllspalten/Zahl-Section-Nummern bei neuen Zeilen fehlten.
7. Teile der Messung waren selbst falsch (11 von 21 „stillen Fehlern“ waren Messartefakte).

## 4. Die fünf Bausteine
| Baustein | Inhalt | Dateien | Stand |
|---|---|---|---|
| **A Gegenprobe** | A1: automatisch nur bei exakt gleichem alten Absatz (Normalform); A3: nach dem Speichern jeden Absatz gegen das PDF prüfen; Absatz-Sperre bei unklaren Signalen | `src/gegenprobe.py`, `src/planner.py`, `src/applier.py` | **fertig** (Etappe 1) |
| **B Absatz-Abgleich** | Änderungen zeichengenau aus Alt/Neu-Läufen, Excel-eigene Zeichen bleiben | `src/abgleich.py` | **fertig** |
| **C Sehen** | Leser-Treue (Trenner, Zeichen-Läufe, Hochzahlen, Absatzgrenzen, Striche, Farbräume, Silbentrennung), zweiter Leser (pdfium), Farb-Inventur | `src/pdf_reader.py`, `src/sicht.py` | **fertig**; **offen:** neue Legende (siehe 5) |
| **D Reißverschluss** | Reihenfolge-Kette PDF ↔ Excel: Dubletten, Seitenumbruch, Anker für neue Absätze | `src/abgleich.py` | **fertig** |
| **E Excel-Zeilen & Markierung** | rot/blau-Markierung (ADR 0005), Zeilen einfügen mit Formeln/Füllspalten/Konstanten, Listen-Summen, gestrichene Section bleibt blau durchgestrichen | `src/markierung.py` (Grundlage fertig), `src/zeilen.py`, `src/applier.py` | **offen** |

## 5. Was als Nächstes kommt (Etappe 2a, dann 2b) – jeweils 2–3 Helfer parallel, strikte Besitzlisten, danach Messung
**Vorab (Orchestrator):** Steckbrief-Ergebnisse von Nico auswerten → ADR 0005 mit echten Werten (Hex-Farben, Trenner, Darstellung neuer/gestrichener Zeilen) bestätigen/korrigieren.
**Etappe 2a**
| Helfer | Besitz | Auftrag |
|---|---|---|
| **LEGENDE** | `src/pdf_reader.py` (Status/Bedeutung), `config/farben.yaml` (`legende`), `src/planner.py` (Sperrliste), Tests | Strich ⇒ entfällt (jede Farbe); rot **oder** blau ohne Strich ⇒ neu; `gemischt_blau_strich` ist normal (aus der Sperrliste); Dokument-Legende im Ergebnis; „steht schon da“-Schutz für farbigen Text, der exakt an der Stelle steht; neue Red-Team-Fälle |
| **MARKIERUNG** | `src/applier.py`, `src/gegenprobe.py` (additiv), `src/excel_model.py`, Tests | Applier schreibt Rich-Text-Läufe (ersetzt = Alt blau+durchgestrichen + Leerzeichen + Neu rot; eingefügt = rot; gelöscht = blau+durchgestrichen bleibt; neue Zeile rot; gestrichene Section bleibt, Inhalt blau+durchgestrichen); Basisschrift erhalten; **A3 markiert:** `alt_ansicht` = Original (nichts geht verloren), `neu_ansicht` = PDF-Neu; Excel-Modell matcht auf `neu_ansicht`, erkennt markierte Stellen (Idempotenz) |
| **GEN v3 / NACHBAU** | `tests/fixtures/*`, `scripts/erzeuge_paare.py`, `src/messen.py` | Generator v3: Legende je Dokument (blau-neu / rot-neu / gemischt, Strich in schwarz/rot/blau), **nachher.xlsx markiert**; Messung „Format-Treue“ je Lauf; **Nachbau-Generator aus dem Steckbrief** (`--steckbrief`) |
**Etappe 2b:** ZEILEN (`src/zeilen.py`: Excel-Semantik beim Einfügen/Löschen für Formeln, Namen, bedingte Formatierung, Datenprüfung, Filter, Tabellen, Merges; Vorlagenzeile; Listen-Summen) ·
UI (Ansicht „PDF sagt / Excel hat / Ergebnis“, Gründe aus `pruefcodes` in einfachem Deutsch, einzelne Automatik **abwählbar**, Gruppen-Freigabe) ·
LLM (`src/llm.py`, `LLMClient` mit `AusClient` Standard / `OllamaClient` nur `urllib` / `FakeClient`; KI-Lesevorschlag, KI-Vorschlag bei Abweichung, KI-Tipp, KI-Lektor; `messen --llm` bei Nico) ·
Kalibrierung nur an Entwicklungsdaten · **ungesehener Test** (Seeds ab 2001, Generator v3, 2 neue Familien) **genau einmal** messen · unabhängiger Prüfer · Version 0.3.0.

## 6. Messung (nur erfundene Daten; Regel 11)
Korpora: `entwicklung` (Seeds 1–60 und 101–160, gerade ab 102 mit LibreOffice), `ex_test` (1001–1030, früherer Test, jetzt offen Entwicklung), später `nachbau_*` (aus Nicos Steckbrief) und der ungesehene Test.
**Stand Etappe 1 (16.10.2026):** Entwicklung 120 Paare: 4 stille Fehler, 98,1 % automatisch richtig (99,2 % je Änderung), Vorschläge nach Freigabe 87,1 %, Abdeckung ohne Nachfrage (Haupttext) 87,4 %,
alle Zellen 57,8 %, Zweiter-Lauf-Test 120/120, Lese-Lücke 2,2 %. `ex_test`: 0 (nach Messkorrektur; gemeldete Farb-Lücken zählen nicht als „still“), 95,8 % richtig, Vorschläge 48,3 %, Lese-Lücke 7,0 %.
**Zielwerte** und Ehrlichkeitshinweise: `docs/ziel_v2.md`. Echte Dokumente sind der eigentliche Test.

## 7. Arbeitsregeln (gelten auch lokal)
`CLAUDE.md` und `docs/regeln/entwicklungsregeln.md` lesen. Nur erfundene Daten in Git (Regel 16), keine API-Keys und keine externen Dienste mit Kundendaten (Regel 17), LLM nur über `LLMClient` mit `temperature 0` (Regel 12/22),
Unsicheres nie automatisch (Regel 27), keine hartcodierten Pfade (Regel 28), Läufe append-only (Regel 29), Schemaänderung ⇒ `SCHEMA_VERSION` und ADR (Regel 24), nie `git add -A`.
Sonderregel Etappe 2: Stille Fehler dürfen von Merge zu Merge nie steigen. Die Konfigurationsweiche `markierung.aenderungen: keine` schützt die alten Tests, Tests mit `@pytest.mark.markierung_an` prüfen den neuen Standard.
