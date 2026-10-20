# ADR 0008: Vertrauensmodus (Stichprobe, Dokument-Ampel, Freigabe-Protokoll) und Eingabe Word oder PDF

Autor: Nico Lehr · Stand: 20.10.2026 · Auftrag „Vertrauen“

## Entscheidung

Nico lernt dem Tool schrittweise zu vertrauen. Das Tool bleibt dabei vorsichtig (Regel 27): Nichts davon ändert, **was** eingetragen wird.
Neu sind Nachfragen, eine Zusammenfassung und eine Auswertung. Kein Schema-Wechsel: Es kommen nur **Zusatzfelder** dazu (alte Läufe bleiben lesbar).

1. **Eingabe Word oder PDF.** Das Dokument mit Änderungen darf eine `.docx` (bevorzugt) oder ein PDF sein (Dashboard, API, CLI `--docx` / `--pdf`).
   Geprüft werden Endung **und** Inhalt (PDF: `%PDF-`; Word: ZIP mit `[Content_Types].xml` und `word/document.xml`), Meldungen in einfachem Deutsch.
   Eine zentrale Funktion `runs.lade_analyse(pfad)` wählt nach Endung `pdf_reader.analysiere_pdf` oder `docx_reader.analysiere_docx` (Import erst beim Aufruf).
   Eine Analyse aus Word hat `legende_modus "docx"`, Seite 0, Bounding-Box (0,0,0,0): Das Dashboard zeigt dann nur den Text (Ergebnis wie in der Excel),
   keinen PDF-Ausschnitt. Der Feldname `pdf_name` bleibt; neu ist `quelle_art` (`"docx"` | `"pdf"`) in `lauf.json` und `plan.json`.
2. **Stichprobe.** Aus den automatisch eingetragenen Änderungen zieht die Analyse eine feste Stichprobe (`stichprobe_anteil` 10 %, mindestens `stichprobe_min` 3,
   höchstens alle). Der Zufall ist aus der Lauf-ID abgeleitet (SHA-256, nicht `hash()`): gleiche Lauf-ID und Änderungen ergeben immer dieselbe Auswahl.
   Gespeichert als Liste von Änderungs-IDs im Feld `stichprobe` von `plan.json`. Im Dashboard steht sie zuerst, einklappbar, nur wenn vorhanden:
   „Passt“ zählt als Bestätigung, „Falsch“ wählt die Änderung ab (wie „Trotzdem nicht eintragen“) und wird als Ablehnung vermerkt. Es blockiert nichts.
3. **Dokument-Ampel** (`src/vertrauen.py`, Schwellen in `config/vertrauen.yaml`). Rot: Bilanz stimmt nicht, Lese-Lücken, Spec unklar oder mehr als
   `rot_offene_je_100` offene Punkte je 100 Änderungen. Gelb: offene Punkte (unentschiedene Vorschläge, nicht eingetragene Änderungen, Codes mit Anfang
   `konsens_`) oder Stichprobe nicht angesehen, oder unbekannt, ob alles gelesen wurde (älterer Lauf). Sonst grün. Das Dashboard rechnet nichts nach: Es fragt
   `GET /api/lauf/<id>/ampel?offen=..&stichprobe_offen=..`; so wird die Ampel beim Entscheiden besser, Bilanz-/Lücken-Rot bleibt.
4. **Freigabe-Protokoll** `entscheidungen.jsonl` im Laufordner, nur angehängt (Regel 29), eine Zeile je Entscheidung beim Erstellen einer Excel:
   `zeit, run_id, ergebnis, id, art, entscheidung, codes`. `art`: `automatisch | stichprobe | pruefen | vorschlag` (Vorschlag, der nur pauschal über
   „Alle übernehmen“ / `--alle-vorschlaege` freigegeben wurde). `entscheidung`: `uebernommen | bestaetigt | abgelehnt | abgewaehlt`.
   **Nie Text aus Dokumenten oder Zellen** (Regel 16): nur Kennungen; Codes, die nicht wie Kennungen aussehen, werden zu `unbekannt`.
   Unentschiedene Vorschläge und nicht eingetragene Änderungen sind keine Entscheidung und stehen nicht darin. Ein Schreibfehler lässt die Excel bestehen und warnt.
   `ergebnis.json` bekommt die Zusatzfelder `abgelehnt`, `pauschal`, `stichprobe {passt, falsch}` (damit ein erneutes Öffnen die Klicks wieder vorbelegt).
5. **Auswertung** `python -m src.cli vertrauen`: liest alle `entscheidungen.jsonl` (nur lesen), je Lauf zählt nur die zuletzt erstellte Excel. Je Prüfcode:
   Anzahl, angesehen (nur Entscheidungen mit eigenem Klick), Freigabe- und Ablehnquote, Empfehlung in einem Satz. „Sicher genug zum Automatisieren“ erst ab
   `min_angesehen` (20) angesehenen Entscheidungen, `freigabe_ab` (98 %) und höchstens `ablehnung_max` (2 %). Das ist nur ein Hinweis: Das Tool stellt nie von selbst um.

## Folgen und Grenzen

- Die Klicks landen im Protokoll erst, wenn eine Excel erstellt wird (wie die übrigen Entscheidungen im Dashboard). Wer nur klickt und dann abbricht, hinterlässt nichts.
- Die Ampel-Schwelle `rot_offene_je_100: 40` ist ein Startwert (gemessen: typisch rund 30 offene Punkte je 100 Änderungen); sie gehört an echten Paaren eingestellt.
- Das Word-Lesen steht in ADR 0009 (`src/docx_reader.py`). Die Tests in `tests/test_vertrauen.py` benutzen einen Fake; ein Test mit dem echten Leser steht in `tests/test_word_dashboard.py`.
- Ältere Läufe ohne `stichprobe`/`quelle_art` bleiben lesbar (keine Stichprobe, PDF).
