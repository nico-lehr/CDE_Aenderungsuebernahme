# Ziel v1 – CDE-Änderungsübernahme

Autor: Nico Lehr · Stand: 15.10.2026 · Status: Schritt 0 aus `docs/auftrag/cc_auftrag_001.md` (festgelegt vor dem ersten Code)

Dieses Dokument legt fest, **was „gut“ heißt**, woran wir es **messen**, **in welcher Reihenfolge** wir arbeiten und **was wir bewusst nicht tun**.
Alle Zahlen gelten für einen **erfundenen (synthetischen) Korpus**. Sie sind ein Zwischenziel, **kein Versprechen für echte Dokumente**.

## (a) Zielbild: Was muss ein Team glauben können?

Ein Team, das dem Tool traut, braucht drei Dinge:

1. **Es sagt nie „fertig“, wenn etwas fehlt oder falsch ist.** Das Schlimmste ist ein stiller Fehler: Das Protokoll sagt „übernommen, ok“, aber die Excel ist
   falsch oder unverändert. Das ist schlimmer als jede Nachfrage.
2. **Es sagt ehrlich, was es nicht verstanden hat.** Scan-Seiten, Zwei-Spalten-Text, Tabellen, unbekannte Farben, ungewöhnliche Excel-Aufbauten
   erzeugen eine sichtbare Warnung. Die Bilanz zeigt dann „unvollständig“ und nie ein grünes „stimmt“.
3. **Es fasst nur an, was es sicher weiß.** Unsichere Treffer bleiben auf „prüfen“ (Regel 27). Die Original-Excel bleibt unverändert (Regel 29, Hash-Prüfung).
   Läuft das Tool zweimal auf demselben Ergebnis, passiert beim zweiten Mal nichts.

Erst danach zählt Bequemlichkeit (weniger Nachfragen, schönere Oberfläche).

## (b) Messbare Erfolgskriterien

Gemessen mit `python -m src.cli messen` (Paket 2) an Paaren aus **PDF + Excel vorher + Excel nachher**. Das Soll kommt aus dem Vergleich vorher↔nachher,
**nie** aus dem PDF-Leser. Berichte enthalten nur Zahlen, Positionen und Fehlerarten, **keine Textinhalte** (Regel 16).

**Korpus (erfunden, vom Generator erzeugt, Version in `KORPUS.json` festgehalten):**
- Entwicklung: mindestens 60 Paare und 400 Soll-Zellen. Hier darf geprüft und eingestellt werden.
- Test: mindestens 30 Paare und 200 Soll-Zellen, **eigene Seeds** und **2 Störer-Familien, die nur dort vorkommen**. Wird höchstens zweimal gemessen
  (nach Welle 2 und zum Abschluss), nie zum Einstellen benutzt (Regel 11).
- Mindestens 12 verschiedene Störer-Familien (Fußnoten, Silbentrennung, Spalten, Tabellen/Bilder, Scan, Farbtöne/Farbräume, Mischmarkierungen,
  Aufzählungen, Seitenumbruch im Absatz, doppelte Absätze, irreführende Excel-Spalten, mehrere Blätter, Formeln/verbundene Zellen, Sonderzeichen).

| Nr | Kriterium | Zielwert | Bedeutung in einfachen Worten |
|---|---|---|---|
| 1 | **Bilanz stimmig** | 100 % der Paare. Scan/Zwei-Spalten/Tabelle wird als „unvollständig“ gemeldet, nie grün | Die Wörter-Summe passt, und was das Tool nicht lesen konnte, steht sichtbar da |
| 2 | **Stille Fehler** | **0** auf Entwicklung und Test | Nichts wird als „übernommen“ gemeldet, was in der Excel falsch oder gar nicht steht |
| 3 | **Treffergenauigkeit** | Automatisch übernommen: 100 %. Nach Freigabe von „prüfen“: ≥ 95 % | Was ohne Nachfrage geht, ist richtig |
| 4 | **Abdeckung ohne Nachfrage** | ≥ 90 % bei Standard-Aufbauten, ≥ 70 % insgesamt. 100 % der Nicht-Übernahmen haben einen Grund | So wenig Rückfragen wie möglich, aber nie auf Kosten von Nr. 2/3 |
| 5 | **Lese-Lücke** | ≤ 2 % der Soll-Wörter, der Rest als Warnung sichtbar | Was im PDF steht, aber nicht erkannt wurde, bleibt nicht unbemerkt |
| 6 | **Idempotenz, Kollateralschaden, Original** | 100 % / 0 / Hash unverändert | Zweiter Lauf ändert nichts; keine Zelle ohne Grund verändert; Original unberührt |
| 7 | **Geschwindigkeit** | Excel mit 20 000 Zeilen < 60 s; eine formatierte Einzelzelle weit unten ≤ +2 s | Auch große Dateien laufen durch |
| 8 | **Tests** | Alle grün. Erwartete Fehler (xfail) der Klasse „stiller Fehler“ = 0. Jeder andere xfail steht in `docs/status.md` | Keine versteckte Rotphase |

Fehlerarten (damit klar ist, *wo* etwas schiefgeht): PDF falsch gelesen → Layout falsch erkannt → Zuordnung falsch → Plan falsch → Bilanz kaputt → unklar
(erster Treffer gewinnt).

**Ehrliche Grenze der Zahl 0:** Bei 200 Test-Zellen bedeutet „0 stille Fehler“ nur, dass die Fehlerrate mit etwa 95 % Sicherheit unter rund 1,5 % liegt
(Dreierregel). Echte Dokumente sind anders als unsere erfundenen. Erst die Messung an Nicos echten Paaren zeigt, wie viel davon trägt.

## (c) Prioritätenliste (Risiko × Nutzen)

Regel: **zuerst, was „übernommen“ sagt und falsch liegt; dann, was Änderungen unsichtbar verliert; zuletzt Komfort.** Belege stammen aus der Code-Prüfung
vom 15.10.2026 (Dateinamen und Zeilen in `docs/status.md`).

| Rang | Paket | Warum an dieser Stelle |
|---|---|---|
| 0 | **Fundament:** `.gitignore` (fehlte komplett), `.venv` mit Python 3.11, Baseline-Tests, dieses Dokument | Ohne `.gitignore` könnten echte PDFs/Excel versehentlich in Git landen (Regel 16/30). Ohne Baseline keine Vergleichbarkeit |
| 1 | **Stille Fehler beim Planen/Anwenden:** Folgeabsätze neuer Sections als „übernommen“ ohne Eintrag; doppeltes Einfügen im zweiten Lauf; Nachkontrolle zu schwach; Zelle leeren prüft nur Anfang/Ende; Formel-/Steuerzeichen/Ligaturen im Neutext | Hier kann die Excel falsch werden und das Protokoll sagt „ok“. Höchstes Risiko |
| 2 | **Messwerkzeug + Generator** (Pakete 2 und 1) | Ohne Messung ist jeder Fix ein Gefühl. Läuft sofort auch mit echten Paaren |
| 3 | **Zuordnung:** doppelte Absätze, Anfangs-Treffer, Section-Nummern als Zahl, falsches Spalten-Layout, mehrere Blätter, Performance | Falsche Stelle getroffen = stiller Fehler im Kern |
| 4 | **PDF lesen – stille Verluste:** verschmolzene Ersetzung, blauer Tabellenkopf/Link, Blau+Strich, Farbbilanz, „unvollständig“-Meldung | Hier gehen Änderungen verloren oder werden erfunden, bevor die Bilanz sie sieht |
| 5 | **PDF lesen – Qualität:** Fußnoten, Silbentrennung, Seitenumbruch, Aufzählungen, Kopf-/Fußzeile, Schwellen in die Config | Weniger Rückfragen, weniger Fehlzuordnungen |
| 6 | **Excel-Sicherheit** (Paket 5): Formeln, Tabellen, Validierung, verbundene Zellen, Bilder erkennen und warnen | Beschädigte Excel ist teuer; Reparieren ist riskant, Warnen + „prüfen“ ist billig und sicher |
| 7 | **Läufe + API-Härtung:** atomarer Lauf-Zähler, atomares Schreiben, Upload-Limit, Host-Prüfung, einfache Fehlertexte, „Archivieren“ statt Löschen | Das Tool ist für ein anderes Team gedacht; es darf nicht an Kleinigkeiten kippen |
| 8 | **Layout-Profile + Kalibrierung** (Pakete 3 und 7) | Brauchbar erst, wenn echte Paare zeigen, wo die Automatik scheitert; Vorbereitung jetzt, Ausbau später |
| 9 | **Dashboard-Qualität** (Paket 6) | Wichtig für Akzeptanz, aber nachrangig zu Richtigkeit |
| 10 | **Doku** (Paket 8) | Läuft nebenher, Abschluss zuletzt |

## (d) Bewusst NICHT

- **Keine echten Dokumente, keine Kundendaten, keine API-Keys, keine externen Dienste** (Regel 16/17). Kein LLM im Kern; später höchstens als Prüfer über `LLMClient` (Regel 22).
- **Keine neuen Pakete** (Regel 10). Zufallstests mit `random.Random(seed)`, kein Hypothesis, kein reportlab.
- **Nicht lösen, nur erkennen und melden:** Zwei-Spalten-Text, Tabellen und Bilder im PDF, Scan-Seiten (kein OCR).
- **Nicht reparieren, nur warnen und auf „prüfen“ setzen:** Formeln, bedingte Formatierung, Datenvalidierung, Filter, Excel-Tabellen, benannte Bereiche, verbundene Zellen an der Einfügestelle.
- **Teilformatierung in Zellen** (einzelne fette Wörter) nicht erhalten; Warnung.
- **Layout im Dashboard bearbeiten** und **PDF-Ausschnitt neben Excel vorher/nachher**: nach v1. Jetzt nur Anzeige des erkannten Aufbaus und Vorher-/Nachher-Text.
  Neue Layout-Profile gehen nie direkt live: erst Vorschau, dann Bestätigung durch einen Menschen, dann aktiv.
- **Keine Umbenennungen/Umbauten** ohne Not; bestehende Schnittstellen bleiben (nur Felder mit Standardwert kommen dazu, `SCHEMA_VERSION` 1 → 2, ADR 0002).
- **Kein Tuning an echten Beispielen** (Regel 11): Entwicklung und Test bleiben getrennt.
- **`bess-pipeline` wird nicht angefasst**, Port bleibt 5100.
- **Kein Löschen von Läufen mehr** (Regel 29): „Archivieren“ statt Löschen.

## Entscheidungen auf einen Blick

| Frage | Entscheidung | Grund |
|---|---|---|
| Unsicher oder gefährlich für die Excel-Struktur? | Auf „prüfen“ + Warnung, nicht automatisch | Lieber nicht übernehmen als falsch (Regel 27) |
| Doppelter oder nur anfangs passender Absatz? | Höchstens „mittel“ | Zuordnung nicht eindeutig |
| Schema ändern? | Ja, Version 2, nur Felder mit Standardwert, alte Läufe bleiben lesbar | Regel 24 |
| Schwellen im Code? | In `config/*.yaml`, Verhalten vorher/nachher durch Test identisch | Regel 28 |
| Wie Red-Team-Tests ohne Dauer-Rot? | `xfail(strict=True)` mit Grund; Reparatur macht den Test rot, bis der Marker fällt | Standardlauf bleibt grün, nichts wird vergessen |
