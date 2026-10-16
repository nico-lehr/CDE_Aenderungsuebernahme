# ADR 0004 – Gegenprobe gegen das PDF, Absatz-Sperre, Schema 3

Autor: Nico Lehr · Datum: 16.10.2026 · Status: angenommen (Auftrag 002, Etappe 1)

## Problem
Die Nachkontrolle nach dem Speichern verglich die Excel nur mit dem **eigenen Plan** des Tools. Hatte das Tool das PDF falsch gelesen
oder falsch zugeordnet, „bestätigte“ es seinen eigenen Fehler (stiller Fehler). Die Zuordnung war unscharf: Bis zu 10 % der alten
PDF-Wörter durften in der Excel fehlen, und die Änderung galt trotzdem als sicher.

## Entscheidung
1. **A1 – vorher exakt:** Eine Änderung wird nur automatisch eingetragen, wenn der alte Absatz aus dem PDF (bleibt + entfällt)
   in der „Normalform“ genau dem Text an der Excel-Stelle entspricht. Sonst „prüfen“ mit Grund (`pruefcodes`).
2. **A3 – nachher gegen das PDF:** Nach dem Speichern muss jeder betroffene Absatz in der Excel genau dem Zustand entsprechen, den das PDF
   für die angewendeten Änderungen verlangt (`gegenprobe.zustand`), und der Text außerhalb muss unverändert sein. Sonst „prüfen“
   (Wirkung „abweichung“, Code `gegenprobe_*`).
3. **Absatz-Sperre:** Ein unerklärtes Signal irgendwo im Absatz (unbekannte Farbe, halbe Markierung, unklare Linie, Tabelle, zwei Spalten,
   unklare Trennung, unerklärte Farbe im Seitenbild …) stoppt die Automatik für den ganzen Absatz. Geschlossene Liste in
   `config/zuordnung.yaml` (`planer.unsichere_hinweise`).
4. **Normalform („streng, aber fair“):** Groß/Klein und Satzzeichen zählen; Leerraum-Varianten, Anführungs-/Strich-Varianten,
   weiches Trennzeichen, Hochzahlen und ein Aufzählungszeichen am Anfang zählen nicht. Tabellen in `config/zuordnung.yaml` (`normalform:`).
5. **Schema 3** (`SCHEMA_VERSION = 3`, lesbar: 1, 2, 3). Nur neue Felder mit Standardwert:
   - plan.json `absaetze`: je PDF-Absatz `{id, blatt, zellen: [[zeile, spalte, start, ende]], laeufe: [[art, text, aenderung, optional]]}`
     (art: k bleibt, d entfällt, a neu, s Trenner). Damit kann A3 auch später im Dashboard ohne PDF prüfen.
   - `Aenderung`: `gruppe`, `pruefcodes`, `soll_absatz` („PDF sagt“), `ist_absatz` („Excel hat“), `zweitmeinung` (optionale KI, nur Hinweis).
   - `Operation`: `formeln`, `leer_spalten` (Etappe 2, neue Zeilen).
   - Lücken-Art `unerklaerte_markierung`.
   - Nur intern (nicht gespeichert): `PdfToken.trenner/teile/trennstelle`, `PdfBlock.fortsetzung`.
6. **Invariante:** Eine automatische Operation ohne gespeicherten Absatz wird bei Schema-3-Plänen nie automatisch angewendet.
   Pläne mit Schema 1/2 behalten die alte Nachkontrolle (sie haben keine Absätze).
7. **Blatt-Aufbau-Sperre gelockert:** Reine Text-Änderungen in einer sichtbaren Zelle, deren alter Absatz exakt und eindeutig
   (auch über die Reihenfolge) gefunden wurde, brauchen keinen „sicher erkannten“ Blatt-Aufbau mehr. Für neue Zeilen/Sections gilt sie weiter.

## Folgen
- Weniger stille Fehler: Lese- und Zuordnungsfehler werden zu „prüfen“ statt still eingetragen.
- Abdeckung kann bei von Hand bearbeiteten Excel-Dateien sinken (Text weicht ab). Gegenmittel: Gründe zählen (`messen`), Normalform-Regeln
  in der Konfiguration, Kalibrierung nur an Entwicklungsbeispielen (Regel 11).
- `soffice`/Chromium werden nur vom Test-Generator benutzt, nie vom Tool zur Laufzeit.
