# Ziel v2 – „Wie ein sorgfältiger Mensch“

Autor: Nico Lehr · Stand: 16.10.2026 · Auftrag 002 · Vorgänger: `docs/ziel_v1.md` (Ziel „0 stille Fehler“ dort nicht erreicht)

## Korrektur des Zielbilds am 16.10.2026 (Nico)
- **Farben im PDF uneinheitlich:** neu = mal rot, mal blau; **nur Durchstreichen** = entfällt (jede Farbe). Rot ohne Strich bedeutet nie „entfällt“.
- **Excel-Ergebnis markiert:** rot = neuer Inhalt, blau + durchgestrichen = alter Inhalt (Alt direkt vor Neu); gestrichene Section bleibt als Zeile stehen (blau durchgestrichen). ADR 0005.
- Echte Beispiele sieht nur eine lokale KI; die Cloud-Sitzung arbeitet mit Steckbrief (Zahlen/Formen) und erfundenen Nachbauten.

## Zielbild
Für einen Menschen ist die Aufgabe einfach: markierten Absatz lesen, die Stelle in der Excel finden, alt durch neu ersetzen, kurz kontrollieren.
Das Tool soll genau so arbeiten und **anhalten, sobald etwas nicht zusammenpasst**:

| Ein sorgfältiger Mensch … | Baustein im Tool |
|---|---|
| prüft, ob in der Excel wirklich der alte Text steht, und kontrolliert das Ergebnis am PDF | **Gegenprobe** A1/A3 + Absatz-Sperre (ADR 0004) |
| liest den ganzen Absatz, Zeichen für Zeichen (auch Satzzeichen) | **Absatz-Abgleich** (`src/abgleich.py`) |
| schaut genau hin (Strich über die ganze Zeile, Farbe, Fußnoten, Silbentrennung) | **Sehen**: genauerer Leser, zweiter Leser, Farb-Inventur |
| geht PDF und Excel parallel von oben nach unten durch | **Reißverschluss** (Reihenfolge-Kette) |
| legt neue Zeilen wie die Nachbarzeile an, löscht gestrichene Zeilen sauber | **Excel-Zeilen richtig** (Etappe 2) |

Ollama (lokal, 12 GB GPU) ist ein **optionales zweites Paar Augen**: Vorschläge und Bremse, nie Entscheider (Regel 22/27).

## Messbare Kriterien (nur erfundene Daten; Regel 11)
Korpora: **Entwicklung** = v1 Seeds 1–60, **ex_test** = früherer Testkorpus 1001–1030 (offen als Entwicklung erklärt, nicht mehr ungesehen),
v2 Seeds 101–160 (Hälfte mit LibreOffice erzeugt). **Test** = v2 Seeds 2001–2040 mit 2 neuen, nur dort vorkommenden Störer-Familien, genau einmal gemessen.

| Kennzahl | Entw. v1 | ex_test | v2 LibreOffice | ungesehener Test v2 |
|---|---|---|---|---|
| Stille Fehler | **0** | ≤ 1 | 0 | ≤ 2 (Ziel 0) |
| Treffergenauigkeit automatisch | ≥ 99,8 % | ≥ 99,5 % | ≥ 99,5 % | ≥ 99 % |
| Vorschläge exakt nach Freigabe | ≥ 85 % | – | – | – |
| Abdeckung ohne Nachfrage (Haupt-Spalten) | ≥ 90 % (Standard ≥ 97 %) | ≥ 82 % | ≥ 85 % | ≥ 78 % |
| Abdeckung ableitbare Zellen / alle Zellen (Etappe 2) | ≥ 85 % / ≥ 80 % | – | – | – |
| Lese-Lücke (alle gemeldet) | ≤ 1,5 % | ≤ 3 % | ≤ 2 % | ≤ 3 % |
| Idempotenz / Regel-27-Verstöße / Kollateral | 60 von 60 / 0 / 0 | – / 0 / 0 | – / 0 / 0 | – / 0 / 0 |
| Restarbeit je 100 Änderungen (neu) | berichten, Richtwert ≤ 12 | berichten | berichten | berichten |

Ehrlich: Bei ~300 automatischen Zellen heißt „0 stille Fehler“ nur „Fehlerquote vermutlich unter ~1 %“. Echte PDFs bleiben der eigentliche Test.

## Prioritäten (Risiko × Nutzen)
1. Absatz-Sperre und Gegenprobe A1/A3 (kleiner Aufwand, größte Wirkung gegen stille Fehler).
2. Leser-Treue: Satzzeichen behalten ihre eigene Farbe, Strich über ganze Zeile, Hochzahlen, Absatzgrenzen, Seitenumbruch, Silbentrennung.
3. Absatz-Abgleich und Reißverschluss (holt die Abdeckung zurück, die die strengeren Regeln kosten; Dubletten, Seitenumbruch).
4. Zweiter Leser (pdfium) und Farb-Inventur (findet Markierungen, die der Textleser übersieht).
5. Etappe 2: Excel-Zeilen richtig (Formeln, Füllspalten, Zeile löschen), Dashboard (Abwahl, „PDF sagt / Excel hat / Ergebnis“), Ollama, ungesehene Endmessung.

## Bewusst NICHT
- Keine KI als Entscheider; die KI schreibt nie selbst Text in die Excel.
- Kein OCR im Kern; Scans/Tabellen/zwei Spalten bleiben „erkennen und melden“ (optional KI-Lesevorschlag, nur Vorschlag).
- Keine neuen Python-Pakete (pip im Firmennetz gesperrt); LibreOffice/Chromium nur im Test-Generator.
- Keine echten Dokumente; Port 5100; `bess-pipeline` unberührt.
