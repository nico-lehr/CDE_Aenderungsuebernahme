# Messung Baseline (Ausgangszustand, vor allen Reparaturen)

Autor: Nico Lehr · Stand: 15.10.2026 · nur Zahlen, keine Dokumentinhalte (Regel 16)

Gemessen mit `python -m src.cli messen --korpus entwicklung --mit-vorschlaegen` auf dem **erfundenen** Entwicklungskorpus
(Generator-Version 1, Seeds 1–60, 60 Paare, 755 Soll-Zellen). Kern (pdf_reader, excel_model, matcher, planner, applier) noch unverändert
gegenüber `baseline-v0.1`; nur Schema 2 (leere Felder), Messwerkzeug und API sind dazugekommen. Der Testkorpus wurde noch **nicht** gemessen (Regel 11).

| Nr | Kriterium | Ziel | Baseline | Ergebnis |
|---|---|---|---|---|
| 1 | Bilanz stimmig / vollständig | alle Paare, Lücken gemeldet | 60 von 60 stimmig, aber 17 erwartete Lücken nicht gemeldet | NICHT ok |
| 2 | **Stille Fehler** | 0 | **64** (47 „übernommen, aber falsch“ und Nachkontrolle „ok“, 15 nicht gemeldete Lücken, 2 Kollateral) | NICHT ok |
| 3 | Treffergenauigkeit automatisch | 100 % | 92,3 % (589 von 638) | NICHT ok |
| 3 | Genauigkeit Vorschläge nach Freigabe | ≥ 95 % | 60,9 % (28 von 46) | NICHT ok |
| 4 | Abdeckung ohne Nachfrage | ≥ 70 % gesamt | 77,1 % (aber nur, weil auch falsche Treffer mitzählen) | ok |
| 4 | Nicht-Übernahmen ohne Grund | 0 | 0 | ok |
| 5 | Lese-Lücke | ≤ 2 % | 5,8 % (413 von 7155 Soll-Wörtern) | NICHT ok |
| 6 | Zweiter Lauf ändert nichts | alle Paare | 0 von 60 | NICHT ok |
| 6 | Kollateralschaden | 0 | 2 | NICHT ok |
| 6 | Original unverändert | alle | 60 von 60 | ok |

Fehlerarten (Zählung der verfehlten Stellen): pdf_falsch_gelesen 99 · layout_falsch 34 · plan_falsch 20 · zuordnung_falsch 16 · bilanz_kaputt 15 · unklar 0.

Lesehilfe: Die Störer-Familien überlappen (ein Paar hat mehrere), deshalb sind Zahlen je Familie nicht addierbar. Am meisten stille Fehler in Paaren mit:
Kopf-/Fußzeile (24), CMYK/Grau (23), Bild (23), Farbtöne (21), Lfd-Spalte (19), Silbentrennung / nummerierte Listen / Sonderzeichen (je 18), Zwei-Spalten (17).

Nachmessen: `python -m src.cli messen --korpus entwicklung --mit-vorschlaegen`; Korpus neu erzeugen: `python scripts/erzeuge_paare.py --korpus entwicklung --seeds 1-60`.
