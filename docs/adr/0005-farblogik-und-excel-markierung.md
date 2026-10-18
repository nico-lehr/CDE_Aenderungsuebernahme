# ADR 0005 – Farblogik im PDF und farbige Markierung der Änderungen in der Excel

Autor: Nico Lehr · Datum: 16.10.2026 · Status: angenommen (Auftrag 002b, Antworten von Nico vom 16.10.2026)

## Neue Fakten (von Nico)
1. Die Farben in den PDFs sind **nicht einheitlich**: Neuer Inhalt ist mal rot, mal blau. **Nur das Durchstreichen** zeigt zuverlässig, dass Inhalt entfällt.
   Es gibt **keine** PDFs, in denen Rot ohne Durchstreichung „entfällt“ bedeutet.
2. In der **CDE-Excel** müssen die Änderungen **markiert** werden: **rot = neuer Inhalt**, **blau + durchgestrichen = alter Inhalt**.
   Ersetzte Stelle: Alt direkt vor Neu in derselben Zelle. Eine ganz gestrichene Section behält ihre Zeile (Inhalt blau + durchgestrichen).
3. Echte Beispiele darf nur eine lokale KI sehen. Wir arbeiten mit einem **Steckbrief** (Zahlen und Formen, kein Text) und einem Nachbau-Generator.

## Entscheidung
**PDF lesen (Legende):** Durchgestrichen ⇒ entfällt (jede Farbe). Rot **oder** blau ohne Strich ⇒ neu. Blau + Strich ist damit normal (alt) und kein
Widerspruch mehr (`gemischt_blau_strich` verschwindet aus der Sperrliste). Andere Farben bleiben „unbekannt“ und sperren den Absatz.
**Excel schreiben:** `markierung.aenderungen: farbe` (Standard; `keine` = frühere Reinigungs-Logik, nur für Regressionstests und Sonderfälle).
- ersetzt ⇒ `Alt (blau, durchgestrichen)` + ein Leerzeichen + `Neu (rot)` in derselben Zelle,
- eingefügt ⇒ rot, gelöscht ⇒ blau + durchgestrichen **bleibt stehen**,
- neue Zeile/Section/Absatz ⇒ rot, gestrichene Section ⇒ Zeile bleibt, Nummer/Titel/Inhalt blau + durchgestrichen (`geloeschte_section: markieren`,
  Alternative `zeile_entfernen`),
- Zahlen/Formeln in markierten Zeilen: Schrift färben, Wert bleibt.
**Rich Text:** Eine gemischte Zelle ist ein Rich-Text-Wert; jedes Stück kopiert die Schrift des Originals (Name, Größe, fett …), sonst fiele es auf die Standardschrift.
Ein Spike (openpyxl 3.1.5) zeigt: Rundreise über Datei bleibt verlustfrei, Zahlenformat/Füllung/Ausrichtung bleiben, LibreOffice stellt Blau, Rot und Durchstreichung richtig dar
(`tests/test_markierung.py::test_libreoffice_zeigt_rot_blau_und_strich`).
**Zwei Ansichten** (neu in `src/markierung.py`): `alt_ansicht` (normal + alt) muss der Original-Absatz sein – nichts geht verloren; `neu_ansicht` (normal + neu) muss der neue
Absatz laut PDF sein. Die Gegenprobe A3 wird dafür erweitert (`pruefe_absatz_markiert`), die Idempotenz stützt sich auf die erkannte Markierung („steht schon so da“).
**Lesen vorhandener Zellen:** Matching und A1 laufen auf der `neu_ansicht` (Alt-Stücke zählen nicht mit). Bedeutung eines Stücks: durchgestrichen = alt (jede Farbe),
rote Schrift ohne Strich = neu, sonst normal.

## Folgen
- Die frühere Entscheidung „gestrichene Section ⇒ Zeile löschen“ ist ersetzt (nur noch Option `zeile_entfernen`).
- Generator, Messung (Format-Treue je Lauf), Gegenprobe, Excel-Modell und Dashboard müssen mit markierten Zellen umgehen (Auftrag 002b, Etappe 2a/2b).
- Alle älteren Tests laufen mit `markierung.aenderungen: keine` (conftest-Weiche); Tests mit `@pytest.mark.markierung_an` prüfen den neuen Standard.

## Nachtrag 18.10.2026 – bestätigt an Nicos echten Beispielen (Auftrag 003)
Zwei echte Paare (PDF + fertige CDE von Hand) wurden in der Sitzung nur gemessen, nie gespeichert. Ergebnis (nur Formen, keine Inhalte):
- **Farben in der CDE:** alt = `FF00B0F0` (Excel-Standardfarbe „Hellblau“) **+ durchgestrichen**, neu = `FFFF0000`. Keine Unterstreichung nötig (beim Kopieren aus Word
  mitgekommene Unterstreichung ist Zufall). In einem Beispiel passen 46 geänderte Absätze genau zu dieser Regel.
- **Kein Trenner:** Die Zelle zeigt den PDF-Absatz „inline“: Alt und Neu stehen genau so hintereinander wie im PDF (Reihenfolge und Leerzeichen aus dem PDF),
  `markierung.trenner: ""` (Nico, 18.10.: „Wie im PDF“). Das ersetzt „ein Leerzeichen“ oben.
- **Grundschrift** (z. B. fett) steckt in jedem Lauf; **Markierungen früherer Runden** (andere Farben, z. B. dunkelblau/dunkelrot) bleiben unangetastet und zählen beim
  Lesen als normaler Text, außer sie sind durchgestrichen (dann alt).
- **Ganz gestrichener Absatz/Listenpunkt:** Zeile bleibt, Text hellblau durchgestrichen (Nico, 18.10.); Nummern-Änderungen der Folgepunkte wie im PDF markiert.
- **Excel weicht im Absatz anderswo vom PDF ab:** fertiger Vorschlag zum Abhaken, nie automatisch (Nico, 18.10.; Regel 27).
- Das PDF-Lesen dazu (Linie entscheidet, Word-Markup-Leiste) steht in ADR 0006, das verlustfreie Schreiben in ADR 0007.
