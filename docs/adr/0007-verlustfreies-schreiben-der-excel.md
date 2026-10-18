# ADR 0007 – Verlustfreies („chirurgisches“) Schreiben der Excel

Autor: Nico Lehr · Datum: 18.10.2026 · Status: angenommen (Auftrag 003)

## Ausgangslage
Bisher lädt der Applier die Excel mit openpyxl, ändert sie im Speicher und speichert sie neu. Ein Test mit Nicos echten CDE-Dateien (nur in der Sitzung) zeigt:
openpyxl **lässt beim Speichern weg**: Bilder/Logos und Formen (Zeichnungen), die Vertraulichkeits-Kennzeichnung der Datei (`customXml`, `docMetadata/LabelInfo.xml`),
Druckereinstellungen und das Blattschutz-Element. Damit wäre das Ergebnis nicht „wie Nico“ und verletzt das Bilanz-Prinzip (nichts geht still verloren).
Neue Python-Pakete (z. B. lxml) sind im Firmennetz nicht installierbar.

## Entscheidung
- openpyxl bleibt **Lese- und Planungsmodell** (Excel-Modell, Planer, Applier-Logik, Gegenprobe A3 lesen weiter mit openpyxl).
- Geschrieben wird mit dem neuen Modul `src/xlsx_patch.py`: Ergebnis = **Kopie des Original-ZIP**, in der nur die Teile neu geschrieben werden, die sich ändern müssen:
  das betroffene Blatt-XML, `xl/sharedStrings.xml` (neue Einträge hinten anfügen, Zähler anpassen), bei neuen Zeilen zusätzlich `xl/workbook.xml` (Druckbereich, Filter-Namen)
  und Zeichnungs-Anker; `xl/calcChain.xml` wird entfernt (samt Verweis), Excel baut sie neu. **Alle anderen Teile bleiben inhaltlich bytegleich** (gleiche Reihenfolge).
- Arbeit nur auf Textebene des XML (reguläre, eng begrenzte Ersetzungen auf maschinell erzeugtem XML), damit Namensräume, `mc:Ignorable` und unbekannte Erweiterungen erhalten bleiben.
- **Zellen:** Rich-Text als neuer `sharedStrings`-Eintrag mit Läufen `<r><rPr>…</rPr><t xml:space="preserve">…</t></r>`; die Schrift jedes Laufs kommt aus der Zellschrift
  (`styles.xml` → `cellXfs[s]` → `fonts`) bzw. aus dem vorhandenen Lauf, ergänzt um Farbe und `<strike/>` (ADR 0005). Zahlen bleiben Zahlen.
- **Neue Zeile:** alle Zeilen darunter rücken; mitgezogen werden Zellbezüge, `dimension`, verbundene Zellen, bedingte Formatierung und Datenprüfung (`sqref`, auch x14),
  Autofilter und `_xlnm._FilterDatabase`, Druckbereich/-titel, Hyperlinks, Seitenumbrüche, Zeichnungs-Anker und Formeln (Zeilenbezüge ≥ Einfügestelle). Vorlage = Nachbarzeile
  gleichen Typs (Stil-Index je Zelle, Höhe, Konstanten). Was nicht sicher verschoben werden kann (geteilte Formel über die Einfügestelle, Tabellenobjekte, Pivot,
  externe/3D-Bezüge, INDIREKT/BEREICH.VERSCHIEBEN), führt zu „prüfen“ statt zu einer kaputten Datei.
- **Prüfung nach dem Schreiben:** Teile-Liste gleich (bis auf `calcChain`), unberührte Teile entpackt bytegleich, openpyxl lädt die Datei, Gegenprobe A3 (alt-Ansicht = vorher,
  neu-Ansicht = PDF-neu, Formatklasse je Zeichen). Schlägt etwas fehl: keine Ausgabe-Datei, klare Meldung.
- Konfiguration `anwenden.schreiber: patch | openpyxl` (Standard `patch`; `openpyxl` nur für alte Tests und Notfälle, mit Warnung „Bilder/Kennzeichnung können verloren gehen“).

## Folgen
- Kommentar- und Gelb-Modus (`markierung.geaenderte_zellen`) und das Zeilen-Löschen haben im Patch-Modus keinen Weg; sie bleiben nur im openpyxl-Modus (Nico braucht sie nicht).
- Tests für das alte Speichern laufen über die Weiche weiter; neue Tests prüfen Bytegleichheit, LibreOffice-Öffnen und Bezugsverschiebung.
