# ADR 0006 – Word-Änderungsverfolgung lesen, CDE-Register zuordnen, Bereich der CDE

Autor: Nico Lehr · Datum: 18.10.2026 · Status: angenommen (Auftrag 003)

## Ausgangslage
Nicos echte Beispiele (nur in der Sitzung gemessen, nie gespeichert; hier nur Formen) sind **Word-PDFs mit Änderungsverfolgung** in der Ansicht „Markup“:
- **Neu und entfällt haben dieselbe Farbe** (Word-Rot `#d13438`). Den Unterschied macht die **Linie**: neu = unterstrichen (dünnes Rechteck auf ~90 % der Zeichenhöhe),
  entfällt = durchgestrichen (Rechteck auf ~50–60 %). Geänderte Listennummern erscheinen grün (`#498205`) als „c.a.“ (alte Nummer gestrichen, neue unterstrichen).
- Rechts eine graue **Markup-Leiste** (`#f2f2f2`, ab ~68 % der Seitenbreite bis zum Rand) mit Sprechblasen („hat formatiert: …“, „Kommentiert …“) und gestrichelten,
  farbigen **Verbindungslinien** vom Text in die Leiste; links schwarze **Änderungsbalken**; gelbe Hinterlegung auf Platzhaltern.
- Kopf- und Fußzeilen enthalten auf jeder Seite eigene Änderungen; die ersten Seiten sind Vorspann (Deckblatt, Prüftabelle, gestrichene Revisionstabelle, Hinweiskästen).
Das bisherige Tool las jedes Rot als „entfällt“, klebte Sprechblasen an Absätze, zerriss Absätze neben Sprechblasen und erkannte Kopf-/Fußzeilen nicht.
Ergebnis auf einem echten Beispiel: 1 von ~60 Änderungen automatisch, über 200 Einträge „nicht übernommen“.

Die **CDE** ist ein Register: Kopfbereich über mehrere Zeilen, eine Beispielzeile, dann **eine Zeile je Absatz/Listenpunkt** mit Section-Spalte („1.1“, „Part 1“),
Sub-Section-Spalte (A…Z, AA…), Textspalte (Listennummer + Tab/geschützte Leerzeichen **im** Text), Konstanten (Spezifikation, Revision, Datum) und Prüfer-Spalten.
Überschriftzeilen tragen nur den Titel (fett) und ein festes Kennwort in einer Nachbarspalte.

## Entscheidung
**1. Legende: „Linie entscheidet“** (`config/farben.yaml legende.modus`):
- `auto` (Standard): erkennt Word-Markup (Markup-Leiste **oder** farbige Unter- und Durchstreichung in derselben Farbe) ⇒ `linie`, sonst `farbe`.
- `linie`: Durchstreichung ⇒ entfällt (jede Farbe), **farbige Unterstreichung ⇒ neu**, schwarze Unterstreichung ⇒ normal, Farbe ohne Linie ⇒ normal mit Hinweis.
- `farbe` (Nicos Regel für von Hand markierte PDFs, ADR 0005): Strich ⇒ entfällt, Rot/Blau ohne Strich ⇒ neu.
- `alt` (nur für die erfundenen Korpora von Auftrag 001/002): blau ⇒ neu, rot ⇒ entfällt. Jeder Korpus nennt seine Legende in `KORPUS.json`.
Farbe + Strich ist kein Widerspruch mehr; Widerspruch = Strich **und** farbige Unterstreichung am selben Zeichen (Absatz-Sperre).
**2. Markup-Leiste:** wird je Seite erkannt und **vor** dem Zeilenbau abgeschnitten; Sprechblasen-Text und Verbindungslinien sind nie Absatztext und nie Strich.
Sprechblasen landen als Info-Liste im Protokoll (`PdfAnalyse.randnotizen`). Änderungsbalken, gelbe Hinterlegung und Verbindungslinien gelten auch in der Farb-Inventur als erklärt.
**3. Kopf-/Fußzeilen** über Randbänder, die auf ≥ 50 % der Seiten an derselben Höhe wiederkehren (Schlüssel aus unveränderten Wörtern, Ziffern vereinheitlicht) ⇒ ignoriert.
**4. Absätze:** jede Listennummer am Zeilenanfang (1. / a. / 1) / a) / A. / AA.) beginnt einen eigenen Absatz; „1.1 TITEL“ und „PART 1 TITEL“ ⇒ Section + Titel.
**5. CDE-Register** (`src/excel_model.py`): Aufbau-Erkennung nur aus der Struktur (keine echten Spaltennamen im Code): Kopfzeile ≠ 1, Beispielzeile, Section-Spalte,
Sub-Section-Spalte, Textspalte mit Label im Text, Überschrift-Kennung; Profil-Datei bleibt möglich (`zeile_pro_absatz` + neue Schlüssel). Die Normalform ignoriert Label und Polster.
Zellen mit Markierung werden über ihre **Neu-Ansicht** gelesen (durchgestrichene Läufe zählen nicht).
**6. Bereich der CDE:** Änderungen vor dem ersten bzw. nach dem letzten sicheren Reißverschluss-Anker (mindestens `planer.umfang_min_anker`) gehören nicht zur CDE
(Vorspann, Revisionstabelle). Sie stehen als **eine Zeile** im Protokoll (`Aenderung.bereich = "ausserhalb"`), zählen in der Bilanz, sind aber **kein** offener Punkt.
**7. Unsicheres bleibt „prüfen“** (Regel 27); weicht die Excel im Absatz anderswo vom PDF ab, gibt es einen fertigen Vorschlag zum Abhaken (Nico, 18.10.).

## Folgen
- `SCHEMA_VERSION` 4: `Aenderung.bereich`, `.stuecke` (fertige Ergebnis-Läufe für die Anzeige), `.grund_einfach`; `PdfAnalyse.randnotizen` (intern).
- Tests, die die alte Legende voraussetzen, laufen über eine Weiche weiter (nicht gelöscht, nicht abgeschwächt).
- Ein Generator v3 (`word_markup`) baut erfundene Paare in genau dieser Struktur; echte Dokumente bleiben außerhalb von Git (Regel 16).

## Nachtrag 19.10.2026 – Spec-Filter (Auftrag 003, Welle 3; Info von Nico)
**Eine CDE enthält mehrere Specs.** Nico filtert nach der Spec, die zum PDF gehört, und bearbeitet nur deren Zeilen – nie die ganze CDE.
- **Spec-Spalte** (`src/excel_model.py`, `BlattLayout.spec_spalte`): im Register die Spalte, deren Werte fast alle spec-artig sind
  (`excel.register.spec_muster`, Anteil `spec_min_anteil`), bevorzugt links der Section-Spalte. Einheitliche Form: Ziffergruppen mit
  Leerzeichen; Leerzeichen, Unterstrich, Punkt und Bindestrich zählen gleich, Zusätze (z. B. eine Region) und Leerzeichen am Ende sind egal.
- **Spec des PDFs** (`src/planner.py`): sicher aus der Titelzeile („SECTION NN NN NN …“) oder der Fußzeile („NN NN NN … - Seite“),
  Muster `planer.spec_pdf_muster`; sonst zählt eine einzige im PDF genannte Spec, die es in der CDE gibt. Der Nutzer kann sie vorgeben
  (`plane(..., spec=...)`).
- **Wirkung:** Enthält die CDE mehrere Specs, benutzen Kandidaten, Reißverschluss, Bereich der CDE, Dubletten, Vorlagezeilen, Polster und
  neue Zeilen nur die Zeilen dieser Spec; Zeilen anderer Specs werden nie angefasst (auch keine „schon eingetragen“-Treffer dort).
  Keine oder mehrere passende Specs ⇒ Warnung „Spec des PDFs nicht in der CDE gefunden“ und nichts automatisch (Code `spec_unklar`,
  alles „prüfen“). Nur eine Spec in der Mappe ⇒ Verhalten wie zuvor.
- **Plan:** `Plan.spec = {"erkannt", "quelle" (vorgabe | pdf | einzige), "zeilen": [von, bis], "kandidaten", "pdf"}` (nur Nummern und
  Zeilen, keine Inhalte); die Übergabe an `plan.json` und Dashboard/CLI bindet der Orchestrator an.

## Nachtrag 19.10.2026 – Leerzeichen an Einfügung/Streichung (Auftrag 003, Welle 3)
Word markiert das Leerzeichen an einer Einfügung oder Streichung mit (` {+neu+}` oder `{+neu +}`). Sichtbar markierter Leerraum wird
vom Leser als `trenner_art` geliefert und im Abgleich ein eigener d-/a-Lauf. **Am Zeilen- oder Seitenumbruch** zeichnet das PDF
kein Leerzeichen, dort ist nichts Markiertes zu sehen (`abgleich._umbruch_leerraum`): Bei einer Gruppe aus nur d- oder nur a-Läufen
ohne sichtbar markiertes Leerzeichen gehört ihr das unsichtbare am Umbruch (davor, sonst dahinter, wenn davor ein sichtbares
unmarkiertes steht); Trenner am Umbruch im Inneren einer Gruppe gehören ihr. Ersetzungen („altneu“) bekommen kein markiertes
Leerzeichen; zwischen „alt“ und „neu“ nur, wenn das Leerzeichen hinter „neu“ sichtbar markiert ist. Der Text ändert sich dadurch nie,
nur die Zuordnung des Leerzeichens zu alter bzw. neuer Fassung.
