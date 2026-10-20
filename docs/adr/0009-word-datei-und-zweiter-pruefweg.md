# ADR 0009: Word-Datei (.docx) lesen und zweiter Prüfweg (Konsens)

Autor: Nico Lehr · Stand: 20.10.2026 · Auftrag 003, Welle 5

## Ausgangslage
Nico hat zu den PDFs meistens die Word-Datei mit Änderungsverfolgung. Im PDF muss das Tool Linien und Farben deuten (ADR 0006);
im Word-XML stehen Einfügungen und Streichungen eindeutig. Außerdem sollen stille Fehler (automatisch eingetragen, aber falsch)
von einem zweiten, unabhängigen Weg abgefangen werden.

## Entscheidung 1: Word-Leser (`src/docx_reader.py`, Einstellungen `config/docx.yaml`)
- `analysiere_docx(pfad)` liefert dieselbe Datenklasse wie der PDF-Leser (`PdfAnalyse`); Planer und Applier bleiben unverändert.
  Kennzeichen: `legende_modus "docx"`, Seite 0, keine Bildkoordinaten, keine Lese-Lücke (das XML ist vollständig).
- Gelesen mit der Standardbibliothek (`zipfile`, `xml.etree`): `w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`, Zeichen-genaue Änderungen,
  markierte Leerzeichen (`trenner_art`), Tabellen (eine Zelle = ein Block, nie allein automatisch), Kopf- und Fußzeilen (getrennt,
  für die Spec-Erkennung), Kommentare (nur Info).
- **Listennummern** („A.“, „1.“, „a.“, „1)“) werden aus `numbering.xml` selbst berechnet, im Original- und im Endstand. Eine Nummernänderung
  wird wie „c.a.“ im PDF als alt/neu abgebildet. Buchstabenfolge ab 27 wie in Word (aa, bb); die Reihe aa, ab ist in der Config umschaltbar.
- **Nie still:** Verbundene oder getrennte Absätze (Änderung an der Absatzmarke) sperren den Absatz (`absatzgrenze_unklar`, Lücke, Warnung).
  Textfelder, Fußnoten, Formeln und Unbekanntes erscheinen als Warnung, mit Änderung auch als Lücke.
- **Kreuzprobe:** `vergleiche_mit_pdf(docx, pdf)` zählt Wortunterschiede zwischen Word- und PDF-Lesung (nur Zahlen und Positionen).
  `python -m src.cli steckbrief --docx <datei>` liefert nur Zahlen und Formen einer Word-Datei, keinen Text (Regel 16).
- Beweis in den Tests: Dasselbe erfundene Dokument als .docx und als Word-PDF ergibt dieselben Änderungslisten (5 Szenarien, Nummern
  aus `numbering.xml` und getippt). **Grenze:** Es gab keine echte .docx zum Messen; der Leser ist nur an selbstgebauten Dateien getestet.

## Entscheidung 2: Zweiter Prüfweg (`src/konsens.py`, Abschnitt `konsens:` in `config/zuordnung.yaml`)
Zwei Prüfungen, die weder die Operationen noch die Hilfsfunktionen des Planers benutzen. Jede Abweichung macht die Änderung zu „prüfen“, nie still.
- **(a) Wort-Bildprobe** (nur PDF, vor dem Schreiben, im Plan): Seite rendern, an jedem Wort einer automatischen Änderung prüfen, ob die Markierung
  sichtbar ist (neu = farbige Unterstreichung bzw. Farbe, entfällt = Durchstreichung, bleibt = nichts). Wörter, die nicht sicher beurteilbar sind,
  zählen als „nicht prüfbar“, nie als Abweichung. Code `konsens_bild`. Eingehängt in `runs.analysiere` und `messen`.
- **(b) Konsens-Text** (nach dem Schreiben, bei PDF und Word): Aus den PDF-/Word-Absätzen wird ein eigenes Soll (Zeichen + Klasse alt/neu/normal)
  berechnet und mit der geschriebenen Zelle verglichen; außerdem darf sich keine Zelle ändern, die unberührt bleiben soll (andere Zeilen, andere Specs).
  Bei Abweichung wird die Datei **ohne** die betroffenen Änderungen neu geschrieben; sie stehen auf „prüfen“ (Code `konsens_text`).
  Leerraum vergleicht die Prüfung bewusst nicht. Neue Section-Zeilen prüft sie nicht.
- `konsens.aktiv: false` schaltet beides ab. Die Gegenprobe A3 bleibt unverändert; sie fängt die gleichen Fälle weiter ab, wenn der zweite Weg aus ist.

## Messung (erfundene Daten, 20.10.2026)
- Fehlalarm auf `nachbau_word` (60 Paare, 3768 automatische Änderungen): 0 zusätzliche „prüfen“; stille Fehler unverändert (12 Zellen, alles Leerraum).
- Eingebaute Fehler (Mutationstests): Konsens-Text erkennt 54 von 56 (96,4 %), die Bildprobe 95 % und mehr je Fehlerart.
  Nicht prüfbar waren auf `nachbau_word` rund 22 % der Wörter (zu klein, Hintergrund, Satzzeichen).

## Folgen
- Der Vertrauensmodus (ADR 0008) zählt `konsens_*`-Codes wie offene Punkte.
- Offen: echte Word-Dokumente sind nicht gemessen; die Bildprobe prüft nur ungefähr drei Viertel der Wörter.
