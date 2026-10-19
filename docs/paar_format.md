# Paar-Format und Messbericht (Vertrag zwischen Daten-Bauer und Mess-Bauer)

Autor: Nico Lehr · Stand: 18.10.2026 (Generator v2 und v3 `word_markup`, Messung v2 + Auftrag 003)

Ein **Paar** ist ein Beispiel mit bekannter Wahrheit: ein PDF mit Änderungen, die Excel **vorher** und die von Hand erstellte Excel **nachher**.
Das Messwerkzeug (`python -m src.cli messen`) lässt das Tool auf PDF + vorher laufen und vergleicht das Ergebnis mit „nachher“.

## Ordner (alles unter `data/`, von Git ignoriert – `.gitignore`)

```
data/paare/
  entwicklung/                 ← hier darf geprüft und eingestellt werden (Seeds < 1000)
    KORPUS.json                ← nur bei erzeugten Paaren: Generator-Version, Seeds, Erzeuger, LibreOffice-Version
    synth_0001_standard/       ← Namen beginnen mit synth_ (erfunden) oder echt_ (Nicos Beispiele)
      aenderungen.pdf
      vorher.xlsx
      nachher.xlsx
      erwartet.json            ← NUR bei synth_, optional (Zusatzwissen des Generators)
  ex_test/                     ← früherer Testkorpus (Seeds 1001–1030), seit 16.10.2026 OFFEN als Entwicklung erklärt
  test/                        ← reserviert für die ungesehene Endmessung (nur Seeds ≥ 2000, Regel 11), nie zum Einstellen
  archiv/                      ← verschobene ältere Korpora (z. B. entwicklung_v1/, test_v1/), nie gelöscht (Regel 29)
data/messung/
  001_entwicklung/bericht.json, bericht.md     ← append-only, nie überschreiben (Regel 29)
```

### Korpora (Generator v2, Regel 11)
| Korpus | Seeds | Inhalt | Zweck |
|---|---|---|---|
| `entwicklung` | 1–60 (writer) und 101–160 (gerade Seeds LibreOffice, ungerade writer) | erfundene Paare v2 | Entwicklung und Einstellen |
| `ex_test` | nur 1001–1030 | derselbe Zufallsstrom wie der frühere Testkorpus v1 (gleiche Änderungen, gleiche Familien inkl. der dort offengelegten Familien `gedrehte_seite`, `gruene_markierung`) | offen als Entwicklung erklärt – nicht mehr ungesehen |
| `test` | nur ≥ 2000 | reserviert (später mit neuen, nur dort vorkommenden Familien) | genau einmal messen; **jetzt nicht erzeugen** |
| `nachbau_word` | 301–360 (erlaubt: < 1000) | Generator **v3** `word_markup`: Word-PDF mit Änderungsverfolgung + CDE-Register (siehe unten) | Entwicklung für Auftrag 003 (Legende `auto`, nachher `markiert`) |

`python scripts/erzeuge_paare.py --korpus entwicklung|ex_test|test|nachbau_word --seeds … [--erzeuger writer|libreoffice|wechselnd]` prüft die
Seed-Bereiche (Exit-Code 2 bei Verstoß). Die v1-Paare liegen in `data/paare/archiv/entwicklung_v1/` und `data/paare/archiv/test_v1/`.
Pfade kommen aus `config/paths.yaml` (`paare_dir`, `messung_dir`), nie hartcodiert (Regel 28). Der Name `echt_…` ist nur für Nicos echte Beispiele, die **lokal** bleiben.

## Das Messwerkzeug braucht `erwartet.json` NICHT
Das **Soll** wird immer aus `vorher.xlsx` ↔ `nachher.xlsx` abgeleitet (Zellvergleich, Absatzausrichtung). So läuft dasselbe Werkzeug später unverändert mit echten Paaren.
`erwartet.json` liefert nur Zusatzwissen für die Auswertung (Störer-Familien, Lücken, Anzahl markierter Wörter).

### `erwartet.json` (nur erzeugte Paare; alle Inhalte erfunden)
```json
{
  "format": 1,
  "generator_version": "2",
  "seed": 12,
  "profil": "standard",
  "korpus": "entwicklung",
  "erzeuger": "writer",
  "stoerer": ["fussnote", "silbentrennung"],
  "woerter_markiert": 37,
  "aenderungen": [
    {"nr": 1, "typ": "text_ersetzt", "blatt": 1, "zeile": 12, "spalte": 3, "woerter": 4,
     "soll_status": "uebernommen"}
  ],
  "luecken": [{"art": "scan_seite", "seite": 2}]
}
```
- `typ`: `text_ersetzt | text_eingefuegt | text_geloescht | absatz_neu | absatz_geloescht | section_neu | section_geloescht` (wie `Aenderung.typ`).
- `blatt` = Nummer des Blatts (1 = erstes), `zeile`/`spalte` beziehen sich auf `vorher.xlsx`; bei neuen Zeilen ist `zeile` die Zeile, **nach** der eingefügt wird.
  (Auch im Messbericht ist `fehler[].zeile` die Zeile in vorher. Seit Welle 3 stehen daneben `zeile_nachher` und `zeile_ergebnis` – dort
  liegen Soll und Ist nach eingefügten Zeilen wirklich –, `neue_zeile` und `stille_arten`. Im bericht.md: „Z167 (N168/E168, neu)“.)
- `soll_status`: was ein sauberes Tool tun sollte: `uebernommen` (eindeutig, automatisch), `pruefen` (absichtlich mehrdeutig) oder `nicht_uebernommen` (absichtlich nicht machbar, z. B. Tabelle).
- `luecken`: was das Tool als Lücke **melden muss** (`models.LUECKEN_ARTEN`).
- Der Generator schreibt die Wahrheit aus seinem **logischen Modell**, nie aus dem PDF-Leser.

- `erzeuger`: `writer` (eigener PDF-Schreiber) oder `libreoffice` (seit v2). Fehlt das Feld (v1), war es `writer`.
- Nur bei `libreoffice`: Block `"libreoffice": {"libreoffice_version", "schrift", "neu_unterstrichen", "aenderungsbalken", "variante",
  "listen", "tabellen", "bilder"}` und `pdf_info.seiten_geschaetzt: true` – die Seitenangaben von Änderungen und Lücken stammen aus dem
  Satz des eigenen Schreibers und sind bei LibreOffice nur geschätzt. Änderungen, Lücken, Excel-Dateien und `woerter_markiert` sind für
  beide Erzeuger **gleich** (Wahrheit aus dem logischen Modell, nie aus dem PDF-Leser).

### `KORPUS.json` (je Korpus-Ordner, nur erzeugte Paare)
```json
{"generator_version": "2", "korpus": "entwicklung", "seeds": [1, 2, 101, 102], "erstellt": "2026-10-16",
 "erzeuger": {"1": "writer", "2": "writer", "101": "writer", "102": "libreoffice"},
 "je_erzeuger": {"writer": 3, "libreoffice": 1}, "libreoffice_version": "24.2.7.2"}
```
Zahlen verschiedener Generator-Versionen sind **nicht vergleichbar** (das Skript verweigert, v2-Paare in einen v1-Ordner zu legen).

Seit Auftrag 003 nennt `KORPUS.json` außerdem, **wie das Tool den Korpus lesen und schreiben soll**:
`"legende": "alt" | "auto" | "linie" | "farbe"` (ADR 0006) und `"nachher_format": "sauber" | "markiert"` (ADR 0005, Nachtrag).
Fehlt die Angabe (alle Korpora v1/v2), gilt beim Lesen `legende: "alt"`, `nachher_format: "sauber"`. Das Skript schreibt beide Felder
(`nachbau_word`: `auto` / `markiert`; sonst `alt` / `sauber`).

### Generator v2 (16.10.2026)
* **Geschütztes Leerzeichen wie ein Mensch es tippt:** v1 setzte es auch hinter `+}`/`-}`; aus „{-100-} {+120+} kWh“ wurde vorher
  „100 ␣<NBSP>kWh“, nachher „120<NBSP>kWh“ – zwei Texte, die kein Mensch so schreibt und die das PDF nicht zeigt (3 Phantom-„stille
  Fehler“). v2 setzt es nur, wenn es im alten UND im neuen Stand direkt zwischen Zahl und Einheit steht (`stoerer.nbsp_setzen`). Sonst sind
  v1 und v2 je Seed gleich (gleicher Zufallsstrom). `erzeuge_paar(..., version="1")` baut v1-Paare weiter nach (nur für Red-Team R3).
* **Zweiter PDF-Erzeuger LibreOffice** (`tests/fixtures/pdf_lo.py`): dasselbe Modell als Flat-ODT (von Hand geschriebenes XML) →
  `soffice --headless -env:UserInstallation=file:///<tmp>/lo_profil_<pid> --convert-to pdf` (eigenes Profil je Aufruf, Zeitlimit).
  Wie ein Textprogramm: eingebettete Schriften (Carlito, Liberation Sans/Serif), Blocksatz, weiche Trennzeichen (U+00AD) in langen
  Wörtern, echte Aufzählungen/Nummerierungen (Listen-Stile), Kopf-/Fußzeile mit Seitenzahl-Feld, natürliche Seitenumbrüche (auch mitten
  im Absatz), Fußnoten, Tabellen, Bilder, Zwei-Spalten-Bereich, Scan-Seite (eigene Seitenvorlage ohne Kopfzeile), Hyperlinks.
  Markierung nur per Zeichenformat: blau = neu (je Paar teils zusätzlich unterstrichen), rot + durchgestrichen = entfällt (der Strich läuft
  über ganze Zeilen), Mischformen wie beim eigenen Schreiber; je Paar zufällig Änderungsbalken am linken Rand. Nicht nachgebildet:
  PDF-Kommentare (Variante `annot` → farbig), CMYK (LibreOffice schreibt RGB), `/Rotate` (`gedrehte_seite` gibt es nur beim eigenen
  Schreiber), „Änderungen nachverfolgen“. Fehlt `soffice`, überspringen Skript und Tests die LibreOffice-Paare sauber.

### Generator v3 `word_markup` (18.10.2026, Auftrag 003 – Korpus `nachbau_word`)
Eigenes Modul `tests/fixtures/word_markup.py` (Zeichnen: `tests/fixtures/pdf_word.py`), Einstieg wie bisher:
`erzeuge_paar(seed, "word_markup", ziel, "nachbau_word")` oder mit kleinerem Profil `{"basis": "word_markup", "name": "klein",
"zeilen": (45, 60)}` (weitere Schalter: `leiste`, `zwei_autoren`, `alte_markierung`, `doppelbuchstaben`, `aenderungen`). Das Profil
gehört nur zu `nachbau_word` (und umgekehrt), Erzeuger nur `writer`. Gleicher Seed ⇒ **bytegleiche** Dateien (feste ZIP-Zeitstempel).
Nachgebaut sind die am echten Beispiel gemessenen **Formen** (Inhalte erfunden):
* **PDF (Word-Markup):** 612 × 792 pt, Times 8,4 pt, Zeilenabstand 10 pt, Einzüge 55/82/109/137/164/191/218 (Label auf Ebene n, Text
  ab n+1, hängender Einzug), rechter Rand ~395. Neu = Word-Rot `#d13438` + Unterstreichung (Rechteck 0,36 pt auf ~90 % der
  Zeichenhöhe), entfällt = dieselbe Farbe + Durchstreichung (0,46 pt auf ~55 %), **zeichengenau** (nur `[` und `]`; `[40]` gestrichen
  direkt vor `42,9`; einzelne Satzzeichen; Wörter; halbe Sätze; ganze Absätze/Listenpunkte). Geänderte Listennummern grün `#498205`
  („c.a.“). Gelbe Hinterlegung (9,6 pt) hinter Platzhaltern `[…]`, schwarze Änderungsbalken (x = 27,3, Breite 0,53), graue Markup-Leiste
  (x 415–612, y 95–696) mit Sprechblasen („hat formatiert:“/„Formatiert:“, „Kommentiert [XY1]:“, Helvetica 6,1 pt) und gestrichelten
  Verbindungslinien (dash [0.14 0.14], 0,14 pt) samt Ankerpunkt – oder **nur inline** (ohne Leiste, ~25 % der Seeds). Kopf- und
  Fußzeile auf jeder Seite (mit eigenen Änderungen, Seitenzahl „11 22 33 - n“), Vorspann (Deckblatt in Helvetica ganz neu,
  Prüftabelle, gestrichene Revisionstabelle, gestrichene Bearbeiter-Hinweiskästen mit `*****`, ggf. Inhaltsverzeichnis), dann
  „SECTION … – Titel“, „PART 1 …“, „1.1 TITEL“, „A.“, Listen 1./a./1)/a), … „END OF SECTION“. Zweite Autorfarbe Blau `#0078d4`.
  Keine PDF-Annotationen. Absätze laufen über Seitenumbrüche.
* **vorher.xlsx (CDE-Register):** Blätter Hinweise, Deckblatt, Unterschriften, Hauptblatt (Register/CDE/Anforderungen), verstecktes
  Blatt Listen. Kopfbereich Zeilen 1–4 (Verbünde A1:A4 … H1:P4), Zählformeln U2:Z5, Spaltenköpfe 5/6 (G5:G6 …), Beispielzeile 7, Daten ab
  Zeile 8 (300–550 Zeilen). Eine Zeile je Absatz/Listenpunkt: E Section („Part 1“, „1.1“, je Seed teils mit Leerzeichen am Ende),
  F Sub-Section (A … Z, AA, BB …), G Text (Listenpunkt mit Label + Tab bzw. 13–19 geschützten Leerzeichen + Leerzeichen),
  Überschriftzeilen fett mit Kennwort in I und ohne Status in H; B/C/D Konstanten, Markierungen früherer Runden in C/D (dunkelblau,
  Zeilenumbruch, dunkelrot), zellenweise fette Grundschrift. Bedingte Formatierung und Datenprüfung in zerstückelten Bereichen,
  Autofilter A6:AB…, Druckbereich, Drucktitel 5:6, `<sheetProtection formatCells="0"/>`, Zeichnung (Form + 1–2 Bilder, Anker Zeile 3/7).
  Per ZIP-Nachbearbeitung: **sharedStrings** statt Inline-Strings (wie Excel), `customXml/` (2 Einträge samt Rels), `docMetadata/LabelInfo.xml`,
  `xl/printerSettings/printerSettings1.bin`, `xl/calcChain.xml` – Content-Types und Rels stimmen, openpyxl und LibreOffice öffnen die Datei.
* **nachher.xlsx (wie Nico):** geänderte Zelle G = PDF-Absatz inline – unverändert normal, entfällt `FF00B0F0` + durchgestrichen,
  neu `FFFF0000`, **kein Trenner**, Grundschrift (Arial 11, ggf. fett) in jedem Lauf; gestrichener Absatz: Zeile bleibt, ganz hellblau
  durchgestrichen; neuer Absatz/Listenpunkt/Sub-Section: neue Zeile direkt nach dem Vorgänger (B/C/D wie Nachbar, E/F aus der Gliederung,
  G rot, Rest leer); Nummern-Änderung „c.“ → „a.“ als „c.“ alt + „a.“ neu. Einfarbige Zellen je Seed als Zellschrift oder als Rich Text mit
  einem Lauf. Kopf-/Fuß-/Vorspann-Änderungen und Sprechblasen kommen nicht in die Excel. Weicht die Excel im Absatz anderswo vom PDF ab,
  **bleibt die Excel-eigene Abweichung stehen** und nur die PDF-Änderung wird markiert eingetragen (Nico, 19.10.; in erwartet.json
  „pruefen“ mit fertigem Vorschlag = Zelle in nachher). Neue Zeilen: nur die Textspalte rot, B/C/D aus der Zeile davor (ihr Spec-Block),
  E/F normal, alle Spalten rechts der Textspalte leer.
* **Mehrere Specs** (Störer `mehrere_specs`, gerade Seeds; Schalter `mehrere_specs`, Größe `fremd_zeilen`): das Register enthält 2–3
  weitere Specs als zusammenhängende Blöcke vor UND nach der Ziel-Spec (= Spec des PDF), je mit eigener Spec-Nr. in Spalte B (Schreibweise
  uneinheitlich: „11_22_33“, „11_22_33 “, „11 22 33“), eigener Revision, eigenem Datum und eigener Gliederung 1.1 …; ein Teil ihrer Sätze
  steht genau so oder fast so auch in der Ziel-Spec (`dubletten_fremde_spec`). Soll: nur Zeilen der Ziel-Spec ändern sich, neue Zeilen
  innerhalb ihres Blocks. erwartet.json: `specs` = [{rolle ziel|fremd, spez, revision, zeilen_vorher [von, bis], zeilen_nachher [von, bis]}].

**`erwartet.json` v3** (zusätzlich zu den Feldern oben): `generator_version: "3"`, `pdf_stil: "word_markup"`, `legende`, `nachher_format`,
`hauptblatt` (Name), `abweichungen_ohne_aenderung` ([{blatt, zeile, spalte}] – Excel weicht ab, PDF ändert dort nichts: das Tool darf die
Zelle nicht anfassen), `excel` ({datenstart, zeilen_vorher, zeilen_nachher, zusatzteile}), `pdf_info` ({seiten, leiste, sprechblasen,
kommentare, formatiert, umbrueche_im_absatz}). Je Änderung zusätzlich:
* `bereich`: `cde` oder `ausserhalb` (Kopf-/Fußzeile, Vorspann – „gehört nicht zur CDE“; dann `blatt`/`zeile`/`spalte` = 0),
* `art`: feinere Kennung (`wort_ersetzt`, `zahl_ersetzt`, `eingefuegt`, `satz_eingefuegt`, `geloescht`, `satzteil_geloescht`,
  `satzzeichen`, `klammern`, `platzhalter_ersetzt`, `listennummer`, `absatz_neu`, `subsection_neu`, `absatz_geloescht`, `kopf_fuss`, `vorspann`),
* `soll_status` zusätzlich `ausserhalb` (kein offener Punkt, zählt aber in der Bilanz, ADR 0006 Punkt 6),
* `nachher_zeile` (Zeile in nachher.xlsx), `absatz` (Kennung im Modell), bei `pruefen` `vorschlag` = fertige Zelle als Läufe
  `[["normal"|"alt"|"neu", Text], …]` (= Zelle in nachher.xlsx).
Typische Störer-Kennungen: `markup_leiste`/`nur_inline`, `sprechblase_formatiert`, `sprechblase_kommentar`, `kopf_fuss_aenderung`,
`vorspann`, `bearbeiter_hinweise`, `inhaltsverzeichnis`, `seitenumbruch_im_absatz`, `listennummer_gruen`, `absatz_neu`, `absatz_geloescht`,
`subsection_neu`, `subsection_doppelbuchstabe`, `klammer_zeichengenau`, `platzhalter_gelb`, `excel_abweichung`,
`excel_abweichung_ohne_aenderung`, `alte_markierung_cd`, `zweite_autorfarbe`, `ueberschrift_aenderung`, `listenebenen`, `fette_zelle`.

### Messung (Auftrag 003): Legende und Format je Korpus, Format-Treue, Rückbau-Messung
* **Legende/Format je Korpus:** `messen` liest `legende` und `nachher_format` aus der `KORPUS.json` des Paar-Ordners (Vorrang: Aufruf
  `--legende`, `--nachher-format` › KORPUS.json › Vorgabe des Korpus `entwicklung`/`ex_test`/`test` = alt/sauber, `nachbau_word` =
  auto/markiert › Konfiguration wie sie ist, z. B. echte Paare unter `--ordner` ohne KORPUS.json). Sie gelten **nur während der Messung
  eines Paars** (`messen.konfig_ueberschreibung`: `legende.modus`; `markiert` ⇒ `markierung.aenderungen: farbe`, `sauber` ⇒ `keine`),
  danach gilt wieder die Konfiguration. Bericht und Konsole nennen beide (`bericht.legende`, `bericht.nachher_format`, je Paar `legende`,
  `nachher_format`, `markiert`), dazu die vom Leser gewählte Legende (`legende_erkannt` je Paar, bei `auto` je Dokument entschieden;
  `gesamt.legende_erkannt` als Häufigkeit) und ob er eine Markup-Leiste erkannt hat (`markup_leiste_erkannt`).
  Die Formatklassen der Messung sind bewusst eine eigene, enge Definition (unabhängig vom Tool): neu = genau `markierung.neu_farbe`.
* **Format-Treue** (nur `markiert`): Zellen werden mit **Formatklasse je Zeichen** gelesen – `alt` = durchgestrichen (jede Farbe),
  `neu` = Neu-Farbe der Markierung (`markierung.neu_farbe`) ohne Strich, sonst `normal` (auch Markierungen früherer Runden). Die strenge
  Vergleichsform einer Zelle enthält die Klassen: gleicher Text mit anderer Markierung ist eine andere Zelle (so wird auch ein nur
  durchgestrichener Absatz zur Soll-Zelle). Wörter zählen je Ansicht (alt = normal + alt, neu = normal + neu); eine offene Änderung in
  einer markierten Zelle „wartet“ wie bisher. Eine **automatisch geschriebene Zelle mit falschem Text oder falscher Formatklasse** ist ein
  stiller Fehler (bei richtigem Text zusätzlich Art `format_falsch`). Kennzahl je Paar und gesamt: `format_treue_zellen_geprueft`
  (Soll-Zellen der Haupt-Spalten, deren Text das Tool richtig geschrieben hat), `…_gleich`, `format_treue_zeichen`, `…_gleich`;
  `gesamt.format_treue` = gleich ÷ geprüft. Kriterium Nr. 3 „Format-Treue“ (Ampel `ok`/`Hinweis`/`nicht gemessen`; der Exit-Code hängt
  an den stillen Fehlern).
* **PDF-blind und Listen-Polster:** Ein Polster aus geschützten Leerzeichen zwischen Listennummer und Text, das genau so schon in
  vorher am Zellanfang vorkommt, wird in neuen Zeilen streng verglichen (ein Mensch übernimmt es von der Nachbarzeile).
* **Bereiche bis zur letzten Excel-Zeile** (`R8:R1048576`) wachsen beim Einfügen nicht (Struktur-Prüfung).
* **Fremde Spec** (Welle 3): Die Messung liest die Spec-Nummern („11 22 33“) der ersten Seiten des PDF und sucht die erste Spalte, deren Wert
  in allen laut Soll geänderten Zeilen diese Nummer ist (Schreibweise vereinheitlicht). Blöcke von mindestens 3 aufeinanderfolgenden Zeilen mit
  einer anderen Nummer dort sind fremde Specs – abgeleitet nur aus vorher/nachher/PDF, also auch für echte Paare. Jede Änderung des Tools in
  einer fremden Spec ist Kollateralschaden und stiller Fehler (Art `fremde_spec`); Kennzahlen `fremde_spec_zeilen`, `fremde_spec_beruehrt`,
  `gesamt.paare_mit_fremder_spec`, Kriterium Nr. 6 „Fremde Spec berührt“ (Rückbau: `tool_extra_fremde_spec`).
* **`keine_wirkung`** zählt nur noch, wenn das Soll die Zelle wirklich ändert (sonst ist „nichts getan“ richtig).
* **Rückbau-Messung** `python -m src.cli messen --nachher [--ordner PFAD] [--paar NAME] [--legende L]` für Paare mit nur
  `aenderungen.pdf` + `nachher.xlsx` (ein vorher.xlsx wird ignoriert): vorher wird aus der Markierung der **letzten Runde** zurückgebaut
  (`markierung.alt_farbe` + Strich → normaler Text, `markierung.neu_farbe` ohne Strich → weg; Zeilen, deren längste Zelle ganz neu ist und
  die nichts Altes tragen, entfallen; alle anderen Zellen bleiben unangetastet), dann läuft das Tool (nachher-Format `markiert`), dann
  wird Zelle für Zelle verglichen: `gleich`, `gleich_bis_leerraum` (gleiche sichtbare Zeichen mit gleicher Formatklasse),
  `tool_fehlt_vorschlag_offen` (eine zugehörige Änderung steht auf „prüfen“), `tool_fehlt_gar_nicht`, `tool_extra` (Tool ändert eine
  Zelle, die nachher nicht ändert), `verschieden`. Ausgabe nur Zahlen (je Kategorie, alle Zellen und Haupt-Spalten) und Schablonen
  (Formen je Lauf wie `--schablonen`, höchstens 40 je Paar, Positivliste) in `messung_dir/NNN_<korpus>_rueckbau/` (append-only).

## Messbericht (`bericht.json`) – enthält NIE Text aus Dokumenten (Regel 16)
```json
{
  "format": 1, "korpus": "entwicklung", "generator_version": "1", "schema_version": 2,
  "gesamt": {"paare": 60, "soll_zellen": 412, "bilanz_stimmig": 60, "vollstaendig": 55, "stille_fehler": 0,
             "auto_richtig": 301, "auto_gesamt": 301, "vorschlaege_richtig": 52, "vorschlaege_gesamt": 55,
             "abdeckung_ohne_nachfrage": 0.73, "lese_luecke_woerter": 12, "idempotent": 60, "kollateral": 0},
  "fehlerarten": {"pdf_falsch_gelesen": 0, "layout_falsch": 0, "zuordnung_falsch": 0, "plan_falsch": 0, "bilanz_kaputt": 0, "unklar": 0},
  "je_stoerer": {"fussnote": {"paare": 5, "stille_fehler": 0, "abdeckung_ohne_nachfrage": 0.8}},
  "kriterien": [{"nr": 2, "name": "Stille Fehler", "wert": "0", "einheit": "Zellen und Befunde", "ziel": "0", "ampel": "ok",
                 "detail": "0 Zellen, 0 Paar-Befunde, 0 Regel-27-Verstöße; 0 Änderungen betroffen; davon nur Leerraum/Zeichenform: 0"}],
  "paare": [
    {"name": "synth_0001_standard", "bilanz_stimmt": true, "vollstaendig": true, "stille_fehler": 0,
     "fehler": [{"blatt": 1, "zeile": 12, "spalte": 3, "aenderung": "a7", "art": "zuordnung_falsch"}]}
  ]
}
```
Der Bericht nennt nur Zahlen, Paar-Namen, Blatt-Nummer, Zeile, Spalte, Änderungs-ID und Fehlerart. Bestehende Felder bleiben, neue kommen nur dazu.

### Zähleinheit (wichtig beim Lesen der Zahlen)
Jede Kriterien-Zeile hat das Feld `einheit` (im `bericht.md` die Spalte „Einheit“; Erläuterungen stehen im Feld `detail`), und `gesamt.zaehleinheiten` fasst es zusammen. Verschiedene Einheiten sind nicht vergleichbar:
**Zellen** (Soll-Zellen, Treffergenauigkeit, Abdeckung, Kollateral), **Änderungen** (je Änderung: `aenderungen_uebernommen_korrekt`, `…_falsch`, `aenderungen_zu_pruefen`, `aenderungen_nicht_uebernommen`, `regel27_verstoesse`), **Paare** (Bilanz, Idempotenz, Originale), **Wörter** (Lese-Lücke), **Seiten** (Plausibilität), **Befunde** (Struktur).
„Stille Fehler“ ist eine gemischte Zahl und wird deshalb aufgeteilt: `stille_fehler_zellen` + `stille_fehler_paar_befunde` (Bilanz, nicht gemeldete Lücke) + `regel27_verstoesse`; dazu `stille_fehler_aenderungen` (betroffene Änderungen) und `stille_fehler_nur_textform` (davon Zellen, die sich nur in Leerraum, unsichtbaren Zeichen oder Absatzgrenzen vom Soll unterscheiden).

### So wird gezählt (Messwerkzeug)
* **Je Änderung bewertet:** Steht neben einer richtig angewendeten Änderung in derselben Zelle eine offene („prüfen“, ohne Freigabe), ist die Zelle kein Fehler, sondern wartet auf die Freigabe (`zellen_warten_auf_freigabe`, zählt unter „zu prüfen“).
* **Strenger Zellvergleich:** doppelte und geschützte Leerzeichen, weiche und unsichtbare Zeichen, Ligaturen und Absatzgrenzen zählen. Nur ein Zeichen, das im PDF gar nicht vorkommen kann (z. B. geschütztes Leerzeichen im Soll), wird als „PDF-blind“ nur lose verglichen und ausgewiesen (`soll_zellen_pdf_blind`, `…_erreicht`); steht es im Ergebnis, aber nicht im Soll, ist es ein Fehler.
* **Regel 27:** Wendet das Tool eine „prüfen“-Änderung ohne Freigabe an, zählt das als `regel27_verstoesse` und als stiller Fehler (Art `regel27_verstoss`); die Zelle zählt nicht als Treffer und nicht als „ohne Nachfrage erreicht“.
* **Struktur:** `struktur_geaendert` (mit `struktur_nach_art`) zählt UNERWARTETE Abweichungen. Die Arten sind: `blatt`, `sichtbarkeit`, `zeile_versteckt`, `spalte_versteckt`, `verbunden`, `validierung`, `zahlenformat`, `schrift` (Schrift und Füllung unveränderter Zellen) und seit K-MESS2 `datentyp`, `bedingte_formatierung`, `hyperlink`, `kommentar`, `benannter_bereich`, `autofilter`, `tabelle`, `fixierung` (Freeze Panes), `spaltenbreite`, `zeilenhoehe` (nur an unveränderten Zeilen). Zeilen werden über die Ausrichtung abgebildet; Bereiche mit Zeilenbezug (Filter, Tabellen, bedingte Formatierung, benannte Bereiche, Validierung, Verbünde, Fixierung) werden um die Zahl der oberhalb eingefügten Zeilen verschoben erwartet. Das legitime Verschieben zählt nicht. Das Tool zieht nur Zeilenhöhe, Ausblenden/Gliederung, Hyperlink-Adressen und Verbünde unterhalb mit; alles andere hält es vorab auf „prüfen“, ein Befund entsteht also nur, wenn trotzdem etwas passiert.
* **Was „Struktur ok“ heißt:** nur, dass bei diesen Eigenschaften nichts Unerwartetes gefunden wurde (`gesamt.struktur_geprueft`). **Nicht geprüft** (`gesamt.struktur_nicht_geprueft`): Rahmen, Ausrichtung und Zellschutz, Blattschutz, Druckbereich und Seitenlayout, Kopf-/Fußzeilen, Bilder, Diagramme und Formen, Gliederungsebenen, Makros, Dokumenteigenschaften, Teilformatierung innerhalb einer Zelle und ob Formeln nach dem Verschieben noch richtig zeigen. Der `bericht.md` sagt das mit dem Satz „Struktur ok heißt nur: …“, die Ampelzeile mit „geprüft: …; nicht geprüft: …“.
* **Datentyp:** Der Zelltext macht aus der Zahl `31` und dem Text `"31"` dasselbe. Deshalb wird zusätzlich der Typ (Zahl, Text, Formel, Datum, Wahrheitswert) verglichen. (a) An einer Soll-Zelle, die das Tool mit richtigem Text befüllt hat (auch in einer neuen Zeile), muss der Typ dem von `nachher.xlsx` entsprechen; sonst gilt die Zelle als nicht erreicht und nicht richtig, in einer verwalteten Spalte außerdem als stiller Fehler der Art `datentyp_gewechselt`. (b) An Zellen, deren Text unverändert blieb, muss der Typ gleich dem in `vorher.xlsx` bleiben: in einer verwalteten Spalte ebenfalls stiller Fehler `datentyp_gewechselt`, sonst Strukturbefund der Art `datentyp` (so auch bei Füll-, Formel- und sonstigen Zellen außerhalb der Haupt-Spalten).
* **Struktur der Excel „alle Vorschläge freigegeben“** (nur mit `--mit-vorschlaegen`): `struktur_vorschlaege_geaendert` (+ `…_nach_art`) zählt Abweichungen, die der Plan **nicht** vorab gemeldet hatte; sie sind ein Befund (eigene Ampelzeile „Vorschlags-Excel (alle freigegeben): Struktur unverändert“). Hatte eine Änderung des Plans das Risiko vorab als Hinweis genannt (z. B. „Autofilter reicht bis an oder unter diese Stelle“), zählt die Abweichung als `struktur_vorschlaege_angekuendigt` und ist kein Befund. Erkannt wird die Vorab-Meldung am Stichwort im Hinweistext, nicht blattgenau.
* **Soll-Zellen außerhalb der Haupt-Spalten:** `soll_zellen` zählt nur die verwalteten Spalten (ohne Formeln). Getrennt: `soll_zellen_formel` (+ `…_verfehlt`), `soll_zellen_unverwaltet_sonstige` (+ `…_verfehlt`, z. B. neu nummerierte Spalte) und `soll_zellen_unverwaltet` (nur kurze Füllung neuer/entfernter Zeilen, aufgeschlüsselt in `…_aufschluesselung`). `nicht_verwaltet_verfehlt_orte` nennt nur Blatt/Zeile/Spalte/Klasse.
* **Abdeckung über ALLE Soll-Zellen:** Die Ampelzeile „Abdeckung ohne Nachfrage (nur Haupt-Spalten)“ zählt nur `soll_zellen`. Die eigene Zeile „Abdeckung über ALLE Soll-Zellen (inkl. Füllspalten, Formeln)“ teilt `erreicht_alle` durch `soll_zellen_alle` (Haupt-Spalten + harmlose Füllung + sonstige + Formelzellen; Einheit Zellen) und hat die Grenze **mindestens 70 %** (Kriterium 4 in `docs/ziel_v1.md`). „Erreicht“ heißt: Text und Datentyp stimmen mit `nachher.xlsx` überein. Die Formelzellen sind im `detail` getrennt ausgewiesen (`erreicht_formel` von `soll_zellen_formel`) und es steht eine Zahl ohne sie dabei (`abdeckung_alle_ohne_formeln`).
* **Absturz eines Messteils:** Stürzt nur der Wiederholungslauf (Idempotenz), die Vorschläge oder die Struktur-Prüfung ab, bleiben die Funde des ersten Laufs erhalten. Das Paar zählt dann als „teilweise gemessen“ (`teilweise_gemessen`, `ausgefallene_messteile`, `paare_teilweise_gemessen`) und in `paare_abgestuerzt`. Bei abgestürzten Paaren steht keine Ampel auf „ok“, sondern auf „nicht vollständig gemessen“ (außer sie ist schon „NICHT ok“).

### Messung v2 (16.10.2026) – neue Kennzahlen (nur Zahlen und feste Kennungen)
* **Klassen der Soll-Zellen** (`soll_klassen`, `soll_klassen_erreicht`, je Paar und gesamt): `haupt` (Nummer/Titel/Inhalt), `formel`,
  `fuell_konstant` (Zelle einer neuen/entfernten Zeile in einer Spalte, die in `vorher` für ALLE Zeilen desselben Typs – gleiche belegte
  Haupt-Spalten, ohne die Kopfzeile – denselben nicht leeren Wert hat; das prüft die Messung selbst) und `nicht_ableitbar` (Kommentar,
  laufende Nummer, geänderte Zelle außerhalb der Haupt-Spalten …). **Abdeckung ableitbare Zellen** = erreicht / (haupt + fuell_konstant +
  formel) (`abdeckung_ableitbar`). Die alten Klassen (`harmlos`, `sonstige`) und Felder bleiben unverändert.
* **Restarbeit je 100 Änderungen** (`restarbeit_je_100`, Teile in `restarbeit`) = 100 × (Änderungen „prüfen“ + Änderungen „nicht
  übernommen“ + leere Zellen in neuen Zeilen, die laut nachher gefüllt sein müssten + Wörter in gemeldeten Lücken ÷ 10) ÷ Änderungen.
  Richtwert ≤ 12 (`docs/ziel_v2.md`).
* **Prüfcodes** (`pruefcodes`): Histogramm von `Aenderung.pruefcodes` (Schema 3; Anzahl Änderungen je Code). Codes, die nicht wie eine
  Maschinen-Kennung aussehen, erscheinen nur als `sonstiger_code` (Regel 16).
* **Sperr-Quote je Code** (`sperr_quote`): je Code die gesperrten Änderungen (Plan-Status „prüfen“), davon messbar und davon
  „nach Freigabe richtig“ (Ziel-Zellen bzw. neue Zeilen stimmen in der Excel mit allen freigegebenen Vorschlägen mit nachher überein;
  nur mit `--mit-vorschlaegen`). Hohe Quote = die Sperre bremst oft unnötig. Gesperrte Änderungen ohne Code stehen unter `ohne_code`.
* **Je Erzeuger** (`je_erzeuger`): Paare, stille Fehler, Regel 27, Kollateral, Treffergenauigkeit, Abdeckung, Lese-Lücke, Idempotenz,
  Restarbeit getrennt für `writer`, `libreoffice` (und `unbekannt` = Paare ohne `erwartet.json`).
* **Kriterien Nr. 7** („Abdeckung ableitbare Zellen“, „Restarbeit je 100 Änderungen“) haben nur die Ampeln `ok` / `Hinweis` / `nicht
  gemessen` – sie ändern den Exit-Code nicht.
* **`messen --vergleich [--gegen NNN]`**: Deltas jeder numerischen Kennzahl (`gesamt`), je Störer-Familie und je Erzeuger gegen den letzten
  Bericht desselben Korpus (oder gegen Bericht `NNN`) – im Feld `vergleich` des Berichts, als Tabelle im `bericht.md` und in der Konsole
  (steigen die stillen Fehler, steht dort „ACHTUNG Stop-Regel“). `vergleichbar: false` bei anderem Korpus oder anderer Generator-Version.
  Unbekannte Nummer bei `--gegen`: Exit-Code 2.
* **Korpora:** `--korpus entwicklung|ex_test|test`. Bei `ex_test` steht nur der Hinweis „offen als Entwicklung erklärt“ (kein
  Regel-11-Warnton mehr); bei `test` bleibt der Regel-11-Hinweis.

### Ampel-Werte
`ok` · `NICHT ok` · `nicht gemessen` · `nicht vollständig gemessen` (es gab abgestürzte Paare) · `Hinweis` (nur Plausibilität).

### Exit-Code von `python -m src.cli messen`
* **0** = die Messung ist gelungen und **kein** Kriterium steht auf „NICHT ok“ (`gesamt.kriterien_nicht_ok` = 0, `gesamt.ziel_verfehlt` = false). „nicht gemessen“ und „nicht vollständig gemessen“ zählen dabei **nicht** als „NICHT ok“; sie stehen aber deutlich in Konsole und Bericht, die Gesamtzeile sagt nur „kein Kriterium verfehlt“.
* **1** = die Messung ist gelungen, aber mindestens ein Kriterium steht auf „NICHT ok“ (Ziel verfehlt). Ein Skript oder eine CI-Stufe kann damit auf Erfolg prüfen.
* **2** = kaputte Eingabe (Ordner oder Paar nicht gefunden, unbekannte Option). Es wird dann kein Bericht geschrieben.

### Paare ohne Profil: Standard-Aufbau wird angenommen
Die Regel „Abdeckung ohne Nachfrage mindestens 90 % bei Standard-Aufbauten“ galt früher nur für Paare mit `erwartet.json` und `profil: "standard"`. Echte Paare haben keine `erwartet.json`, also auch kein Profil; dieselben 75 % waren dort „ok“ und bei einem erfundenen Standard-Paar „NICHT ok“, die Messung war bei echten Paaren milder. **Entscheidung (K-MESS2):** Ein Paar ohne `erwartet.json` oder ohne Feld `profil` wird wie ein Standard-Paar behandelt und bekommt dieselbe Abdeckungs-Ampel (Konsole und `detail` sagen „Standard-Aufbau angenommen“; `gesamt.standard_paare_angenommen` zählt sie). Nur `profil: "schwer"` ist davon ausgenommen. Das ist bewusst die strengere Annahme: Ob ein echtes Paar ein „schwerer“ Aufbau ist, kann das Werkzeug nicht wissen, und im Zweifel soll die Messung nicht besser aussehen als sie ist.

### Ohne `erwartet.json` (alle echten Paare)
Die Prüfung „erwartete Lücke nicht gemeldet“ liest `erwartet.json`. Fehlt die Datei (oder das Feld `luecken`), kann eine nicht gemeldete Lücke (Scan, Zwei-Spalten, Tabelle, Bild) **nie auffallen**: Kriterium 1 steht dann auf „nicht gemessen“ (nicht „ok“), Konsole und Bericht sagen das ausdrücklich (`luecken_pruefbar`, `luecken_nicht_pruefbar_paare`).
Zusätzlich prüft die Messung das PDF **unabhängig vom Tool** direkt mit `pdfplumber`: Seiten ohne Textzeichen und Seiten mit einem großen eingebetteten Bild (`plausibilitaet_hinweise`: Seite, Art, ob das Tool dazu eine Lücke gemeldet hat; `plausibilitaet_ungemeldet`). Das ist **nur ein Hinweis**, nie ein Fehler.
