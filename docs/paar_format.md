# Paar-Format und Messbericht (Vertrag zwischen Daten-Bauer und Mess-Bauer)

Autor: Nico Lehr · Stand: 15.10.2026

Ein **Paar** ist ein Beispiel mit bekannter Wahrheit: ein PDF mit Änderungen, die Excel **vorher** und die von Hand erstellte Excel **nachher**.
Das Messwerkzeug (`python -m src.cli messen`) lässt das Tool auf PDF + vorher laufen und vergleicht das Ergebnis mit „nachher“.

## Ordner (alles unter `data/`, von Git ignoriert – `.gitignore`)

```
data/paare/
  entwicklung/                 ← hier darf geprüft und eingestellt werden
    KORPUS.json                ← nur bei erzeugten Paaren: Generator-Version, Seeds
    synth_0001_standard/       ← Namen beginnen mit synth_ (erfunden) oder echt_ (Nicos Beispiele)
      aenderungen.pdf
      vorher.xlsx
      nachher.xlsx
      erwartet.json            ← NUR bei synth_, optional (Zusatzwissen des Generators)
  test/                        ← nur zur Endmessung (Regel 11), nie zum Einstellen
    ...
data/messung/
  001_entwicklung/bericht.json, bericht.md     ← append-only, nie überschreiben (Regel 29)
```
Pfade kommen aus `config/paths.yaml` (`paare_dir`, `messung_dir`), nie hartcodiert (Regel 28). Der Name `echt_…` ist nur für Nicos echte Beispiele, die **lokal** bleiben.

## Das Messwerkzeug braucht `erwartet.json` NICHT
Das **Soll** wird immer aus `vorher.xlsx` ↔ `nachher.xlsx` abgeleitet (Zellvergleich, Absatzausrichtung). So läuft dasselbe Werkzeug später unverändert mit echten Paaren.
`erwartet.json` liefert nur Zusatzwissen für die Auswertung (Störer-Familien, Lücken, Anzahl markierter Wörter).

### `erwartet.json` (nur erzeugte Paare; alle Inhalte erfunden)
```json
{
  "format": 1,
  "generator_version": "1",
  "seed": 12,
  "profil": "standard",
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
- `soll_status`: was ein sauberes Tool tun sollte: `uebernommen` (eindeutig, automatisch), `pruefen` (absichtlich mehrdeutig) oder `nicht_uebernommen` (absichtlich nicht machbar, z. B. Tabelle).
- `luecken`: was das Tool als Lücke **melden muss** (`models.LUECKEN_ARTEN`).
- Der Generator schreibt die Wahrheit aus seinem **logischen Modell**, nie aus dem PDF-Leser.

### `KORPUS.json` (je Korpus-Ordner, nur erzeugte Paare)
```json
{"generator_version": "1", "korpus": "entwicklung", "seeds": [1, 2, 3], "erstellt": "2026-10-15"}
```
Zahlen verschiedener Generator-Versionen sind **nicht vergleichbar**.

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
* **Struktur:** `struktur_geaendert` (mit `struktur_nach_art`: blatt, sichtbarkeit, zeile_versteckt, spalte_versteckt, verbunden, validierung, zahlenformat, schrift) zählt UNERWARTETE Abweichungen von Blättern, ausgeblendeten Zeilen/Spalten, verbundenen Zellen, Datenvalidierung, Zahlenformat und Schrift/Füllung unveränderter Zellen. Zeilen werden über die Ausrichtung abgebildet, das legitime Verschieben beim Einfügen zählt nicht.
* **Soll-Zellen außerhalb der Haupt-Spalten:** `soll_zellen` zählt nur die verwalteten Spalten (ohne Formeln). Getrennt: `soll_zellen_formel` (+ `…_verfehlt`), `soll_zellen_unverwaltet_sonstige` (+ `…_verfehlt`, z. B. neu nummerierte Spalte) und `soll_zellen_unverwaltet` (nur kurze Füllung neuer/entfernter Zeilen, aufgeschlüsselt in `…_aufschluesselung`). `nicht_verwaltet_verfehlt_orte` nennt nur Blatt/Zeile/Spalte/Klasse.
* **Absturz eines Messteils:** Stürzt nur der Wiederholungslauf (Idempotenz), die Vorschläge oder die Struktur-Prüfung ab, bleiben die Funde des ersten Laufs erhalten. Das Paar zählt dann als „teilweise gemessen“ (`teilweise_gemessen`, `ausgefallene_messteile`, `paare_teilweise_gemessen`) und in `paare_abgestuerzt`. Bei abgestürzten Paaren steht keine Ampel auf „ok“, sondern auf „nicht vollständig gemessen“ (außer sie ist schon „NICHT ok“).

### Ampel-Werte
`ok` · `NICHT ok` · `nicht gemessen` · `nicht vollständig gemessen` (es gab abgestürzte Paare) · `Hinweis` (nur Plausibilität).

### Ohne `erwartet.json` (alle echten Paare)
Die Prüfung „erwartete Lücke nicht gemeldet“ liest `erwartet.json`. Fehlt die Datei (oder das Feld `luecken`), kann eine nicht gemeldete Lücke (Scan, Zwei-Spalten, Tabelle, Bild) **nie auffallen**: Kriterium 1 steht dann auf „nicht gemessen“ (nicht „ok“), Konsole und Bericht sagen das ausdrücklich (`luecken_pruefbar`, `luecken_nicht_pruefbar_paare`).
Zusätzlich prüft die Messung das PDF **unabhängig vom Tool** direkt mit `pdfplumber`: Seiten ohne Textzeichen und Seiten mit einem großen eingebetteten Bild (`plausibilitaet_hinweise`: Seite, Art, ob das Tool dazu eine Lücke gemeldet hat; `plausibilitaet_ungemeldet`). Das ist **nur ein Hinweis**, nie ein Fehler.
