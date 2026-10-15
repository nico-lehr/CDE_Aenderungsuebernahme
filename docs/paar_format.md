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
  "paare": [
    {"name": "synth_0001_standard", "bilanz_stimmt": true, "vollstaendig": true, "stille_fehler": 0,
     "fehler": [{"blatt": 1, "zeile": 12, "spalte": 3, "aenderung": "a7", "art": "zuordnung_falsch"}]}
  ]
}
```
Der Bericht nennt nur Zahlen, Paar-Namen, Blatt-Nummer, Zeile, Spalte, Änderungs-ID und Fehlerart.
