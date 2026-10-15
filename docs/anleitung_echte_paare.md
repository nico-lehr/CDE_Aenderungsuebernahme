# Anleitung: Echte Beispiele (Paare) sicher lokal einspeisen und messen

Autor: Nico Lehr · Stand: 15.10.2026 · für Windows, in einfachen Worten

Ziel: Das Tool an **echten** Dokumenten messen, ohne dass echte Inhalte irgendwo landen, wo sie nicht hingehören.
Ein **Paar** ist ein echtes Beispiel mit bekannter Wahrheit: das PDF mit den Änderungen, die Excel **vorher** und die von Hand erstellte Excel **nachher**.

## Die drei Sicherheitsregeln (bitte vorher lesen)

1. **Echte Dokumente bleiben lokal** auf deinem Rechner (Regel 16, Regel 17, Regel H). Nie in Git, nie in die Cloud, nie in einen Chat, nie als Anhang einer E-Mail an ein KI-Werkzeug.
   Der Ordner `data/` ist von Git ignoriert (`.gitignore`), PDF- und Excel-Dateien ebenso. Verlasse dich trotzdem nicht darauf: **Kopiere echte Dateien nur in `data/paare/…`.**
2. **Cloud-Sitzungen von Claude sehen nur erfundene Daten.** Schreibe in einen Prompt nie Inhalte aus echten Paaren. Beschreibe Probleme nur mit Zahlen
   (Blattnummer, Zeile, Spalte, Fehlerart aus dem Bericht), nie mit Textstellen.
3. **Teile die Paare auf, bevor du sie ansiehst.** Sonst stellst du das Tool unbewusst auf genau die Beispiele ein, an denen du es später testest (Regel 11).

## Schritt 1: Ordner anlegen

Im Projektordner (dort, wo `README.md` liegt) gibt es den Ordner `data`. Falls es ihn noch nicht gibt, legt ihn das Tool beim ersten Lauf an. Du brauchst:

```
data\paare\entwicklung\     ← hier darf geprüft und eingestellt werden
data\paare\test\            ← nur zur Endmessung; nie zum Einstellen
```

Beide Ordner kannst du im Windows-Explorer anlegen (Rechtsklick, Neu, Ordner).

## Schritt 2: Paare aufteilen (bevor du sie ansiehst)

Du hast 3 bis 5 echte Paare (mehr ist besser, aber 3 sind das Minimum).

- Würfle oder lose, **ohne** die Inhalte anzusehen. Beispiel bei 4 Paaren: Paar A und B kommen nach `entwicklung`, Paar C und D nach `test`. Bei 3 Paaren: zwei nach `entwicklung`, eines nach `test`.
- Wichtig: **Nicht** nach „schwierig/einfach“ aufteilen. Das wäre schon ein Ansehen. Eine Münze oder die Reihenfolge der Dateinamen genügt.
- Notiere dir die Aufteilung nur mit den Namen `echt_01`, `echt_02` … (keine Kundennamen, keine Projektnamen).

## Schritt 3: Dateien benennen und ablegen

Je Paar ein Ordner. Der Ordnername beginnt mit `echt_` und einer Nummer: `echt_01`, `echt_02`, … Darin liegen **genau diese drei Dateien mit genau diesen Namen**:

```
data\paare\entwicklung\echt_01\aenderungen.pdf    ← das PDF mit den farbigen Änderungen
data\paare\entwicklung\echt_01\vorher.xlsx        ← die Excel vor den Änderungen (Original)
data\paare\entwicklung\echt_01\nachher.xlsx       ← die von Hand erstellte Excel nach den Änderungen
```

Dasselbe für die Paare im Ordner `test`, z. B. `data\paare\test\echt_03\…` (die Nummern laufen einfach weiter, damit kein Name doppelt vorkommt).

Hinweise:
- **Kopiere** die Dateien (Original behalten), verschiebe sie nicht.
- Nur `.xlsx` (nicht `.xls`). Falls eine Datei `.xls` ist: in Excel „Speichern unter“ → „Excel-Arbeitsmappe (*.xlsx)“.
- Die `nachher.xlsx` muss wirklich das **Soll** sein: die Excel, wie sie nach dem manuellen Einarbeiten aussieht. Das Tool vergleicht sein Ergebnis damit.
- Keine Datei darf geöffnet sein, während das Tool läuft (Excel schließen).
- Eine Beschreibung des Formats steht in `docs/paar_format.md`. Erfundene Paare heißen `synth_…`; sie bleiben davon getrennt.

## Schritt 4: Messen

Öffne die Eingabeaufforderung oder PowerShell im Projektordner und schalte die Python-Umgebung ein (immer in der `.venv` arbeiten):

```
.venv\Scripts\activate
python -m src.cli messen --korpus entwicklung
```

Mit Vorschlägen (misst zusätzlich, wie gut die „prüfen“-Vorschläge sind, wenn man sie alle freigibt):

```
python -m src.cli messen --korpus entwicklung --mit-vorschlaegen
```

Nur ein Paar:

```
python -m src.cli messen --korpus entwicklung --paar echt_01
```

Die Messung verändert **keine** deiner Dateien. Sie dauert je nach Größe ein paar Sekunden bis Minuten.

Die **Endmessung** auf den Test-Paaren macht man erst ganz am Schluss und höchstens zweimal (Regel 11):

```
python -m src.cli messen --korpus test
```

Davor nie auf `test` messen, auch nicht „nur zum Schauen“.

## Schritt 5: Den Bericht lesen

Am Ende der Messung steht eine Zeile `Bericht: …`. Der Bericht liegt unter

```
data\messung\NNN_entwicklung\bericht.md      ← zum Lesen (öffnen mit Editor oder VS Code)
data\messung\NNN_entwicklung\bericht.json    ← dieselben Zahlen für Programme
```

`NNN` zählt hoch (001, 002, …). Alte Berichte werden **nie überschrieben**, so kannst du Vorher und Nachher vergleichen.
Der Bericht enthält **keine Texte aus deinen Dokumenten**, nur Zahlen, Paar-Namen, Blatt-Nummer, Zeile, Spalte, Änderungs-ID und Fehlerart. Du darfst die Zahlen daraus bedenkenlos weitergeben.

Oben im Bericht steht die **Ampel** (Tabelle „Kriterien“). Jede Zeile hat „ok“ oder „NICHT ok“ (oder „nicht gemessen“). Das bedeuten die Kriterien:

| Nr | Kriterium | Ziel | In einfachen Worten |
|---|---|---|---|
| 1 | Bilanz stimmig | alle Paare | Die Wörter-Summe passt, und was das Tool nicht lesen konnte (Scan-Seite, Tabelle …), steht sichtbar als Lücke da. „NICHT ok“: Das Tool hat etwas übersehen und nichts davon gesagt, oder ein Paar konnte nicht gemessen werden. |
| 2 | **Stille Fehler** | **0** | Das Wichtigste. Ein stiller Fehler heißt: Das Tool meldet „übernommen“, aber die Excel ist an der Stelle falsch oder unverändert. Alles über 0 ist ein Problem, das zuerst gelöst wird. |
| 3 | Treffergenauigkeit automatisch | 100 % | Was das Tool **ohne Nachfrage** einträgt, muss richtig sein. |
| 3 | Genauigkeit der Vorschläge | mindestens 95 % | Wenn du alle „prüfen“-Vorschläge freigibst: Wie viele sind richtig? (Nur mit `--mit-vorschlaegen`.) |
| 4 | Abdeckung ohne Nachfrage | mindestens 70 % (Standard-Aufbauten 90 %) | Wie viel der nötigen Änderungen das Tool von allein richtig erledigt. Hoch ist gut, aber nie auf Kosten von Nr. 2 und 3. |
| 4 | Nicht-Übernahmen mit Grund | 0 ohne Grund | Wenn das Tool etwas nicht übernimmt, muss es sagen, warum. |
| 5 | Lese-Lücke | höchstens 2 % | Anteil der Wörter, die im PDF stehen, aber nicht erkannt wurden. Der Rest muss als Warnung sichtbar sein. |
| 6 | Zweiter Lauf ändert nichts | alle Paare | Läuft das Tool ein zweites Mal auf seinem eigenen Ergebnis, darf nichts mehr passieren (kein doppeltes Einfügen). |
| 6 | Kollateralschaden | 0 | Keine Zelle wurde ohne Grund verändert. |
| 6 | Originale unverändert | alle Paare | Deine Original-Excel bleibt Byte für Byte gleich. |

Weiter unten im Bericht:
- **Fehlerarten** sagen, *wo* etwas schiefging: `pdf_falsch_gelesen` (PDF falsch gelesen), `layout_falsch` (Excel-Aufbau falsch erkannt), `zuordnung_falsch` (falsche Stelle getroffen),
  `plan_falsch` (falsch geplant), `bilanz_kaputt` (Bilanz stimmt nicht), `unklar`.
- **Je Störer-Familie** gibt es nur bei erfundenen Paaren (bei `echt_…` fehlt die Angabe).
- Die Tabelle **Paare** zeigt je Paar, welche Stelle (Blatt, Zeile, Spalte, Änderungs-ID) betroffen ist. Mit dieser Angabe findest du die Stelle selbst in deiner `vorher.xlsx`, ohne dass ein Text weitergegeben wird.

## Schritt 6: Was tun mit dem Ergebnis?

- **Alles „ok“?** Gut, aber lies es ehrlich: Mit wenigen Paaren sagt „0 stille Fehler“ nicht, dass es in Zukunft nie passiert (`docs/ziel_v1.md`, „Ehrliche Grenze“).
- **Etwas „NICHT ok“?** Schau zuerst in die **Fehlerart**:
  - Excel-Aufbau falsch erkannt (`layout_falsch`)? Prüfe mit `python -m src.cli profil vorschau --excel data\paare\entwicklung\echt_01\vorher.xlsx`, was erkannt wird (zeigt keine Zellinhalte).
    Wenn es falsch ist, kannst du in `config\layouts.local.yaml` ein **Profil** festlegen (Vorlage: `config\layouts.yaml`). Neue Profile gehen nie direkt live: erst Vorschau, dann Bestätigung
    (`docs/adr/0003-layout-profile.md`).
  - Farben im PDF falsch gedeutet (`pdf_falsch_gelesen`)? `python -m src.cli farben --pdf data\paare\entwicklung\echt_01\aenderungen.pdf` zeigt die Farben und welche unbekannt blieben (nur Zahlen und Farbcodes).
  - Danach ggf. Schwellen einstellen, **nur an `entwicklung`** und immer nur einen Wert auf einmal (`docs/kalibrierung.md`).
- **Probleme melden** (z. B. an die Entwicklung): nur Zahlen und Fehlerart, nie Textstellen, nie echte Namen.

## Kurzliste zum Abhaken

- [ ] Paare **vor dem Ansehen** aufgeteilt (entwicklung / test)
- [ ] Ordner `echt_01` … mit `aenderungen.pdf`, `vorher.xlsx`, `nachher.xlsx`
- [ ] Nichts davon in Git, Cloud, Chat oder E-Mail
- [ ] `python -m src.cli messen --korpus entwicklung` läuft
- [ ] Bericht gelesen (Ampel, Fehlerarten)
- [ ] Test-Ordner unberührt bis zur Endmessung
