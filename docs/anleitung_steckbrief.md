# Anleitung: Steckbrief eines echten Beispiels (nur Zahlen, Farben und Formen)

Autor: Nico Lehr · Stand: 16.10.2026 · für Windows, in einfachen Worten

## Wozu das Ganze?

Wir wollen das Tool an **echten** CDE-Beispielen verbessern. Die echten Dokumente dürfen aber nur bei dir bleiben (Regel 16 und 17).
Der **Steckbrief** ist ein Werkzeug, das **bei dir** läuft und ein Beispiel beschreibt, **ohne ein einziges Wort aus den Dokumenten zu verraten**.
Darin stehen nur:

- **Zahlen** (wie viele Seiten, Zeilen, Wörter, Änderungen ...),
- **Farbcodes** (z. B. `#ff0000` für Rot),
- **Formen** (der Text wird durch Zeichen ersetzt: Großbuchstabe → `A`, Kleinbuchstabe → `a`, Ziffer → `9`).

Mit dem Steckbrief erkennen wir, **wie** echte Änderungen aussehen: Welche Farben hat das PDF? Was ist durchgestrichen? Wie sieht die Änderung in deiner
„nachher“-Excel aus (Alt blau und durchgestrichen? Neu rot? Was steht dazwischen?). Daraus bauen wir danach erfundene Nachbau-Beispiele.

Der Steckbrief **misst nur, was wirklich so ist**. Er rät nichts.

## Einmal vorbereiten

Öffne die Eingabeaufforderung (oder PowerShell) im Projektordner (dort, wo `README.md` liegt) und schalte die Python-Umgebung ein:

```
.venv\Scripts\activate
```

Falls es die Umgebung noch nicht gibt: einmal `scripts\setup.bat` ausführen (siehe README).

## Die drei Befehle

**1. Einen Steckbrief für ein Beispiel (ein Projekt) erstellen**

```
python -m src.cli steckbrief --pdf C:\...\aenderungen.pdf --vorher C:\...\vorher.xlsx --nachher C:\...\nachher.xlsx --name projekt1
```

- `--pdf`: das PDF mit den Änderungen
- `--vorher`: die Excel **vor** den Änderungen
- `--nachher`: die von Hand erstellte Excel **nach** den Änderungen
- `--name`: ein **anonymer** Name für den Ordner, z. B. `projekt1`. Erlaubt sind nur kleine Buchstaben a–z, Ziffern, `_` und `-`.
  Wähle **nie** einen Kundennamen oder Projektnamen (siehe unten, „Namen“).

**2. Den Steckbrief prüfen, bevor du ihn teilst**

```
python -m src.cli steckbrief --pruefe data\steckbrief\001_projekt1\steckbrief.json
```

Antwort bei einer sauberen Datei:

```
OK: enthält nur Zahlen, Farbcodes und Formen
```

Sonst steht dort, **wo** etwas nicht passt (ohne den Inhalt zu wiederholen). Dann bitte die Datei **nicht** weitergeben.

**3. Mehrere Beispiele auf einmal**

Lege die Beispiele in Unterordner (je Unterordner genau diese drei Dateien: `aenderungen.pdf`, `vorher.xlsx`, `nachher.xlsx`, genau wie bei `messen`,
siehe `docs/paar_format.md`) und starte:

```
python -m src.cli steckbrief --ordner C:\...\meine_beispiele --name sammlung1
```

Das Tool schreibt für jedes Beispiel einen eigenen Steckbrief (`paar_01`, `paar_02` ...) **und** eine `zusammenfassung.json` mit den Summen und Häufigkeiten über alle Beispiele.
Die echten Ordner- und Dateinamen kommen nirgends vor.

## Wo liegt die Datei?

```
data\steckbrief\001_projekt1\steckbrief.json     ← alle Zahlen, für Programme
data\steckbrief\001_projekt1\steckbrief.md       ← dieselben Zahlen zum Lesen (öffnen mit Editor oder VS Code)
```

Bei mehreren Beispielen:

```
data\steckbrief\001_sammlung1\paar_01\steckbrief.json (und .md)
data\steckbrief\001_sammlung1\paar_02\...
data\steckbrief\001_sammlung1\zusammenfassung.json (und .md)
```

Die Nummer (`001`, `002` ...) zählt hoch. Alte Steckbriefe werden **nie überschrieben**. Der Ordner `data\` ist von Git ignoriert.
Mit `--ausgabe C:\anderer\ordner` kannst du einen anderen Zielordner wählen.

## Was steht drin? (und was nicht)

**Aus dem PDF:** Seitenzahl und -größen, Zeichen und Wörter je Seite, Schriftfamilien (nur gängige Namen wie `Arial`, sonst `anderer`), Schriftgrößen,
**Farben** (Farbraum + Farbcode + Zahl der Wörter, davon durchgestrichen / unterstrichen / hinterlegt), die **Legende**
(wie oft kommt „rot ohne Strich“, „blau mit Strich“ ... vor), Form der Striche (Länge, Höhe, Dicke), Kommentare und Links im PDF,
Tabellen/Bilder/mehrere Spalten/Scan-Seiten, Kopf- und Fußzeilen, Silbentrennung, Aufzählungszeichen, hochgestellte Zahlen, gedrehte Seiten.

**Aus der Excel „vorher“:** Zahl der Blätter (nur „Blatt 1, Blatt 2 ...“, nie die Namen), erkannte Spalten (Section = `A`, Titel = `B`, Inhalt = `C` ...),
wie sicher der Aufbau erkannt wurde, Zeilenarten, Zellen mit mehreren Absätzen, Wortzahlen je Absatz, Formeln (nur als **Muster**, z. B. `=LEN(C{r})`),
Spalten mit immer gleichem Wert (nur als Form), verbundene Zellen, versteckte Zeilen/Spalten, Datenprüfung, Filter, Tabellen, benannte Bereiche, Zahlenformat der Section-Spalte.

**Aus dem Vergleich „vorher“ ↔ „nachher“ (der wichtigste Teil):**

- die **Reihenfolge der Stücke** in geänderten Zellen, als Muster, z. B. `normal,alt,normal,neu,normal` (17-mal),
- die **Farbcodes** der Alt- und Neu-Stücke und ob sie durchgestrichen sind,
- was **zwischen Alt und Neu** steht (Leerzeichen, Zeilenumbruch, nichts) und ob Alt vor Neu steht,
- wie **neue Zeilen** aussehen (Farbe, welche Spalten gefüllt sind, ob gleichbleibende Werte übernommen wurden),
- was mit **gestrichenen Zeilen** passiert (bleiben sie stehen? Farbe? Inhalt noch da?),
- welche Zellen **ohne Markierung** einfach ersetzt wurden, welche **andere Farben** (z. B. gelb hinterlegt) haben,
- ob sich Zeilenhöhen oder Spaltenbreiten geändert haben.

**Beispiel-Auszug** (erfundene Zahlen):

```
"muster": { "normal,alt,normal,neu,normal": 17, "normal,alt,neu,normal": 2 },
"stueck_farben": { "alt": { "#0000ff": 19 }, "neu": { "#ff0000": 19 } },
"trenner": { " ": 17, "kein": 2 },
"schablonen": [ { "typ": "text_ersetzt",
    "excel_nachher": [ { "f": "-", "form": "Aaa Aaaaaaaa aaaaaa " }, { "f": "a", "form": "Aaaa" },
                       { "f": "-", "form": " " }, { "f": "n", "form": "Aaaaaaaa" }, { "f": "-", "form": " aa Aaaaaa." } ] } ]
```

Die **Formen** („Schablonen“, höchstens 40 Beispiele) zeigen die Umgebung einer Änderung nur als Zeichenmuster.
Bedeutung: `A` = Großbuchstabe, `a` = Kleinbuchstabe, `9` = Ziffer (Umlaute und ß zählen als Buchstaben), `·` = geschütztes Leerzeichen,
`⇥` = Tab, `¶` = Zeilenumbruch, Satzzeichen bleiben, jedes andere Zeichen wird `?`, ein Wort mit 30 oder mehr Zeichen wird `W…`.
Das Kürzel `f` sagt das Format: `-` normal, `a` alt (durchgestrichen), `n` neu.
**Ehrlich gesagt:** Aus einer Form sieht man die **Länge der Wörter, Satzzeichen und Groß-/Kleinschreibung**, aber nicht, welche Wörter es sind.

**Was NICHT drinsteht:** Texte aus PDF oder Excel, Dateinamen, Ordnernamen, Blattnamen, Spaltenköpfe, Zellinhalte, Kommentartexte,
Links, Autoren, Titel oder andere Metadaten. Das ist technisch abgesichert (siehe nächster Abschnitt).

## Wie prüfst du den Steckbrief?

1. **Selbst ansehen:** Öffne `steckbrief.md` (oder `steckbrief.json`) mit dem Editor. Du solltest nur Zahlen, Farbcodes wie `#ff0000`, Begriffe wie
   `text_ersetzt` und Formen wie `Aaa aaaa` sehen. Wenn dir irgendwo ein echtes Wort auffällt: **nicht weitergeben** und melden.
2. **Automatisch prüfen:** `python -m src.cli steckbrief --pruefe ...\steckbrief.json` (siehe oben). Die Prüfung vergleicht **jedes Feld und jeden Schlüssel** mit einer festen
   Positivliste (erlaubt ist nur, was ausdrücklich erlaubt ist). Alles andere fällt durch.
   Das Tool prüft **schon beim Schreiben** selbst und schreibt bei einem Verstoß **gar nichts**.

## Was schickst du uns?

**Nur** diese Dateien:

- `steckbrief.json` und `steckbrief.md` (und bei mehreren Beispielen `zusammenfassung.json`/`.md`),
- die **Zahlen** der Messung (`messen`, siehe unten) als Zahlen aus `bericht.md`/`bericht.json`.

**Nie** das PDF, nie die Excel-Dateien, nie Screenshots davon.

## Fehler-Schablonen in der Messung (`messen --schablonen`)

Wenn du die Messung laufen lässt, kannst du zusätzlich die **Formen der Fehler** bekommen:

```
python -m src.cli messen --korpus entwicklung --mit-vorschlaegen --schablonen
python -m src.cli messen --korpus entwicklung --paar echt_01 --schablonen
```

Im Bericht (`bericht.md`/`bericht.json`) steht dann ein Abschnitt **„Fehler-Schablonen“**: für jeden stillen Fehler und jeden falschen Vorschlag (höchstens 40) die Form von
**Soll** (so sollte die Zelle aussehen) und **Ist** (so sieht sie im Ergebnis aus), dazu Fehlerart, Änderungstyp und Erzeuger. Wieder **ohne Klartext**.
Ohne `--schablonen` bleibt der Bericht genau wie bisher.

## Wichtig: Namen

- **Dateinamen, Ordnernamen und Blattnamen können echte Namen enthalten** (Kunde, Projekt). Der Steckbrief schreibt sie **nicht** in die Ergebnisdateien und
  gibt sie auch **nicht** in der Konsole aus. Trotzdem: Wähle bei `--name` einen **anonymen** Namen (`projekt1`, `bsp_a`), denn dieser Name steht im Ordnernamen und in der Datei.
- Gib **keine** Fehlermeldungen oder Bildschirmfotos weiter, auf denen Pfade zu sehen sind.
- Die Sicherheitsregeln aus `docs/anleitung_echte_paare.md` gelten unverändert: echte Dokumente bleiben lokal, nie in Git, Cloud, Chat oder E-Mail.

## Wenn etwas schiefgeht

| Meldung | Was tun |
|---|---|
| „Dieses Paar konnte nicht gelesen werden“ | Sind es unbeschädigte Dateien (PDF, `.xlsx`)? Excel schließen? |
| „Der Steckbrief verstößt gegen die Positivliste und wird NICHT geschrieben“ | Es wurde nichts geschrieben. Bitte melden (den Text der Meldung, nicht die Dateien). |
| „Der Name darf nur ...“ | `--name` nur mit kleinen Buchstaben, Ziffern, `_`, `-`. |
| `--pruefe` sagt „NICHT in Ordnung“ | Datei nicht weitergeben, melden. |

## Kurzliste zum Abhaken

- [ ] `--name` ist anonym gewählt
- [ ] Steckbrief erstellt, `steckbrief.md` angesehen: nur Zahlen, Farbcodes, Formen
- [ ] `--pruefe` zeigt „OK: enthält nur Zahlen, Farbcodes und Formen“
- [ ] Weitergegeben werden nur `steckbrief.json`/`.md` (und die Zahlen aus `messen`), nie PDF oder Excel
