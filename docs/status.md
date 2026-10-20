# Status – was geht, was nicht, was fehlt

Autor: Nico Lehr · Stand: 20.10.2026 · **Version 0.4.0** (Auftrag 003, Welle 5) · Alle Zahlen stammen von erfundenen Nachbauten oder von zwei echten Beispielpaaren, die nur gemessen und nie gespeichert wurden.

## Stand Welle 5 (20.10.2026) – „Word-Datei, zweiter Prüfweg, Vertrauensmodus“

**Was neu ist (ADR 0008, 0009; `docs/so_gehts.md`):**
- **Word-Datei (.docx) lesen** (`src/docx_reader.py`): Einfügungen, Streichungen, Verschiebungen, Zeichen-genaue Änderungen, Listennummern aus `numbering.xml`, Tabellen, Kopf-/Fußzeilen. Dasselbe erfundene Dokument als .docx und als Word-PDF ergibt dieselben Änderungen.
  Zusammengeführte oder getrennte Absätze, Tabellen und Unbekanntes werden nie automatisch eingetragen, sondern gemeldet.
- **Zweiter Prüfweg** (`src/konsens.py`): unabhängige Prüfung des Ergebnisses (Zeichen für Zeichen; bei PDF zusätzlich im Seitenbild). Abweichung ⇒ „Bitte prüfen“, die Datei wird ohne die Änderung neu geschrieben.
- **Vertrauensmodus** (Dashboard, `src/vertrauen.py`): Dokument-Ampel, Stichprobe (etwa 10 %) zum Ansehen, Eingabe Word oder PDF, Freigabe-Protokoll je Lauf (nur Codes, keine Texte) und Auswertung `python -m src.cli vertrauen`.
- **Besser lesen und vorschlagen:** Listen-Einrückungen galten fälschlich als „mehrspaltig“ (487 richtige Änderungen waren gesperrt); Vorschläge bei abweichender Excel entstehen nur noch, wenn sie zeichengenau stimmen (sonst Code `kein_vorschlag`).

**Messung (erfundene Nachbauten):**

| Kennzahl | Ziel | `nachbau_word` (60, Entwicklung) vorher → jetzt | frischer Satz (Seeds 3001–3030): erste Messung → nach den Korrekturen |
|---|---|---|---|
| Stille Fehler | 0 | 12 → **0** | 3 → **0** |
| Automatisch richtig, Zellen / je Änderung | 100 % | 99,8 % / 99,4 % → **100 % / 100 %** (7131 von 7131; 4314 von 4314) | 99,9 % / 100,0 % → 100,0 % / 100,0 % (2820 von 2820) |
| Format-Treue (Farbe, Strich je Zeichen) | 100 % | 100 % → 100 % | 100,0 % → 100,0 % |
| Vorschläge exakt nach Freigabe | ≥ 95 % | 87,8 % → 98,8 % | 95,8 % → 97,2 % |
| Abdeckung ohne Nachfrage, Haupt-Spalten | ≥ 70 % | 85,9 % → 97,8 % | 97,5 % → 97,6 % |
| Bitte prüfen (Änderungen) | berichten | 619 → 115 (2,1 %) | 79 (2,0 %) |
| Kollateral (fremde Zelle geändert) | 0 | 0 → 0 | 2 → **0** |
| Fremde Spec berührt / Idempotenz | 0 / alle | 0 / 60 von 60 | 0 / 30 von 30 |
| Lese-Lücke | ≤ 2 % | 4,1 % → 4,1 % (davon nur 2 Wörter in Zellen, die im Ergebnis nicht stimmen) | 3,6 % → 3,6 % |
| Restarbeit je 100 Änderungen | berichten | 29,3 → 16,2 | 14,0 → 14,0 |

**Echte Beispielpaare** (nur Zahlen, im Scratchpad): keine Verschlechterung durch die neuen Teile. Der zweite Prüfweg schlug bei Beispiel B nicht an (517 Wörter geprüft, 0 Abweichungen) und stufte bei Beispiel A eine falsch eingetragene Änderung auf „Bitte prüfen“ zurück.

**Ehrlich lesen:**
- Der frische Satz ist eine **einmalige** Messung vor jeder Korrektur (3 stille Fehler, davon 2 „fremde Zelle geändert“). Danach wurden die Fehlerklassen **nur als Klassen** an einen Helfer gegeben, der sie auf Entwicklungsdaten
  (Seeds < 1000) gesucht hat. Ergebnis: 7 Zellen mit falsch gefärbtem Leerzeichen (zwischen neuer Listennummer und neuem erstem Wort) wurden im Tool behoben; die „fremde Zelle“ und eine „keine Wirkung“-Zelle waren **Messfehler**
  (die Ergebnis-Excel gleicht dem Soll Zelle für Zelle; die Messung hatte eine neue Listenzeile mit der ähnlichen alten gepaart). Die Messung wurde dafür in zwei engen Stellen korrigiert (`messen.py`: gepaarte neue Zeile, zwei Listennummern vor dem Polster).
  Die zweite Messung am frischen Satz (0 stille Fehler) ist deshalb **eine Gegenprobe**, kein ungesehener Test mehr. Ein wirklich ungesehener Satz fehlt noch (neue Seeds erzeugen, einmal messen).
- Die Lese-Lücke ist überwiegend ein Messartefakt: Rund 97 % der „nicht erkannten“ Wörter sind Section-Nummern neuer oder entfernter Zeilen, die das PDF nicht markiert und die das Tool aus dem Kontext richtig einträgt. Die Messung nennt deshalb zusätzlich die Zahl nur für Zellen, die im Ergebnis nicht stimmen (die alte Zahl bleibt).
- Der Word-Leser ist nur an selbstgebauten Dateien getestet (keine echte .docx vorhanden). Sicherheitsnetz: `steckbrief --docx` auf einer eigenen Datei und der Abgleich mit dem PDF (`vergleiche_mit_pdf`).
- Die Bildprobe prüft nur rund drei Viertel der Wörter (kleine Wörter, farbiger Grund und Satzzeichen sind „nicht prüfbar“, nie „falsch“).
- Die Ampel-Schwelle „rot ab 40 offenen Punkten je 100 Änderungen“ ist ein Startwert und gehört an echten Paaren eingestellt.

## Stand Auftrag 003 (19.10.2026) – „Word-PDFs, CDE wie Nico, einfaches Dashboard“

**Was neu ist (ADR 0005 Nachtrag, 0006, 0007; `docs/so_gehts.md`):**
- **Word-Änderungsverfolgung lesen:** Linie entscheidet (durchgestrichen = entfällt, farbig unterstrichen = neu, jede Farbe). Markup-Leiste mit Sprechblasen, Verbindungslinien, Änderungsbalken,
  Kopf-/Fußzeilen und Vorspann stören nicht mehr; markierte Leerzeichen, geänderte Listennummern („c.a.“) und hochgestellte Einheiten („mm²“) werden richtig gelesen.
- **CDE als Register:** Aufbau wird aus der Struktur erkannt (Kopfzeile, Beispielzeile, Section, Sub-Section, Text mit Listennummer). **Mehrere Specs in einer CDE:** nur die Zeilen der Spec des PDFs werden bearbeitet
  (Spec aus Titel-/Fußzeile; bei Unklarheit nichts automatisch, im Dashboard Auswahl). Änderungen aus Vorspann/Revisionstabelle sind „nicht in der CDE“ und kein offener Punkt.
- **Schreiben wie Nico:** alt = hellblau (00B0F0) + durchgestrichen, neu = rot, kein zusätzlicher Trenner, Grundschrift bleibt, ältere Markierungen bleiben. Die Datei wird **verlustfrei** geschrieben
  (`src/xlsx_patch.py`): Logos, Formen, Vertraulichkeits-Kennzeichnung, Druckeinstellungen bleiben, nur die geänderten Zellen ändern sich. Neue Zeilen: nur die Textspalte rot, Spezifikation/Revision/Datum aus dem Block der Spec.
- **Dashboard in 3 Schritten** (Dateien → Prüfen → Fertig), Karten wie bei „Überarbeiten“ in Word, Abwahl automatischer Einträge, Spec-Auswahl.

**Messung (erfundene Nachbauten in der Struktur der echten Dateien):**

| Kennzahl | Ziel | `nachbau_word` (60, Entwicklung) | `nachbau_test` (30, **ungesehen**, einmal gemessen) |
|---|---|---|---|
| Stille Fehler | 0 (ungesehen ≤ 1) | 12 (Start: über 2000) | **3** |
| Automatisch richtig, Zellen / je Änderung | 100 % | 99,8 % / 99,4 % | 99,9 % / 99,9 % |
| Format-Treue (Farbe, Strich je Zeichen) | 100 % | 100 % | 100 % |
| Vorschläge exakt nach Freigabe | ≥ 95 % | 87,8 % | 90,8 % |
| Abdeckung ohne Nachfrage, Haupt-Spalten | ≥ 70 % | 85,9 % | 84,9 % |
| Fremde Spec berührt / Idempotenz | 0 / alle | 0 / 60 von 60 | 0 / 30 von 30 |
| Lese-Lücke | ≤ 2 % | 4,1 % | 3,6 % |
| Restarbeit je 100 Änderungen | berichten | 29,3 | 29,2 |

**Zwei echte Beispielpaare** (nur Zahlen; Rückbau des Vorher-Stands aus Nicos fertiger CDE, dann Tool, dann Zellvergleich): Beispiel B: 51 von 60 geänderten Zellen exakt wie Nico (bis auf die Lage von Leerzeichen),
der Rest sind Abweichungen von Nicos Vorlage selbst bzw. Rückbau-Effekte; offene Prüfpunkte im Haupttext 2–6 je Dokument (vorher über 40); Beispiel A: 9 gleich, dazu 8 fertige Vorschläge zum Abhaken
(Excel weicht vom alten PDF-Text ab). Datei-Treue geprüft (Teile bytegleich, LibreOffice öffnet). Laufzeit etwa 10 Sekunden je Dokument.

**Ehrlich lesen:** Ungesehen = 3 stille Fehler (Ziel ≤ 1) und 1 Kollateralzelle: Die Fehler liegen bei den für den Test neu gebauten Sonderfällen (Silbentrennung am Zeilenende, Nummernwechsel über Seitenumbruch).
Lese-Lücke und Vorschlagsgenauigkeit liegen unter dem Ziel. Der ungesehene Korpus ist erfunden; echte Dokumente bleiben der eigentliche Test: Jede automatisch eingetragene Änderung am Anfang stichprobenartig prüfen.

**Bekannte Grenzen:** Tabellen, zwei Spalten, Scans im PDF werden nur erkannt und gemeldet; Word-Autorfarbe „Schwarz“ (RT-313), Hyperlink ohne Link-Kommentar (RT-305), einseitige PDFs ohne Kopfzeilen-Wiederholung (RT-306) als strikte xfail dokumentiert.
Ollama (Zweitmeinung) ist noch nicht angeschlossen. Kommentare/Gelb-Markierung und Zeilen-Löschen gibt es nur im alten openpyxl-Modus.

---
(Älterer Stand, Version 0.2.0 und Auftrag 002 Etappe 1, siehe unten.)

## In einem Satz

Das Tool ist von „kann still falsch sein“ zu „meldet fast alles ehrlich“ gekommen: Auf dem erfundenen **Entwicklungskorpus** (an dem wir gearbeitet haben)
sind die stillen Fehler von **64 auf 7** gesunken (Ziel war 0). Auf dem **Testkorpus** (nie zum Einstellen benutzt, einmal am Ende gemessen) sind es dagegen
**23 stille Fehler** bei 89,7 % Treffergenauigkeit. Der Abstand zeigt: Ein Teil der Reparaturen passt noch zu stark auf bekannte Fälle.
**Das Ziel „0 stille Fehler“ ist nicht erreicht.** Das Tool ist ein Prototyp, der ehrlich warnt, aber noch nicht „blind vertrauenswürdig“ ist:
Jede automatisch übernommene Änderung sollte stichprobenartig geprüft werden. Wie es bei echten Dokumenten abschneidet, zeigt erst die Messung an deinen
echten Paaren (`docs/anleitung_echte_paare.md`).

## Zwischenstand Auftrag 002, Etappe 1 (16.10.2026) – „wie ein sorgfältiger Mensch“

**Was neu ist (ADR 0004, `docs/ziel_v2.md`):** Das Tool trägt eine Änderung nur noch automatisch ein, wenn in der Excel genau der alte Absatz aus dem PDF steht (Gegenprobe A1).
Nach dem Speichern wird jeder betroffene Absatz gegen das PDF geprüft, nicht nur gegen den eigenen Plan (A3). Ein unklares Signal irgendwo im Absatz (unbekannte Farbe, halbe Markierung,
Tabelle, zwei Spalten …) stoppt die Automatik für den ganzen Absatz. Der PDF-Leser behält Satzzeichen, Trenner und Fußnoten-Hochzahlen richtig, erkennt Striche über ganze Zeilen,
gedrehte Seiten und Seitenumbrüche; ein zweiter Leser (pdfium) und eine Farb-Inventur finden Markierungen, die der Textleser übersieht. Schema 3 (`plan.json` speichert die Absätze).

**Neue, realistischere Testdaten:** Generator v2 (Phantom-Fehler mit geschütztem Leerzeichen behoben) und ein zweiter PDF-Erzeuger über LibreOffice (eingebettete Schriften, Blocksatz,
weiche Trennung, Kopf-/Fußzeilen). Korpora: `entwicklung` = 120 Paare (Seeds 1–60 und 101–160, 30 davon LibreOffice), `ex_test` = 30 Paare (früherer Testkorpus, jetzt offen als Entwicklung
erklärt). Ein ungesehener Test (Seeds 2001–2040) kommt erst am Ende von Etappe 2.

| Kennzahl | Ziel | alter Stand auf neuem Korpus¹ | jetzt: Entwicklung (120) | jetzt: ex_test (30) |
|---|---|---|---|---|
| Stille Fehler (nach Korrektur der Messartefakte²) | 0 | 31 | **4** | **0**² |
| Automatisch richtig, Zellen | 100 % | 96,2 % | 98,1 % | 95,8 % |
| Automatisch richtig, je Änderung | 100 % | – | 99,2 % (1562 von 1574) | – |
| Vorschläge exakt nach Freigabe | ≥ 95 % | 78,0 % | 87,1 % | 48,3 % |
| Abdeckung ohne Nachfrage, Haupt-Spalten | ≥ 90 % (Standard) | 83,7 % | 87,4 % | 85,6 % |
| Abdeckung über alle Soll-Zellen | ≥ 70 % | 51–60 % | 57,8 % | 62,8 % |
| Lese-Lücke | ≤ 2 % | 1,8 % | 2,2 % | 7,0 % |
| Zweiter Lauf ändert nichts | alle | 111/120 | 120/120 | 30/30 |
| Regel-27-Verstöße / Kollateral | 0 / 0 | 0 / 0 | 0 / 1 | 0 / 0 |
| Restarbeit je 100 Änderungen | berichten | 24,6 | 22,6 | 28,9 |

¹ Vor Etappe 1, gemessen mit dem alten Tool auf dem neuen Korpus (v2 + LibreOffice).
² **Ehrlich lesen:** Die 0 bei `ex_test` entsteht auch dadurch, dass Wörter in einer **gemeldeten** Farb-Lücke (grüne Schrift) nicht mehr als „still“ zählen – sie sind gemeldet und der Absatz ist gesperrt.
Die Genauigkeit der automatischen Einträge liegt dort trotzdem nur bei 95,8 %, und die Lese-Lücke bei 7,0 % (grüner Text wird weiter nicht gelesen, nur gemeldet).

**Noch offen:** die 4 stillen Fehler der Entwicklung (u. a. ein abgetrennter Satzteil als eigene Zeile, eine Füllzelle in der falschen neuen Zeile); Formeln und Füllspalten in neuen Zeilen
(Etappe 2: ZEILEN); schwache Vorschläge bei grüner/unbekannter Farbe (Leser, nicht Planer); Dashboard zeigt „PDF sagt / Excel hat“ und Abwahl noch nicht (Etappe 2: UI); Ollama (Etappe 2).
Alle Zahlen stammen von **erfundenen** Dokumenten.

## Messwerte (erfundener Entwicklungskorpus, 60 Paare, Generator-Version 1)

| Nr | Kriterium (`docs/ziel_v1.md`) | Ziel | Baseline | Jetzt | Ampel |
|---|---|---|---|---|---|
| 1 | Bilanz stimmig / Lücken gemeldet | alle Paare | 60 von 60 stimmig, **17 Lücken nicht gemeldet** | 60 von 60, 0 ungemeldet | ok |
| 2 | **Stille Fehler** | 0 | **64** | **7** (4 nur Leerzeichen/Zeichenform, 2 Zahl-wird-Text, 1 sonstige; Stand nach strengerer Messung) | NICHT ok |
| 2 | Regel 27: Vorschlag ohne Freigabe angewendet | 0 | nicht gemessen | 0 | ok |
| 3 | Treffergenauigkeit automatisch (Zellen) | 100 % | 92,3 % | 98,2 % (599 von 610) | NICHT ok |
| 3 | je Änderung (automatisch übernommen) | 100 % | nicht gemessen | 98,2 % (763 von 777) | NICHT ok |
| 3 | Genauigkeit der Vorschläge nach Freigabe | ≥ 95 % | 60,9 % | 72,3 % | NICHT ok |
| 4 | Abdeckung ohne Nachfrage (nur Haupt-Spalten, 681 von 995 Soll-Zellen; ohne Profil gilt jetzt auch die 90-%-Regel) | ≥ 70 % (Standard 90 %) | 77,1 % | 88,0 % (599 von 681; Standard-Paare 97,5 %) | ok |
| 4 | **Abdeckung über ALLE Soll-Zellen** (Haupt-Spalten, Füllung, sonstige, Formeln; Stand K-MESS2) | ≥ 70 % | nicht gemessen | **60,2 %** (599 von 995: Füllung 0 von 226, sonstige 0 von 6, Formeln 0 von 82) | **NICHT ok** |
| 4 | Nicht-Übernahmen ohne Grund | 0 | 0 | 0 | ok |
| 5 | Lese-Lücke | ≤ 2 % | 5,8 % | 1,4 % | ok |
| 6 | Zweiter Lauf ändert nichts (Idempotenz) | alle Paare | 0 von 60 | 57 von 60 (nach K-FINAL; vorher 51) | NICHT ok |
| 6 | Kollateralschaden | 0 | 2 | 0 | ok |
| 6 | Struktur unverändert. **Geprüft:** Blätter, versteckte Zeilen/Spalten, Verbünde, Validierung, bedingte Formatierung, Hyperlink-Adressen, Kommentare, benannte Bereiche, Filter, Excel-Tabellen, Fixierung, Spaltenbreiten, Zeilenhöhen, Zahlenformat, Schrift/Füllung, Datentyp. **Nicht geprüft:** Rahmen, Ausrichtung, Schutz, Druckeinstellungen, Bilder, Diagramme, Gliederung, Makros | 0 | nicht gemessen | 0 (Automatik und Excel mit allen freigegebenen Vorschlägen; 3 angekündigte Abweichungen dort sind kein Befund) | ok |
| 6 | Datentyp-Wechsel (Zahl wird Text) an richtig befüllten Zellen, neu seit K-MESS2 | 0 | nicht gemessen | 2 Zellen in 2 Paaren (neue Section-Nummern stehen im Soll als Zahl, das Tool schreibt Text); zählen als stille Fehler | NICHT ok |
| 6 | Originale unverändert | alle | 60 von 60 | 60 von 60 | ok |
| – | Formelzellen in neuen/verschobenen Zeilen | – | nicht gemessen | 82 von 82 **nicht angepasst** (bewusst, siehe unten) | NICHT ok |
| 7 | Geschwindigkeit | 20 000 Zeilen < 60 s | Einzelzelle weit unten: 6,7 s / 525 MB | 0,01 s / 27 MB; 20 000 Zeilen Analyse 2,4 s | ok |
| 8 | Tests | alle grün | 26 | **980 grün**, kein offener erwarteter Fehler | ok |

**Wichtig beim Lesen:** Die Zahlen der Baseline und von „Jetzt“ stammen aus demselben Messwerkzeug-Prinzip, aber das Werkzeug wurde zwischendurch
**strenger und genauer** gemacht (es zählte anfangs 16 Phantom-Fehler und versteckte echte). „Soll-Zellen“ sind deshalb nicht 1:1 vergleichbar
(755 früher, 681 jetzt, weil Formelzellen und harmlose Füllspalten getrennt gezählt werden). Die Richtung stimmt trotzdem: weniger stille Fehler bei strengerer Messung.

**Testkorpus (Seeds ≥ 1000, nie zum Einstellen benutzt, Regel 11):** am Ende genau einmal gemessen. Ergebnis: siehe Abschnitt „Testkorpus“ unten.

## Was geht (mit erfundenen Daten getestet)

- 4 PDF-Markierungsarten (Textfarbe, Hintergrund, Durchstreichen, Annotation), 3 Excel-Aufbauten, Wort-/Satz-/Absatz-/Section-Änderungen, Umnummerierung, mehrseitige PDFs.
- Verschmolzene Ersetzungen (`alt` direkt vor `neu`), Silbentrennung (auch in farbigen Wörtern), Fußnoten, Aufzählungen, Kopf-/Fußzeilen, gedrehte Seiten, CMYK/Grau/Sonderfarben.
- **Ehrliche Lücken:** Scan-Seiten, Zwei-Spalten-Text, Tabellen, Bilder, unbekannte Farben und unlesbare Zeichen werden als „Lücke“ gemeldet; die Bilanz heißt dann „unvollständig“.
- Unsichere Stellen werden **nie automatisch** eingetragen: doppelte Absätze, Anfangs-Treffer, unsicheres Excel-Layout, ausgeblendete Zeilen/Blätter, Formel-/Steuerzeichen im Neutext, Trennmuster, Excel-Eigenschaften an der Einfügestelle (Formeln, Tabellen, Validierung, Filter, benannte Bereiche, Merges durch die Stelle).
- Jede übernommene Änderung wird nach dem Speichern **nachgelesen** (`wirkung`): ohne nachweisbare Wirkung nie „übernommen“.
- Läufe sind append-only; **Archivieren statt Löschen**; atomares Schreiben; Schema 1 und 2 lesbar; Host-/Origin-Schutz; Upload-Limits; Fehlermeldungen in einfachem Deutsch.
- Layout-Profile (`config/layouts.yaml`) mit Vorschau → Bestätigung → aktiv, nie direkt live.
- Messwerkzeug `python -m src.cli messen`, Farbwerkzeug `farben`, Generator `scripts/erzeuge_paare.py`.
- Dashboard: Banner „unvollständig“, Hinweise/Wirkung je Änderung, Suche/Filter/Sortierung, Tastatur, hell/dunkel, Smartphone-Breite.

## Was nicht geht (bewusst, siehe `docs/ziel_v1.md`)

- **Keine Verarbeitung** von Zwei-Spalten-Text, Tabellen, Bildern und Scans (nur erkennen und melden; kein OCR).
- **Formeln werden nie angepasst** (weder verschoben noch neu geschrieben). Steht eine Formel unterhalb der Einfügestelle, wird die Änderung nur zum Vorschlag („prüfen“).
  Deshalb zeigt die Messung bei Formelzellen in neuen/verschobenen Zeilen „82 von 82 nicht angepasst“. Das ist erwartet, soll aber sichtbar bleiben.
- Teilformatierung in Zellen (einzelne fette Wörter) geht beim Ersetzen verloren (Warnung).
- Kein LLM; kein Zugriff auf echte Dokumente; Port 5100 fix.
- Automatisch Übernommenes lässt sich im Dashboard nicht einzeln abwählen (die Schnittstelle kennt nur Freigaben für Vorschläge).

## Was noch fehlt / offen

| Thema | Stand | Wirkung |
|---|---|---|
| 7 stille Fehler im Entwicklungskorpus | alle `plan_falsch`/Typ: 4 nur Leerzeichen/Zeilenumbruch-Form (ein Satz am Zeilenumbruch wird als eigener Absatz eingetragen; Leerzeichen vor geschütztem Leerzeichen), 2 neue Section-Nummern stehen im Soll als Zahl, das Tool schreibt Text | Text meist inhaltlich richtig, Form/Typ stimmt nicht exakt |
| Idempotenz 57 von 60 (Seeds 1–60); ungesehene Seeds 61–120: 54 von 60 | Behoben (K-FINAL): Steht der alte Satz in einer **falschen** Section eindeutig, wurde er dort automatisch als „übernommen / hoch / Wirkung ok“ geschrieben (Prüfer hat das bewiesen); die richtige Section blieb unberührt. Passt jetzt die Section des PDF-Blocks (alte Nummer) nicht zur gefundenen Excel-Stelle, ist der Treffer höchstens „mittel“ (Vorschlag); Umnummerierung ist nicht betroffen. Messung vorher → nachher: Seeds 1–60 51 → 57 von 60, Seeds 61–120 49 → 54 von 60; Abdeckung ohne Nachfrage und stille Fehler unverändert. **Rest 3 bzw. 6 Paare:** derselbe Satz/dasselbe Wort steht **in derselben Section** (meist in derselben Zelle) mehrmals; der erste Lauf trifft die richtige Stelle, der zweite Lauf ändert die andere. | Ein zweiter Lauf mit demselben PDF auf dem **Ergebnis** kann automatisch eine andere Stelle ändern, die unverändert bleiben soll (nicht nur Doppel-Einträge). Ein Ergebnis nie ein zweites Mal mit demselben PDF bearbeiten; für eine neue Fassung die ursprüngliche Excel nehmen. Offen (Planer). |
| Vorschlags-Genauigkeit 72 % | Vorschläge („prüfen“) sind nach Freigabe oft nicht exakt | Mensch muss prüfen — dafür sind sie Vorschläge |
| `RT-88` | erledigt (K-FINAL): der Testaufbau war fehlerhaft (schwarzer Text statt Sonderfarbe) und ist repariert; der Test prüft jetzt die Separation-Farbe wirklich, der xfail-Marker ist entfernt | kein Tool-Fehler |
| Zwei-Spalten/Tabellen/Bilder/Scan ohne `erwartet.json` | bei **echten Paaren** kann die Messung „Lücke nicht gemeldet“ nie prüfen (nur Scan-Seiten und große Bilder per PDF-Plausibilität) | Kriterium 1 steht bei echten Paaren auf „nicht gemessen“ |
| Messung vertraut `bilanz()['stimmt']` des Tools | nicht unabhängig geprüft | Rest-Risiko |
| Gruppen-Freigabe bei Section-Löschung | die Kopplung gilt beim Planen; gibt ein Mensch nur Überschrift oder nur einen Absatz frei, kann das den Rest trennen | bräuchte ein Gruppenfeld (Schema 3) |
| Test-only-Familien nicht mehr ungesehen | „gedrehte Seite“ und „grüner Text“ (nur im Testkorpus) wurden als Red-Team-Fälle behandelt (RT-87, RT-93/94) | Endmessung belegt dort weniger Allgemeingültigkeit |
| Generator ist von uns | echte Dokumente sind anders; Schwellen (Farbe, Zeilenabstand, Absatz) sind nicht an echten PDFs kalibriert | Nicos echte Paare nötig (`docs/kalibrierung.md`) |

## Wie ihr es selbst nachmessen könnt

```
python scripts/erzeuge_paare.py --korpus entwicklung --seeds 1-60
python -m src.cli messen --korpus entwicklung --mit-vorschlaegen
```
Bericht (nur Zahlen, nie Inhalte): `data/messung/NNN_entwicklung/bericht.md`. Echte Paare: `docs/anleitung_echte_paare.md`.

## Testkorpus (einmalige Endmessung, Regel 11)

Erfundene Paare mit eigenen Seeds (1001–1030, 30 Paare, 312 Soll-Zellen in den Haupt-Spalten), nie zum Einstellen benutzt, **genau einmal** gemessen (Bericht lokal: `data/messung/007_test/`).
Danach wurde am Tool **nichts** mehr geändert. Wer den Testkorpus noch einmal als „ungesehen“ verwenden will, muss einen neuen mit neuen Seeds (und neuer Generator-Version) erzeugen.

| Kriterium | Ziel | Testkorpus | Entwicklungskorpus | Ampel (Test) |
|---|---|---|---|---|
| Bilanz stimmig | alle Paare | 30 von 30 | 60 von 60 | ok |
| **Stille Fehler** | 0 | **23** (alle „übernommen, aber falsch“; 4 nur Leerzeichen/Zeichenform) | 7 | NICHT ok |
| Treffergenauigkeit automatisch (Zellen / je Änderung) | 100 % | 89,7 % (252 von 281) / 89,2 % (330 von 370) | 98,2 % / 98,2 % | NICHT ok |
| Genauigkeit der Vorschläge | ≥ 95 % | 62,1 % | 72,3 % | NICHT ok |
| Abdeckung ohne Nachfrage (Haupt-Spalten) | ≥ 70 % (Standard 90 %) | 80,8 % | 88,0 % | NICHT ok (90-%-Regel) |
| Abdeckung über alle Soll-Zellen | ≥ 70 % | 58,6 % (252 von 430) | 60,2 % | NICHT ok |
| Lese-Lücke | ≤ 2 % | 7,0 % (237 von 3405 Wörtern) | 1,4 % | NICHT ok |
| Zweiter Lauf ändert nichts | alle Paare | 30 von 30 | 57 von 60 | ok |
| Regel 27 / Kollateral / Struktur / Originale | 0 / 0 / 0 / alle | 0 / 0 / 0 / 30 von 30 | 0 / 0 / 0 / 60 von 60 | ok |

**Wie ist das zu lesen?** Der Testkorpus ist deutlich schlechter als der Entwicklungskorpus. Das ist der Zweck eines getrennten Testkorpus: Er zeigt, wie viel der Reparaturen nur auf
bekannte Fälle passt. Die Fehler sind in Paaren mit grünem Text (eine Farbe, die nur im Testkorpus vorkommt), Silbentrennung, Kopf-/Fußzeilen und CMYK/Grau am häufigsten
(Familien überlappen, die Zahlen sind nicht addierbar). Es gab **keine** nicht gemeldeten Lücken und **keinen** Verstoß gegen Regel 27 — die Warnungen funktionieren, die Genauigkeit der
automatischen Übernahme tut es noch nicht in allen Fällen. Ich habe diese Ergebnisse bewusst **nicht** zum Nachbessern benutzt (Regel 11).

**Empfehlung:** (1) Bis die Genauigkeit an echten Paaren gemessen ist, alle automatisch übernommenen Änderungen im Protokoll stichprobenartig prüfen. (2) Die Fälle des Testkorpus
sind jetzt „gesehen“: als nächste Entwicklungsrunde nutzen und danach einen neuen, ungesehenen Testkorpus erzeugen. (3) Mit 3–5 echten Paaren messen und kalibrieren.
