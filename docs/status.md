# Status – was geht, was nicht, was fehlt

Autor: Nico Lehr · Stand: 15.10.2026 · Version 0.2.0 · **Alle Zahlen stammen von erfundenen Beispielen.** Sie sind kein Versprechen für echte Dokumente.

## In einem Satz

Das Tool ist von „kann still falsch sein“ zu „meldet fast alles ehrlich“ gekommen: Auf dem erfundenen Entwicklungskorpus sind die stillen Fehler
von **64 auf 5** gesunken (Ziel war 0), und die Messung selbst wurde so gehärtet, dass sie Fehler nicht mehr versteckt. **Das Ziel 0 ist noch nicht erreicht.**
Ob es bei echten Dokumenten trägt, zeigt erst die Messung an deinen echten Paaren (`docs/anleitung_echte_paare.md`).

## Messwerte (erfundener Entwicklungskorpus, 60 Paare, Generator-Version 1)

| Nr | Kriterium (`docs/ziel_v1.md`) | Ziel | Baseline | Jetzt | Ampel |
|---|---|---|---|---|---|
| 1 | Bilanz stimmig / Lücken gemeldet | alle Paare | 60 von 60 stimmig, **17 Lücken nicht gemeldet** | 60 von 60, 0 ungemeldet | ok |
| 2 | **Stille Fehler** | 0 | **64** | **5** (davon 4 nur Leerzeichen/Zeichenform) | NICHT ok |
| 2 | Regel 27: Vorschlag ohne Freigabe angewendet | 0 | nicht gemessen | 0 | ok |
| 3 | Treffergenauigkeit automatisch (Zellen) | 100 % | 92,3 % | 98,5 % (601 von 610) | NICHT ok |
| 3 | je Änderung (automatisch übernommen) | 100 % | nicht gemessen | 98,5 % (765 von 777) | NICHT ok |
| 3 | Genauigkeit der Vorschläge nach Freigabe | ≥ 95 % | 60,9 % | 72,3 % | NICHT ok |
| 4 | Abdeckung ohne Nachfrage | ≥ 70 % (Standard 90 %) | 77,1 % | 88,2 % | ok |
| 4 | Nicht-Übernahmen ohne Grund | 0 | 0 | 0 | ok |
| 5 | Lese-Lücke | ≤ 2 % | 5,8 % | 1,4 % | ok |
| 6 | Zweiter Lauf ändert nichts (Idempotenz) | alle Paare | 0 von 60 | 51 von 60 | NICHT ok |
| 6 | Kollateralschaden | 0 | 2 | 0 | ok |
| 6 | Struktur unverändert (Blätter, versteckte Zeilen, Merges …) | 0 | nicht gemessen | 0 | ok |
| 6 | Originale unverändert | alle | 60 von 60 | 60 von 60 | ok |
| – | Formelzellen in neuen/verschobenen Zeilen | – | nicht gemessen | 82 von 82 **nicht angepasst** (bewusst, siehe unten) | NICHT ok |
| 7 | Geschwindigkeit | 20 000 Zeilen < 60 s | Einzelzelle weit unten: 6,7 s / 525 MB | 0,01 s / 27 MB; 20 000 Zeilen Analyse 2,4 s | ok |
| 8 | Tests | alle grün | 26 | **914 grün**, 1 bekannter Fehlfall (xfail) | ok |

**Wichtig beim Lesen:** Die Zahlen der Baseline und von „Jetzt“ stammen aus demselben Messwerkzeug-Prinzip, aber das Werkzeug wurde zwischendurch
**strenger und genauer** gemacht (es zählte anfangs 16 Phantom-Fehler und versteckte echte). „Soll-Zellen“ sind deshalb nicht 1:1 vergleichbar
(755 früher, 681 jetzt, weil Formelzellen und harmlose Füllspalten getrennt gezählt werden). Die Richtung stimmt trotzdem: weniger stille Fehler bei strengerer Messung.

**Testkorpus (Seeds ≥ 1000, nie zum Einstellen benutzt, Regel 11):** wird zum Abschluss genau einmal gemessen. Ergebnis: siehe Abschnitt „Testkorpus“ unten.

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
| 5 stille Fehler im Entwicklungskorpus | alle `plan_falsch`; 4 davon nur Leerzeichen/Zeilenumbruch-Form (ein Satz am Zeilenumbruch wird als eigener Absatz eingetragen; Leerzeichen vor geschütztem Leerzeichen) | Text ist inhaltlich richtig, Form stimmt nicht exakt |
| Idempotenz 51 von 60 | in 9 Paaren ändert ein zweiter Lauf auf dem Ergebnis noch etwas | Risiko für Doppel-Einträge, wenn jemand dasselbe PDF zweimal auf dasselbe Ergebnis anwendet |
| Vorschlags-Genauigkeit 72 % | Vorschläge („prüfen“) sind nach Freigabe oft nicht exakt | Mensch muss prüfen — dafür sind sie Vorschläge |
| `RT-88` (xfail) | der Reader ist repariert (Separation-Farben); der Test selbst ist fehlerhaft gebaut (setzt im PDF schwarzen Text) | kein Tool-Fehler |
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

*(wird beim Abschluss eingetragen)*
