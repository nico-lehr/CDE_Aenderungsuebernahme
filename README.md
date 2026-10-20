# CDE-Änderungsübernahme

Autor: **Nico Lehr** · Stand: 20.10.2026 · Version 0.4.0 (Auftrag 003: Word-Datei und Word-PDF, CDE wie Nico, mehrere Specs, zweiter Prüfweg, Vertrauensmodus, einfaches Dashboard; nur mit erfundenen Daten getestet; Zahlen siehe `docs/status.md`, Bedienung `docs/so_gehts.md`)

Seitenprojekt, **getrennt von der Masterarbeit (bess-pipeline)**. Läuft lokal, ohne Internet, ohne KI-Dienst, ohne API-Key.

## Was macht das Tool?

Du hast ein **Word-Dokument (.docx) mit Änderungsverfolgung** (bevorzugt) oder ein **PDF mit sichtbaren Änderungen** (neu = farbig unterstrichen, entfällt = durchgestrichen, bis hin zu ganzen
Abschnitten) und eine **bestehende CDE-Excel** (nach Sections geordnet, Textblöcke dahinter, kein einheitliches Format).
Das Tool findet die Stellen in der Excel, übernimmt **nur die inhaltlichen Änderungen** und schreibt eine **neue** Excel.
Die Original-Excel bleibt unberührt. Dazu gibt es ein Protokoll, in dem jede Änderung steht.

**Grundsatz: lieber nicht übernehmen als falsch übernehmen.** Jede erkannte Änderung ist entweder *übernommen*,
*zu prüfen* oder *nicht übernommen (mit Grund)*. Zusätzlich meldet das Tool ehrlich, was es **nicht lesen konnte**
(Scan-Seiten, Zwei-Spalten-Text, Tabellen, Bilder, unbekannte Farben): dann steht in der Bilanz **„unvollständig“**, nie ein grünes „alles ok“.

## Starten (Windows)

1. Python **3.11** installiert? `py -3.11 --version`
2. Einmalig: `scripts\setup.bat` (legt `.venv` an und installiert die Pakete). Wenn pip im Firmennetz Zertifikatsfehler meldet:
   **nicht** mit `--trusted-host` umgehen, sondern IT fragen.
3. Dashboard: `scripts\start_dashboard.bat` → http://localhost:5100/ (bess-pipeline nutzt einen anderen Port, beide können parallel laufen).
4. Tests: `python -m pytest -q` (immer in der `.venv`).

## Befehle auf der Kommandozeile (`python -m src.cli …`)

| Befehl | Wofür |
|---|---|
| `analyse --docx A.docx --excel B.xlsx` (oder `--pdf A.pdf`) | nur auswerten, Bilanz anzeigen, nichts schreiben |
| `lauf --docx A.docx --excel B.xlsx` (oder `--pdf A.pdf`) | auswerten und alle **sicheren** Änderungen eintragen (`--alle-vorschlaege` nimmt auch Vorschläge mit, `--abwahl A0003,A0007` lässt diese automatischen Änderungen aus, `--spec "11 22 33"` wählt bei einer CDE mit mehreren Specs die Spec des PDFs) |
| `messen --korpus entwicklung` | **Messwerkzeug**: Paare (PDF + Excel vorher + nachher) ausführen und mit „nachher“ vergleichen — Treffergenauigkeit, Abdeckung, **stille Fehler**, Fehlerarten. Ausgabe nur Zahlen, nie Dokumentinhalt |
| `vertrauen` | wertet die Freigabe-Protokolle aller Läufe aus: je Prüfcode Anzahl, Freigabe- und Ablehnquote, Empfehlung (nur eine Hilfe, das Tool stellt nie von selbst um) |
| `steckbrief --docx A.docx` (oder `--pdf`) | Zahlen und Formen einer Datei, **kein Text** (zum Weitergeben ohne Inhalte) |
| `farben --pdf A.pdf` | welche Farben im PDF vorkommen und wie sie gedeutet werden (nur Zahlen und Farbcodes) |
| `profil liste / vorschau / bestaetigen / analyse` | Excel-Aufbau von Hand festlegen, wenn die Automatik unsicher ist. **Neue Profile gehen nie direkt live**: erst Vorschau, dann bewusste Bestätigung |

Echte Beispiele sicher einspeisen: `docs/anleitung_echte_paare.md`. Schwellen verstehen/kalibrieren: `docs/kalibrierung.md`.
Testdaten erzeugen (erfunden): `python scripts/erzeuge_paare.py --korpus entwicklung --seeds 1-60`.

## Dashboard

Eine Seite in **drei Schritten** (Anleitung für den Alltag: `docs/so_gehts.md`):
**① Dateien** (PDF und CDE-Excel ablegen, „Abgleichen“; darunter „Zuletzt“ mit den letzten Abgleichen, **Archivieren statt Löschen**) ·
**② Prüfen** (ein Satz mit Fortschritt; Karten „Bitte prüfen“ mit der Vorschau, wie die Zelle in der Excel aussieht – alt hellblau durchgestrichen,
neu rot –, PDF-Ausschnitt und Grund, je Karte „Übernehmen“ / „Nicht übernehmen“, per Tastatur J/K, A, N, ?; eingeklappt „Automatisch eingetragen“
mit „Trotzdem nicht eintragen“ und „Nicht in der CDE“; „Excel erstellen“) · **③ Fertig** (Excel und Protokoll herunterladen).
Enthält die CDE **mehrere Specs**, steht oben „Spec: 11 22 33 · Zeilen 120–380“ mit „Andere Spec wählen“ (neuer Abgleich mit denselben Dateien);
ist die Spec unklar, fragt das Dashboard zuerst „Welche Spec gehört zu diesem PDF?“ und trägt bis dahin nichts ein.
Alles Technische (Wort-Bilanz, Lücken, erkannter Aufbau, Farben im PDF, Prüfcodes) steht unten in „Details für Fachleute“; Warnungen (bei einer
Bilanz, die nicht stimmt, oder einem unvollständig gelesenen PDF) erscheinen oben. Bedienbar per Tastatur, hell und dunkel, Meldungen in einfachem Deutsch.

## Ordner

```
config/       paths.yaml, server.yaml (Port 5100), farben.yaml, zuordnung.yaml, layouts.yaml (Vorlage)  (alle Einstellungen, kein Code)
src/          pdf_reader, excel_model, matcher, planner, applier, protokoll, runs, messen, layouts, cli, api/
frontend/     Dashboard (HTML/CSS/JS)
scripts/      setup.bat, start_dashboard.bat, erzeuge_paare.py (erfundene Testpaare)
tests/        pytest; fixtures/ = Generator für erfundene PDFs/Excel; redteam/ = gemeine Fälle (erwartete Fehler sind als xfail markiert)
docs/         ziel_v1.md (Ziel + Kriterien), status.md (was geht / was nicht), konzept_v1.md, paar_format.md, kalibrierung.md,
              anleitung_echte_paare.md, messung_baseline.md, adr/ (Entscheidungen), auftrag/, regeln/
data/         Läufe (runs/), Paare (paare/), Messberichte (messung/), Profile (profile/) — append-only, wird NICHT versioniert
```

## Wie gut ist es?

**Kurz und ehrlich:** Auf erfundenen Nachbauten in der Struktur deiner echten Word-PDFs und CDE-Dateien gibt es jetzt **0 stille Fehler** (Entwicklungsdaten, 60 Paare; auf einem wirklich ungesehenen Satz von 30 Paaren, einmal gemessen, ebenfalls 0). Vorschläge stimmen zu 98,8 %, 97,8 % der Zellen gehen ohne Nachfrage durch. Das ist **kein Versprechen für echte Dokumente**: Die Lese-Lücke liegt bei 4,1 % (Ziel 2 %), der Word-Leser ist nur an selbstgebauten Dateien getestet, und die zwei echten Beispielpaare (nur gemessen, nie gespeichert) zeigen, dass noch Handarbeit bleibt. Automatisch übernommene Änderungen bitte am Anfang stichprobenartig prüfen; das Dashboard legt dir dafür eine Stichprobe und eine Ampel vor.

Aktuelle Messwerte und ehrliche Grenzen stehen in **`docs/status.md`**. Die Ziele und Kriterien stehen in **`docs/ziel_v1.md`**.
Alle Zahlen stammen von **erfundenen** Beispielen; sie sind kein Versprechen für echte Dokumente.

## Bekannte Grenzen (ehrlich)

- Bei PDFs werden nur **einspaltige Text-PDFs** gelesen. Tabellen, Bilder, Zwei-Spalten-Text und **Scans** werden **erkannt und gemeldet**, aber nicht verarbeitet (kein OCR).
- Teil-Formatierung (z. B. einzelne fette Wörter) in geänderten Excel-Zellen kann verloren gehen.
- Formeln, bedingte Formatierung, Datenprüfung, Filter, Excel-Tabellen und verbundene Zellen an der Einfügestelle werden **nicht repariert**, sondern führen zu „prüfen“ mit Warnung.
- Der **Word-Leser** (.docx) kennt noch keine echte Word-Datei von Nico; `steckbrief --docx` und der Vergleich mit dem PDF sind das Sicherheitsnetz. Die Ampel-Schwelle im Vertrauensmodus ist ein Startwert.
- Automatisch Übernommenes lässt sich im Dashboard einzeln abwählen („Trotzdem nicht eintragen“) und auf der Kommandozeile mit `lauf --abwahl ID,ID`. Eine **fertige Excel** wird dadurch nie verändert: jede Abwahl oder Freigabe erzeugt ein neues Ergebnis.

## Stand Auftrag 002 und lokale Fortsetzung

Etappe 1 („wie ein sorgfältiger Mensch“: Gegenprobe gegen das PDF, genauerer Leser, Reißverschluss-Abgleich, zweiter Farb-Leser) ist fertig. Entwicklung: 4 stille Fehler bei 120 erfundenen Paaren
(vorher 31), 98 % der automatischen Einträge richtig. **Neu festgelegt, noch nicht eingebaut:** Farblogik (durchgestrichen = entfällt, rot oder blau ohne Strich = neu) und farbige Markierung in der Excel
(rot = neu, blau durchgestrichen = alt), siehe `docs/adr/0005-farblogik-und-excel-markierung.md`.
Roadmap: `docs/plan_auftrag_002.md` · Übergabe für die lokale Arbeit: `docs/uebergabe_lokal.md` · Steckbrief statt echter Inhalte: `docs/anleitung_steckbrief.md` · Fragebogen für Ollama: `docs/fragebogen_ollama.md` · Updates ohne Git: `docs/arbeiten_ohne_github_sync.md` · Prompt für eine lokale Claude-Code-Sitzung: `docs/prompt_lokale_sitzung.md`.

## Nächste Schritte

1. 3–5 echte Beispielpaare (PDF + Excel vorher + Excel nachher) vom anderen Team, **nach Freigabe**, nur lokal, getrennt in Entwicklungs- und Testbeispiele (`docs/anleitung_echte_paare.md`).
2. Messen (`python -m src.cli messen`), Schwellen nur an den Entwicklungsbeispielen kalibrieren, einmal mit den Testbeispielen messen.
3. Farblegende (später), Layout-Profile für Excel-Typen, die die Automatik nicht sicher erkennt, optional lokales LLM nur als Prüfer.

## Für Claude Code (online/Cloud)

Erst `CLAUDE.md` und `.claude/entwicklungsregeln.md` (bzw. `docs/regeln/`) lesen. **Nur erfundene Daten verwenden.** Keine echten Dokumente anfordern,
keine API-Keys einbauen, `bess-pipeline` nicht anfassen, Port 5100 behalten. Im Cloud-Container ist Python evtl. neuer als 3.11:
Code darf keine 3.12+-Syntax benötigen.
