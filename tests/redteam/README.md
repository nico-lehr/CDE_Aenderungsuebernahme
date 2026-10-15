# Red-Team-Tests (erfundene Daten, Regel 16)

> **Übersicht über alle Red-Team-Dateien** (Stand 15.10.2026). Die Tabelle weiter unten (RT-01 bis RT-53) stammt von Red-Team R1 (Excel, Planer, API).
> Dort steht in der Spalte „Kurz-Fehlermeldung“, **was damals** gemeldet wurde; die meisten dieser Fälle sind inzwischen **repariert** (Marker entfernt, der Test ist ein normaler Schutz-Test).
> Welche Fälle **heute noch offen** sind, steht in `docs/status.md`. Weitere Tabellen:
> - `README_pdf.md` – R2, PDF-Leser (RT-60 bis RT-109)
> - `README_korpus.md` – R3, Ursachen der stillen Fehler im Entwicklungskorpus und Schwächen der Messung (RT-110 bis RT-142)
>
> Besitzer: **K-PL** = planner/applier/protokoll · **K-XL** = excel_model/matcher · **K-API** = runs/api · **K-PDF** = pdf_reader · **Mess-Bauer/K-MESS** = messen.

Diese Tests brechen das Tool absichtlich mit gemeinen, **erfundenen** Fällen. Sie reparieren nichts: `src/` bleibt unberührt.

## So funktioniert es

- Jeder Fehlerfall ist ein Test mit `@pytest.mark.redteam` und `@pytest.mark.xfail(strict=True, raises=AssertionError, reason="RT-nn …; Besitzer K-…")`.
- Normaler Lauf (`python -m pytest -q`): xfail zählt als „erwartet“, alles bleibt grün.
- **Wird ein Fehler repariert, wird der Test XPASS und damit rot.** Dann muss der Besitzer (K-PL, K-XL oder K-API) den `xfail`-Marker dieses Tests löschen. Danach ist es ein normaler Schutz-Test.
- Beweis, dass ein Test aus dem richtigen Grund fällt: `python -m pytest --runxfail tests/redteam -q --tb=line` (oder einzeln mit `-k rt07`).
- Fehler im Testaufbau werfen bewusst `RuntimeError` (nicht `AssertionError`). So kann ein kaputter Aufbau nicht als „erwarteter Fehler“ durchgehen.
- Nur gesamt: `python -m pytest -q -m redteam` (nur Fehlerfälle) oder `-m "not redteam"` (ohne).
- Hilfsfunktionen: `_bau.py` (baut Excel/PDF-Fälle, ändert nichts in `tests/fixtures/`). `conftest.py` meldet den Marker `redteam` an.
- Dateien: `test_pl_planer_applier.py`, `test_pl_excel_eigenschaften.py` (K-PL) · `test_xl_modell_matcher.py` (K-XL) · `test_api_laeufe.py` (K-API) · `test_waechter.py` (Wächter, ohne xfail).

Besitzer: **K-PL** = planner/applier/protokoll · **K-XL** = excel_model/matcher · **K-API** = runs/api.

## Fehlerfälle (xfail)

Kurz-Fehlermeldung = was `--runxfail` heute meldet (leicht gekürzt).

### K-PL

| RT | Fall (ein Satz) | Besitzer | Datei::Test | Kurz-Fehlermeldung mit `--runxfail` |
|---|---|---|---|---|
| RT-01 | Neue Section, nur der Folgeabsatz wird freigegeben: er gilt als „übernommen/ok“, steht aber nicht in der Excel. | K-PL | test_pl_planer_applier.py::test_rt01_folgeabsatz_allein_freigegeben_ohne_wirkung | Folgeabsatz hat Status 'uebernommen' (Nachkontrolle: 'ok'), steht aber nicht in der Excel |
| RT-02 | Neue Section, nur die Überschrift wird freigegeben: der Folgeabsatz steht trotzdem in der Excel, sein Status bleibt „pruefen“. | K-PL | test_pl_planer_applier.py::test_rt02_nur_ueberschrift_freigegeben_folgeabsatz_schleicht_mit | Folgeabsatz steht in der Excel, obwohl er nicht freigegeben wurde (Status 'pruefen') |
| RT-03 | Zweiter Lauf mit demselben PDF auf dem Ergebnis: ein eingefügter Satz wird nochmal eingetragen (hoch/ok). | K-PL | test_pl_planer_applier.py::test_rt03_zweiter_lauf_text_eingefuegt_nicht_doppelt | Satz steht nach dem zweiten Lauf 2-mal in der Excel (Status uebernommen/hoch) |
| RT-04 | Zweiter Lauf: neuer Absatz (Modell „Zelle“) steht danach doppelt. | K-PL | test_pl_planer_applier.py::test_rt04_zweiter_lauf_absatz_neu_in_zelle_nicht_doppelt | Neuer Absatz steht nach dem zweiten Lauf 2-mal in der Excel |
| RT-05 | Zweiter Lauf: neuer Absatz (Modell „eine Zeile je Absatz“) steht danach doppelt. | K-PL | test_pl_planer_applier.py::test_rt05_zweiter_lauf_absatz_neu_als_zeile_nicht_doppelt | Neuer Absatz steht nach dem zweiten Lauf 2-mal in der Excel |
| RT-06 | Gestrichenes „•“ mitten im Text: „übernommen/ok“, aber nur ein Leerzeichen eingefügt, das „•“ bleibt. | K-PL | test_pl_planer_applier.py::test_rt06_gestrichener_punkt_mitten_im_text | Status 'uebernommen'/ok, aber Zelle ist jetzt 'Das System läuft stabil im Betrieb  • und liefert Daten.' |
| RT-07 | Gestrichenes „•“ am Absatzanfang: dasselbe, nur Doppel-Leerzeichen. | K-PL | test_pl_planer_applier.py::test_rt07_gestrichener_punkt_am_anfang | Status 'uebernommen'/ok, aber Zelle ist jetzt '•  Das System läuft stabil …' |
| RT-08 | Neuer Absatz „=SUMME(A1:A9)“ wird in der Excel zur Formel. | K-PL | test_pl_planer_applier.py::test_rt08_neutext_mit_gleichheitszeichen_wird_keine_formel | In der Excel steht jetzt eine Formel statt Text: [('C4', '=SUMME(A1:A9)')] |
| RT-09 | Ersetzter Zellanfang „=Dieses …“ wird zur Formel. | K-PL | test_pl_planer_applier.py::test_rt09_ersetzter_zellanfang_mit_gleichheitszeichen_wird_keine_formel | In der Excel steht jetzt eine Formel statt Text: [('C2', '=Dieses Dokument …')] |
| RT-10 | Steuerzeichen (BEL) im Neutext einer plan.json: Absturz beim Übernehmen, leerer `ergebnis_001` bleibt liegen. | K-PL | test_pl_planer_applier.py::test_rt10_steuerzeichen_im_neutext_stuerzt_nicht_ab | Absturz beim Uebernehmen: IllegalCharacterError … cannot be used in worksheets.; halbfertige Ordner: ['ergebnis_001'] |
| RT-11 | Neutext mit „(cid:12)“ (kaputte PDF-Schrift) wird ungeprüft „übernommen/hoch“. | K-PL | test_pl_planer_applier.py::test_rt11_cid_muell_wird_nicht_automatisch_uebernommen | '(cid:12)' steht jetzt in der Excel, Status uebernommen/hoch, Grund '' |
| RT-12 | Neutext mit Ligatur „ﬁ“ wird roh und ungeprüft „übernommen/hoch“. | K-PL | test_pl_planer_applier.py::test_rt12_ligatur_wird_nicht_ungeprueft_geschrieben | Ligatur 'ﬁ' steht roh in der Excel (Status uebernommen/hoch, Grund '') |
| RT-13 | Ganzer Absatz gestrichen (Excel-Zelle mit mehreren Absätzen): ungestrichene Wörter zwischen erstem und letztem Wort gehen verloren. | K-PL | test_pl_planer_applier.py::test_rt13_14_zelle_leeren_verliert_nicht_gestrichene_woerter[L1-13] | Status 'uebernommen'/hoch, aber die nicht gestrichenen Woerter 'zusätzlich'/'noch' sind weg |
| RT-14 | Dasselbe, wenn der Absatz eine eigene Zeile hat (`zelle_leeren`). | K-PL | test_pl_planer_applier.py::test_rt13_14_zelle_leeren_verliert_nicht_gestrichene_woerter[L3-14] | Status 'uebernommen'/hoch, aber die nicht gestrichenen Woerter 'zusätzlich'/'noch' sind weg |
| RT-15 | Einfügen mit veraltetem Plan landet mitten im Wort; Nachkontrolle meldet „ok“, weil der kurze Neutext („s“, „2.2“) irgendwo in der Zelle vorkommt. | K-PL | test_pl_planer_applier.py::test_rt15_nachkontrolle_erkennt_falsche_stelle[s] und [2.2] | Einfuegung mitten im Wort ('Das Gehä suse steht …'), aber Status 'uebernommen' und Nachkontrolle 'ok' |
| RT-16 | Einfügen (seq 2) und Ersetzen (seq 1) mit gleichem `start`: die Einfügung geht verloren, der Text wird zerstückelt. | K-PL | test_pl_planer_applier.py::test_rt16_ersetzen_und_einfuegen_gleicher_start | … ist aber: 'Das Gerätystem muss laufen und stabil bleiben.' |
| RT-17 | Neue Zeile direkt unter der Überschriftszeile übernimmt das Überschriftsformat (fett, groß, Füllfarbe). | K-PL | test_pl_planer_applier.py::test_rt17_neue_zeile_nach_ueberschrift_hat_absatzformat | Neue Zeile 3 hat Ueberschriftsformat: fett=True, Groesse=14.0, Fuellung=00CCCCFF |
| RT-18 | Zeile einfügen: eine verbundene Zelle über der Einfügestelle (D2:D4) wird nicht mitgezogen, Section 3 fällt heraus. | K-PL | test_pl_excel_eigenschaften.py::test_rt18_verbundene_zelle_ueber_der_einfuegestelle | Verbundene Zelle D2:D4 deckt Section 3 nicht mehr ab -- neue Zeile wurde trotzdem automatisch eingetragen (Status uebernommen/hoch) |
| RT-19 | Zeile einfügen: das `hidden`-Flag bleibt an der alten Zeilennummer, danach ist die falsche Zeile versteckt. | K-PL | test_pl_excel_eigenschaften.py::test_rt19_versteckte_zeile_bleibt_versteckt | Die versteckte Zeile (Section 5) ist nicht mehr versteckt / eine andere Zeile ist versteckt -- … (Status uebernommen/hoch) |
| RT-20 | Zeile einfügen: der Hyperlink bleibt an der alten Zelladresse. | K-PL | test_pl_excel_eigenschaften.py::test_rt20_hyperlink_wandert_mit | Der Hyperlink steht nicht mehr an der Zelle von Section 5 -- … (Status uebernommen/hoch) |
| RT-21 | Zeile einfügen: bedingte Formatierung (D2:D7) deckt die letzte Zeile nicht mehr ab. | K-PL | test_pl_excel_eigenschaften.py::test_rt21_bedingte_formatierung_wird_verschoben | Bedingte Formatierung D2:D7 deckt die letzte Zeile (Section 6) nicht mehr ab -- … |
| RT-22 | Zeile einfügen: Datenvalidierung (Dropdown D2:D7) deckt die letzte Zeile nicht mehr ab. | K-PL | test_pl_excel_eigenschaften.py::test_rt22_datenvalidierung_wird_verschoben | Datenvalidierung D2:D7 deckt die letzte Zeile (Section 6) nicht mehr ab -- … |
| RT-23 | Zeile einfügen: Autofilter (A1:D7) endet vor der letzten Zeile. | K-PL | test_pl_excel_eigenschaften.py::test_rt23_autofilter_wird_erweitert | Autofilter A1:D7 endet vor der letzten Zeile -- … |
| RT-24 | Zeile einfügen: Excel-Tabelle (`Table`, A1:D7) endet vor der letzten Zeile. | K-PL | test_pl_excel_eigenschaften.py::test_rt24_excel_tabelle_wird_erweitert | Excel-Tabelle A1:D7 endet vor der letzten Zeile -- … |
| RT-25 | Zeile einfügen: benannter Bereich „Daten“ (A2:D7) zeigt auf zu wenige Zeilen. | K-PL | test_pl_excel_eigenschaften.py::test_rt25_benannter_bereich_wird_erweitert | Benannter Bereich 'Daten' (A2:D7) endet vor der letzten Zeile -- … |
| RT-26 | Zeile einfügen: eine Formel `=C7` zeigt danach auf die falsche Zeile; es gibt nur eine allgemeine Warnung, die Änderung läuft trotzdem automatisch. | K-PL | test_pl_excel_eigenschaften.py::test_rt26_formel_nach_unten_wird_angepasst | Formel =C7 zeigt nach dem Einfuegen nicht mehr auf Section 6 -- … Warnungen: ['1 Formelzellen werden nicht verändert; …'] |
| RT-53 | Ersetztes Zeichen ohne Normform („•“ → „-“, „.“ → „!“): das Neue wird nur eingefügt, das Alte bleibt stehen (übernommen/ok). | K-PL | test_pl_planer_applier.py::test_rt53_ersetztes_zeichen_ohne_normform_bleibt_nicht_stehen[•--] und [.-!] | Status 'uebernommen'/ok, aber das alte Zeichen steht noch da: '… Betrieb - • und liefert Daten.' |

### K-XL

| RT | Fall (ein Satz) | Besitzer | Datei::Test | Kurz-Fehlermeldung mit `--runxfail` |
|---|---|---|---|---|
| RT-27 | Derselbe Absatz steht zweimal in einer Zelle, im PDF ändert sich der zweite: es wird der erste „hoch“ geändert. | K-XL | test_xl_modell_matcher.py::test_rt27_gleicher_absatz_zweimal_in_derselben_zelle | Status uebernommen/hoch: erster Absatz '… gemeinsam geprüft.', zweiter Absatz '… einzeln geprüft.' |
| RT-28 | PDF-Block ist nur der Anfang eines längeren Excel-Absatzes: „hoch“, der neue Absatz landet mitten im Satz. | K-XL | test_xl_modell_matcher.py::test_rt28_pdf_block_nur_anfang_des_excel_absatzes | Status uebernommen/hoch: neuer Absatz steht mitten im alten Absatz: '… danach\nNeuer Absatz nach dem Anfang. dokumentiert; …' |
| RT-29 | Section als Zahl: 2.1 und 2.10 (Excel zeigt 2.10 per Zahlenformat) werden beide zu „2.1“. | K-XL | test_xl_modell_matcher.py::test_rt29_section_2_1_und_2_10_als_zahl_bleiben_verschieden | Zeile 3 (angezeigt 2.1) und Zeile 4 (angezeigt 2.10) haben dieselbe Section '2.1' |
| RT-30 | Eine Zahl-Zelle (Section 2.2) wird beim Ersetzen zu Text; das Zahlenformat geht verloren. | K-XL | test_xl_modell_matcher.py::test_rt30_zahl_bleibt_beim_ersetzen_eine_zahl | Zelle A5 war eine Zahl, ist jetzt 's' ('2.3'), Format 0.0 |
| RT-31 | Kommentarspalte mit mehr Text gilt als Inhaltsspalte: der neue Absatz landet im Kommentar (hoch). | K-XL | test_xl_modell_matcher.py::test_rt31_kommentarspalte_ist_nicht_die_inhaltsspalte | Neuer Absatz steht in der Kommentar-Zelle D3 (… ), Status uebernommen/hoch |
| RT-32 | Spalte „Lfd“ (1..n) schlägt die echte Section-Spalte, keine Warnung. | K-XL | test_xl_modell_matcher.py::test_rt32_lfd_spalte_ist_nicht_die_section_spalte | Aufbau falsch erkannt: {'section_spalte': 1, 'inhalt_spalte': 2}, richtig waere {'section_spalte': 2}; keine Warnung ([]) |
| RT-33 | Aufzählungen „1. Das System …“ im Text werden als Sections gelesen (Text- und Nummernspalte vertauscht), keine Warnung. | K-XL | test_xl_modell_matcher.py::test_rt33_aufzaehlung_im_text_ist_keine_section | Aufbau falsch erkannt: {'section_spalte': 2, 'inhalt_spalte': 1}, richtig waere {'section_spalte': 1, 'inhalt_spalte': 2}; keine Warnung ([]) |
| RT-34 | Gleicher Text in zwei Blättern (Section 3 in Blatt A, Section 2 in Blatt B): die Section-Nummer wird blattübergreifend verglichen, Treffer „hoch“. | K-XL | test_xl_modell_matcher.py::test_rt34_section_nummer_gilt_nur_innerhalb_eines_blattes | Der Satz steht in zwei Blaettern; das PDF sagt nicht, welches gemeint ist, trotzdem uebernommen/hoch in Blatt B!C3 |
| RT-35 | Die Änderung wird automatisch in ein verstecktes Tabellenblatt geschrieben, ohne Warnung. | K-XL | test_xl_modell_matcher.py::test_rt35_verstecktes_blatt_wird_nicht_automatisch_geaendert | Aenderung automatisch (uebernommen/hoch) in das versteckte Blatt Archiv!C3 geschrieben, ohne Warnung |
| RT-36 | Die Änderung wird automatisch in eine versteckte Zeile geschrieben, ohne Warnung. | K-XL | test_xl_modell_matcher.py::test_rt36_versteckte_zeile_wird_nicht_automatisch_geaendert | Aenderung automatisch (uebernommen/hoch) in die versteckte Zeile CDE!C3 geschrieben, ohne Warnung |
| RT-37 | Eine formatierte Zelle bei Zeile 60000/Spalte 40: Einlesen dauert ~9 s statt unter 3 s. | K-XL | test_xl_modell_matcher.py::test_rt37_leeres_riesiges_blatt_wird_schnell_gelesen | Einlesen dauerte 8.8 s (Grenze 3 s) |
| RT-38 | Dieselbe Datei braucht ~525 MB Arbeitsspeicher beim Einlesen. | K-XL | test_xl_modell_matcher.py::test_rt38_leeres_riesiges_blatt_braucht_wenig_speicher | Spitzenspeicher 525 MB (Grenze 250 MB) |

### K-API

| RT | Fall (ein Satz) | Besitzer | Datei::Test | Kurz-Fehlermeldung mit `--runxfail` |
|---|---|---|---|---|
| RT-39 | Zwei gleichzeitige Übernahmen: die Ordnernummer „Anzahl+1“ ist nicht atomar. | K-API | test_api_laeufe.py::test_rt39_zwei_gleichzeitige_uebernahmen | Fehler: ["FileExistsError: [Errno 17] File exists: '…/ergebnis_001'"]; angelegte Ordner: ['ergebnis_001'] |
| RT-40 | Abgebrochene Übernahme lässt einen leeren `ergebnis_NNN` liegen, er zählt in der Lauf-Liste als Ergebnis. | K-API | test_api_laeufe.py::test_rt40_abgebrochene_uebernahme_hinterlaesst_keinen_halben_ordner | Halbfertige Ordner ['ergebnis_001']; Lauf-Liste zaehlt 1 Ergebnis(se), obwohl keines fertig ist |
| RT-41 | `plan.json` wird nicht atomar geschrieben: ein Leser sieht eine halbe Datei. | K-API | test_api_laeufe.py::test_rt41_plan_json_nicht_atomar | Leser sieht halb geschriebene plan.json -> JSONDecodeError: Unterminated string … |
| RT-42 | `DELETE /api/lauf/<id>` löscht den Lauf von der Platte (Regel 29). | K-API | test_api_laeufe.py::test_rt42_delete_loescht_lauf_nicht_von_der_platte | Nach DELETE (Status 200) ist nichts mehr auf der Platte (kein Archiv) |
| RT-43 | Eine fehlgeschlagene Analyse löscht den ganzen Lauf samt hochgeladenen Dateien (Regel 29). | K-API | test_api_laeufe.py::test_rt43_fehlgeschlagene_analyse_loescht_sich_nicht | Nach der fehlgeschlagenen Analyse (Status 500) sind die hochgeladenen Dateien weg |
| RT-44 | Upload über dem Größenlimit: englische HTML-Fehlerseite statt JSON mit deutscher Meldung. | K-API | test_api_laeufe.py::test_rt44_zu_grosser_upload_liefert_json | Antwort ist text/html; charset=utf-8: b'<!doctype html> … 413 Request Entity Too Large' |
| RT-45 | Fehlermeldung der Analyse enthält den Rohtext der Ausnahme. | K-API | test_api_laeufe.py::test_rt45_analysefehler_zeigt_keinen_rohtext | Meldung enthaelt Rohtext: 'Die Analyse ist fehlgeschlagen: ROHTEXT_XYZ No /Root object! …' |
| RT-46 | Fehlermeldung beim Übernehmen enthält den Rohtext der Ausnahme. | K-API | test_api_laeufe.py::test_rt46_uebernahmefehler_zeigt_keinen_rohtext | Meldung enthaelt Rohtext: 'Das Übernehmen ist fehlgeschlagen: ROHTEXT_XYZ unexpected keyword argument …' |
| RT-47 | `schema_version` 99 in der plan.json wird nicht geprüft: 500 mit Python-Fehlertext statt klarer Meldung. | K-API | test_api_laeufe.py::test_rt47_plan_json_aus_der_zukunft_gibt_klare_meldung | Status 500, Meldung: "Das Übernehmen ist fehlgeschlagen: Aenderung.__init__() got an unexpected keyword argument …" |
| RT-48 | POST mit fremdem `Origin:` (fremde Webseite) wird angenommen. | K-API | test_api_laeufe.py::test_rt48_fremder_origin_wird_abgelehnt | Anfrage von Origin http://boese.example wurde angenommen (Status 200) |
| RT-49 | Anfrage mit fremdem `Host:` (DNS-Rebinding) wird angenommen. | K-API | test_api_laeufe.py::test_rt49_fremder_host_wird_abgelehnt | Anfrage mit Host boese.example wurde angenommen (Status 200) |
| RT-50 | Windows-Reservenamen (CON, NUL, COM1, AUX, LPT1, PRN) werden als Dateiname gespeichert. | K-API | test_api_laeufe.py::test_rt50_windows_reservenamen[3 Varianten] | Gespeicherte Dateinamen sind unter Windows nicht anlegbar: ['CON.pdf', 'NUL.xlsx'] |
| RT-51 | Sehr langer Dateiname beim Upload: 500 (HTML) und ein halber Lauf-Ordner bleibt liegen. | K-API | test_api_laeufe.py::test_rt51_sehr_langer_dateiname | Status 500 (text/html; charset=utf-8); Lauf-Ordner ohne Analyse: ['2026…'] |
| RT-52 | Eine einzige kaputte `plan.json` bringt die ganze Lauf-Liste zum Absturz (500), auch die guten Läufe sind weg. | K-API | test_api_laeufe.py::test_rt52_ein_kaputter_lauf_zerstoert_nicht_die_liste | Lauf-Liste: Status 500 (text/html; charset=utf-8), sichtbare Laeufe: [] |

## Wächter (laufen heute grün und müssen grün bleiben)

In `test_waechter.py` (34 Tests, ohne xfail):

- Original-Excel bleibt byte-gleich (Hash) – direkt über `uebernehmen` (3 Layouts) und über die API.
- Bilanz stimmt in schwierigen Fällen (Aufzählungszeichen, Zelle leeren, doppelte Absätze, `(cid:…)`, unbekannter Absatz); jede „nicht übernommen“-Änderung hat einen Grund.
- Determinismus (zweimal laufen = gleiches Ergebnis) und 5 Zufallsdokumente mit festem Seed (`random.Random(seed)`): Bilanz stimmt, Ergebnis = Lösung, „übernommen“ heißt „Nachkontrolle ok“.
- Zweiter Lauf mit Ersetzungen ändert nichts mehr (Gegenstück zu RT-03..05).
- Zeilenhöhe und eine verbundene Zelle *unter* der Einfügestelle wandern korrekt mit.
- Einfügen (seq 1) vor Ersetzen (seq 2) am gleichen Start ist richtig (Gegenstück zu RT-16).
- Gleicher Text in zwei Blättern mit **gleicher** Section-Nummer bleibt „pruefen/mittel“ (Gegenstück zu RT-34).
- 20 000 gefüllte Zeilen: Gesamtlauf (Einlesen, Planen, Eintragen, Nachkontrolle) in unter 60 s (heute ≈ 12 s).
- API: zweite Übernahme legt neuen Ordner an und lässt den ersten und `plan.json` unberührt (append-only); Schema-1-`plan.json` (auch mit weniger optionalen Feldern) lädt und übernimmt; eigener Host/Origin (127.0.0.1:5100, localhost) wird angenommen; Pfad-Tricks im Upload-Dateinamen (`../../`, `..\\`, absolute Pfade, `:*?`) bleiben im Lauf-Ordner; Download-Pfad-Tricks werden abgelehnt.

## Fälle, die nicht (so) reproduziert wurden

- **Kein Limit je Datei beim Upload:** nicht getestet; es gibt nur das Gesamtlimit (200 MB). Ein Test bräuchte Uploads von über 100 MB. Der Teil „HTML statt JSON“ ist in RT-44 belegt.
- **20 000 gefüllte Zeilen:** fällt nicht; der Gesamtlauf dauert ≈ 12 s (Wächter). Auch 50 Sections mit Änderungen auf 20 000 Zeilen: ≈ 19 s.
- **Section „2“ in Blatt B bekommt Bonus von Blatt A:** bei *gleicher* Section-Nummer in beiden Blättern ist der Treffer korrekt „pruefen/mittel“ (Wächter). Der Fehler zeigt sich nur, wenn der Text in Blatt A unter einer *anderen* Section steht (RT-34).
- **Steuerzeichen aus dem PDF:** der PDF-Leser macht aus einem BEL-Zeichen den Text „(cid:7)“. Der Absturz (RT-10) ist deshalb nur mit einer von Hand veränderten `plan.json` erreichbar, nicht direkt aus einem PDF.
- **Ersetzen und Einfügen am gleichen `start` (RT-16) und Nachkontrolle bei falscher Stelle (RT-15):** über den Planer nicht erreichbar gefunden; die Tests bauen die Änderungen von Hand (veralteter Plan bzw. zwei Operationen am gleichen Start).
