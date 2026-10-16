# Red-Team R4 – Ende-zu-Ende-Fälle für Etappe 1 (RT-200 bis RT-216)

Stand 16.10.2026 (Welle 0). Nur erfundene Daten (Regel 16). `src/` bleibt unberührt.

**Weg in jedem Test** (wie das Tool): `analysiere_pdf` → `ExcelModell` → `plane` →
`uebernehmen(aend, vorher, ziel, freigaben, absaetze=plan.absaetze)`, danach Ergebnis-Excel und Status prüfen.

**Prüfregel: „richtig ODER sichtbar unsicher, nie still falsch“** (`_bau_r4.py`):
- `richtig_oder_unsicher`: Zelle ist exakt das erwartete Ergebnis, ODER die betroffene Änderung steht auf „pruefen“
  bzw. „nicht_uebernommen“ mit Grund und die Zelle ist unverändert. „pruefen“ zählt auch ohne Grundtext als sichtbar
  (ein Vorschlag mit Konfidenz „mittel“ hat heute keinen Grund).
- `erkannt_oder_gemeldet` (für Markierungen, die der Leser evtl. gar nicht sieht): wie oben, oder – wenn keine
  Änderung erkannt wurde – der Plan meldet eine Lücke bzw. eine „pruefen“-Änderung. Die allgemeine Warnung
  „keine Änderungen erkannt“ zählt nicht (in den SICHT-Tests steht deshalb zusätzlich eine normale blaue Änderung).
- Aufbau-Fehler werfen `RuntimeError`, nie `AssertionError`. Die SICHT-PDFs werden im Aufbau mit pdfium gerendert:
  die Stelle muss wirklich blau bzw. rot aussehen (`wirkt`).
- Fälle, die `tests/fixtures/pdf_writer.py` nicht bauen kann (Separation, DeviceN mit Funktion Typ 4, Kurven-Fläche,
  Render-Modus 1, Hochzahl, genau gesetzter Zeilenumbruch), baut der kleine Roh-PDF-Schreiber `RohDok` in `_bau_r4.py`.

Laufzeit aller R4-Tests: etwa 2–3 s. Beweis: `python -m pytest --runxfail tests/redteam/test_r4_*.py -q --tb=line`.

Besitzer (Welle 1/2/3 aus dem Plan): **PDF** = `pdf_reader` · **PLAN** = planner/matcher/abgleich · **SICHT** = zweiter
Leser/Farb-Inventur · **APPLY** = applier (Gegenprobe A3) · **ZEILEN** = neue Zeilen (Etappe 2).

| RT | Fall (ein Satz) | Besitzer | Datei::Test | Kurz-Fehlermeldung mit `--runxfail` |
|---|---|---|---|---|
| RT-200 | Blaues Wort mit direkt folgendem schwarzem Punkt („… die Anlage insgesamt.“): der Punkt wird mit eingefügt. | PDF/PLAN | test_r4_pdf.py::test_rt200_blaues_wort_vor_schwarzem_punkt | still falsch -- Status ['uebernommen'], Wirkung ['ok'], Zelle '… die Anlage. insgesamt.' |
| RT-201 | Rotes Wort mit schwarzem Punkt: der Punkt wird mit gelöscht. | PDF/PLAN | test_r4_pdf.py::test_rt201_rotes_wort_vor_schwarzem_punkt | still falsch -- Status ['uebernommen'], Wirkung ['ok'], Zelle '… die Anlage' statt '… die Anlage.' |
| RT-202 | Schwarze Linie (≥ 70 % Seitenbreite) quer durch alle Zeilen eines Absatzes: Absatz gilt still als unverändert. | PDF | test_r4_pdf.py::test_rt202_strich_ueber_ganze_zeilen | nicht erkannt und nicht gemeldet -- Zelle bleibt still '…'; Aenderungen: keine; Luecken: [] |
| RT-203 | Grünes Wort (unbekannte Farbe) im selben Absatz wie eine blaue Einfügung: die blaue wird trotzdem automatisch eingetragen (keine Absatz-Sperre). | PLAN | test_r4_plan.py::test_rt203_gruenes_wort_sperrt_den_absatz | blaue Aenderung trotz unerklaerter Farbe im selben Absatz automatisch eingetragen (…uebernommen/hoch…) |
| RT-204 | Excel weicht außerhalb der Änderung um ein Wort ab („Druckwert“ statt „Druckwerte“): trotzdem automatisch (A1 fehlt). | PLAN | test_r4_plan.py::test_rt204_excel_weicht_ausserhalb_ab_nicht_automatisch | alter PDF-Absatz steht so nicht in der Excel ('Druckwert'), trotzdem automatisch: …uebernommen/hoch |
| RT-205 | Geschütztes Leerzeichen in der Excel direkt an der Einfügestelle (3 Varianten) bleibt erhalten. | PLAN | test_r4_plan.py::test_rt205_nbsp_an_der_einfuegestelle_bleibt | **besteht schon** (Schutz-Test) |
| RT-206 | Blauer neuer Satz über einen Seitenumbruch: zwei Änderungen „hoch“, Zeilenumbruch mitten im Satz in der Zelle. | PDF/PLAN | test_r4_pdf.py::test_rt206_neuer_satz_ueber_seitenumbruch | still falsch -- Status ['uebernommen', 'uebernommen'], Zelle '… Zusätzlich\nwird jede einzelne Prüfung …' |
| RT-207 | Rote Hinterlegung als Separation-Farbe „Rot“ mit Ton 1,0: nicht erkannt, nicht gemeldet (Kontrolle in RGB besteht). | SICHT | test_r4_sicht.py::test_rt207_hinterlegung_in_separation_farbe[separation] | nicht erkannt und nicht gemeldet -- Zelle bleibt still '… Anlage täglich und …'; Luecken: [] |
| RT-208 | Blauer Text in DeviceN-Farbe mit 2 Kanälen: gilt still als unverändert (Kontrolle in CMYK besteht). | SICHT | test_r4_sicht.py::test_rt208_text_in_devicen_farbe[devicen] | nicht erkannt und nicht gemeldet -- Zelle bleibt still '… Anlage und meldet …'; Luecken: [] |
| RT-209 | Blaue Hinterlegung als Kurven-Pfad (abgerundetes Feld) statt Rechteck: gilt still als unverändert (Kontrolle mit Rechteck besteht). | SICHT | test_r4_sicht.py::test_rt209_markierung_als_kurvenflaeche[kurve] | nicht erkannt und nicht gemeldet -- …; Luecken: [] |
| RT-210 | Nur-Kontur-Text (Render-Modus 1) mit blauer Kontur: gilt still als unverändert (Kontrolle mit Füllfarbe besteht). | SICHT | test_r4_sicht.py::test_rt210_nur_kontur_text[kontur] | nicht erkannt und nicht gemeldet -- …; Luecken: [] |
| RT-211 | `=LEN(Cn)` je Zeile und `=SUM(D2:D5)` darunter, neue Section in der Mitte: heute „pruefen“ mit Grund (Formeln-Hinweis), Excel unverändert. | ZEILEN | test_r4_excel.py::test_rt211_formeln_und_summe_bei_neuer_section | **besteht schon** (Schutz-Test; prüft bei automatischem Eintrag auch alle Formeln und die Summe) |
| RT-212 | Section-Nummern als Zahl (Format „0.0“): neue Section 2.3 wird als Text „2.3“ geschrieben. | APPLY | test_r4_excel.py::test_rt212_neue_section_nummer_bleibt_zahl | neue Section-Nummer ist '2.3' (str, Format '0.0'), soll die Zahl 2.3 mit Format '0.0' sein |
| RT-213 | Derselbe Satz zweimal in einer Zelle, Änderung am zweiten Vorkommen; zweiter Lauf auf dem Ergebnis ändert nichts. | PLAN | test_r4_plan.py::test_rt213_dublette_in_derselben_zelle_und_zweiter_lauf | **besteht schon** (Schutz-Test) |
| RT-214 | Fußnoten-Hochzahl (7 pt, angehoben) am blauen Wort landet im Excel-Text. | PDF | test_r4_pdf.py::test_rt214_hochzahl_am_geaenderten_wort | still falsch -- … Zelle '… an die Zentrale1 weiter.' |
| RT-215 | Silbentrennung im neuen Text („E-“ am Zeilenende, „Mail“ danach): wird „E- Mail“. | PDF | test_r4_pdf.py::test_rt215_trennung_e_mail_im_neuen_text | still falsch -- … Zelle '… zuverlässig per E- Mail an die Leitstelle weiter.' |
| RT-216 | Von Hand gebauter falscher Plan (ersetzt „Werte“ statt „alte“, Absatz korrekt in `absaetze`) bleibt „uebernommen/ok“. | APPLY | test_r4_excel.py::test_rt216_falscher_plan_besteht_die_gegenprobe_nicht | falscher Plan bleibt 'uebernommen' mit Wirkung 'ok' (Nachkontrolle 'ok'); Zelle '… meldet alle neue an die alte Leitstelle.' |

**Wird ein Fall repariert**, wird sein Test XPASS (rot): dann den `xfail`-Marker (bzw. die `marks=` im `pytest.param`)
entfernen – er bleibt als Schutz-Test.
