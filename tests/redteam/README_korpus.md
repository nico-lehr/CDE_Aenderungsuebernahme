# Red-Team R3: Analyse der 43 „stillen Fehler“ im Entwicklungskorpus

Alles **erfunden** (Regel 16). `src/`, `config/` und `tests/fixtures/` sind unverändert; es wurde nur Neues unter `tests/redteam/` angelegt.
Der Testkorpus (Seeds ab 1000) wurde **nicht** erzeugt oder gemessen (Regel 11); `_bau_korpus.paar()` bricht bei Seeds ≥ 1000 ab.

Ausgangslage: `python -m src.cli messen --korpus entwicklung --mit-vorschlaegen` (60 Paare, Seeds 1–60) meldet **43 stille Fehler**
(24× „übernommen, aber falsch“ mit Nachkontrolle „ok“, 4× Kollateral, 15× erwartete Lücke nicht gemeldet).
Alle 43 wurden einzeln durchgesehen (Paar, Blatt, Zeile, Spalte, Änderungs-ID) und nach **Ursache** gruppiert (nicht nach Zelle).

## Kurzergebnis

| Klasse | Einträge | Anteil | Wo liegt es |
|---|---|---|---|
| **Echte Tool-Fehler** | **27** | 63 % | `pdf_reader` **26**, `excel_model`/`matcher` **1**, `planner`/`applier` **0** (ein nicht stiller Befund, RT-122) |
| **Messartefakte** (Zählfehler in `messen.py`) | **16** | 37 % | Mess-Bauer |
| **Generator-Fehler** | **0** | 0 % | Keiner bewiesen; zwei Hinweise für den Daten-Bauer (unten) |

Wichtig für das Verständnis: Ein „Eintrag“ ist nicht immer ein Ereignis. Die 28 Zell-Einträge zählen **Zellen** (eine falsch gesetzte
Zeile mit 4 Spalten = bis zu 4 Einträge), die 15 Lücken-Einträge zählen **Paare** (zwei fehlende Lücken im selben Paar = 1 Eintrag).

### Verteilung der 43 auf Ursachen

| Ursache | Einträge | Klasse | Modul / Besitzer | RT |
|---|---|---|---|---|
| G1 Lücke (Scan, Zwei-Spalten, Tabelle, Bild) wird **nie** gemeldet | **15** Paare (17 Lücken) | Tool | `pdf_reader` / K-PDF | RT-110 bis 113 |
| G2 Neues **blaues** Wort mit Silbentrennung steht zerrissen in der Excel | **7** | Tool | `pdf_reader` / K-PDF | RT-114, 115 |
| G3 Unveränderte **Farbtabelle** (blauer Kopf) wird in die Zelle geschrieben | **3** | Tool | `pdf_reader` / K-PDF | RT-116, 117 |
| G4 **Zwei-Spalten**-Satz wird zerrissen mitten in andere Sätze geschrieben | **1** | Tool | `pdf_reader` / K-PDF | RT-118, 119 |
| G5 Neuer **Listenpunkt** „4. …“ wird in die **Nummernspalte** geschrieben | **1** | Tool | `excel_model` (Kandidaten) / K-XL | RT-120, 121 |
| MA-1 Offene „pruefen“-Änderung in **derselben Zelle** | **8** | Messartefakt | `messen.py` / Mess-Bauer | RT-125, 126 |
| MA-2 **Phantom-Kollateral** bei teilweise geleerter Zeile | **4** | Messartefakt | `messen.py` / Mess-Bauer | RT-127, 128 |
| MA-3 **Anker** nach geleerter statt gelöschter Zeile | **4** | Messartefakt | `messen.py` / Mess-Bauer | RT-129, 130 |
| **Summe** | **43** | | | |

Doppelte Ursache: die neue Zeile in Seed 57 (Zelle 9/3) hat zwei Gründe (Silbentrennung **und** der Anker aus MA-3); sie ist bei G2 gezählt, die Zellen 9/1 und 9/2 bei MA-3. Die Paare 33, 39, 45 (Tabelle) und 42 (Zwei-Spalten) stehen zweimal: einmal als fehlende Lücke (G1), einmal als falsche Zelle (G3/G4) –
beides sind echte, aber getrennte Folgen derselben Lese-Ursache.

## Tabelle der Tests

Aufruf: `python -m pytest -q tests/redteam` (alles grün, xfail = erwartet). Beweis, dass jeder Test aus dem richtigen Grund fällt:
`python -m pytest --runxfail -q tests/redteam/test_korpus_*.py --tb=line` → 60 Fehlschläge, **alle** `AssertionError`, kein Aufbau-Fehler (`RuntimeError`).
Kurz-Fehlermeldung = was `--runxfail` heute meldet (gekürzt). „Zellen“ = Zahl der betroffenen Einträge im Korpus.

### A. Echte Tool-Fehler (`test_korpus_analyse.py`)

| RT | Ursache (ein Satz) | Klasse | Modul / Besitzer | Zellen | Datei::Test | Meldung mit `--runxfail` |
|---|---|---|---|---|---|---|
| RT-110 | Scan-Seiten (Seeds 6, 12, 27) werden nicht als Lücke `scan_seite` gemeldet | Tool | pdf_reader / K-PDF | 3 Paare | test_korpus_analyse.py::test_rt110_korpus_scan_seiten_ohne_luecke | Luecke 'scan_seite' nicht gemeldet in 3 von 3 Paaren (synth_0006, synth_0012, synth_0027); erwartet: 3; die Bilanz sagt 'vollstaendig' |
| RT-111 | Zwei-Spalten-Abschnitte (Seeds 3, 9, 18, 42, 48) ohne Lücke `mehrspaltig` | Tool | pdf_reader / K-PDF | 5 Paare | …::test_rt111_korpus_zwei_spalten_ohne_luecke | Luecke 'mehrspaltig' nicht gemeldet in 5 von 5 Paaren (…) |
| RT-112 | Tabellen (Seeds 33, 39, 45) ohne Lücke `tabelle` | Tool | pdf_reader / K-PDF | 3 Paare | …::test_rt112_korpus_tabellen_ohne_luecke | Luecke 'tabelle' nicht gemeldet in 3 von 3 Paaren (…) |
| RT-113 | Bilder im Fließtext (Seeds 6, 21, 48, 51, 54, 60) ohne Lücke `bild`; 4 Paare davon nur wegen Bild (kein Wortverlust) | Tool | pdf_reader / K-PDF | 4 Paare (+2 schon bei RT-110/111 gezählt) | …::test_rt113_korpus_bilder_ohne_luecke | Luecke 'bild' nicht gemeldet in 6 von 6 Paaren (…) |
| RT-114 | Neues blaues Wort am Zeilenende getrennt („Netzanschlussbe-“ „dingungen“) steht zerrissen in der Excel, übernommen/hoch/ok (Minimalfall; ascii und weich) | Tool | pdf_reader / K-PDF (Leser: RT-63/64) | Minimalfall | …::test_rt114_blaues_getrenntes_wort_landet_zerrissen_in_der_excel[ascii] und [weich] | Wort zerrissen in der Excel, Status 'uebernommen'/hoch, Nachkontrolle 'ok': 'Sicherheitsarchitektur Netzanschlussbe- dingungen geprueft …' |
| RT-115 | Dasselbe im Korpus: Seeds 3, 9 (2×), 51, 57 (+ je eine neue Zeile in Seed 9 und 57, im Test nicht einzeln erreichbar) | Tool | pdf_reader / K-PDF | 5 Zellen (+2 neue Zeilen = **7**) | …::test_rt115_korpus_zerrissenes_wort_in_der_excel[5 Fälle] | Seed 3, Zelle B2Z39S3: Wort in zwei Stuecken in der Excel ('Betrie- bsführungskonzept'); A0009 uebernommen/hoch, Nachkontrolle 'ok' |
| RT-116 | Unveränderte Tabelle mit blauem Kopf und hellblauer Bänderung: Kopf-/Datenwörter werden in die Zelle eingefügt (Minimalfall) | Tool | pdf_reader / K-PDF (Leser: RT-76/77) | Minimalfall | …::test_rt116_unveraenderte_farbtabelle_wird_in_die_zelle_geschrieben | Unveraenderte Tabelle in die Zelle geschrieben: A0001 uebernommen/hoch, Nachkontrolle 'ok', Zelle jetzt: 'Pos. | Bauteil | … Stueck Pos. Bauteil Menge Einheit 1 Netzteil 24 Stueck …' |
| RT-117 | Dasselbe im Korpus: Seeds 33, 39, 45 | Tool | pdf_reader / K-PDF | 3 | …::test_rt117_korpus_tabellenkopf_in_der_excel[3 Fälle] | Seed 33: Tabellenkopf als Fliesstext in der Excel-Zelle B3Z13S2 (A0013 uebernommen/hoch, Nachkontrolle 'ok') |
| RT-118 | Zwei-Spalten: neuer Satz der rechten Spalte wird in zwei Stücken mitten in Sätze der linken Spalte geschrieben, ohne Lücke (Minimalfall) | Tool | pdf_reader / K-PDF | Minimalfall | …::test_rt118_zwei_spalten_satz_wird_zerrissen_eingetragen | Satz der rechten Spalte zerrissen in der Excel (2 Teile uebernommen/hoch), keine Luecke: 'Der Nennwert von 20 Monate gilt bei 40 Grad Bei Ueberstrom wird ein Alarm an die Leitstelle Umgebungstemperatur. …' |
| RT-119 | Dasselbe im Korpus: Seed 42 | Tool | pdf_reader / K-PDF | 1 | …::test_rt119_korpus_zwei_spalten_satz_zerrissen_seed42 | Seed 42: der neue Satz steht zerrissen in der Excel-Zelle B2Z3S3: … |
| RT-120 | Neuer Listenpunkt „4. {+Sicherung tauschen+}“: Nummer „4.“ (schwarz) trifft die Nummern-Zelle „4“, der Text steht in der Nummernspalte, übernommen/hoch/ok (Minimalfall) | Tool | excel_model._kandidaten_bauen + matcher / K-XL | Minimalfall | …::test_rt120_listenpunkt_text_landet_in_der_nummernspalte | Der Text steht in der Nummernspalte: '4 Sicherung tauschen' (A0001 uebernommen/hoch, Nachkontrolle 'ok', Operation auf Spalte 1) |
| RT-121 | Dasselbe im Korpus: Seed 30 | Tool | excel_model + matcher / K-XL | 1 | …::test_rt121_korpus_listenpunkt_in_nummernspalte_seed30 | Seed 30: Nummern-Zelle B2Z18S1 ist jetzt '4 Sicherung tauschen' statt '4' (A0006 uebernommen/hoch, Nachkontrolle 'ok') |
| RT-122 | Section-Löschung (Nummer + Titel) läuft automatisch, der Inhalt bleibt „pruefen“: verwaiste Zeile mit altem Inhalt. **Kein stiller Fehler** (der Inhalt steht sichtbar offen), aber Ursache von MA-2 in Seeds 2 und 37 | Tool | planner / K-PL (Abhängigkeit wie RT-01/02) | 0 (2 der 4 Phantom-Einträge) | …::test_rt122_section_geloescht_automatisch_inhalt_offen_laesst_verwaiste_zeile | Verwaiste Zeile 3: Nummer und Titel sind geleert (A0001 uebernommen/hoch), der Inhalt steht noch da, weil A0002 auf 'pruefen' steht |

### B. Messartefakte (`test_korpus_messartefakte.py`; Besitzer **Mess-Bauer**)

Jeder Test beweist mit `aufbau(...)` (RuntimeError bei Abweichung), dass das **Tool** in diesem Fall nichts Falsches getan hat, und meldet dann per `AssertionError`, was die Messung falsch zählt.

| RT | Ursache (ein Satz) | Klasse | Modul / Besitzer | Zellen | Datei::Test | Meldung mit `--runxfail` |
|---|---|---|---|---|---|---|
| RT-125 | Zelle mit zwei Änderungen, eine sicher (richtig angewendet), eine offen („pruefen“): die richtige gilt als „übernommen, aber falsch“ (Minimalfall) | Messartefakt | messen.py / Mess-Bauer | Minimalfall | test_korpus_messartefakte.py::test_rt125_offene_aenderung_in_derselben_zelle_ist_kein_stiller_fehler | 1 stiller Fehler {'uebernommen_aber_falsch': 1, 'nachkontrolle_ok_aber_falsch': 1}, obwohl die angewendete Aenderung richtig ist und die offene Aenderung sichtbar auf 'pruefen' steht (Treffergenauigkeit 0 von 1) |
| RT-126 | Dasselbe im Korpus: Zelle weicht vom Soll **nur** um die Wörter der offenen Änderung ab (Seeds 11, 22 (2×), 37, 42, 49, 57, 58) | Messartefakt | messen.py / Mess-Bauer | **8** | …::test_rt126_korpus_zelle_weicht_nur_um_offene_aenderung_ab[8 Fälle] | Seed 11, Zelle B3Z3S2 zaehlt als stiller Fehler (A0002), obwohl nur A0001 (pruefen) fehlt |
| RT-127 | Zeile nur teilweise geleert (Nummer/Titel weg, ein Absatz bleibt): die Zeilen-Ausrichtung findet sie nicht wieder und meldet eine „neue Zeile“ = Kollateral (Minimalfall) | Messartefakt | messen.py (`_aehnlich`, `_diff`) / Mess-Bauer | Minimalfall | …::test_rt127_teilweise_geleerte_zeile_ist_kein_kollateral | 2 Kollateralschaeden an Zeilen, die das Ergebnis gar nicht neu angelegt hat (Positionen: [(1, 2, 3), (1, 2, 4)]) |
| RT-128 | Dasselbe im Korpus: Seeds 2 und 37, je Zelle 3 und 4 – Zellen sind in Vorher, Soll und Ergebnis **gleich** | Messartefakt | messen.py / Mess-Bauer | **4** (= alle „Kollateral“) | …::test_rt128_korpus_phantom_kollateral_an_unveraenderter_zelle[4 Fälle] | Seed 2: Zelle B2Z7S3 ist unveraendert und zaehlt trotzdem als '{'kollateral': 2}' |
| RT-129 | Alte Zeile wird geleert statt gelöscht (Bemerkung bleibt), neue Zeile steht dahinter: gleicher Inhalt, gleiche Reihenfolge, aber anderer „Anker“ → alles falsch (Minimalfall) | Messartefakt | messen.py (`_neu_abgleich`) / Mess-Bauer | Minimalfall | …::test_rt129_geleerte_statt_geloeschte_zeile_verschiebt_den_anker | 3 stille Fehler {'kollateral': 3} bei gleichem Inhalt und gleicher Reihenfolge (nur eine leere Restzeile mit Bemerkung bleibt stehen) |
| RT-130 | Dasselbe im Korpus: Seed 43 (Zelle 4/3), Seed 57 (Zeile 9, Spalte 1 und 2) | Messartefakt | messen.py / Mess-Bauer | **4** (Seed 43 zählt doppelt) | …::test_rt130_korpus_anker_nach_geleerter_zeile[3 Fälle] | Seed 57: Zelle B2Z9S1 der NEUEN Zeile zaehlt als stiller Fehler (A0014), nur weil die alte Zeile geleert statt geloescht ist (Rest: 'Status offen') |
| RT-131 | **Lese-Lücke:** Layout „eine Zeile je Absatz“, kurzer Absatz + langer Zusatz → „Zeile gelöscht + neue Zeile“; die unveränderten Wörter zählen als „vom Leser nicht erkannt“ (Leser und Tool fehlerfrei) | Messartefakt | messen.py (`_aehnlich`) / Mess-Bauer | – (Kriterium 5) | …::test_rt131_kurzer_absatz_mit_langem_zusatz_ist_keine_lese_luecke | Lese-Luecke 8 Woerter von 27 im Layout 'eine Zeile je Absatz' (derselbe Inhalt im Layout 'eine Zeile je Section': 0 von 19); … |
| RT-132 | **Lese-Lücke:** Formel-Wörter (`=LEN(C5)`) zählen als „im PDF nicht erkannt“ (Seed 60: 29 von 57) | Messartefakt | messen.py (`_verwaltet`, Lese-Lücke) / Mess-Bauer | – (Kriterium 5) | …::test_rt132_korpus_formelzellen_sind_keine_pdf_woerter_seed60 | 29 von 57 'nicht erkannten' Woertern in Seed 60 sind Formeln (z. B. ['=len(c5', '=len(c6', '=len(c7']); die Lese-Luecke ist zu schlecht |

### C. Schwächen der Messung: Sabotage-Tests (`test_korpus_messluecken.py`)

Teil A sind **Wächter** (normale Tests, grün): diese Sabotage bemerkt die Messung heute. Teil B sind `xfail`-Tests (Besitzer **Mess-Bauer**): diese Sabotage bemerkt sie **nicht**.
Die Sabotage verfälscht nur im Messwerkzeug das Ergebnis des Tools (per `unittest.mock`), ein sauberes Basis-Paar liefert vorher 0 stille Fehler.

| RT | Was wird sabotiert, und was sieht die Messung? | Klasse | Besitzer | Datei::Test | Meldung mit `--runxfail` |
|---|---|---|---|---|---|
| Wächter | Basis-Paar sauber; falsches Wort; Zahl vertauscht; Änderung zurückgenommen („keine Wirkung“); Absätze vertauscht; Text in Bemerkungsspalte; Zeile doppelt; Zeile gelöscht; anderes Blatt verändert – **alle werden bemerkt** | Wächter | – | test_korpus_messluecken.py::test_waechter_* (10 Tests) | grün |
| RT-135 | Nur Leerraum/Absatzgrenze/weiches oder unsichtbares Zeichen/Ligatur verfälscht (6 Varianten) → **0 stille Fehler, Treffergenauigkeit 100 %** | Messlücke | Mess-Bauer (`_k`) | …::test_rt135_messung_bemerkt_unsichtbare_textverfaelschung[6] | Verfaelschung 'ligatur_fi' bleibt unbemerkt: stille Fehler 0, Kollateral 0, Treffergenauigkeit 4 von 4 = 100 % |
| RT-136 | Blatt umbenannt/ausgeblendet, Zeile/Spalte ausgeblendet, Zellen verbunden, Datenvalidierung, Schriftformat (7 Varianten) → unbemerkt (nur Zellwerte werden verglichen; R1: RT-18…26, RT-35/36) | Messlücke | Mess-Bauer (`_lese_excel`) | …::test_rt136_messung_bemerkt_strukturschaden[7] | Strukturschaden 'blatt_umbenannt' bleibt unbemerkt (stille Fehler 0, Kollateral 0, Hinweise []) |
| RT-137 | **Regel 27:** Das Tool wendet eine „pruefen“-Änderung ohne Freigabe an → die Messung sieht **besser** aus (stille Fehler 1 → 0, Abdeckung 0 → 100 %) und meldet den Verstoß nicht | Messlücke | Mess-Bauer | …::test_rt137_pruefen_ohne_freigabe_angewendet_wird_gemeldet | Der Verstoss gegen Regel 27 bleibt unbemerkt und sieht BESSER aus: stille Fehler 1 -> 0, Abdeckung ohne Nachfrage 0 -> 1 von 1 Soll-Zellen |
| RT-138 | Derselbe Verstoß an Korpus-Paaren (Seeds 11, 58) | Messlücke | Mess-Bauer | …::test_rt138_korpus_pruefen_ohne_freigabe_angewendet[11] und [58] | Seed 11: 1 'pruefen'-Aenderung(en) ohne Freigabe angewendet -- die Messung zeigt 0 stille Fehler, Treffergenauigkeit 4 von 4 |
| RT-139 | Absturz im **Wiederholungslauf** (Idempotenz) → das ganze Paar „nicht messbar“ mit **0** stillen Fehlern; der schon gefundene stille Fehler des ersten Laufs geht verloren | Messlücke | Mess-Bauer (`messe_paar`) | …::test_rt139_absturz_im_wiederholungslauf_versteckt_stillen_fehler | Der stille Fehler des ersten Laufs (1) ist weg: Bericht zeigt 0 stille Fehler, Paar 'nicht messbar' |
| RT-140 | Bei einem abgestürzten Paar steht Kriterium 2 „Stille Fehler“ auf **„ok“** (Wert 0); nur Kriterium 1 wird rot | Messlücke | Mess-Bauer (`fasse_zusammen`) | …::test_rt140_abgestuerztes_paar_laesst_kriterium_stille_fehler_gruen | Kriterium 2 'Stille Fehler' steht auf 'ok' (Wert 0), obwohl 1 von 1 Paaren nicht gemessen werden konnten |
| RT-141 | **Ohne `erwartet.json`** (= bei allen echten Paaren) kann eine verschluckte Scan-Seite nie auffallen; Kriterium 1 bleibt „ok“ | Messlücke | Mess-Bauer | …::test_rt141_scan_seite_ohne_erwartet_json_bleibt_unbemerkt | Eine Scan-Seite mit Aenderungen wurde nicht als Luecke gemeldet, die Messung zeigt 0 stille Fehler (Bilanz stimmt: True, vollstaendig: True) |
| RT-142 | **Nicht verwaltete Spalten:** 6 von 7 geänderten Soll-Zellen fehlen im Ergebnis, trotzdem „Abdeckung 100 %“ und alle Kriterien „ok“; kein Kriterium nennt die Zahl | Messlücke | Mess-Bauer | …::test_rt142_unverwaltete_soll_zellen_tauchen_in_keinem_kriterium_auf | Abdeckung 1.0 und alle Kriterien 'ok', obwohl 6 geaenderte Soll-Zellen (nicht verwaltete Spalte) im Ergebnis fehlen; kein Kriterium nennt sie |

Freie Nummern: RT-123, 124, 133, 134 (nicht benutzt).

## Wie `messen.py` richtig zählen müsste (Vorschläge, KEIN Eingriff)

1. **MA-1 (8):** Eine Zelle, an der auch eine offene („pruefen“/„nicht übernommen“) Änderung hängt, darf nicht als „übernommen, aber falsch“ gelten, wenn sie sich vom Soll
   **nur** um die Wörter der offenen Änderung unterscheidet. Praktisch: Zelle gegen „Soll ohne die offenen Änderungen“ prüfen (Wort-Diff gegen `alt`/`neu` der offenen
   Änderungen; `_wort_diff` und `_zugehoerig` gibt es schon) und die Zelle als „teilweise, wartet auf Freigabe“ zählen – nicht als stillen Fehler.
   `nur_durch_offene_aenderungen_verschieden` in `_bau_korpus.py` zeigt die Prüfung.
2. **MA-2 (4):** Die Zeilen-Ausrichtung (`_aehnlich`, `ZEILEN_AEHNLICH = 0,6`) findet eine Zeile nicht wieder, wenn zwei Zellen geleert sind und der Inhalt halbiert ist.
   Vorschlag: Zeilen auch dann paaren, wenn alle belegten Zellen der **neuen** Zeile in der alten stehen (Teilmenge, auch mit gekürztem Text),
   oder die Ähnlichkeit über **verwaltete** Spalten rechnen (Bemerkung/Prio ignorieren) und eine Zeile mit geleerten verwalteten Zellen als „gelöscht“ behandeln.
3. **MA-3 (4):** Eine Zeile, deren **verwaltete** Zellen alle leer sind (nur Bemerkung/Prio übrig), zählt als „gelöscht“ – im Ergebnis genauso wie im Soll. Dann stimmt der
   Anker (`_neu_abgleich`), und die Reihenfolge ist der richtige Vergleich. Zusätzlich sinnvoll: den Anker als „Ort stimmt, wenn keine behaltene Zeile dazwischen liegt“ definieren.
4. **Lese-Lücke (RT-131/132):** Formelzellen (`=…`) und Zellen unverwalteter Spalten nie als Soll-Wörter zählen; Zeilen-Ausrichtung wie bei MA-2 verbessern.
   Schon die beiden nachgewiesenen Gruppen (Seed 60: 29 Formel-Wörter; Seed 56: 12 Wörter eines unveränderten Satzes) machen 41 der 160 „fehlenden“ Wörter aus;
   ohne sie läge die Lese-Lücke bei etwa 1,7 % statt 2,35 % (Rechnung 119 von 6820, nicht neu gemessen).

Erwartete Wirkung auf die Zahl: 43 → **27** stille Fehler, und die 27 sind dann alle echt.

## Generator-Fehler: Befund

Keiner der 43 Einträge ist ein bewiesener Generator-Fehler. Geprüft wurde auch die Treue der Wahrheit: Von 670 Änderungen aus den `erwartet.json` finden sich **668** im aus
Vorher/Nachher abgeleiteten Soll wieder (die 2 Ausnahmen sind Fehlausrichtungen der Messung in Seed 15, siehe MA-2/RT-131). Zwei **Hinweise** für den Daten-Bauer:

* **Neuer Listenpunkt (Seed 30, G5):** Der Generator setzt die Nummer „4.“ **schwarz** und nur den Text blau. In einem echten Änderungs-PDF wäre bei einem eingefügten Absatz
  auch die Nummer markiert. Das Tool darf trotzdem nie in die Nummernspalte schreiben (RT-120 bleibt ein Tool-Fehler), aber der Korpus überschätzt diesen Fall vermutlich.
* **Lücken ohne Inhalt:** Bei 5 der 15 Lücken-Paare (Seeds 21, 39, 51, 54, 60) hat die verpasste Lücke `mit_aenderung: false` und 0 Wörter (Bild ohne Text, unveränderte Tabelle).
  Kein Wortverlust – das Tool soll sie trotzdem melden (ziel_v1 (d)), aber die Messung zählt sie gleich schwer wie eine Scan-Seite mit Änderungen.
  Vorschlag: getrennte Zähler „Lücke mit Änderung“ / „Lücke ohne Änderung“.

## Schwächen der Messung (Treffergenauigkeit und Abdeckung): wo sie zu GUT aussieht

Reihenfolge nach Gefährlichkeit. Jede Zeile hat einen Test in Teil C.

| # | Schwäche | Wirkung | Test |
|---|---|---|---|
| 1 | **Regel 27 wird belohnt:** „Abdeckung ohne Nachfrage“ und „Treffergenauigkeit“ zählen jede richtige Zelle, egal ob die Änderung `uebernommen` war. Ein Tool, das „pruefen“ heimlich anwendet, sieht besser aus (stille Fehler 1 → 0). | zu gut | RT-137, 138 |
| 2 | **Text wird stark normalisiert** (`_k`: NFKC, unsichtbare Zeichen weg, Leerraum zusammen): doppelte Leerzeichen, geschützte Leerzeichen, weiche Trennstriche, Ligaturen, **verschmolzene Absätze** sind unsichtbar. Das sind genau R1-Fehler wie RT-06/07/12. | zu gut | RT-135 |
| 3 | **Nur Zellwerte werden verglichen:** Blatt umbenannt/ausgeblendet, versteckte Zeilen/Spalten, verbundene Zellen, Validierung, Formate, Formel-Typ sind unsichtbar (R1: RT-18…26, 35/36). | zu gut | RT-136 |
| 4 | **Absturz im Wiederholungslauf** löscht die Funde des ersten Laufs; Kriterium 2 bleibt „ok“ (Wert 0) bei abgestürzten Paaren. | zu gut | RT-139, 140 |
| 5 | **Lücken nur mit `erwartet.json`:** Bei echten Paaren kann „Lücke nicht gemeldet“ nie auftreten. Die ganze Kategorie (15 von 43 heute) wäre bei Nicos echten Beispielen **unsichtbar**. | zu gut | RT-141 |
| 6 | **Nicht verwaltete Spalten:** 305 von 1004 Soll-Zellen (30 %) fehlen in Abdeckung und Fehlerzählung. Zusammensetzung (Entwicklungskorpus): 210 Zellen neuer/entfernter Zeilen in Bemerkung/Prio (harmlos), **56 Formelzellen** (verschobene Formeln = echte Gefahr, R1 RT-26; das Tool warnt, die Messung nennt sie nirgends), 39 sonstige (u. a. neu durchnummerierte Lfd-Spalte in Seeds 36, 42, 48, 51; Kommentartexte gelöschter Zeilen in Seed 6). Nur als Zahl im md-Bericht, in keinem Kriterium. | zu gut | RT-142 |
| 7 | **Zähleinheit gemischt:** „stille Fehler“ zählen Zellen (1 falsche 4-Spalten-Zeile = 4; Wächter `zeile_geloescht`), Lücken aber Paare (2 fehlende Lücken = 1). Die 43 sind also keine 43 Ereignisse. Auch „Treffergenauigkeit“ (590 von 618) zählt Zellen: große Strukturänderungen dominieren einzelne falsche Wörter. | verzerrt | Wächter-Tests (`zeile_geloescht` = 4 stille Fehler) |
| 8 | **Abdeckung zählt Zellen, nicht Wörter oder Änderungen:** Eine Zelle mit 10 Änderungen zählt wie eine mit einer. (Streng, nicht zu gut: die Zelle zählt erst, wenn **alles** stimmt – das erzeugt die Artefakte MA-1.) | verzerrt | RT-125/126 |
| 9 | **Soll-Ausrichtung per Ähnlichkeit** (0,6): Fehlausrichtungen erzeugen Phantom-Soll und -Kollateral (zu schlecht), können aber auch ein falsches Ergebnis zufällig als „richtig“ paaren (nicht beobachtet). Treue zur Generator-Wahrheit: 668 von 670. | zu schlecht, selten zu gut | RT-127…131 |
| 10 | **Lese-Lücke global gezählt:** `Counter`-Differenz über das ganze Dokument; ein Wort, das an anderer Stelle noch einmal vorkommt, deckt eine verlorene Stelle zu (nicht nachgemessen); Formeln und Fehlausrichtung zählen dagegen zu viel (RT-131/132). | beides | RT-131, 132 |
| 11 | **„Bilanz stimmig“ vertraut dem Tool:** `bilanz()['stimmt']` kommt aus dem Tool selbst; die Messung rechnet nicht unabhängig nach. | blind für Fehler in `protokoll.py` | (nicht getestet) |

Die Messung **bemerkt** (Wächter, grün): falsches Wort/Zahl in einer verwalteten Zelle, Text in die Bemerkungsspalte, doppelte oder gelöschte Zeile, anderes Blatt verändert,
zurückgenommene Änderung („keine Wirkung“), vertauschte Absätze.

## Empfehlung zur Reihenfolge der Reparatur

1. **Zuerst die Messung (Mess-Bauer, klein, schnell):** MA-1, MA-2, MA-3 (43 → 27), danach RT-137 (Regel 27), RT-139/140 (Absturz) und RT-141 (Lücken aus dem PDF selbst).
   Ohne diese Schritte bedeutet „0 stille Fehler“ nicht „0 Fehler“. RT-135/136 (Text und Struktur) danach, aber vor dem ersten Test-Korpus-Lauf (Regel 11: nur zweimal messen!).
2. **K-PDF, ein Arbeitspaket „Lücken erkennen und den Bereich sperren“** (G1 + G3 + G4 = **19 von 27** echten Fehlern): Scan-Seite (Seite ohne Textebene), Bild-Objekt, Tabelle (Linien/Rechtecke/blauer Kopf),
   Zwei-Spalten (Wörter in zwei x-Bändern) als `luecken` melden **und** die Wörter dieses Bereichs mit einem Hinweis versehen, damit der Planer sie auf „pruefen/nicht übernommen“ setzt.
   Das ist laut `ziel_v1` ohnehin vorgesehen („nicht lösen, nur erkennen und melden“) und deckt R2 RT-95…99, 109 mit ab.
3. **K-PDF Silbentrennung farbiger Wörter** (G2, 7 Einträge, schon R2 RT-63/64/108): Zusammenfügen auch für blaue/rote Wörter und für weiche Trennstriche.
   Zusätzliche Sicherung im Planer (`neutext_probleme`): Neutext mit „Wort- Rest“-Muster oder Pipe-Tabellentext nie automatisch (wie bei `(cid:…)` und Ligaturen, RT-11/12).
4. **K-XL** (G5, 1 Eintrag): `excel_model._kandidaten_bauen` – Nummern-/Titelspalte nicht als Absatz-Kandidat für Fließtext-Blöcke zulassen (nur für Überschriften-Blöcke).
5. **K-PL** (RT-122, kein stiller Fehler): `section_geloescht` nur automatisch, wenn die zugehörigen Inhalts-Löschungen es auch sind; sonst alles auf „pruefen“.
   Das entfernt auch die Ursache der verwaisten Zeilen in Seeds 2, 37 und 60.

Der Daten-Bauer muss nichts ändern. Offen bleiben für ihn die beiden Hinweise oben.

## Reproduzieren

```
python scripts/erzeuge_paare.py --korpus entwicklung --seeds 1-60 --ziel <TEMP>
CDE_DATA_ROOT=<TEMP2> python -m src.cli messen --korpus entwicklung --ordner <TEMP>/entwicklung --mit-vorschlaegen
python -m pytest --runxfail -q tests/redteam/test_korpus_*.py --tb=line
```
Die Korpus-Tests erzeugen ihre Paare selbst (`tests/redteam/_bau_korpus.py`, Temp-Ordner, nach dem Lauf weg). Sie sind an Generator-Version 1 gebunden; ändert sich ein Fall durch eine
Reparatur so, dass der Testaufbau nicht mehr stimmt, meldet der Test einen **Aufbau-Fehler (RuntimeError)** statt eines stillen Bestehens – dann Fall neu bewerten.

## Dateien

| Datei | Inhalt |
|---|---|
| `_bau_korpus.py` | Hilfen: Korpus-Paare (Seeds < 1000), Automatik-Lauf mit Zugriff auf Plan/Ergebnis, Zell-Vergleich Vorher/Soll/Ist, handgebaute Paare, Sabotage |
| `test_korpus_analyse.py` | RT-110 bis RT-122: echte Tool-Fehler (20 Läufe, alle xfail) |
| `test_korpus_messartefakte.py` | RT-125 bis RT-132: Messartefakte (20 Läufe, alle xfail, Besitzer Mess-Bauer) |
| `test_korpus_messluecken.py` | Wächter (10 Tests, grün) und RT-135 bis RT-142: Schwächen der Messung (20 Läufe, xfail, Besitzer Mess-Bauer) |
