# Kalibrierung: Schwellen verstehen und vorsichtig einstellen

Autor: Nico Lehr · Stand: 15.10.2026

Das Tool entscheidet an vielen Stellen mit **Schwellen** (Zahlen in `config/farben.yaml` und `config/zuordnung.yaml`). Dieses Dokument erklärt jede Schwelle in
einfachen Worten: was sie bedeutet, was passiert, wenn man sie **erhöht** oder **senkt**, und welchen **Standardwert** sie hat.
Danach steht, **wie** man kalibriert, ohne sich selbst zu täuschen.

## Die eiserne Regel (Regel 11)

- Schwellen werden **nur an den Entwicklungsbeispielen** eingestellt (`data/paare/entwicklung/`).
- Die **Test-Beispiele** (`data/paare/test/`) werden nie zum Einstellen benutzt, nicht einmal zum „nur mal schauen“. Sie dienen der Endmessung (höchstens zweimal).
  Sonst sieht der Test gut aus, weil wir ihn „auswendig gelernt“ haben, und sagt nichts mehr über neue Dokumente.
- Keine echten Inhalte in Git, Code, Tests, Kommentaren oder Commits (Regel 16). Im Repo stehen nur erfundene Beispiele.
- **Ein Wert nach dem anderen.** Nach jeder Änderung neu messen und die Zahlen vorher/nachher vergleichen (Ablauf unten).
- **Sicherheit vor Bequemlichkeit.** Eine Schwelle, die mehr Treffer „automatisch“ macht, ist nur dann gut, wenn die **stillen Fehler bei 0 bleiben**
  (`docs/ziel_v1.md`). Im Zweifel lieber mehr Nachfragen („prüfen“) als ein falscher automatischer Treffer (Regel 27).

Zum Messen gibt es zwei Werkzeuge:

- `python -m src.cli farben --pdf DATEI.pdf` zeigt, welche Farben im PDF vorkommen und wie sie gedeutet werden (nur Zahlen und Farbcodes, **keine Wörter**).
- `python -m src.cli messen --korpus entwicklung` misst das ganze Tool an den Entwicklungspaaren (Bericht unter `data/messung/…/bericht.md`).

---

## 1. `config/farben.yaml` – Wie Änderungen im PDF erkannt werden

### `legende` (Bedeutung der Farben)
| Eintrag | Standard | Bedeutung |
|---|---|---|
| `blau` | `neu` | Blauer Text (oder blau hinterlegt) ist neu hinzugekommen. |
| `rot` | `entfaellt` | Roter Text entfällt. |
| `durchgestrichen` | `entfaellt` | Durchgestrichener Text entfällt. |

Erlaubte Werte: `neu`, `entfaellt`, `ignorieren`. Mit `ignorieren` zählt diese Farbe nicht als Änderung. Das ist keine Schwelle, sondern die Vereinbarung mit dem
Auftraggeber: Nur ändern, wenn sich die Legende der Dokumente wirklich ändert.

### `farbton` (welche Farben zählen als „blau“ und „rot“)
Farbtöne werden in Grad auf dem Farbkreis angegeben (0 = rot, 120 = grün, 240 = blau).

| Schwelle | Standard | Bedeutung und Wirkung |
|---|---|---|
| `blau: [185, 265]` | von 185° bis 265° | Alles in diesem Bereich gilt als „blau“. **Bereich weiten:** auch Türkis und Violett zählen als blau (mehr Treffer, aber auch Links und Zierfarben werden zu Änderungen). **Bereich verengen:** nur reines Blau zählt; leicht andere Blautöne werden übersehen und stehen dann bei „Unbekannte Farben“. |
| `rot_bis` | 18 | Rot reicht von 0° bis zu diesem Wert (Orange-Rot). **Erhöhen:** auch Orange gilt als rot. **Senken:** nur ein sattes Rot zählt. |
| `rot_ab` | 335 | Rot beginnt bei diesem Wert und reicht bis 360° (Purpur-Rot). **Senken:** auch Magenta und Pink gelten als rot. **Erhöhen:** nur reines Rot zählt. |

### `text` (Mindest-Kraft der Textfarbe)
| Schwelle | Standard | Bedeutung und Wirkung |
|---|---|---|
| `min_saettigung` | 0,35 | Wie „bunt“ die Farbe mindestens sein muss (0 = grau, 1 = ganz kräftig). **Erhöhen:** blasse Blau-/Rottöne werden ignoriert (weniger Fehlalarme bei Grau mit Farbstich, aber echte, blasse Markierungen gehen verloren). **Senken:** auch fast graue Farben können als Änderung zählen. |
| `min_helligkeit` | 0,25 | Wie hell die Farbe mindestens sein muss (0 = schwarz, 1 = hell). **Erhöhen:** sehr dunkles Blau/Rot (das fast schwarz aussieht) zählt nicht mehr. **Senken:** auch sehr dunkle Farben können als Änderung zählen. |

### `hintergrund` (farbig hinterlegte Wörter, Textmarker)
| Schwelle | Standard | Bedeutung und Wirkung |
|---|---|---|
| `aktiv` | `true` | Ob Hinterlegungen überhaupt ausgewertet werden. `false`: nur die Textfarbe zählt. |
| `min_saettigung` | 0,15 | Wie bunt eine Hinterlegung mindestens sein muss. Niedriger als bei Text, weil Textmarker oft blass sind. **Erhöhen:** blasse Marker werden übersehen. **Senken:** hellgraue Tabellenschattierung kann als Markierung zählen. |
| `min_helligkeit` | 0,40 | Mindest-Helligkeit der Hinterlegung. **Erhöhen:** nur helle Marker zählen. **Senken:** auch dunkle Flächen zählen. |
| `max_flaeche_anteil` | 0,25 | Rechtecke, die größer sind als dieser Anteil der Seite, gelten als Seiten- oder Tabellenhintergrund und werden ignoriert. **Erhöhen:** auch sehr große Flächen zählen als Markierung (Gefahr bei farbigen Tabellen). **Senken:** auch mittelgroße Marker über mehrere Zeilen fallen heraus. |

### `durchstreichung` (durchgestrichene Wörter)
| Schwelle | Standard | Bedeutung und Wirkung |
|---|---|---|
| `aktiv` | `true` | Ob Durchstreichungen ausgewertet werden. |
| `band: [0.28, 0.72]` | 28 % bis 72 % der Zeichenhöhe | Eine Linie zählt nur als Durchstreichung, wenn sie in diesem Höhenband des Buchstabens liegt (Mitte). **Band weiten:** auch Unterstreichungen oder Linien nahe der Grundlinie zählen als Durchstreichung (Gefahr!). **Band verengen:** nur Linien genau in der Mitte; leicht verrutschte Streichungen werden übersehen. |
| `min_abdeckung` | 0,5 | Anteil der Zeichenbreite, den die Linie überdecken muss. **Erhöhen:** die Linie muss fast den ganzen Buchstaben queren. **Senken:** schon ein kurzer Strich über einem Buchstaben reicht. |
| `max_linienstaerke_pt` | 2,0 | Dickere Linien (in Punkt) sind Rahmen, keine Streichung. **Erhöhen:** auch dicke Linien zählen. **Senken:** dünne Hilfslinien nur noch. |
| `max_breite_anteil` | 0,6 | Linien, die breiter als dieser Anteil der Seite sind, sind Trennlinien. **Erhöhen:** auch sehr lange Linien zählen als Streichung (Gefahr bei Tabellenrändern). **Senken:** auch normale lange Streichungen werden verworfen. |
| `wort_anteil` | 0,6 | Ab diesem Anteil durchgestrichener Buchstaben gilt das **ganze Wort** als durchgestrichen. **Erhöhen:** das Wort muss fast vollständig gestrichen sein. **Senken:** schon wenige gestrichene Buchstaben streichen das ganze Wort. |

### `annotationen` (PDF-Kommentare: Hervorheben, Durchstreichen)
| Schwelle | Standard | Bedeutung |
|---|---|---|
| `aktiv` | `true` | Ob die Markierungen aus PDF-Kommentaren (Textmarker und Durchstreich-Kommentar) ausgewertet werden. `false`: nur Farben und Linien zählen. |

### `layout` (Wörter, Zeilen und Absätze im PDF)
| Schwelle | Standard | Bedeutung und Wirkung |
|---|---|---|
| `zeilen_toleranz` | 0,5 | Zwei Buchstaben gelten als „gleiche Zeile“, wenn ihre Höhen um höchstens diesen Anteil der Schriftgröße abweichen. **Erhöhen:** hoch- und tiefgestellter Text wird eher in die Zeile gezogen, aber eng gesetzte Zeilen werden eher vermischt. **Senken:** Zeilen werden genauer getrennt, hochgestellte Zeichen können abfallen. |
| `wort_luecke` | 0,18 | Ein Abstand ab diesem Anteil der Schriftgröße gilt als Leerzeichen (neues Wort). **Erhöhen:** Wörter werden eher zusammengezogen („einWort“). **Senken:** Wörter werden eher zerrissen („ein Wo rt“). |
| `absatz_faktor` | 1,3 | Ist der Zeilenabstand größer als dieser Faktor mal der typische Zeilenabstand, beginnt ein neuer Absatz. **Erhöhen:** Absätze werden eher zusammengefasst. **Senken:** Absätze werden eher zerteilt. |
| `ueberschrift_faktor` | 1,08 | Eine Zeile mit Gliederungsnummer und Schrift mindestens so viel größer als der Fließtext (oder fett) ist eine Überschrift. **Erhöhen:** nur deutlich größere Schrift zählt. **Senken:** schon fast gleich große Zeilen mit Nummer werden zu Überschriften. |
| `kopf_fuss_anteil` | 0,07 | Oberer und unterer Seitenrand (Anteil der Seitenhöhe), in dem Kopf- und Fußzeilen gesucht werden. **Erhöhen:** mehr Text am Rand wird als Kopf-/Fußzeile ignoriert (Gefahr: echter Text am Rand geht verloren). **Senken:** Kopf-/Fußzeilen werden eher als Inhalt gelesen. |
| `silbentrennung_zusammenfuegen` | `true` | Ob getrennte Wörter am Zeilenende („Kapazi-“ / „tät“) wieder zusammengesetzt werden. |
| `keine_silbentrennung_vor` | `[und, oder, bzw, sowie, bis, to, and, or]` | Wörter, nach denen ein Bindestrich **keine** Silbentrennung ist („Strom- und …“). Liste ergänzen, wenn ein falsches Zusammenfügen auffällt. |

---

## 2. `config/zuordnung.yaml` – Wie Änderungen in der Excel gefunden werden

Das Tool vergleicht jeden PDF-Absatz mit den Absätzen der Excel und vergibt einen **Wert von 0 bis 1** (1 = gleicher Text). Je nach Wert und Eindeutigkeit gilt ein Treffer als
„hoch“ (wird automatisch eingetragen), „mittel“ (Vorschlag, ein Mensch prüft) oder gar nicht gefunden.

### `schwellen`
| Schwelle | Standard | Bedeutung und Wirkung |
|---|---|---|
| `hoch` | 0,90 | Ab diesem Wert wird automatisch übernommen. **Erhöhen:** weniger automatische Treffer, mehr „prüfen“ (sicherer, mehr Handarbeit). **Senken:** mehr wird automatisch eingetragen, aber falsche Treffer rutschen eher durch. **Vorsicht:** Das ist die wichtigste Sicherheits-Schwelle; nur senken, wenn die stillen Fehler (Kriterium 2) dabei sicher bei 0 bleiben. |
| `mittel` | 0,65 | Zwischen `mittel` und `hoch` entsteht ein Vorschlag („prüfen“). Darunter gibt es keine Zuordnung („nicht übernommen“). **Erhöhen:** weniger Vorschläge, mehr „nicht übernommen“. **Senken:** mehr (auch schlechtere) Vorschläge. |
| `eindeutig_abstand` | 0,08 | Der zweitbeste Treffer muss mindestens so viel schlechter sein, sonst gilt die Stelle als nicht eindeutig (höchstens „prüfen“). **Erhöhen:** bei ähnlichen Absätzen wird früher nachgefragt (sicherer). **Senken:** bei ähnlichen Absätzen wird eher automatisch gewählt (Verwechslungsgefahr). |
| `kurz_tokens` | 4 | Absätze mit höchstens so vielen Wörtern sind „kurz“ und brauchen eine passende Section, sonst höchstens „prüfen“. **Erhöhen:** mehr kurze Texte werden nur vorgeschlagen (sicherer). **Senken:** auch sehr kurze Texte werden automatisch gesetzt (Verwechslungsgefahr, etwa bei „Siehe Anhang“). |

### `kandidaten_pro_block`
| Schwelle | Standard | Bedeutung und Wirkung |
|---|---|---|
| `kandidaten_pro_block` | 10 | So viele Excel-Absätze (die besten laut Wortindex) werden genau verglichen. **Erhöhen:** gründlicher, aber langsamer (bei großen Excel-Dateien). **Senken:** schneller, aber die richtige Stelle kann in der Vorauswahl fehlen. |

### `matcher` (Bewertung eines Treffers)
| Schwelle | Standard | Bedeutung und Wirkung |
|---|---|---|
| `gewicht_abdeckung` | 0,9 | Wie stark zählt, welcher Anteil der PDF-Wörter in der Excel in gleicher Reihenfolge steht. |
| `gewicht_dice` | 0,1 | Wie stark zählt, dass der Excel-Absatz knapp ist (Absatz statt ganzer Zelle). Beide Gewichte zusammen ergeben den Wert: `Wert = gewicht_abdeckung × Abdeckung + gewicht_dice × Dice`. **Mehr Gewicht auf Dice:** kürzere Excel-Absätze werden bevorzugt. Die beiden Gewichte sollten zusammen 1 ergeben. |
| `section_bonus` | 0,03 | Zuschlag, wenn die Section-Nummer in PDF und Excel gleich ist. **Erhöhen:** die Section entscheidet stärker bei gleichen Texten. **Senken/0:** die Section hilft nicht. |
| `untergrenze` | 0,4 | Unter diesem Wert gibt es gar keinen Treffer, auch keinen Vorschlag. **Erhöhen:** weniger schwache Vorschläge. **Senken:** auch sehr schwache Treffer werden gezeigt. |
| `max_kandidaten_section` | 400 | So viele Absätze derselben Section werden zusätzlich geprüft. **Erhöhen:** gründlicher bei sehr langen Sections, langsamer. **Senken:** schneller, aber bei sehr langen Sections kann die Stelle übersehen werden. |
| `excel_abdeckung_min` | 0,75 | Deckt der PDF-Text weniger als diesen Anteil eines längeren Excel-Absatzes ab (Seitenumbruch?), wird dort kein **neuer** Absatz automatisch angesetzt. **Erhöhen:** früher Nachfrage. **Senken:** auch bei kleinem Ausschnitt wird automatisch gesetzt. |
| `rest_min_woerter` | 5 | So viele Wörter des Excel-Absatzes müssen im PDF fehlen, damit die Regel oben greift. **Erhöhen:** die Regel greift seltener. **Senken:** schon bei wenigen fehlenden Wörtern wird nachgefragt. |

### `excel` (Erkennung des Excel-Aufbaus)
Wenn der Aufbau falsch oder unsicher erkannt wird, ist oft ein **Layout-Profil** besser als eine Schwelle (siehe `docs/adr/0003-layout-profile.md`, `config/layouts.yaml`).

| Schwelle | Standard | Bedeutung und Wirkung |
|---|---|---|
| `min_inhalt_zeichen` | 25 | Ab so vielen Zeichen gilt eine Zelle als Inhaltstext (nicht als Titel oder Notiz). **Erhöhen:** nur lange Texte zählen als Inhalt. **Senken:** auch kurze Zellen zählen. |
| `section_spalte_min` | 2 | So viele Gliederungsnummern braucht eine Spalte mindestens, um „Section-Spalte“ zu sein. **Erhöhen:** kleine Blätter ohne Gliederung werden nicht als Gliederung missverstanden. **Senken:** schon zwei Zahlen machen eine Spalte zur Section-Spalte. |
| `modell_schwelle` | 0,6 | Steht in mindestens diesem Anteil der Section-Zeilen schon der Inhalt, gilt „zeile_pro_section“, sonst „zeile_pro_absatz“. **Erhöhen:** es muss fast überall Inhalt in der Section-Zeile stehen. **Senken:** „zeile_pro_section“ wird eher gewählt. |
| `titel_max_zeichen` | 160 | Längere Zellen sind kein Titel und bei „1. Das System …“ keine Section. **Erhöhen:** auch lange Überschriften zählen. **Senken:** nur kurze Titel. |
| `kopf_max_woerter` | 12 | Eine Zelle wie „1. Das System …“ mit mehr Wörtern (oder Satzpunkt am Ende) ist eine Aufzählung, keine Section. **Erhöhen:** auch lange Zeilen werden zu Section-Titeln. **Senken:** nur kurze Titel. |
| `titel_fliesstext_anteil` | 0,5 | Ist eine Spalte neben der Nummer in mehr als diesem Anteil der Folgezeilen lang, ist sie Fließtext, kein Titel. **Erhöhen:** die Spalte bleibt eher „Titel“. **Senken:** sie wird früher als Fließtext angesehen. |
| `inhalt_zweifel_anteil` | 0,5 | Hat eine zweite Textspalte mindestens diesen Anteil des Texts der ersten, ist der Aufbau **unsicher** (z. B. Kommentarspalte). **Senken:** früher unsicher (mehr „prüfen“, sicherer). **Erhöhen:** später unsicher. |
| `gemischt_anteil` | 0,25 | Steht der Text zu mindestens diesem Anteil in Section-Zeilen **und** in Folgezeilen, ist der Aufbau unsicher. **Senken:** früher unsicher. **Erhöhen:** später unsicher. |
| `lfd_min` | 3 | Eine Spalte 1, 2, 3 … ohne Punkte gilt ab so vielen Zeilen als laufende Nummer, nicht als Gliederung. **Erhöhen:** erst bei längeren Reihen. **Senken:** schon kurze Reihen. |
| `warn_ab_zellen` | 6 | Blätter ohne Gliederung werden erst ab so vielen Textzellen gemeldet (damit Deckblätter nicht warnen). **Erhöhen:** weniger Warnungen. **Senken:** mehr Warnungen. |
| `lfd_kopfwoerter`, `kommentar_kopfwoerter`, `section_kopfwoerter` | Wortlisten | Spaltenüberschriften, die auf „laufende Nummer“, „Kommentar-Spalte“ bzw. „Section-Spalte“ hindeuten. **Ergänzen**, wenn ein Excel-Typ andere Überschriften benutzt (Begriffe klein und ohne Endungen schreiben). Hier dürfen nur allgemeine Begriffe stehen, keine Namen aus echten Dokumenten im Repo (Regel 16); firmenspezifische Begriffe gehören in lokale Profile. |

### `einfuegen`, `geloeschtes`, `markierung` (Einstellungen statt Schwellen)
| Einstellung | Standard | Bedeutung |
|---|---|---|
| `einfuegen.neuer_absatz` | `auto` | Wohin ein neuer Absatz kommt: `zelle` (in dieselbe Zelle), `zeile` (neue Zeile) oder `auto` (je nach Aufbau). |
| `geloeschtes.inhalt` | `leeren` | Gestrichener Inhalt: `leeren` (Zelle leeren, Zeile bleibt) oder `zeile_entfernen` (Zeile löschen, wenn sie leer wird). Löschen hat mehr Risiko für Formeln und Verknüpfungen. |
| `markierung.geaenderte_zellen` | `keine` | Geänderte Zellen markieren: `keine`, `kommentar` oder `gelb`. |

### `planer`
| Einstellung | Standard | Bedeutung und Wirkung |
|---|---|---|
| `verdaechtige_zeichen` | Ligaturen und Ersatzzeichen | Zeichen, die beim PDF-Lesen oft verfälscht werden (z. B. „ﬁ“). Neutext mit solchen Zeichen wird nie automatisch übernommen. **Ergänzen**, wenn weitere verdächtige Zeichen auftauchen. |
| `verdaechtige_muster` | `(cid:12)`-Muster | Dasselbe als Muster für kaputte PDF-Schrift. |
| `doppelt_min_woerter` | 3 | Ab so vielen Wörtern gilt „der gleiche Absatz steht schon in der Excel“ als sicher (die Änderung war schon eingetragen und wird nicht noch einmal gemacht); kürzere Absätze sind nur ein Vorschlag. **Erhöhen:** mehr kurze Absätze bleiben Vorschlag (sicherer gegen Doppeltes). **Senken:** schon sehr kurze Absätze zählen als „schon vorhanden“ (kann eine echt neue Zeile verschlucken). |

### `anwenden`
| Einstellung | Standard | Bedeutung |
|---|---|---|
| `markierung_farbe` | `FFF2CC` | Hintergrundfarbe, wenn `markierung.geaenderte_zellen: gelb` gilt. |
| `kommentar_autor` | `CDE-Tool` | Autor der Zellkommentare, wenn `markierung.geaenderte_zellen: kommentar` gilt. |

---

## 3. Ablauf zum Kalibrieren (nur an Entwicklungsbeispielen)

1. **Ausgangsmessung festhalten.** `python -m src.cli messen --korpus entwicklung --mit-vorschlaegen`. Den Bericht (`data/messung/NNN_entwicklung/bericht.md`) nicht löschen;
   Berichte bleiben liegen (Regel 29). Notiere die Nummer des Berichts als „Vorher“.
2. **Das Problem eingrenzen.** Was sagt der Bericht? Die **Fehlerarten** zeigen, wo es hakt:
   - `pdf_falsch_gelesen`: Farben, Linien, Zeilen oder Absätze im PDF → Abschnitt 1 (`farben.yaml`). Mit `python -m src.cli farben --pdf …` an einem **Entwicklungs**-PDF nachsehen:
     Steht eine kräftige Farbe unter „Unbekannte Farben“? Dann wird sie nicht als Änderung gedeutet.
   - `layout_falsch`: Excel-Aufbau falsch erkannt → zuerst ein **Layout-Profil** erwägen (`profil vorschau`), erst dann `excel.*`-Schwellen.
   - `zuordnung_falsch`, `plan_falsch`: Abschnitt 2 (`schwellen`, `matcher`).
3. **Genau einen Wert ändern**, in `config/farben.yaml` oder `config/zuordnung.yaml`. Notiere: welcher Wert, alt → neu, warum.
4. **Neu messen**: derselbe Befehl wie in Schritt 1. Vergleiche mit „Vorher“:
   - **Stille Fehler** (Kriterium 2) müssen **0** bleiben. Steigt die Zahl, sofort zurück zum alten Wert.
   - **Treffergenauigkeit automatisch** (Kriterium 3) darf nicht sinken.
   - **Abdeckung ohne Nachfrage** (Kriterium 4) darf steigen, aber nie auf Kosten von 2 und 3.
   - Schau auch die Tabelle **je Störer-Familie**: Hat sich eine andere Familie verschlechtert?
5. **Besser?** Behalten und weiter mit dem nächsten Wert. **Nicht besser oder schlechter?** Zurücksetzen. Immer nur eine Änderung gleichzeitig.
6. **Nicht für ein Beispiel „zurechtbiegen“.** Wenn ein Wert nur ein einziges Paar rettet und andere verschlechtert, ist er falsch. Besser: die Ursache verstehen (Regel 20)
   oder die Aufgabe an die Entwicklung geben, statt die Schwelle zu verdrehen.
7. **Tests grün halten:** `python -m pytest -q`. Einige Tests stützen sich auf die Standardwerte; schlägt einer fehl, ist das ein Hinweis, dass der neue Wert das Verhalten
   spürbar verändert. Dann bewusst entscheiden, nicht einfach den Test anpassen.
8. **Dokumentieren:** Wenn ein Standardwert dauerhaft geändert wird, hier den neuen Standardwert eintragen und kurz begründen (mit Zahlen aus dem Bericht, nie mit Textinhalten).
9. **Endmessung erst am Schluss:** `python -m src.cli messen --korpus test` höchstens zweimal (nach Welle 2 und zum Abschluss). Danach **nicht** mehr an Schwellen drehen, um das Ergebnis zu verbessern.
