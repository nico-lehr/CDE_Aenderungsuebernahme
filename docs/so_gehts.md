# So geht's – CDE-Änderungsübernahme in 3 Schritten

Für Nico · Stand 18.10.2026 · eine Seite, ohne Fachbegriffe

## Starten

1. Im Projektordner doppelklicken: `scripts\start_dashboard.bat`
   (oder in der `.venv`: `python run_server.py`).
2. Im Browser öffnen: **http://127.0.0.1:5100**
3. Das schwarze Fenster offen lassen, solange du arbeitest. Alles bleibt auf deinem Rechner, nichts geht ins Internet.

## Die 3 Schritte

**① Dateien** – Zieh das **PDF mit den Änderungen** und deine **CDE-Excel** in die zwei Felder (oder klick drauf).
Dann **Abgleichen**. Das dauert etwa 15 Sekunden. Es wird dabei noch nichts in die Excel geschrieben.
Unter „Zuletzt“ findest du frühere Abgleiche wieder („Öffnen“). Über „⋯“ kannst du einen Abgleich **archivieren** – er ist dann nur
ausgeblendet, nicht weg („Archivierte anzeigen“ holt ihn zurück).

**② Prüfen** – Oben steht in einem Satz, was gefunden wurde, zum Beispiel
„62 Änderungen gefunden · 55 automatisch eingetragen · 7 bitte prüfen (3 erledigt)“.
- **Bitte prüfen**: Hier ist sich das Tool nicht sicher. Jede Karte zeigt, wie die Zelle **danach in der Excel** aussieht, darunter
  den Ausschnitt **im PDF** und den **Grund**. Du entscheidest: **Übernehmen** oder **Nicht übernehmen**. Nochmal klicken hebt die
  Entscheidung wieder auf.
- **Automatisch eingetragen** (eingeklappt): Das ist sicher und kommt rein. Willst du eine davon nicht, schalte
  „Trotzdem nicht eintragen“ ein.
- **Nicht in der CDE** (eingeklappt): z. B. Deckblatt oder Revisionstabelle, oder es gab keine passende Stelle. Mit Grund.
- Dann **Excel erstellen**. Was du noch nicht entschieden hast, wird **nicht** eingetragen – das steht vorher unter dem Knopf.

**③ Fertig** – **Excel herunterladen** und **Protokoll herunterladen**. Mit **Neuer Abgleich** geht es von vorn los.
Über „Entscheidungen ändern …“ kannst du zurück und eine weitere Excel erstellen; die frühere bleibt erhalten.

## Mehrere Specs in einer CDE

Enthält deine CDE mehrere Specs (Spezifikationsnummern), bearbeitet das Tool **nur die Zeilen der Spec, zu der das PDF gehört**.
Die anderen Zeilen bleiben unverändert.
- Oben in Schritt ② steht dann zum Beispiel „Spec: 11 22 33 · Zeilen 120–380“. Stimmt das nicht, klick **„Andere Spec wählen“** und such die richtige aus.
  Das startet einen neuen Abgleich mit denselben Dateien (nichts muss neu abgelegt werden; der erste bleibt unter „Zuletzt“).
- Findet das Tool die Spec nicht sicher heraus, fragt es oben **„Welche Spec gehört zu diesem PDF?“**. Bis du eine wählst, wird **nichts eingetragen**.
- Hat die CDE nur eine Spec, siehst du davon nichts.

## Was die Farben bedeuten

Genau wie in der fertigen Excel:
- hellblau und durchgestrichen = **alter Text, fällt weg**
- rot = **neuer Text, kommt dazu**
- normale Schrift = bleibt, wie es ist

## Tastatur (wie „Überarbeiten“ in Word)

`J` oder `↓` nächste Karte · `K` oder `↑` vorige · `A` übernehmen · `N` nicht übernehmen · `?` Hilfe · `Esc` schließen.
Nach `A` oder `N` springt das Tool zur nächsten offenen Karte.

## Was heißt „Bitte prüfen“?

Das Tool trägt **nur Sicheres** von selbst ein. Ist es unsicher (z. B. die Excel sieht an der Stelle anders aus als das PDF, ein Wort
ist nur halb markiert, der Text kommt zweimal vor), macht es einen **Vorschlag**. Ohne dein „Übernehmen“ passiert damit nichts.

## Wo liegen die Ergebnisse?

- Am einfachsten: in Schritt ③ herunterladen (landet in deinem Download-Ordner).
- Zusätzlich bleibt alles im Projektordner unter `data\runs\<Datum-Uhrzeit-…>\` (Ordner einstellbar in `config\paths.yaml`):
  `eingabe\` = Kopien deiner Dateien, `ergebnis_001\`, `ergebnis_002\` … = je erstellter Excel die neue Excel + das Protokoll.
- **Deine Original-Dateien werden nie verändert.** Es wird nichts überschrieben und nichts gelöscht.

## Wenn oben ein Warnhinweis steht

- **„Die Bilanz stimmt nicht“**: Es kann etwas verloren gegangen sein. Ergebnis bitte nicht verwenden und melden.
- **„PDF unvollständig gelesen“**: Auf den genannten Seiten konnte das Tool nicht alles lesen. Diese Seiten bitte selbst im PDF ansehen.
- Alles Technische (Wörter-Bilanz, erkannter Excel-Aufbau, Farben im PDF, Prüfcodes, Lauf-ID) steht ganz unten unter
  **„Details für Fachleute“** – brauchst du normalerweise nicht.

## Was das Tool nicht kann

- Keine Scans lesen (keine Texterkennung), keine Tabellen, Bilder oder zweispaltigen Text – solche Stellen werden **gemeldet**, nicht übernommen.
- Kommentare und Sprechblasen am Rand des PDFs werden nicht übernommen (nur gezählt).
- Es entscheidet nichts Unsicheres für dich und ersetzt nicht den Blick auf das Ergebnis: Stichproben in der neuen Excel lohnen sich.
- Läuft nur auf diesem Rechner; andere Personen im Netz können es nicht öffnen.
