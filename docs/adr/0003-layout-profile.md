# ADR 0003 – Layout-Profile: Excel-Aufbau von Hand festlegen, aber nie direkt live

Autor: Nico Lehr · Datum: 15.10.2026 · Status: angenommen (Regel 24: nur zusätzliche Felder mit Standardwert, keine neue Schema-Version)

## Anlass
Die Erkennung des Excel-Aufbaus (Section-, Titel-, Inhaltsspalte; Modell „zeile_pro_section“ oder „zeile_pro_absatz“) ist eine Heuristik
(`src/excel_model.py`). Bei ungewöhnlichen Excel-Typen ist sie unsicher („mittel“/„niedrig“, Treffer gehen dann nur als Vorschlag durch) oder liegt falsch.
Das Team braucht einen Weg, den Aufbau für einen Excel-Typ **von Hand festzulegen**. Der Auftrag verlangt aber: **neue Vorlagen nie direkt live,
erst Vorschau und Prüfung.**

## Entscheidung
1. **Profil = YAML.** Ein Profil hat einen `name`, optional ein Erkennungsmuster (`erkennung`: Blattnamen, Kopfzeilen-Begriffe) und je Tabellenblatt
   entweder `ignorieren: true` oder `section_spalte`, `titel_spalte`, `inhalt_spalte`, `modell`, `kopfzeile`. Spalten sind Buchstaben („C“).
   Format und Beispiel: `config/layouts.yaml` (Vorlage, nur erfundene Namen, `aktiv: false`).
2. **Speicherorte.** `config/layouts.yaml` (im Repo, nur Vorlage) und `config/layouts.local.yaml` (von Git ignoriert, hier liegen **echte** Profile,
   weil echte Blatt- und Spaltennamen nie ins Repo dürfen, Regel 16). Beide Dateien werden gelesen; ein Name darf nur einmal vorkommen
   (sonst wird keines der beiden benutzt). Vorschauen und Bestätigungen liegen unter `data/profile/` (Pfad `profile_dir` in `config/paths.yaml`,
   Regel 28; mit `CDE_DATA_ROOT` unter dem Test-Datenordner). Der Ordner ist **append-only** (Regel 29): Dateien werden mit „x“ neu angelegt, nie
   überschrieben oder gelöscht.
3. **Nie direkt live.** Ein Profil wird nur benutzt, wenn
   (a) es in einer Datei steht und `aktiv: true` hat **und**
   (b) zu **genau diesem Inhalt** eine bestätigte Vorschau existiert.
   Weg: `profil vorschau` (zeigt Spalten, Modell, Sicherheit, Zahl der Sections und Kandidaten, Warnungen – **keine Zellinhalte, keine Blattnamen**;
   schreibt einen Vorschau-Datensatz) → `profil bestaetigen --vorschau ID` (die bewusste Handlung; schreibt eine Bestätigung) → erst dann nimmt
   `runs.analysiere(run_id, profil=NAME)` das Profil an. Sonst gibt es eine Fehlermeldung in einfachem Deutsch, und es entsteht keine Analyse.
4. **Hash-Bindung.** Die Bestätigung gilt für den **Hash (SHA-256) des Profil-Inhalts** (Name, Erkennung, alle Blatt-Angaben; Reihenfolge der Blätter
   egal). `aktiv` und `beschreibung` gehören nicht zum Hash. Ändert jemand nach der Bestätigung eine Spalte, passt der Hash nicht mehr: Das Profil wird
   abgelehnt („nach der Bestätigung geändert“), bis eine neue Vorschau bestätigt ist. Auch `bestaetigen` lehnt ab, wenn die Datei seit der Vorschau
   geändert wurde. Bestätigungsdateien, deren Inhalt nicht mehr zu ihrem Hash passt, zählen nicht.
5. **Haken im Excel-Modell.** `ExcelModell(pfad, profil=None)` (Signatur kompatibel). Das Profil **ersetzt die Erkennung je Blatt** (die Erkennungslogik
   selbst bleibt unverändert). Die Sicherheit eines festgelegten Blatts ist „hoch“ (die Angabe kommt vom Menschen); passen die Daten nicht dazu
   (z. B. keine Gliederungsnummer in der Section-Spalte), gibt es eine Warnung. Eine Spalte außerhalb des Blatts ist ein Fehler. Blätter, die das Profil
   nicht nennt, werden wie bisher automatisch erkannt; `ignorieren: true` heißt: keine Kandidaten aus diesem Blatt. Der Zwischenspeicher
   (`_MODELLE`) berücksichtigt den Profil-Hash. `beschreibung()` weist „Profil: NAME“ aus (nur bei Profil; ohne Profil bleibt alles wie vorher).
6. **Was in `plan.json` steht.** Neuer Schlüssel `"profil": {"name": ..., "hash": ...}` oder `null`. Alte Läufe ohne den Schlüssel bleiben lesbar.
7. **Eintragen bleibt konsistent.** `uebernehmen_lauf` lädt das Profil aus der **Bestätigung** (nicht aus der YAML) und stellt es für die Dauer des
   Eintragens als Standard-Profil bereit (`excel_model.profil_kontext`, ein `ContextVar`), weil der Applier die Excel selbst nochmals einliest. So gilt beim
   Eintragen derselbe Aufbau wie bei der Analyse, auch wenn die YAML inzwischen geändert oder auf `aktiv: false` gesetzt wurde.
8. **Zurücknehmen** heißt `aktiv: false` (nichts wird gelöscht). Das Erkennungsmuster (`erkennung`) liefert nur einen **Vorschlag** in `profil vorschau`
   ohne `--profil`; es wendet nie etwas an.

## Warum so
- **Hash statt Name oder Zeitstempel:** Ein Name sagt nichts darüber, *was* geprüft wurde. Der Hash macht „genau das, was ich in der Vorschau gesehen
  habe“ prüfbar; ein später geänderter Aufbau kann nicht unbemerkt live gehen.
- **Bestätigung speichert den Inhalt mit:** Läufe bleiben reproduzierbar, auch wenn die YAML später geändert wird.
- **Vorschau ohne Zellinhalte:** Regel 16 – die Ausgabe darf bedenkenlos geteilt werden (nur Zahlen, Buchstaben, Blattnummern).
- **Profil nur als Ergänzung der Heuristik:** Die Erkennung bleibt der Standard; wer nichts einstellt, merkt keinen Unterschied.

## Folgen
- API und Dashboard können `profil` an `runs.analysiere` durchreichen und `plan.json["profil"]` anzeigen (Antrag an die Besitzer von `api/` und `frontend/`).
- Der Applier könnte `profil` später als ausdrücklichen Parameter bekommen; bis dahin genügt der Kontext aus Punkt 7.
- Ein Widerruf mit Protokolleintrag (statt `aktiv: false`) und Profile je Korpus sind denkbar, aber nicht Teil von v1.
