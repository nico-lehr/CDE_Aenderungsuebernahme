# Fragebogen für deine lokale Ollama (Beispielprojekte verstehen, ohne Inhalte zu teilen)

Autor: Nico Lehr · Stand: 16.10.2026

**Zweck:** Die Cloud-Sitzung von Claude darf deine echten Projekte nicht sehen. Deine lokale Ollama darf. Sie beantwortet einen festen Fragebogen **in Worten**
(Struktur, Gewohnheiten, Sonderfälle) – so versteht Claude das Zielbild besser, ohne dass Inhalte das Haus verlassen.
Das ist die **Ergänzung** zum Steckbrief (`docs/anleitung_steckbrief.md`), der Zahlen und Formen liefert. Der Steckbrief ist der Hauptweg; Ollama erklärt, was ein Programm nicht erkennt.

## Regeln (wichtig)
1. **Du liest die Antwort selbst**, bevor du sie weitergibst. Kleine Modelle irren sich und kopieren manchmal echte Sätze in ihre Antwort. Streiche alles Echte
   (Namen, Zahlen mit Bedeutung, Fachbegriffe, die auf ein Projekt hinweisen, ganze Sätze aus dem Dokument).
2. Beispiele in der Antwort müssen **erfundene** Wörter und Zahlen sein (z. B. „Pumpe“, „Ventil“, „Zebrafink“), im selben **Stil** wie das Original (Satzlänge, Zeichensetzung, Schreibweise).
3. Gib Claude nur: die bereinigte Antwort und den Steckbrief. **Nie** PDF, Excel oder Auszüge daraus.
4. Wenn Ollama etwas nicht weiß, soll sie „weiß ich nicht“ sagen und nicht raten.

## So gehst du vor
1. Ollama starten (z. B. `ollama run <dein-modell>`; ein Textmodell mit 8–14 B reicht; für Seitenbilder ein Bildmodell).
2. Den **Prompt** unten einfügen, dann Auszüge aus **einem** Projekt (Text aus dem PDF und aus der Excel – nur lokal, in Ollama) dazugeben.
   Bei langen Dokumenten: ein Kapitel nach dem anderen und die Antworten zusammenführen.
3. Antwort lesen, bereinigen, hier in die Cloud-Sitzung einfügen. Je Projekt eine eigene Antwort.

## Prompt zum Einfügen in Ollama
```
Du hilfst mir, ein Änderungs-Tool zu verbessern. Ich zeige dir gleich Auszüge aus einem echten Projekt (PDF mit Änderungen, Excel vorher, Excel nachher).
Beantworte NUR die Fragen unten, in einfachem Deutsch, kurz und sachlich.
STRENGE REGELN:
- Zitiere NIE echte Texte, Namen, Zahlen oder Fachbegriffe aus den Auszügen. Wenn du ein Beispiel brauchst, erfinde eines im gleichen STIL
  (gleiche Satzlänge, Zeichensetzung, Schreibweise), aber mit frei erfundenen Wörtern (z. B. Pumpe, Ventil, Zebrafink) und erfundenen Zahlen.
- Wenn du etwas nicht sicher weißt, schreibe "weiß ich nicht". Rate nicht.
- Keine Bewertung des Inhalts, nur Aufbau und Gewohnheiten.

FRAGEN
1. Aufbau der Excel: Gibt es eine Zeile je Section, eine Zeile je Absatz oder beides gemischt? Welche Spalten gibt es und wofür (Nummer, Titel, Inhalt, Status, Kommentar, laufende Nummer, Formeln)?
2. Wie sehen Section-Nummern aus (z. B. 1, 1.1, 1.1.1)? Werden sie als Zahl oder als Text geführt? Gibt es Besonderheiten (Gliederung springt, Nummern doppelt, Buchstaben)?
3. Was steht in den Zellen: ganze Absätze, mehrere Absätze mit Zeilenumbruch, Aufzählungen, Tabellen als Text, Verweise auf andere Kapitel?
4. Wie sieht die Änderungsmarkierung im PDF aus? Welche Farben für NEU und ENTFÄLLT? Immer gleich oder je Dokument verschieden? Gibt es Durchstreichungen, Unterstreichungen, Änderungsbalken am Rand, Hinterlegungen?
5. Wie sieht eine ERSETZTE Stelle in der Excel NACHHER aus? Wo steht der alte Text (blau durchgestrichen?), wo der neue (rot?), wie sind sie getrennt (Leerzeichen, Zeilenumbruch, nichts)? Gilt das für Wörter, Sätze und ganze Absätze gleich?
6. Wie sieht eine NEUE Section / ein NEUER Absatz in der Excel NACHHER aus (neue Zeile? welche Spalten gefüllt? Farbe? Wird eine laufende Nummer vergeben? Werden Formeln oder feste Werte aus der Nachbarzeile übernommen)?
7. Wie sieht eine GANZ GESTRICHENE Section in der Excel NACHHER aus (Zeile bleibt? Inhalt blau durchgestrichen? Nummer und Titel auch?)?
8. Welche Sonderfälle kommen im PDF vor: Tabellen, Bilder, mehrere Spalten, Fußnoten, Kopf-/Fußzeilen, Silbentrennung am Zeilenende, Absätze über Seitenumbrüche, Aufzählungszeichen?
9. Gibt es Änderungen im PDF, die in der Excel NICHT übernommen werden (z. B. nur Formatierung, Fußnoten, Kopfzeilen)? Woran erkennt man das?
10. Gibt es Stellen, an denen die Excel nicht genau dem alten PDF-Text entspricht (von Hand nachgebessert, gekürzt, anders formuliert)? Wie häufig (selten/oft)?
11. Was ist für den Menschen, der das heute von Hand macht, der mühsamste oder fehleranfälligste Teil?

Gib am Ende ein Beispiel einer ersetzten Stelle als ERFUNDENE Zeile im Muster:
  vorher: <erfundener Satz>
  PDF:    <derselbe Satz mit {-altem Wort-} und {+neuem Wort+} in erfundenen Wörtern>
  nachher in der Excel: [normal]…[blau durchgestrichen]…[rot]…[normal]…
```

## Was Claude damit macht
Die Antworten fließen in ADR 0005 (echte Farben und Darstellung), in den Nachbau-Generator (gleiche Struktur mit erfundenen Wörtern) und in die Prioritäten der nächsten Etappe.
Widerspricht ein Steckbrief-Wert einer Antwort, gilt der **Steckbrief** (gemessen).
