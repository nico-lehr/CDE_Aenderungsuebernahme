# Konzept v1 – CDE-Änderungsübernahme (in einfachen Worten)

Autor: Nico Lehr · Stand: 15.10.2026 · Status: Entwurf, noch nicht mit echten Beispielen geprüft

## 1. Zielbild

Heute: Jemand liest ein PDF mit blauen und roten/durchgestrichenen Stellen und trägt jede Änderung **von Hand** in die CDE-Excel ein.
Das ist langsam und fehleranfällig.

Künftig: PDF und Excel ins Tool ziehen → das Tool zeigt eine **Liste aller Änderungen** und was es damit machen würde → du prüfst
(vor allem die unsicheren) → es erzeugt eine **neue Excel** und ein **Protokoll**. Die alte Excel bleibt unverändert.

Regeln dabei:
- blau = neu dazu, rot oder durchgestrichen = fällt weg (Wort, Satz, Absatz, ganze Section/Subsection)
- nur **Inhalt** ändern, Aufbau und Formatierung der Excel möglichst lassen
- die Excel hat kein festes Format → das Tool erkennt den Aufbau selbst
- lieber **nicht** übernehmen als falsch übernehmen

## 2. So arbeitet es (ohne KI)

1. **PDF lesen:** Jedes Wort bekommt einen Status: bleibt / neu / entfällt (aus Textfarbe, Hintergrund, Durchstreichlinie oder PDF-Markierung).
   Wörter werden zu Zeilen und Absätzen. Kopf-/Fußzeilen werden ignoriert.
2. **Excel verstehen:** Welche Spalte ist Section-Nummer, Titel, Inhalt? Eine Zeile pro Section oder pro Absatz?
3. **Stelle finden:** Jeder PDF-Absatz wird mit den Excel-Texten verglichen (ähnliche Wörter, Section-Nummer als Hilfe).
   Ergebnis: Sicherheit *hoch*, *mittel* oder *niedrig*.
4. **Plan:** Was genau wird ersetzt, eingefügt, gelöscht, welche Zeile neu?
5. **Anwenden** auf einer Kopie. Vorher wird geprüft, ob an der Stelle wirklich der erwartete alte Text steht; hinterher wird die gespeicherte Datei nochmal gelesen und kontrolliert.

Warum ohne KI: nachvollziehbar, wiederholbar (gleiches Ergebnis bei gleicher Eingabe), läuft ohne Internet und ohne Datenschutzfragen.
Ein lokales Modell (Ollama) wäre höchstens später ein zusätzlicher *Prüfer*.

## 3. Messen: Woran sehen wir, ob es gut ist?

| Messgröße | Bedeutung | Ziel (Vorschlag, mit Beispielen bestätigen) |
|---|---|---|
| **Bilanz stimmt** | Wörter übernommen + zu prüfen + nicht übernommen + ignoriert = markierte Wörter im PDF | immer 100 % |
| **Treffergenauigkeit** | Anteil der automatisch übernommenen Änderungen, nach denen die Excel exakt der echten "nachher"-Excel entspricht | möglichst ≈ 100 % (falsche Auto-Übernahmen sind das Hauptrisiko) |
| **Abdeckung** | Anteil der Änderungen, die ohne Nachfrage übernommen werden | so hoch wie möglich, aber nie auf Kosten der Genauigkeit |
| **Stille Fehler** | Änderungen falsch übernommen, ohne Warnung | Ziel: 0 |
| **Zeitgewinn** | Minuten von Hand vs. mit Tool (inkl. Prüfen) | später schätzen |

Vergleich mit der Wahrheit: Beispielpaare haben eine Excel *vorher* und eine Excel *nachher* (von Hand erstellt). Das Tool läuft auf "vorher" + PDF, das Ergebnis wird mit "nachher" verglichen.
Entwicklungsbeispiele (zum Einstellen) und Testbeispiele (nur zur Endmessung, Regel 11) werden getrennt.

## 4. Machbarkeitstest (sobald echte Beispiele da sind)

1. 3–5 Paare (PDF + Excel vorher + Excel nachher). Möglichst verschieden (andere Aufbauten, mit Section-Löschungen).
2. Pro Paar: Tab "Farben im PDF" prüfen – werden Blau/Rot richtig erkannt?
3. Lauf starten, Bilanz ansehen, Ergebnis mit "nachher" vergleichen.
4. Fehler sortieren: PDF falsch gelesen / Excel-Aufbau falsch erkannt / Stelle falsch zugeordnet / Plan falsch.
5. Schwellen in `config/` anpassen **nur anhand der Entwicklungsbeispiele**, dann einmal mit den Testbeispielen messen.
6. Entscheidung: reicht regelbasiert, oder braucht es zusätzlich ein Layout-Profil je Excel-Typ?

Erfolg = Bilanz immer stimmig, keine stillen Fehler, ein großer Teil ohne Nachfrage richtig.

## 5. Offene Punkte

- Datenschutz / Code-Rechte für die Daten des anderen Teams (noch nicht geklärt) → bis dahin nur erfundene Daten, Cloud-Arbeit nur daran.
- Farblegende (welche Farbe = was) wird später ergänzt; bis dahin: blau = neu, rot/durchgestrichen = entfällt.
- Tabellen/Bilder im PDF, Scans: derzeit nicht unterstützt.
- Wer im anderen Team prüft/freigibt: laut Nico aktuell "irrelevant".
