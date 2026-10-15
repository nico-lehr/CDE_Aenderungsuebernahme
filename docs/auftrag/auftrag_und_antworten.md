# Auftrag und Antworten (bereinigt, ohne echte Inhalte)

Autor: Nico Lehr · Quelle: Übergabe-Text und Antworten von Nico im Chat vom Oktober 2026

## Auftrag
Side-Projekt **CDE-Änderungsübernahme**, komplett getrennt von der Masterarbeit (bess-pipeline).
Ein PDF mit sichtbaren Änderungen wird gelesen, die inhaltlichen Änderungen werden in eine **bestehende CDE-Excel** übernommen.
Das Tool ist für ein anderes Team gedacht.

## Harte Vorgaben
- `bess-pipeline` ist **read only** (nichts bearbeiten, ändern, verschieben). Nur dieser Projektordner ist Arbeitsordner.
- Eigene Struktur erlaubt; Architektur von bess übernehmen, wo sinnvoll (z. B. Dashboard-Frontend).
- **Nicht** auf demselben localhost/Port wie bess; keine Konflikte (Lösung: Port 5100, eigener Datenordner, eigene .venv).
- Immer in einer `.venv` arbeiten, Python 3.11; gleiche Regeln wie bess (siehe `.claude/entwicklungsregeln.md`).
- Alles lokal, keine API-Keys, kein `git add -A`, echte Daten des anderen Teams nur nach Freigabe, Test- und Entwicklungsbeispiele getrennt, Konfiguration in YAML getrennt vom Code, neue Vorlagen nie direkt live, Bilanz-Prinzip.
- Umgebung: Windows, VS Code, Ollama (12 GB VRAM). Laufwerk M: ist langsam (Code besser auf C:). pip im Firmennetz blockiert → **nicht** mit `--trusted-host` umgehen.
- Dokumente, die weitergegeben werden könnten, tragen den Namen Nico Lehr.
- Zusatz: Ordner wird in die Cloud verschoben, Claude Code arbeitet dort weiter → Ordner enthält alles Nötige (Dashboard-Referenz, Regeln, README, .gitignore). Cloud-Arbeit nur mit erfundenen Daten.

## Vorgehen laut Übergabe
Erst Konzept in einfachen Worten (`docs/konzept_v1.md`: Zielbild, Messung, Machbarkeitstest); ein Auftrag an Claude Code für den Bau erst nach echten Beispielen und Nicos Zustimmung.

## Antworten von Nico auf die Rückfragen
1. **PDF-Änderungen:** meist PDFs. **Blau** = Inhalt/Text kommt hinzu. **Rot oder durchgestrichen** = Inhalt entfällt – bis hin zu ganzen Sections/Subsections. Farblegende: im Moment noch nicht, später ja.
2. **Excel:** chronologisch nach Sections aufgebaut, dahinter Inhalt (meist Textblöcke). Kein einheitliches Format; das Tool soll sich dynamisch anpassen und nur den Inhalt überarbeiten, wie es das PDF vorgibt.
3. **Varianten:** eine Variante direkt für alle Fälle.
4. **Team/Entscheider:** irrelevant.
5. **Beispiele:** kann Nico nächste Woche besorgen. Jetzt muss erst alles funktionieren und gebaut werden. Keine geschwärzten Dokumente.
