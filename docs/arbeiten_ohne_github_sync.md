# Arbeiten ohne GitHub-Synchronisierung (Firmennetz blockiert Git)

Autor: Nico Lehr · Stand: 16.10.2026 · für Windows, in einfachen Worten

**Lage:** Die Cloud-Sitzung von Claude kann nicht auf deinen lokalen Ordner zugreifen. Dein Rechner kann Git nicht benutzen, aber im Browser GitHub-Seiten öffnen und Dateien herunterladen.
Darum läuft der Austausch so: **Updates kommen als ZIP-Datei zu dir, Ergebnisse gehen als Text zurück.**

## 1. Update holen und einspielen (Cloud ⇒ dein Rechner)
1. **Sicherung (nur beim allerersten Mal):** Kopiere deinen ganzen Projektordner einmal von Hand (z. B. `CDE_Aenderungsuebernahme_Sicherung`).
2. Im Browser herunterladen: `https://github.com/nico-lehr/CDE_Aenderungsuebernahme/archive/refs/heads/main.zip`
   (Datei heißt dann meist `CDE_Aenderungsuebernahme-main.zip`, liegt in „Downloads“).
3. Eingabeaufforderung im Projektordner öffnen und ausführen:
   `scripts\update_einspielen.bat "C:\Users\DEIN_NAME\Downloads\CDE_Aenderungsuebernahme-main.zip"`
   Das Skript
   - entpackt die ZIP in einen Temp-Ordner,
   - legt **zuerst eine Sicherung** des bisherigen Codes an (`data\update_sicherung\<Datum_Uhrzeit>`),
   - kopiert die neuen Dateien über den Projektordner und **löscht nichts**,
   - lässt **`data\`, `.venv\`, `.claude\`, `.git\` und alle `*.local.yaml` unberührt** (deine Beispiele, Berichte, Steckbriefe, Umgebung und lokale Einstellungen bleiben).
4. Danach: hat sich `requirements.txt` geändert, einmal `scripts\setup.bat`. Dann `.venv\Scripts\python -m pytest -q` (alles grün?).
5. Falls das Skript etwas Falsches überschrieben hat: Der alte Stand liegt in `data\update_sicherung\…`.

*Hinweis:* Wurde in einem neuen Stand eine Datei **gelöscht**, bleibt sie bei dir stehen. Das ist bewusst so (sicher). Wenn eine alte Datei stört, sag Bescheid.
*Hinweis:* Das Skript kopiert sich selbst vor dem Start in den Temp-Ordner, damit es sich beim Update selbst ersetzen kann.

## 2. Ergebnisse an Claude zurückgeben (dein Rechner ⇒ Cloud)
Nur **Text und Zahlen**, nie PDF oder Excel:
- die geprüfte `steckbrief.json` (und `steckbrief.md`) – Anleitung: `docs/anleitung_steckbrief.md` (vorher `--pruefe`!),
- die Zahlen aus `messen` (`bericht.md`, mit `--schablonen` für Fehlerformen),
- bereinigte Antworten aus `docs/fragebogen_ollama.md`.
Füge den Inhalt direkt in den Chat ein (Steckbrief ca. 25 KB Text).

## 3. Optional: Lokale Claude-Code-Sitzung statt Cloud
Wenn deine IT die Verbindung zu Anthropic erlaubt, kannst du Claude Code lokal in VS Code im Projektordner starten (Prompt dafür: `docs/prompt_lokale_sitzung.md`). Dann entfallen ZIP und Skript.
**Datenschutz:** Auch diese Sitzung schickt gelesene Inhalte an Anthropic. Sie darf deshalb **keine echten Dokumente lesen** (Regeln im Prompt). Ollama allein ersetzt Claude Code nicht; Ollama ist nur der Leser für deine echten Daten.

## 4. Wenn etwas schiefgeht
- „Zielordner sieht nicht wie das Projekt aus“: Das Skript im richtigen Projektordner starten oder den Ordner als zweiten Wert angeben.
- „Entpacken fehlgeschlagen“: ZIP vollständig heruntergeladen? Datei nicht per Rechtsklick „Eigenschaften ⇒ Zulassen“ blockiert?
- Unsicher? Erst die Sicherung des Ordners von Hand machen und dann das Skript ausführen.
