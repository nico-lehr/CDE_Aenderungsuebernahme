# ADR 0002 – Schema 2, Lücken, Wirkung, Archivieren

Autor: Nico Lehr · Datum: 15.10.2026 · Status: angenommen (Regel 24: Schema-Änderung → Version erhöht, Notiz hier)

## Anlass
Die Code-Prüfung vom 15.10.2026 (siehe `docs/ziel_v1.md`) zeigte: Das Tool kann „übernommen / ok“ melden, obwohl die Excel falsch oder
unverändert ist, und die Bilanz grün zeigen, obwohl das PDF gar nicht vollständig gelesen wurde. Dafür brauchen die Module gemeinsame,
vorher festgelegte Felder. Sie werden hier **einmal** festgelegt, damit Helfer parallel arbeiten können, ohne sich zu überschreiben.

## Entscheidung
1. **`SCHEMA_VERSION` 1 → 2.** Es kommen nur Felder mit Standardwert dazu. Beim Laden gelten `GELESENE_SCHEMAS = (1, 2)`.
   `protokoll.aenderung_aus_dict` füllt fehlende Felder mit dem Standardwert und ignoriert unbekannte (neuere) Felder.
   Alte Läufe unter `data/runs/` bleiben lesbar (Test: `tests/test_schema2.py`).
2. **Neue Felder (alle mit Standardwert):**
   - `PdfToken.hinweis: str` – Auffälligkeit eines Wortes (z. B. `gemischt_blau_strich`, `verschmolzen`).
   - `PdfBlock.hinweise: list[str]` – Auffälligkeiten des Blocks, einfaches Deutsch.
   - `Aenderung.hinweise: list[str]` – Warnungen zu dieser Änderung, einfaches Deutsch, sichtbar im Dashboard.
   - `Aenderung.wirkung: str` – `"ok"` (Excel nachweislich wie geplant geändert) · `"keine"` (Operation hatte keine Wirkung; dann **nie** „übernommen“) ·
     `"abweichung"` · `""` (nicht angewendet). Gesetzt von `applier.py`; die bestehende `nachkontrolle` bleibt aus Kompatibilität.
3. **Lücken** (`models.luecke(art, seite, beschreibung, woerter)`): Dinge, die das Tool nicht (sicher) lesen konnte. Arten: `scan_seite`, `mehrspaltig`, `tabelle`, `bild`,
   `unbekannte_farbe`, `nicht_lesbarer_text`, `sonstiges`. Die Beschreibung enthält **nie Dokumentinhalt** (Regel 16).
   Weg: `PdfAnalyse.luecken` → `Plan.luecken` → `plan.json["luecken"]` → `bilanz(..., luecken)`.
4. **Bilanz:** `stimmt` bleibt „Wortsumme passt“. Neu: `luecken` (Anzahl) und `vollstaendig` (= keine Lücken). **Ein Lauf mit Lücken ist nie „sauber grün“**:
   Dashboard und Protokoll zeigen „unvollständig“ und nennen die Lücken. Das Protokoll darf den Satz „nichts ging still verloren“ nur bei `vollstaendig` schreiben.
5. **Konfig:** Neue Schlüssel stehen mit den **heutigen** Werten in `config/paths.yaml` (`paare_dir`, `messung_dir`) und `config/zuordnung.yaml`
   (`matcher.*`, `anwenden.*`, `excel.modell_schwelle`, `excel.titel_max_zeichen`). Verhalten ist dadurch unverändert. Die Kern-Module verdrahten sie.
   `CDE_DATA_ROOT` verschiebt auch `paare_dir` und `messung_dir` (für Tests).
6. **Läufe sind append-only (Regel 29):** Die Löschfunktion wird durch **Archivieren** ersetzt (Lauf bleibt auf der Platte, erscheint nicht in der Standardliste).
   Fehlgeschlagene Analysen löschen sich nicht mehr selbst. (Umsetzung: Paket „Läufe + API“.)
7. **Red-Team-Tests ohne Dauer-Rot:** Fehlertests tragen `@pytest.mark.redteam` und `@pytest.mark.xfail(strict=True, reason="RT-nn …, Besitzer …")`.
   Wird der Fehler repariert, ist der Test XPASS und damit rot; dann muss der Marker entfernt werden. `pyproject.toml` setzt `xfail_strict = true`.
8. **Bei Gefahr für die Excel-Struktur** (Formeln, Tabellen, Validierung, Filter, benannte Bereiche, verbundene Zellen an der Einfügestelle) wird nicht repariert,
   sondern die Änderung auf „prüfen“ mit Hinweis gesetzt (Regel 27).

## Warum so
- Nur Felder mit Standardwert → keine Migration nötig, alte Läufe laufen weiter.
- Lücken getrennt von Wortsumme: Die Wortsumme prüft nur, was das Tool *erkannt* hat. Was es nicht erkannt hat, kann sie nie sehen.
- `wirkung` getrennt von `status`: „übernommen“ darf nur stehen, wenn die Excel nachweislich anders ist.

## Folgen
- Frontend zeigt `vollstaendig`, `luecken`, `hinweise`, `wirkung` (Paket Dashboard).
- Weitere Feldwünsche gehen als Antrag an den Orchestrator; höchstens eine Version 3 am Ende.
- Layout-Profile (Paket 3) schreiben Profil und Profil-Hash in `plan.json`; das braucht ein eigenes ADR (0003), keine weitere Schema-Version, wenn nur Felder mit Standardwert.
