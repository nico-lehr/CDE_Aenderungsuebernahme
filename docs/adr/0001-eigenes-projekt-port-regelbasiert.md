# ADR 0001 – Eigenes Projekt, Port 5100, regelbasierter Kern

Autor: Nico Lehr · Datum: 15.10.2026 · Status: angenommen

## Entscheidung
1. Eigenständiges Projekt neben der bess-pipeline; bess bleibt read only.
2. Dashboard auf **127.0.0.1:5100** (`config/server.yaml`), eigener Datenordner `data/`, eigene `.venv`, kein CORS, keine Cookies.
3. Kern **regelbasiert** (pdfplumber + Textvergleich + openpyxl). Kein LLM im Kern.
4. Arbeit immer auf einer **Kopie** der Excel; Ergebnisse append-only (`data/runs/<id>/ergebnis_NNN/`).
5. Unsichere Treffer nie automatisch (Regel 27); **Bilanz-Prinzip** (Wörter-Summe muss zum PDF passen).

## Warum
- Getrennte Ports/Daten verhindern Konflikte mit der Masterarbeit (Regel 26: Masterarbeit hat Vorrang).
- Regelbasiert ist nachvollziehbar, wiederholbar, offline und braucht keine Datenschutzfreigabe für Cloud-Dienste.
- Kopie + Protokoll: ein Mensch kann jede Änderung prüfen; nichts geht still verloren.

## Folgen
- Ein lokales LLM (über `LLMClient`) wäre nur ein späterer Prüfer.
- JSON-Schema versioniert (`SCHEMA_VERSION = 1`); Änderungen brauchen neues ADR/Notiz.
- Einschränkungen siehe README (nur einspaltige Text-PDFs, keine Tabellen/Bilder, keine Scans).
