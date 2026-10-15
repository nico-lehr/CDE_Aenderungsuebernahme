"""Gemeinsame Datenobjekte (kanonisches Format, versioniert -- Regel 24).

SCHEMA_VERSION wird in jedes Ergebnis geschrieben, damit alte Laeufe dem
damals gueltigen Aufbau zugeordnet werden koennen.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

SCHEMA_VERSION = 1

# Status eines Wortes im PDF
KEEP, DEL, ADD, MOD = "keep", "del", "add", "mod"

# Status einer Aenderung im Ergebnis
UEBERNOMMEN = "uebernommen"
PRUEFEN = "pruefen"
NICHT = "nicht_uebernommen"

# Konfidenz
HOCH, MITTEL, NIEDRIG = "hoch", "mittel", "niedrig"


@dataclass
class PdfToken:
    """Ein Wort im PDF. old/new = Wortform im alten bzw. neuen Stand."""
    old: str
    new: str
    status: str
    page: int
    bbox: tuple[float, float, float, float]  # x0, top, x1, bottom (PDF-Punkte)


@dataclass
class PdfBlock:
    idx: int
    page: int
    kind: str                       # "absatz" | "ueberschrift"
    tokens: list[PdfToken]
    bbox: tuple[float, float, float, float]
    section_alt: str | None = None  # Section-Nummer im alten Stand
    section_neu: str | None = None  # ... im neuen Stand
    ignoriert: str | None = None    # z.B. "kopf_fuss"

    @property
    def alt_tokens(self) -> list[PdfToken]:
        return [t for t in self.tokens if t.old]

    @property
    def alt_text(self) -> str:
        return " ".join(t.old for t in self.tokens if t.old)

    @property
    def neu_text(self) -> str:
        return " ".join(t.new for t in self.tokens if t.new)

    @property
    def hat_aenderung(self) -> bool:
        return any(t.status != KEEP for t in self.tokens)

    @property
    def ganz_neu(self) -> bool:
        return bool(self.tokens) and all(t.status == ADD for t in self.tokens)

    @property
    def ganz_geloescht(self) -> bool:
        return bool(self.tokens) and all(t.status == DEL for t in self.tokens)


@dataclass
class Operation:
    """Eine konkrete Aenderung an der Excel (immer bezogen auf den Originalzustand)."""
    art: str            # ersetzen | einfuegen | zelle_leeren | zeile_einfuegen | zeile_loeschen
    blatt: str
    zeile: int
    spalte: int = 0
    start: int = 0      # Zeichenposition in der Zelle (ersetzen/einfuegen)
    ende: int = 0
    text: str = ""
    seq: int = 0        # Reihenfolge bei gleicher Stelle
    zellen: dict[int, str] = field(default_factory=dict)   # zeile_einfuegen: Spalte -> Text
    stil_von: int = 0   # zeile_einfuegen: Zeile, deren Format kopiert wird
    abgedeckt: bool = False  # durch andere Operation erledigt
    aenderung: str = ""      # ID der Aenderung, zu der die Operation gehoert
    vor: bool = False        # zeile_einfuegen: True = vor `zeile`, sonst nach `zeile`
    absatz_weg: bool = False  # ersetzen: ein ganzer Absatz wird entfernt (Leerzeile danach aufraeumen)
    erwartet: str = ""      # ersetzen: Text, der vor der Aenderung an der Stelle stehen muss


@dataclass
class Aenderung:
    id: str
    typ: str
    seite: int
    bbox: tuple[float, float, float, float]
    section: str | None
    alt: str
    neu: str
    woerter: int
    status: str = NICHT
    konfidenz: str = NIEDRIG
    grund: str = ""
    ziel: str = ""                     # z.B. "Anforderungen!C14"
    block_idx: int = -1
    ops: list[Operation] = field(default_factory=list)
    kontext: str = ""
    excel_text: str = ""              # aktueller Zelleninhalt (Auszug) an der Zielstelle
    score: float = 0.0
    bestaetigt: bool = False           # Mensch hat einen Vorschlag (pruefen) freigegeben
    nachkontrolle: str = ""            # ok | abweichung (nach dem Speichern geprueft)

    def zu_dict(self) -> dict[str, Any]:
        from dataclasses import asdict
        return asdict(self)
