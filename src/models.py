"""Gemeinsame Datenobjekte (kanonisches Format, versioniert -- Regel 24).

SCHEMA_VERSION wird in jedes Ergebnis geschrieben, damit alte Laeufe dem
damals gueltigen Aufbau zugeordnet werden koennen.

Schema 2 (ADR 0002): nur NEUE Felder mit Standardwert (hinweis/hinweise/wirkung) und
das Luecken-Format. Schema-1-Laeufe bleiben lesbar.

Schema 3 (ADR 0004): Gegenprobe. plan.json bekommt `absaetze` (die PDF-Absaetze als Laeufe, damit das Ergebnis
NACH dem Speichern gegen das PDF geprueft werden kann, nicht nur gegen den eigenen Plan); Aenderung bekommt
gruppe/pruefcodes/soll_absatz/ist_absatz/zweitmeinung, Operation formeln/leer_spalten. Alles mit Standardwert,
Schema 1 und 2 bleiben lesbar.

Schema 4 (ADR 0006/0007, Auftrag 003): Word-Aenderungsverfolgung und CDE wie Nico. Aenderung bekommt bereich
("" = in der CDE, "ausserhalb" = Vorspann/Revisionstabelle u. a., gehoert nicht in die CDE), stuecke (die fertigen
Ergebnis-Laeufe der Zielzelle fuer die Anzeige: [[format, text], ...] mit format normal|alt|neu) und grund_einfach
(ein Satz in einfachem Deutsch). Alles mit Standardwert, Schema 1-3 bleiben lesbar.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

SCHEMA_VERSION = 4
GELESENE_SCHEMAS = (1, 2, 3, 4)   # Versionen, die beim Laden akzeptiert werden

# Status eines Wortes im PDF
KEEP, DEL, ADD, MOD = "keep", "del", "add", "mod"

# Status einer Aenderung im Ergebnis
UEBERNOMMEN = "uebernommen"
PRUEFEN = "pruefen"
NICHT = "nicht_uebernommen"

# Konfidenz
HOCH, MITTEL, NIEDRIG = "hoch", "mittel", "niedrig"

# Wirkung einer angewendeten Aenderung (hat die Excel sich WIRKLICH wie geplant geaendert?)
WIRKUNG_OK = "ok"                    # Zelle/Zeile ist nachweislich wie geplant geaendert
WIRKUNG_KEINE = "keine"              # Operation hatte keine Wirkung (z.B. nur Leerzeichen) -> nie "uebernommen"
WIRKUNG_ABWEICHUNG = "abweichung"    # Ergebnis weicht vom Plan ab
# "" = nicht angewendet

# Luecken: Dinge, die das Tool nicht (sicher) lesen konnte. Jede Luecke macht die Bilanz "unvollstaendig".
# Format (einfaches dict, JSON-tauglich, KEINE Dokumentinhalte, Regel 16):
#   {"art": <LUECKEN_ARTEN>, "seite": int, "beschreibung": "einfacher deutscher Satz", "woerter": int}
LUECKEN_ARTEN = ("scan_seite", "mehrspaltig", "tabelle", "bild", "unbekannte_farbe",
                 "nicht_lesbarer_text", "unerklaerte_markierung", "sonstiges")   # Schema 3: unerklaerte_markierung


def luecke(art: str, seite: int, beschreibung: str, woerter: int = 0) -> dict[str, Any]:
    """Baut einen Luecken-Eintrag im vereinbarten Format (siehe oben)."""
    return {"art": art, "seite": seite, "beschreibung": beschreibung, "woerter": woerter}


@dataclass
class PdfToken:
    """Ein Wort im PDF. old/new = Wortform im alten bzw. neuen Stand."""
    old: str
    new: str
    status: str
    page: int
    bbox: tuple[float, float, float, float]  # x0, top, x1, bottom (PDF-Punkte)
    hinweis: str = ""               # Schema 2: z.B. "gemischt_blau_strich", "verschmolzen" (leer = nichts auffaellig)
    # Schema 3, nur intern (nicht in plan.json):
    trenner: str = " "              # was zwischen diesem und dem NAECHSTEN Wort steht: " ", "\xa0", "\t" oder "" (geklebt)
    teile: list[tuple[str, str]] = field(default_factory=list)   # Zeichen-Laeufe [(status, text)]; leer = ganzes Wort hat `status`
    trennstelle: int = -1           # Position im Wort, an der eine Zeilenend-Trennung zusammengefuegt wurde (-1 = keine);
                                    # steht dort ein "-", ist er optional (die Excel darf ihn haben oder nicht)


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
    hinweise: list[str] = field(default_factory=list)   # Schema 2: Auffaelligkeiten des Blocks (einfaches Deutsch)
    fortsetzung: str = ""           # Schema 3, intern: der Absatz geht im NAECHSTEN Block weiter ("sicher" | "moeglich" | "")

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
    formeln: dict[int, str] = field(default_factory=dict)    # Schema 3, zeile_einfuegen: Spalte -> Formel (aus der Vorlagenzeile)
    leer_spalten: list[int] = field(default_factory=list)    # Schema 3, zeile_einfuegen: Spalten, die bewusst leer bleiben


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
    hinweise: list[str] = field(default_factory=list)   # Schema 2: Warnungen zu dieser Aenderung (einfaches Deutsch)
    wirkung: str = ""                  # Schema 2: WIRKUNG_OK | WIRKUNG_KEINE | WIRKUNG_ABWEICHUNG | "" (nicht angewendet)
    gruppe: str = ""                   # Schema 3: Aenderungen derselben Gruppe gelten nur gemeinsam (z. B. Section + Inhalt)
    pruefcodes: list[str] = field(default_factory=list)   # Schema 3: maschinenlesbare Gruende ("a1_abweichung", "sperre:unbekannte_farbe", ...)
    soll_absatz: str = ""              # Schema 3: der ganze Absatz, wie ihn das PDF nach der Aenderung verlangt ("PDF sagt")
    ist_absatz: str = ""               # Schema 3: der Absatz, wie er vorher in der Excel steht ("Excel hat")
    zweitmeinung: dict[str, Any] = field(default_factory=dict)   # Schema 3: optionale KI-Zweitmeinung (nur Hinweis/Bremse)
    bereich: str = ""                  # Schema 4: "" = in der CDE | "ausserhalb" = gehoert nicht in die CDE (kein offener Punkt)
    stuecke: list[list[str]] = field(default_factory=list)   # Schema 4: Ergebnis der Zielzelle [[format, text], ...] (Anzeige)
    grund_einfach: str = ""            # Schema 4: Grund/Status in einem Satz, einfaches Deutsch (Dashboard)

    def zu_dict(self) -> dict[str, Any]:
        from dataclasses import asdict
        return asdict(self)


# ---------------------------------------------------------------- Schema 3: Absaetze fuer die Gegenprobe (ADR 0004)
LAUF_ARTEN = ("k", "d", "a", "s")   # k = bleibt, d = entfaellt, a = neu, s = Trenner (Leerraum zwischen Woertern)


@dataclass(frozen=True)
class Lauf:
    """Ein Stueck eines PDF-Absatzes. Zwei Laeufe ohne Trenner dazwischen sind aneinandergeklebt (z. B. Wort + Punkt)."""
    art: str                 # LAUF_ARTEN
    text: str
    aenderung: str = ""      # ID der Aenderung (bei d/a); leer bei k/s
    optional: bool = False   # darf in der Excel fehlen (z. B. Trennstrich an einer Zeilenend-Trennung)

    def zu_liste(self) -> list[Any]:
        return [self.art, self.text, self.aenderung, self.optional]

    @staticmethod
    def aus_liste(x: list[Any]) -> "Lauf":
        art, text, aenderung, optional = (list(x) + ["", False])[:4]
        if art not in LAUF_ARTEN or not isinstance(text, str):
            raise ValueError("ungueltiger Lauf")
        return Lauf(str(art), text, str(aenderung or ""), bool(optional))


@dataclass
class Absatz:
    """Ein PDF-Absatz und seine Stelle in der Excel (Originalzustand), fuer die Gegenprobe nach dem Speichern.

    zellen: [(zeile, spalte, start, ende)] -- Zeichenbereich im ORIGINAL-Zelltext; mehrere Zellen (z. B. Nummer + Titel)
    werden fuer den Vergleich mit einem Leerzeichen verbunden."""
    id: str
    blatt: str
    zellen: list[tuple[int, int, int, int]]
    laeufe: list[Lauf]

    @property
    def aenderungen(self) -> set[str]:
        return {l.aenderung for l in self.laeufe if l.aenderung}

    def zu_dict(self) -> dict[str, Any]:
        return {"id": self.id, "blatt": self.blatt, "zellen": [list(z) for z in self.zellen],
                "laeufe": [l.zu_liste() for l in self.laeufe]}

    @staticmethod
    def aus_dict(d: dict[str, Any]) -> "Absatz":
        zellen = [tuple(int(v) for v in z) for z in d["zellen"]]
        if any(len(z) != 4 for z in zellen):
            raise ValueError("ungueltige Zelle")
        return Absatz(str(d["id"]), str(d["blatt"]), zellen, [Lauf.aus_liste(x) for x in d["laeufe"]])  # type: ignore[arg-type]
