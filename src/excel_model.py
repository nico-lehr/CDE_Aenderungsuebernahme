"""Liest eine bestehende CDE-Excel und beschreibt sie, ohne ein festes Format vorauszusetzen.

Die Excel ist "chronologisch nach Sections aufgebaut, dahinter der Inhalt".
Wie genau, wird je Tabellenblatt aus den Daten erkannt:

  * Section-Spalte  : Spalte mit den meisten Gliederungsnummern (1, 2.1, 3.2.4 ...)
  * Titel-Spalte    : Spalte mit der Ueberschrift neben der Nummer (falls getrennt)
  * Inhalt-Spalte   : Spalte mit dem meisten Text
  * Modell          : "zeile_pro_section"  (Inhalt steht in der Zeile der Section)
                      "zeile_pro_absatz"   (jeder Absatz hat eine eigene Zeile)
  * Sicherheit      : "hoch" | "mittel" | "niedrig" -- wie sicher der Aufbau erkannt wurde
                      (bei "mittel"/"niedrig" gibt es eine Warnung; der Matcher uebernimmt dort nie automatisch)

CDE-Register (ADR 0006 Punkt 5, Konfiguration excel.register): Kopfbereich ueber mehrere Zeilen, Kopfzeile mit vielen kurzen
Spaltenkoepfen, Beispielzeile(n), dann EINE Zeile je Absatz/Listenpunkt. Rollen: Section-Spalte (Nummer in jeder Zeile,
"1.1", "Part 1"), Sub-Section-Spalte (A ... Z, AA, BB ...), Textspalte (Listennummer + Polster IM Text), Kennwort-Spalte
(ein fester kurzer Wert nur in Ueberschriftzeilen). Erkannt wird nur aus der Struktur, nie aus Spaltennamen.
Modell bleibt "zeile_pro_absatz", dazu `BlattLayout.register` und die Zusatzrollen.

Aus allen Textzellen werden "Kandidaten" gebildet (Absatz, ganze Zelle, Kopf aus
Nummer+Titel); im Register nur aus der Textspalte der Datenzeilen. Gegen diese Kandidaten werden die PDF-Bloecke abgeglichen.
Zellen mit Rich Text werden ueber ihre NEU-ANSICHT gelesen (durchgestrichene Laeufe = alter Text frueherer Runden zaehlen
nicht, ADR 0005/0006); die Positionen aller Operationen beziehen sich auf diese Ansicht (`roh_position` bildet ab).
Reine Python-Logik; schreibt nie in die Datei.
"""
from __future__ import annotations

import hashlib
import re
import weakref
import zipfile
from collections import Counter
from array import array
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import TYPE_CHECKING, Any, Iterable, Iterator

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.utils import coordinate_to_tuple, get_column_letter, range_boundaries
from openpyxl.worksheet.formula import ArrayFormula

from . import config
from .textutil import kopf_zerlegen, norm, section_nummer, woerter

if TYPE_CHECKING:                      # nur fuer Typangaben; layouts.py importiert dieses Modul, nicht umgekehrt
    from .layouts import BlattProfil, Profil

HOCH_S, MITTEL_S, NIEDRIG_S = "hoch", "mittel", "niedrig"   # Sicherheit des Aufbaus
_FESTE_STELLEN = re.compile(r"^0\.(0+)$")                    # Zahlenformat "0.0", "0.00", ...
_BEZUG = re.compile(r"(?<![A-Za-z0-9_.$'])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!)?"
                    r"\$?[A-Za-z]{1,3}\$?(\d+)(?![\d(])(?::\$?[A-Za-z]{1,3}\$?(\d+))?")


def zahl_als_text(wert: int | float, zahlenformat: str = "General") -> str:
    """Text einer Zahlzelle so, wie Excel ihn anzeigt bzw. speichert (nie "%g": 2.10 bleibt 2.10).

    * ganze Zahlen: ohne Nachkommastellen ("3")
    * Kommazahlen mit festem Zahlenformat "0.0" / "0.00": mit genau diesen Stellen (2.1 mit "0.00" -> "2.10")
    * sonst: kuerzeste Schreibweise mit 15 gueltigen Stellen (so speichert Excel), z. B. "2.1"
    """
    if isinstance(wert, int) or float(wert).is_integer():
        return str(int(wert))
    m = _FESTE_STELLEN.match(zahlenformat or "")
    if m:
        stellen = len(m.group(1))
        return format(Decimal(repr(float(wert))).quantize(Decimal(1).scaleb(-stellen), rounding=ROUND_HALF_UP), "f")
    return "%.15g" % wert


def neu_ansicht(wert: CellRichText, zellschrift: Any = None) -> tuple[str, list[int] | None, bool]:
    """Neu-Ansicht einer Rich-Text-Zelle (ADR 0005/0006): durchgestrichene Laeufe (alter Text frueherer Runden, jede Farbe)
    zaehlen nicht, andere (auch farbige) Laeufe sind normaler Text. Wie markierung.stuecke_aus_zelle: ein Teil ohne eigene
    Schrift hat die Zellschrift. -> (Text, Abbildung Ansicht -> Rohposition oder None wenn gleich, alle sichtbaren Teile fett)."""
    teile: list[tuple[str, bool, bool]] = []
    for b in wert:
        f = b.font if isinstance(b, TextBlock) else zellschrift
        text = b.text if isinstance(b, TextBlock) else str(b)
        if isinstance(b, TextBlock) and f is None:
            f = zellschrift
        teile.append((text, bool(getattr(f, "strike", None)), bool(getattr(f, "b", None))))
    fett = all(b for t, s, b in teile if t.strip() and not s) and any(t.strip() and not s for t, s, _ in teile)
    if not any(s and t for t, s, _ in teile):
        return "".join(t for t, _, _ in teile), None, fett
    text: list[str] = []
    abb: list[int] = []
    p = 0
    for t, s, _ in teile:
        if not s:
            text.append(t)
            abb.extend(range(p, p + len(t)))
        p += len(t)
    return "".join(text), abb, fett


def _register_cfg(cfg: dict) -> dict:
    """Abschnitt excel.register (ADR 0006 Punkt 5); fehlt er, ist die Register-Erkennung aus."""
    return cfg.get("register") or {"aktiv": False}


def section_kanon(text: str | None, rc: dict) -> str | None:
    """Section-Wert einer Register-Zeile in einheitlicher Form: "1.1 " -> "1.1", "PART 1" / "Part 1 " -> "Part 1"
    (Name aus excel.register.kapitel_name, wie der PDF-Leser); kein Gliederungswert -> None."""
    if text is None:
        return None
    s = str(text).strip()
    m = re.match(rc.get("kapitel_muster", r"^(?:part|teil)\s*(\d{1,2})$"), s, re.I)
    if m:
        return str(rc.get("kapitel_name", "Part {n}")).replace("{n}", str(int(m.group(1))))
    if re.match(rc.get("section_muster", r"^[1-9]\d{0,2}(?:\.\d{1,3}){0,6}$"), s):
        return s
    return None


def listen_label(text: str, rc: dict) -> tuple[str, str, int]:
    """Listennummer + Polster am Anfang eines Register-Texts: "a.\\xa0\\xa0 Text" -> ("a.", "\\xa0\\xa0 ", 4).
    Ohne Listennummer: ("", "", 0)."""
    m = re.match(rc.get("listen_label", r"^(?P<label>\d{1,3}[.)])(?P<polster>[\t ]+)(?=\S)"), text)
    if not m:
        return "", "", 0
    return m.group("label"), m.group("polster"), m.end()


def spec_norm(text: Any, rc: dict) -> str | None:
    """Spezifikationsnummer in einheitlicher Form (Punkt 7, Welle 3): Leerzeichen, Unterstrich, Punkt und Bindestrich zaehlen
    gleich, Zusaetze (z. B. eine Region) und Leerzeichen am Ende sind egal: "11_22_33 " / "11.22.33 EMEA" -> "11 22 33".
    Muster: excel.register.spec_muster (Ziffergruppen). Kein Spec-Wert -> None."""
    if text is None:
        return None
    m = re.search(rc.get("spec_muster", r"(?<![\d])(\d{2})[ _.\-](\d{2})[ _.\-](\d{2})(?:[ _.\-](\d{2}))?(?![\d])"),
                  str(text))
    return " ".join(g for g in m.groups() if g) if m else None


def label_art(label: str) -> str:
    """Art einer Listennummer fuer das Polster neuer Zeilen: "zahl." "zahl)" "klein." "klein)" "gross." "gross)"."""
    if not label:
        return ""
    kern, ende = label[:-1], label[-1:]
    art = "zahl" if kern.isdigit() else ("gross" if kern.isupper() else "klein")
    return art + ende


# Register-Zeilentypen (ADR 0006): Ueberschriften (Kennwort in der Kennwort-Spalte) und Text
REGISTER_UEBERSCHRIFT = ("part", "section", "sub_titel", "liste_titel")


@dataclass(frozen=True)
class RegisterZeile:
    """Eine Datenzeile eines CDE-Registers (nur Formen, fuer Zuordnung und neue Zeilen)."""
    zeile: int
    art: str                 # part | section | sub | sub_titel | liste | liste_titel | absatz
    section: str | None      # einheitliche Section ("1.1", "Part 1")
    section_roh: str         # Zelltext der Section-Spalte (z. B. mit Leerzeichen am Ende)
    sub: str                 # Buchstabe der Sub-Section-Spalte ("" = leer)
    label: str               # Listennummer im Text ("" = keine)
    polster: str             # Tab / geschuetzte Leerzeichen nach der Listennummer
    koerper_start: int       # Position im Text, an der der Text nach Listennummer + Polster beginnt


@dataclass(slots=True)
class XTok:
    text: str
    norm: str
    blatt: str
    zeile: int
    spalte: int
    start: int   # Zeichenposition in der Zelle
    end: int


class Kandidat:
    """Eine Stelle in der Excel, mit der ein PDF-Block verglichen wird (Absatz, ganze Zelle oder Kopf).

    `tokens` (Woerter mit Zeichenposition) werden erst beim ersten Zugriff als XTok-Objekte gebaut; bis dahin
    liegen nur die Normformen und die Positionen kompakt im Speicher (grosse Blaetter bleiben schlank)."""
    __slots__ = ("id", "art", "blatt", "zeile", "spalte", "text", "section", "start", "end", "zelle_text",
                 "zelle_ist_zahl", "versteckt", "section_unsicher", "kopf_spalte", "_tokens", "_norms", "_pos", "_zellen")

    def __init__(self, id: int, art: str, blatt: str, zeile: int, spalte: int, tokens: list[XTok] | None = None,
                 text: str = "", section: str | None = None, start: int = 0, end: int = 0, zelle_text: str = "",
                 zelle_ist_zahl: bool = False, versteckt: str = "", section_unsicher: bool = False,
                 kopf_spalte: bool = False, *,
                 norms: list[str] | None = None, pos: array | None = None, zellen: dict | None = None) -> None:
        self.id, self.art, self.blatt, self.zeile, self.spalte = id, art, blatt, zeile, spalte
        self.text, self.section, self.start, self.end, self.zelle_text = text, section, start, end, zelle_text
        self.zelle_ist_zahl = zelle_ist_zahl      # die Excel-Zelle ist eine Zahl (kein Text); beim Ersetzen Zahlenformat beachten
        self.versteckt = versteckt                # "" | "blatt" | "zeile" | "spalte": Stelle ist ausgeblendet
        self.section_unsicher = section_unsicher  # Section-Nummer stammt aus mehrdeutigen Zahlen (z. B. 2.1 und 2.10)
        self.kopf_spalte = kopf_spalte            # Stelle liegt in der Nummern- oder Titelspalte (Ueberschrift), nicht im Fliesstext
        self._tokens = tokens
        self._norms = [x.norm for x in tokens] if tokens is not None else (norms or [])
        self._pos, self._zellen = pos, zellen

    @property
    def tokens(self) -> list[XTok]:
        if self._tokens is None:
            p, z = self._pos, self._zellen
            self._tokens = [XTok(z[(self.blatt, p[4 * i], p[4 * i + 1])][p[4 * i + 2]:p[4 * i + 3]], n, self.blatt,
                                 p[4 * i], p[4 * i + 1], p[4 * i + 2], p[4 * i + 3])
                            for i, n in enumerate(self._norms)]
        return self._tokens

    @property
    def ziel(self) -> str:
        return f"{self.blatt}!{get_column_letter(self.spalte)}{self.zeile}"

    @property
    def norms(self) -> list[str]:
        return self._norms


@dataclass
class BlattLayout:
    name: str
    section_spalte: int | None = None
    titel_spalte: int | None = None
    inhalt_spalte: int | None = None
    modell: str = "zeile_pro_absatz"
    kopf_zeilen: list[int] = field(default_factory=list)
    max_zeile: int = 0
    max_spalte: int = 0
    kombiniert: bool = False       # Nummer und Titel stehen in einer Zelle ("3.2 Titel")
    sicherheit: str = HOCH_S       # hoch | mittel | niedrig: wie sicher der Aufbau erkannt wurde
    sicherheit_gruende: list[str] = field(default_factory=list)   # Zweifel in einfachem Deutsch
    versteckt: bool = False        # Blatt ist ausgeblendet
    unsichere_sections: set[str] = field(default_factory=set)     # Section-Nummern aus mehrdeutigen Zahlen
    profil: str = ""               # Name des Layout-Profils, das diesen Aufbau festlegt ("" = automatisch erkannt)
    ignoriert: bool = False        # Das Profil sagt: dieses Blatt wird nie abgeglichen
    # ADR 0006 Punkt 5: CDE-Register (Modell zeile_pro_absatz + Zusatzrollen)
    register: bool = False
    untergliederung_spalte: int | None = None   # Sub-Section (A ... Z, AA ...); bei Listenpunkten der Buchstabe der Eltern-Sub-Section
    kennwort_spalte: int | None = None          # festes Kennwort in Ueberschriftzeilen
    kennwort: str = ""                          # (Zellinhalt: nie in Meldungen oder Vorschauen)
    kopfzeile: int = 0                          # Zeile der Spaltenkoepfe (0 = keine)
    daten_ab: int = 0                           # erste Datenzeile (darueber: Kopfbereich, Kopfzeile, Beispielzeilen)
    spec_spalte: int | None = None              # Spezifikationsnummer je Zeile (eine CDE kann mehrere Specs enthalten)

    def beschreibung(self) -> dict[str, Any]:
        b = lambda c: get_column_letter(c) if c else None  # noqa: E731
        d = {"blatt": self.name, "section_spalte": b(self.section_spalte), "titel_spalte": b(self.titel_spalte),
             "inhalt_spalte": b(self.inhalt_spalte), "modell": self.modell,
             "sections": len(self.kopf_zeilen), "zeilen": self.max_zeile,
             "sicherheit": self.sicherheit, "sicherheit_gruende": list(self.sicherheit_gruende),
             "versteckt": self.versteckt}
        if self.profil:                # ohne Profil bleibt die Beschreibung genau wie frueher
            d["profil"] = f"Profil: {self.profil}"
            d["ignoriert"] = self.ignoriert
        if self.register:              # nur Spaltenbuchstaben und Zeilennummern (das Kennwort ist ein Zellinhalt)
            d.update(register=True, untergliederung_spalte=b(self.untergliederung_spalte),
                     kennwort_spalte=b(self.kennwort_spalte), kopfzeile=self.kopfzeile, daten_ab=self.daten_ab)
            if self.spec_spalte:
                d["spec_spalte"] = b(self.spec_spalte)
        return d


@dataclass
class _BlattEigenschaften:
    """Excel-Eigenschaften eines Blatts, die beim Einfuegen/Loeschen von Zeilen brechen koennen."""
    verbunden: list[tuple[int, int]] = field(default_factory=list)       # (erste, letzte Zeile)
    formeln: list[tuple[int, int, str]] = field(default_factory=list)    # (Zeile, Spalte, Text)
    formeln_viele: bool = False
    bedingt: list[tuple[int, int]] = field(default_factory=list)
    pruefung: list[tuple[int, int]] = field(default_factory=list)
    filter: tuple[int, int] | None = None
    tabellen: list[tuple[str, int, int]] = field(default_factory=list)   # (Name, erste, letzte Zeile)
    namen: list[tuple[str, str]] = field(default_factory=list)           # (Name, Bezug)
    links: list[int] = field(default_factory=list)
    gruppiert: list[int] = field(default_factory=list)                   # ausgeblendete oder gruppierte Zeilen
    anker: list[int] = field(default_factory=list)                       # Bilder / Diagramme (erste Zeile)


def _bereiche(ref: Any) -> list[tuple[int, int]]:
    """'A1:C9 E2:E5' -> [(1, 9), (2, 5)] (nur die Zeilen)."""
    out = []
    for teil in str(ref).replace(",", " ").split():
        try:
            _, r0, _, r1 = range_boundaries(teil)
        except (ValueError, TypeError):
            continue
        out.append((r0 or 1, r1 or 1048576))
    return out


def _belegte_zellen(ws) -> list:
    """Nur die Zellen, die es wirklich gibt (nicht die ganze Blattgroesse durchlaufen)."""
    roh = getattr(ws, "_cells", None)
    if roh is not None:
        return [roh[k] for k in sorted(roh)]
    return [c for row in ws.iter_rows(max_row=min(ws.max_row, 20000), max_col=min(ws.max_column, 200)) for c in row]


# Dieselbe Datei (gleicher Inhalt, gleiche Einstellungen) wird nur einmal eingelesen, solange das Modell noch irgendwo
# gebraucht wird (z. B. Analyse und Eintragen im selben Lauf). Das Modell ist nach dem Einlesen schreibgeschuetzt gedacht.
_MODELLE: "weakref.WeakValueDictionary[tuple[str, str, str, str], ExcelModell]" = weakref.WeakValueDictionary()

# Layout-Profil, das gilt, wenn ExcelModell ohne ausdruecklichen `profil` aufgerufen wird. So kann runs.py auch den
# Applier (der die Excel selbst nochmals einliest) denselben Aufbau sehen lassen wie die Analyse. Standard: kein Profil.
_STANDARD_PROFIL: "ContextVar[Profil | None]" = ContextVar("cde_standard_profil", default=None)


@contextmanager
def profil_kontext(profil: "Profil | None") -> Iterator[None]:
    """Innerhalb dieses Blocks benutzt jedes ExcelModell(pfad) ohne eigenen Parameter dieses (bestaetigte) Profil."""
    token = _STANDARD_PROFIL.set(profil)
    try:
        yield
    finally:
        _STANDARD_PROFIL.reset(token)


def _schreiber() -> str:
    """anwenden.schreiber (ADR 0007): patch (Standard, verlustfrei) | openpyxl. Manche Warnungen gelten nur fuer openpyxl."""
    return str((config.lade("zuordnung").get("anwenden") or {}).get("schreiber", "patch"))


def _datei_hash(pfad: Path) -> str:
    """SHA-1 stueckweise berechnen (grosse Dateien nie komplett in den Speicher laden)."""
    h = hashlib.sha1()
    with open(pfad, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


class ExcelModell:
    def __new__(cls, pfad: str | Path, profil: "Profil | None" = None) -> "ExcelModell":
        if profil is None:
            profil = _STANDARD_PROFIL.get()
        try:
            p = Path(pfad)
            schluessel = (str(p.resolve()), _datei_hash(p),
                          repr(sorted(config.lade("zuordnung")["excel"].items())) + "|" + _schreiber(),
                          profil.hash if profil is not None else "")
        except (OSError, TypeError, KeyError):
            schluessel = None
        if schluessel is not None and schluessel in _MODELLE:
            return _MODELLE[schluessel]
        self = super().__new__(cls)
        self._schluessel = schluessel
        self._fertig = False
        return self

    def __init__(self, pfad: str | Path, profil: "Profil | None" = None) -> None:
        if self._fertig:                      # aus dem Zwischenspeicher: schon eingelesen
            return
        self.pfad = Path(pfad)
        # Profil = vom Menschen festgelegter Aufbau je Blatt (src/layouts.py); es ersetzt dort die Erkennung.
        # Ob es bestaetigt ist, prueft layouts.py, nicht dieses Modul.
        self.profil: "Profil | None" = profil if profil is not None else _STANDARD_PROFIL.get()
        self.warnungen: list[str] = []
        self.layouts: dict[str, BlattLayout] = {}
        self.zellen: dict[tuple[str, int, int], str] = {}
        self.zahl_zellen: set[tuple[str, int, int]] = set()      # Zellen, die in der Excel Zahlen sind
        self.zeilen_section: dict[tuple[str, int], str | None] = {}
        self.kandidaten: list[Kandidat] = []
        self.nach_section: dict[str, list[int]] = {}
        self.nach_blatt_section: dict[tuple[str, str], list[int]] = {}   # wie nach_section, aber je Blatt getrennt
        self.versteckte_zeilen: dict[str, set[int]] = {}
        self.versteckte_spalten: dict[str, set[int]] = {}
        self._eig: dict[str, _BlattEigenschaften] = {}
        self._formel_bezuege: dict[str, list[tuple[str, int]]] = {}
        self._norm_cache: dict[str, str] = {}
        # Rich-Text-Zellen mit durchgestrichenen Laeufen: Position in der Neu-Ansicht -> Position im ganzen Zelltext
        self.roh_abbildung: dict[tuple[str, int, int], list[int]] = {}
        self.gestrichen: dict[tuple[str, int, int], str] = {}     # ... und ihr durchgestrichener Text (Idempotenz)
        self.fett_zellen: set[tuple[str, int, int]] = set()      # Textzellen, deren (sichtbarer) Text ganz fett ist
        self._reg_zeilen: dict[tuple[str, int], RegisterZeile | None] = {}
        self._reg_ueber: dict[str, set[int]] = {}                 # Register: Ueberschriftzeilen je Blatt
        self._rc: dict = {}                                       # Konfiguration excel.register
        self._lesen()
        self._fertig = True
        if self._schluessel is not None:
            _MODELLE[self._schluessel] = self

    # ------------------------------------------------------------ Lesen
    def _lesen(self) -> None:
        cfg = config.lade("zuordnung")["excel"]
        self._rc = _register_cfg(cfg)
        suffix = self.pfad.suffix.lower()
        if suffix not in (".xlsx", ".xlsm"):
            raise ValueError("Nur .xlsx / .xlsm werden unterstützt. Bitte die Datei in Excel als .xlsx speichern.")
        alt_schreiber = _schreiber() == "openpyxl"
        self._pruefe_paket(alt_schreiber)
        wb = load_workbook(str(self.pfad), rich_text=True, keep_vba=(suffix == ".xlsm"))
        n_formeln = n_rich = 0
        globale_namen = [(n, str(getattr(d, "attr_text", "") or "")) for n, d in wb.defined_names.items()]
        for ws in wb.worksheets:
            lay = BlattLayout(name=ws.title, versteckt=(ws.sheet_state != "visible"))
            self.layouts[ws.title] = lay
            eig = self._eig[ws.title] = _BlattEigenschaften()
            eig.namen = list(globale_namen) + [(n, str(getattr(d, "attr_text", "") or "")) for n, d in ws.defined_names.items()]
            max_z = max_s = 0
            belegt: list[tuple[int, int]] = []
            for c in _belegte_zellen(ws):
                if getattr(c, "_hyperlink", None) is not None:
                    eig.links.append(c.row)
                v = c.value
                if v is None:
                    continue
                max_z, max_s = max(max_z, c.row), max(max_s, c.column)
                if isinstance(v, ArrayFormula):
                    self._formel_merken(eig, c.row, c.column, str(v.text or ""))
                    continue
                fett = False
                if isinstance(v, CellRichText):
                    n_rich += 1
                    # ADR 0006: Neu-Ansicht (durchgestrichene Laeufe frueherer Runden zaehlen nicht)
                    roh = str(v)
                    v, abb, fett = neu_ansicht(v, c.font)
                    if abb is not None:
                        self.roh_abbildung[(ws.title, c.row, c.column)] = abb
                        sichtbar = set(abb)
                        self.gestrichen[(ws.title, c.row, c.column)] = "".join(z for i, z in enumerate(roh) if i not in sichtbar)
                elif isinstance(v, str):
                    fett = bool(c.font is not None and c.font.b)
                if isinstance(v, str):
                    if v.startswith("="):
                        n_formeln += 1
                        self._formel_merken(eig, c.row, c.column, v)
                        continue
                    if v.strip():
                        self.zellen[(ws.title, c.row, c.column)] = v
                        belegt.append((c.row, c.column))
                        if fett:
                            self.fett_zellen.add((ws.title, c.row, c.column))
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    self.zellen[(ws.title, c.row, c.column)] = zahl_als_text(v, c.number_format)
                    self.zahl_zellen.add((ws.title, c.row, c.column))
                    belegt.append((c.row, c.column))
            lay.max_zeile, lay.max_spalte = max_z, max_s
            self._eigenschaften_lesen(ws, eig)
            self._versteckt_lesen(ws, belegt)
            if ws.merged_cells.ranges and alt_schreiber:
                self.warnungen.append(f"Blatt '{ws.title}' enthält verbundene Zellen (Zeilen einfügen/löschen verschiebt sie nur teilweise).")
        # Diese Warnungen gelten nur fuer das fruehere Speichern mit openpyxl; der Patch-Schreiber (ADR 0007) zieht Formeln und
        # verbundene Zellen mit und behaelt Teilformatierungen.
        if n_formeln and alt_schreiber:
            self.warnungen.append(f"{n_formeln} Formelzellen werden nicht verändert; Zeilen einfügen/löschen passt Formeln nicht an.")
        if n_rich and alt_schreiber:
            self.warnungen.append(f"{n_rich} Zellen mit gemischter Formatierung (z. B. teilweise fett). Geänderte Zellen verlieren diese Teilformatierung.")
        if self.roh_abbildung:
            self.warnungen.append(f"{len(self.roh_abbildung)} Zellen enthalten durchgestrichenen Text früherer Runden. Er zählt beim "
                                  "Vergleich nicht mit und bleibt unverändert stehen.")
        je_blatt: dict[str, dict[tuple[int, int], str]] = {name: {} for name in self.layouts}
        for (b, r, c), t in self.zellen.items():
            je_blatt[b][(r, c)] = t
        if self.profil is not None:
            fehlt = [n for n in self.profil.blaetter if n not in self.layouts]
            if fehlt:
                self.warnungen.append(f"Das Profil '{self.profil.name}' nennt {len(fehlt)} Blatt/Blätter, die es in dieser Excel nicht gibt "
                                      "(sie werden übergangen).")
            self.warnungen.append(f"Der Aufbau wurde durch das Profil '{self.profil.name}' festgelegt "
                                  "(nicht automatisch erkannt), soweit das Profil die Blätter nennt.")
        for name in self.layouts:
            bp = self.profil.blaetter.get(name) if self.profil is not None else None
            if bp is not None:
                self._layout_aus_profil(name, bp, cfg, je_blatt[name])
            else:
                self._layout_erkennen(name, cfg, je_blatt[name])
        self._kandidaten_bauen()

    @staticmethod
    def _formel_merken(eig: _BlattEigenschaften, zeile: int, spalte: int, text: str) -> None:
        if len(eig.formeln) < 50000:
            eig.formeln.append((zeile, spalte, text))
        else:
            eig.formeln_viele = True

    def _versteckt_lesen(self, ws, belegt: list[tuple[int, int]]) -> None:
        name = ws.title
        zeilen = {r for r, d in ws.row_dimensions.items() if d.hidden}
        spalten: set[int] = set()
        for key, d in ws.column_dimensions.items():
            if d.hidden:
                von = d.min or 0
                bis = d.max or von
                spalten.update(range(von, min(bis, 16384) + 1))
        if zeilen:
            self.versteckte_zeilen[name] = zeilen
        if spalten:
            self.versteckte_spalten[name] = spalten
        if not belegt:
            return
        if ws.sheet_state != "visible":
            self.warnungen.append(f"Blatt '{name}' ist ausgeblendet. Treffer dort werden nie automatisch übernommen, "
                                  "nur als Vorschlag zum Prüfen.")
        n_z = len({r for (r, _) in belegt if r in zeilen})
        if n_z and ws.sheet_state == "visible":
            self.warnungen.append(f"Blatt '{name}' hat {n_z} ausgeblendete Zeilen mit Text. Treffer dort werden nie automatisch "
                                  "übernommen, nur als Vorschlag zum Prüfen.")
        n_s = len({c for (_, c) in belegt if c in spalten})
        if n_s and ws.sheet_state == "visible":
            self.warnungen.append(f"Blatt '{name}' hat {n_s} ausgeblendete Spalten mit Text. Treffer dort werden nie automatisch "
                                  "übernommen, nur als Vorschlag zum Prüfen.")

    def _eigenschaften_lesen(self, ws, eig: _BlattEigenschaften) -> None:
        eig.verbunden = [(r.min_row, r.max_row) for r in ws.merged_cells.ranges]
        for cf in ws.conditional_formatting:
            eig.bedingt += _bereiche(cf.sqref)
        for dv in ws.data_validations.dataValidation:
            eig.pruefung += _bereiche(dv.sqref)
        if ws.auto_filter and ws.auto_filter.ref:
            b = _bereiche(ws.auto_filter.ref)
            eig.filter = b[0] if b else None
        for tab in ws.tables.values():
            for r0, r1 in _bereiche(tab.ref):
                eig.tabellen.append((str(tab.name), r0, r1))
        for h in getattr(ws, "_hyperlinks", []):
            for r0, r1 in _bereiche(h.ref):
                eig.links.append(r0)
        eig.links = sorted(set(eig.links))
        eig.gruppiert = sorted(r for r, d in ws.row_dimensions.items() if d.hidden or (d.outlineLevel or 0) > 0)
        for obj in list(getattr(ws, "_images", [])) + list(getattr(ws, "_charts", [])):
            a = getattr(obj, "anchor", None)
            try:
                if isinstance(a, str):
                    eig.anker.append(coordinate_to_tuple(a)[0])
                elif a is not None:
                    eig.anker.append(a._from.row + 1)
            except (AttributeError, ValueError):
                eig.anker.append(1)

    def _pruefe_paket(self, alt_schreiber: bool = True) -> None:
        try:
            with zipfile.ZipFile(self.pfad) as z:
                namen = z.namelist()
        except zipfile.BadZipFile as e:
            raise ValueError("Die Excel-Datei ist beschädigt oder keine .xlsx-Datei.") from e
        if alt_schreiber and any(n.startswith(("xl/drawings/", "xl/charts/", "xl/media/")) for n in namen):
            self.warnungen.append("Die Excel enthält Bilder, Diagramme oder Formen. Beim Speichern können diese verloren gehen – "
                                  "bitte das Ergebnis mit dem Original vergleichen.")
        if any("vbaProject" in n for n in namen):
            self.warnungen.append("Die Datei enthält Makros (VBA). Sie bleiben nur bei .xlsm erhalten.")

    # ------------------------------------------------------------ Layout
    @staticmethod
    def _kopf_art(text: str, cfg: dict) -> tuple[str, str] | None:
        """('nr', '2.1') fuer eine reine Nummer, ('kombi', '2.1') fuer "2.1 Titel"; Aufzaehlungen im Fliesstext sind keine Section."""
        n = section_nummer(text)
        if n is not None:
            return "nr", n
        kz = kopf_zerlegen(text)
        if kz is None:
            return None
        zeile1 = kz[1].split("\n")[0].strip()
        if (len(zeile1) > cfg["titel_max_zeichen"] or len(zeile1.split()) > cfg["kopf_max_woerter"]
                or zeile1.endswith((".", "!", "?", ";", ":"))):
            return None        # "1. Das System muss ... tun." ist eine Aufzaehlung im Text
        return "kombi", kz[0]

    @staticmethod
    def _kopf_hat(text: str | None, woerter_liste: list[str]) -> bool:
        if not text:
            return False
        for tok in re.findall(r"\w+", text.casefold()):
            for w in woerter_liste:
                if tok == w or (len(w) >= 4 and tok.startswith(w)):
                    return True
        return False

    # ------------------------------------------------------------ CDE-Register (ADR 0006 Punkt 5)
    def _register_erkennen(self, blatt: str, cfg: dict, zellen: dict[tuple[int, int], str]) -> bool:
        """Erkennt ein CDE-Register nur aus der Struktur und setzt dann das Layout. True = Register.

        Section-Spalte: Gliederungswerte ("1.1", "Part 1"), die sich wiederholen (sie stehen in JEDER Datenzeile);
        Textspalte: die meisten Zeichen; Sub-Section-Spalte: Buchstabenfolge A, B ... je Section neu ab A; Kennwort-Spalte:
        ein kurzer Wert, der nur neben kurzem oder fettem Text steht; Kopfzeile: viele kurze Koepfe ueber den Beispielzeilen.
        Ohne Sub-Section- und ohne Kennwort-Spalte ist es kein Register (dann gilt die bisherige Erkennung)."""
        rc = _register_cfg(cfg)
        if not rc.get("aktiv", True):
            return False
        min_z = int(rc.get("min_zeilen", 6))
        werte: dict[int, dict[int, str]] = {}
        for (r, c), t in zellen.items():
            k = section_kanon(t, rc)
            if k is not None:
                werte.setdefault(c, {})[r] = k
        spalten = [c for c, w in werte.items()
                   if len(w) >= min_z and len(w) / len(set(w.values())) >= float(rc.get("wiederholung_min", 1.5))]
        if not spalten:
            return False
        s_sp = max(spalten, key=lambda c: (len(werte[c]), -c))
        daten_ab = min(werte[s_sp])
        laengen: Counter = Counter()
        for (r, c), t in zellen.items():
            if r >= daten_ab and c != s_sp:
                laengen[c] += len(t)
        if not laengen:
            return False
        g_sp = laengen.most_common(1)[0][0]
        if sum(1 for (r, c) in zellen if c == g_sp and r >= daten_ab) < min_z:
            return False
        f_sp = self._untergliederung_spalte(zellen, werte[s_sp], daten_ab, {s_sp, g_sp}, rc)
        k_sp, kennwort = self._kennwort_spalte(blatt, zellen, daten_ab, {s_sp, g_sp, f_sp}, g_sp, rc, cfg)
        if f_sp is None and k_sp is None:
            return False
        lay = self.layouts[blatt]
        lay.register, lay.modell = True, "zeile_pro_absatz"
        lay.section_spalte, lay.titel_spalte, lay.inhalt_spalte = s_sp, None, g_sp
        lay.untergliederung_spalte, lay.kennwort_spalte, lay.kennwort = f_sp, k_sp, kennwort
        lay.kopfzeile, lay.daten_ab = self._kopfzeile(zellen, daten_ab, rc), daten_ab
        lay.spec_spalte = self._spec_spalte(zellen, daten_ab, {s_sp, g_sp, f_sp, k_sp}, s_sp, rc)
        self._register_abschluss(blatt, lay, zellen, rc, cfg)
        return True

    @staticmethod
    def _spec_spalte(zellen: dict[tuple[int, int], str], daten_ab: int, belegt: set, s_sp: int, rc: dict) -> int | None:
        """Spalte mit der Spezifikationsnummer je Zeile: fast alle Werte sind spec-artig (excel.register.spec_muster,
        spec_min_anteil); bei mehreren gewinnt die Spalte links der Section-Spalte mit den meisten Zeilen."""
        je: dict[int, list[bool]] = {}
        for (r, c), t in zellen.items():
            if r >= daten_ab and c not in belegt and t.strip():
                je.setdefault(c, []).append(spec_norm(t, rc) is not None)
        gut = [c for c, w in je.items() if len(w) >= int(rc.get("min_zeilen", 6))
               and sum(w) / len(w) >= float(rc.get("spec_min_anteil", 0.8))]
        return max(gut, key=lambda c: (c < s_sp, sum(je[c]), -c)) if gut else None

    def specs(self, blatt: str) -> dict[str, list[int]]:
        """Datenzeilen je Spezifikationsnummer (einheitliche Form) eines Registers mit Spec-Spalte; sonst {}."""
        lay = self.layouts.get(blatt)
        if lay is None or not lay.register or not lay.spec_spalte:
            return {}
        out: dict[str, list[int]] = {}
        for r in range(lay.daten_ab, lay.max_zeile + 1):
            sp = spec_norm(self.zellen.get((blatt, r, lay.spec_spalte)), self._rc)
            if sp is not None:
                out.setdefault(sp, []).append(r)
        return out

    @staticmethod
    def _sub_rang(v: str) -> int:
        """Reihenfolge der Sub-Section-Buchstaben: A=0 ... Z=25, AA=26, BB=27 ... (wie Word), AB (Excel-Art) dahinter."""
        if len(v) == 1:
            return ord(v) - 65
        if v[0] == v[1]:
            return 26 + ord(v[0]) - 65
        return 52 + 26 * (ord(v[0]) - 65) + ord(v[1]) - 65

    def _untergliederung_spalte(self, zellen: dict[tuple[int, int], str], sec: dict[int, str], daten_ab: int,
                                belegt: set, rc: dict) -> int | None:
        """Spalte mit der Buchstabenfolge A, B ... (je Section neu ab A, nicht fallend); Status- oder Kreuz-Spalten mit
        einzelnen Grossbuchstaben fallen durch (zufaellige Folge, kein "A" oder nur ein Wert)."""
        rx = re.compile(rc.get("untergliederung_muster", r"^[A-Z]{1,2}$"))
        je: dict[int, dict[int, str]] = {}
        for (r, c), t in zellen.items():
            if r >= daten_ab and c not in belegt and rx.match(t.strip()):
                je.setdefault(c, {})[r] = t.strip()
        beste, wert = None, 0
        for c, w in je.items():
            if len(w) < int(rc.get("untergliederung_min", 2)) or "A" not in w.values() or len(set(w.values())) < 2:
                continue
            gut = schritte = 0
            vorher: tuple[str | None, str] | None = None
            for r in sorted(w):
                s = sec.get(r)
                if vorher is not None:
                    schritte += 1
                    if (s != vorher[0] and w[r] == "A") or (s == vorher[0] and self._sub_rang(w[r]) >= self._sub_rang(vorher[1])):
                        gut += 1
                vorher = (s, w[r])
            if schritte and gut / schritte >= float(rc.get("untergliederung_folge_min", 0.9)) and len(w) > wert:
                beste, wert = c, len(w)
        return beste

    def _kennwort_spalte(self, blatt: str, zellen: dict[tuple[int, int], str], daten_ab: int, belegt: set, g_sp: int,
                         rc: dict, cfg: dict) -> tuple[int | None, str]:
        """Spalte, in der ein und derselbe kurze Wert nur in Zeilen mit kurzem oder fettem Text steht (Ueberschriftzeilen) und
        nicht in den meisten Zeilen (Konstanten wie Spezifikation/Revision stehen in fast allen Zeilen). Bei mehreren gewinnt
        die Spalte, deren Zeilen am haeufigsten fetten Text haben, dann die mit mehr Zeilen."""
        titel_max = int(cfg.get("titel_max_zeichen", 160))
        daten = {r for (r, c) in zellen if r >= daten_ab and c == g_sp}
        je: dict[int, Counter] = {}
        for (r, c), t in zellen.items():
            if r >= daten_ab and c not in belegt:
                je.setdefault(c, Counter())[t.strip()] += 1
        beste: tuple[float, int, int, str] | None = None
        for c, zaehler in je.items():
            wort, n = zaehler.most_common(1)[0]
            if n < int(rc.get("kennwort_min_zeilen", 2)) or not wort or len(wort) > int(rc.get("kennwort_max_zeichen", 40)) \
                    or n > float(rc.get("kennwort_max_anteil", 0.6)) * max(len(daten), 1):
                continue
            zeilen = [r for (r, cc), t in zellen.items() if cc == c and r >= daten_ab and t.strip() == wort]
            if not all((g := zellen.get((r, g_sp))) is not None
                       and (len(g.strip()) <= titel_max or (blatt, r, g_sp) in self.fett_zellen) for r in zeilen):
                continue
            fett = sum(1 for r in zeilen if (blatt, r, g_sp) in self.fett_zellen) / len(zeilen)
            if fett < float(rc.get("kennwort_min_fett", 0.5)):
                continue                                   # Ueberschriftzeilen haben fetten Text (Status-Spalten nicht)
            if beste is None or (fett, n) > beste[:2]:
                beste = (fett, n, c, wort)
        return (beste[2], beste[3]) if beste else (None, "")

    @staticmethod
    def _kopfzeile(zellen: dict[tuple[int, int], str], daten_ab: int, rc: dict) -> int:
        """Zeile der Spaltenkoepfe: die naechste Zeile ueber den Daten mit vielen kurzen Textkoepfen (dazwischen
        Beispielzeilen mit Platzhaltern und langen Beispielsaetzen). 0 = keine."""
        for r in range(daten_ab - 1, max(0, daten_ab - 1 - int(rc.get("kopf_suche_zeilen", 4))), -1):
            werte = [t.strip() for (rr, _), t in zellen.items() if rr == r]
            kurz = [t for t in werte if len(t) <= int(rc.get("kopf_max_zeichen", 40)) and "\n" not in t]
            if len(kurz) >= int(rc.get("kopf_min_zellen", 3)) and len(kurz) == len(werte):
                return r
        return 0

    def _register_abschluss(self, blatt: str, lay: BlattLayout, zellen: dict[tuple[int, int], str], rc: dict,
                            cfg: dict) -> None:
        """Ueberschriftzeilen, Section je Zeile und Sicherheit eines erkannten (oder per Profil festgelegten) Registers."""
        ueber: set[int] = set()
        if lay.kennwort_spalte and lay.kennwort:
            ueber = {r for (r, c), t in zellen.items() if c == lay.kennwort_spalte and r >= lay.daten_ab
                     and t.strip() == lay.kennwort}
        else:            # ohne Kennwort: erste Zeile einer neuen Section mit kurzem Text ohne Listennummer und Sub-Section
            vorher = None
            for r in sorted({r for (r, c) in zellen if r >= lay.daten_ab}):
                s = section_kanon(zellen.get((r, lay.section_spalte)), rc)  # type: ignore[arg-type]
                g = zellen.get((r, lay.inhalt_spalte), "")                  # type: ignore[arg-type]
                f = zellen.get((r, lay.untergliederung_spalte), "").strip() if lay.untergliederung_spalte else ""
                if s and s != vorher and not f and not listen_label(g, rc)[0] and 0 < len(g.strip()) <= int(cfg.get("titel_max_zeichen", 160)):
                    ueber.add(r)
                vorher = s or vorher
        self._reg_ueber[blatt] = ueber
        aktuell: str | None = None
        for r in range(1, lay.max_zeile + 1):
            if r >= lay.daten_ab:
                aktuell = section_kanon(zellen.get((r, lay.section_spalte)), rc) or aktuell  # type: ignore[arg-type]
            self.zeilen_section[(blatt, r)] = aktuell if r >= lay.daten_ab else None
        lay.kopf_zeilen = [r for r in sorted(ueber) if (rz := self.register_zeile(blatt, r)) and rz.art in ("part", "section")]
        self._zahl_sections(blatt, lay)
        lay.sicherheit, lay.sicherheit_gruende = HOCH_S, []

    def register_zeile(self, blatt: str, zeile: int) -> RegisterZeile | None:
        """Form einer Datenzeile des Registers (None: kein Register oder keine Datenzeile)."""
        key = (blatt, zeile)
        if key in self._reg_zeilen:
            return self._reg_zeilen[key]
        lay = self.layouts.get(blatt)
        rz: RegisterZeile | None = None
        if lay is not None and lay.register and lay.daten_ab <= zeile <= lay.max_zeile:
            rc = self._rc
            g = self.zellen.get((blatt, zeile, lay.inhalt_spalte), "")             # type: ignore[arg-type]
            e = self.zellen.get((blatt, zeile, lay.section_spalte), "")            # type: ignore[arg-type]
            sub = self.zellen.get((blatt, zeile, lay.untergliederung_spalte), "").strip() if lay.untergliederung_spalte else ""
            if not re.match(rc.get("untergliederung_muster", r"^[A-Z]{1,2}$"), sub):
                sub = ""
            if g or e:
                label, polster, ks = listen_label(g, rc)
                ueber = zeile in self._reg_ueber.get(blatt, set())
                if label:
                    art = "liste_titel" if ueber else "liste"
                elif sub:
                    art = "sub_titel" if ueber else "sub"
                elif ueber:
                    art = "part" if re.match(rc.get("kapitel_muster", r"^(?:part|teil)\s*(\d{1,2})$"), e.strip(), re.I) \
                        else "section"
                else:
                    art = "absatz"
                rz = RegisterZeile(zeile, art, self.zeilen_section.get((blatt, zeile)), e, sub, label, polster, ks)
        self._reg_zeilen[key] = rz
        return rz

    def _zeilen(self, blatt: str, zeilen: Iterable[int] | None) -> list[int]:
        lay = self.layouts[blatt]
        return sorted(zeilen) if zeilen is not None else list(range(lay.daten_ab, lay.max_zeile + 1))

    def register_polster(self, blatt: str, zeile: int, label: str, zeilen: Iterable[int] | None = None) -> str:
        """Polster nach einer neuen Listennummer wie in der naechsten Zeile mit derselben Art Listennummer
        (Tab oder Folge geschuetzter Leerzeichen); sonst excel.register.polster_standard. `zeilen`: nur diese Zeilen ansehen
        (z. B. die der Spec des PDFs)."""
        art = label_art(label)
        for r in sorted(self._zeilen(blatt, zeilen), key=lambda r: (abs(r - zeile), r > zeile)):
            rz = self.register_zeile(blatt, r)
            if rz is not None and rz.label and label_art(rz.label) == art:
                return rz.polster
        return str(self._rc.get("polster_standard", "\t"))

    _VERWANDT = {"part": "section", "section": "part", "sub": "sub_titel", "sub_titel": "sub", "liste": "liste_titel",
                 "liste_titel": "liste", "absatz": "sub"}

    def register_vorlage(self, blatt: str, zeile: int, art: str, zeilen: Iterable[int] | None = None) -> int:
        """Vorlagezeile fuer eine neue Zeile: die naechste Zeile gleichen Typs (oben zuerst), sonst eines verwandten Typs;
        `zeilen`: nur unter diesen Zeilen (Spec des PDFs).

        Auftrag 004: Der Schreiber kopiert die Konstanten links der Section-Spalte (Spezifikation, Revision, Datum ...) WOERTLICH
        aus der Vorlagezeile. Ein Mensch kopiert sie aus der Zeile, an der er einfuegt (`zeile`, dem Anker). Darum hat unter den
        naechsten Zeilen gleichen Typs (`excel.register.vorlage_kandidaten`) die erste den Vorrang, deren Konstanten genauso
        geschrieben sind wie beim Anker (auch ein verirrtes Leerzeichen am Ende); gibt es keine, gilt die naechste wie bisher."""
        reihe = sorted(self._zeilen(blatt, zeilen), key=lambda r: (abs(r - zeile), r > zeile))
        lay = self.layouts[blatt]
        spalten = range(1, lay.section_spalte) if lay.register and lay.section_spalte else range(0)
        wie_anker: dict[int, str] = {}                   # Spalte -> Schreibweise beim Anker (nur fuer wirklich konstante Spalten)
        if spalten and reihe and zeile in set(reihe):
            def glatt(r: int) -> tuple[str, ...]:
                return tuple(self.text(blatt, r, c).strip() for c in spalten)
            gleiche = [r for r in reihe if glatt(r) == glatt(zeile)]       # Zeilen derselben Spec (ohne Rand-Leerzeichen gleich)
            for c in spalten:
                wert, n = Counter(self.text(blatt, r, c).strip() for r in gleiche).most_common(1)[0]
                if n >= 0.9 * len(gleiche) and wert == self.text(blatt, zeile, c).strip():
                    wie_anker[c] = self.text(blatt, zeile, c)
        k_max = int(config.lade("zuordnung")["excel"].get("register", {}).get("vorlage_kandidaten", 60))
        for gesucht in (art, self._VERWANDT.get(art, art)):
            erste: int | None = None
            gesehen = 0
            for r in reihe:
                rz = self.register_zeile(blatt, r)
                if rz is None or rz.art != gesucht:
                    continue
                if erste is None:
                    erste = r
                gesehen += 1
                if all(self.text(blatt, r, c) == w for c, w in wie_anker.items()):
                    return r
                if gesehen >= k_max:
                    break
            if erste is not None:
                return erste
        return zeile

    def roh_position(self, blatt: str, zeile: int, spalte: int, pos: int) -> int:
        """Position `pos` im Text des Modells (Neu-Ansicht) -> Position im ganzen Zelltext (mit durchgestrichenen Laeufen)."""
        abb = self.roh_abbildung.get((blatt, zeile, spalte))
        if abb is None:
            return pos
        if pos < len(abb):
            return abb[pos]
        return abb[-1] + 1 if abb else 0

    def _layout_erkennen(self, blatt: str, cfg: dict, zellen: dict[tuple[int, int], str]) -> None:
        lay = self.layouts[blatt]
        if not zellen:
            return
        if self._register_erkennen(blatt, cfg, zellen):
            return
        zweifel: list[str] = []           # hoechstens "mittel"
        schwer: list[str] = []            # "niedrig"
        sp = get_column_letter
        # Kopfzeile (Spaltenueberschriften): nur die erste belegte Zeile, wenn sie keine Section ist
        kopf_text: dict[int, str] = {}
        if zellen:
            erste = min(r for r, _ in zellen)
            zeile1 = {c: t for (r, c), t in zellen.items() if r == erste}
            if len(zeile1) >= 2 and not any(self._kopf_art(t, cfg) for t in zeile1.values()):
                kopf_text = zeile1
        ist_lfd_kopf = lambda c: self._kopf_hat(kopf_text.get(c), cfg["lfd_kopfwoerter"])           # noqa: E731
        ist_kommentar_kopf = lambda c: self._kopf_hat(kopf_text.get(c), cfg["kommentar_kopfwoerter"])  # noqa: E731
        ist_section_kopf = lambda c: self._kopf_hat(kopf_text.get(c), cfg["section_kopfwoerter"])      # noqa: E731

        treffer: dict[int, list[tuple[int, str, str]]] = {}
        for (r, c), t in sorted(zellen.items()):
            k = self._kopf_art(t, cfg)
            if k:
                treffer.setdefault(c, []).append((r, k[0], k[1]))
        kandidaten = {c: h for c, h in treffer.items() if len(h) >= cfg["section_spalte_min"]}
        if kandidaten:
            struktur = {c for c, h in kandidaten.items() if any(art == "kombi" or "." in nr for _, art, nr in h)}
            pool = {c: h for c, h in kandidaten.items() if c in struktur} or dict(kandidaten)
            sauber = {c: h for c, h in pool.items() if not (ist_lfd_kopf(c) or ist_kommentar_kopf(c))}
            pool = sauber or pool

            def lfd_artig(h: list[tuple[int, str, str]]) -> bool:
                if len(h) < cfg["lfd_min"] or any(art != "nr" or "." in nr for _, art, nr in h):
                    return False
                return [int(nr) for _, _, nr in h] == list(range(1, len(h) + 1))

            def schluessel(c: int):
                return (ist_section_kopf(c) and not ist_lfd_kopf(c), lfd_artig(pool[c]) and not struktur, len(pool[c]))

            spalte = max(pool, key=schluessel)
            lay.section_spalte = spalte
            if ist_lfd_kopf(spalte):
                schwer.append(f"Die Spalte {sp(spalte)} ('{kopf_text[spalte]}') sieht wie eine laufende Nummer aus, "
                              "eine Spalte mit echten Section-Nummern (1, 1.1, 2 ...) wurde nicht gefunden.")
            andere = [c for c in pool if c != spalte and len(pool[c]) >= cfg["inhalt_zweifel_anteil"] * len(pool[spalte])]
            if andere:
                zweifel.append(f"Mehrere Spalten enthalten Section-Nummern ({', '.join(sp(c) for c in [spalte, *andere])}); "
                               f"verwendet wird Spalte {sp(spalte)}.")
        if lay.section_spalte:
            hits = treffer[lay.section_spalte]
            kombi = sum(1 for _, art, _ in hits if art == "kombi")
            lay.kombiniert = kombi > len(hits) - kombi
            lay.kopf_zeilen = [r for r, _, _ in hits]
            kopfset = set(lay.kopf_zeilen)
            if not lay.kombiniert:
                titel: Counter = Counter()
                for r in lay.kopf_zeilen:
                    for c in range(lay.section_spalte + 1, lay.max_spalte + 1):
                        t = zellen.get((r, c))
                        if t and len(t) < cfg["titel_max_zeichen"] and "\n" not in t:
                            titel[c] += 1
                            break
                if titel:
                    c = titel.most_common(1)[0][0]
                    folge = [x for (r, cc), x in zellen.items() if cc == c and r not in kopfset]
                    lang = sum(1 for x in folge if len(x) >= cfg["min_inhalt_zeichen"])
                    if not (lang >= 2 and lang / len(folge) > cfg["titel_fliesstext_anteil"]):
                        lay.titel_spalte = c       # sonst ist es Fliesstext neben der Nummer, kein Titel
        laengen: Counter = Counter()
        for (r, c), t in zellen.items():
            if c not in (lay.section_spalte, lay.titel_spalte):
                laengen[c] += len(t) if len(t) >= cfg["min_inhalt_zeichen"] else 0
        laengen = Counter({c: n for c, n in laengen.items() if n > 0})
        if laengen:
            ohne_komm = Counter({c: n for c, n in laengen.items() if not ist_kommentar_kopf(c)})
            kandidat_spalten = ohne_komm or laengen
            lay.inhalt_spalte = kandidat_spalten.most_common(1)[0][0]
            rest = [c for c in kandidat_spalten if c != lay.inhalt_spalte]
            if rest and max(kandidat_spalten[c] for c in rest) >= cfg["inhalt_zweifel_anteil"] * kandidat_spalten[lay.inhalt_spalte]:
                zweifel.append(f"Mehrere Spalten enthalten viel Text ({sp(lay.inhalt_spalte)} und "
                               f"{sp(max(rest, key=lambda c: kandidat_spalten[c]))}); nicht sicher, welche der Anforderungstext ist.")
            else:
                links = min([c for c in kandidat_spalten if c > max(lay.section_spalte or 0, lay.titel_spalte or 0)], default=None)
                if links is not None and links != lay.inhalt_spalte:
                    zweifel.append(f"Die Spalte mit dem meisten Text ({sp(lay.inhalt_spalte)}) steht nicht direkt neben der "
                                   f"Section ({sp(links)} hat auch Text); es könnte ein Kommentar statt der Anforderung sein.")
        if lay.kopf_zeilen and lay.inhalt_spalte:
            kopfset = set(lay.kopf_zeilen)
            mit = sum(1 for r in lay.kopf_zeilen if zellen.get((r, lay.inhalt_spalte)))
            ohne = sum(1 for (r, c) in zellen if c == lay.inhalt_spalte and r > min(kopfset) and r not in kopfset)
            lay.modell = "zeile_pro_section" if mit / len(lay.kopf_zeilen) >= cfg["modell_schwelle"] else "zeile_pro_absatz"
            if mit and ohne and min(mit, ohne) / (mit + ohne) >= cfg["gemischt_anteil"]:
                zweifel.append("Der Text steht teils in der Zeile der Section, teils in eigenen Zeilen darunter.")
        melden = lay.section_spalte is not None or len(zellen) >= cfg["warn_ab_zellen"]
        if lay.section_spalte is None and zellen:
            schwer.append("Es wurden keine Gliederungsnummern (1, 2.1 ...) gefunden.")
        # Section je Zeile
        aktuell: str | None = None
        kopfset = set(lay.kopf_zeilen)
        for r in range(1, lay.max_zeile + 1):
            if r in kopfset:
                t = zellen[(r, lay.section_spalte)]  # type: ignore[index]
                aktuell = section_nummer(t) or (kopf_zerlegen(t) or (None, ""))[0]
            self.zeilen_section[(blatt, r)] = aktuell
        self._zahl_sections(blatt, lay)
        if lay.section_spalte is None:
            lay.sicherheit = NIEDRIG_S           # ohne erkennbare Gliederung nie automatisch
        elif schwer:
            lay.sicherheit = NIEDRIG_S
        elif zweifel:
            lay.sicherheit = MITTEL_S
        lay.sicherheit_gruende = schwer + zweifel
        for g in lay.sicherheit_gruende if melden else ():
            self.warnungen.append(f"Aufbau von Blatt '{blatt}' unsicher: {g} Treffer in diesem Blatt werden nur als Vorschlag zum Prüfen angeboten.")

    def _layout_aus_profil(self, blatt: str, bp: "BlattProfil", cfg: dict, zellen: dict[tuple[int, int], str]) -> None:
        """Aufbau eines Blatts so, wie ihn das Profil festlegt (statt _layout_erkennen). Die Sicherheit ist "hoch",
        weil die Angabe vom Menschen kommt; Unstimmigkeiten mit den Daten werden als Warnung gemeldet."""
        from .layouts import ProfilFehler        # erst hier: layouts.py importiert dieses Modul
        lay = self.layouts[blatt]
        lay.profil = self.profil.name           # type: ignore[union-attr]
        sp = get_column_letter
        nr = list(self.layouts).index(blatt) + 1          # Blattnummer fuer Meldungen (Blattnamen kommen nie in Meldungen)
        if bp.ignorieren:
            lay.ignoriert = True
            for r in range(1, lay.max_zeile + 1):
                self.zeilen_section[(blatt, r)] = None
            return
        genannt = {"Section": bp.section_spalte, "Titel": bp.titel_spalte, "Inhalt": bp.inhalt_spalte,
                   "Sub-Section": getattr(bp, "untergliederung_spalte", None), "Kennwort": getattr(bp, "kennwort_spalte", None)}
        if zellen:
            for art, c in genannt.items():
                if c and c > lay.max_spalte:
                    raise ProfilFehler(f"Im Profil '{self.profil.name}' steht für ein Blatt die {art}-Spalte {sp(c)}, "   # type: ignore[union-attr]
                                       f"aber das Blatt (Blatt {nr}) reicht nur bis Spalte {sp(lay.max_spalte)}. "
                                       "Bitte die Spalte im Profil korrigieren.")
        else:
            self.warnungen.append(f"Blatt {nr} ist laut Profil festgelegt, enthält aber keinen Text.")
        lay.section_spalte, lay.titel_spalte, lay.inhalt_spalte = bp.section_spalte, bp.titel_spalte, bp.inhalt_spalte
        lay.modell = bp.modell
        if bp.modell == "register":
            self._register_aus_profil(blatt, bp, lay, cfg, zellen)
            return
        treffer: list[tuple[int, str, str]] = []
        if bp.section_spalte:
            for (r, c), t in sorted(zellen.items()):
                if c == bp.section_spalte and r > bp.kopfzeile:
                    k = self._kopf_art(t, cfg)
                    if k:
                        treffer.append((r, k[0], k[1]))
            kombi = sum(1 for _, art, _ in treffer if art == "kombi")
            lay.kombiniert = (not bp.titel_spalte) and kombi > len(treffer) - kombi
            lay.kopf_zeilen = [r for r, _, _ in treffer]
            if not treffer and zellen:
                self.warnungen.append(f"Laut Profil steht die Section in Spalte {sp(bp.section_spalte)}, dort wurde aber keine "
                                      f"Gliederungsnummer gefunden (Blatt {nr}).")
        if bp.inhalt_spalte and zellen and not any(c == bp.inhalt_spalte for (_, c) in zellen):
            self.warnungen.append(f"Laut Profil steht der Inhalt in Spalte {sp(bp.inhalt_spalte)}, dort steht aber kein Text "
                                  f"(Blatt {nr}).")
        aktuell: str | None = None
        kopfset = set(lay.kopf_zeilen)
        for r in range(1, lay.max_zeile + 1):
            if r in kopfset:
                t = zellen[(r, bp.section_spalte)]  # type: ignore[index]
                aktuell = section_nummer(t) or (kopf_zerlegen(t) or (None, ""))[0]
            self.zeilen_section[(blatt, r)] = aktuell
        self._zahl_sections(blatt, lay)
        lay.sicherheit, lay.sicherheit_gruende = HOCH_S, []

    def _register_aus_profil(self, blatt: str, bp: "BlattProfil", lay: BlattLayout, cfg: dict,
                             zellen: dict[tuple[int, int], str]) -> None:
        """Profil mit `modell: register`: Rollen aus dem Profil; erste Datenzeile = `daten_ab` oder die erste Zeile unter der
        Kopfzeile mit Gliederungswert (Beispielzeilen haben Platzhalter); Kennwort = haeufigster Wert der Kennwort-Spalte."""
        rc = {**_register_cfg(cfg), "aktiv": True}
        lay.register, lay.modell, lay.titel_spalte = True, "zeile_pro_absatz", None
        lay.untergliederung_spalte, lay.kennwort_spalte, lay.kopfzeile = bp.untergliederung_spalte, bp.kennwort_spalte, bp.kopfzeile
        sec_zeilen = sorted(r for (r, c), t in zellen.items() if c == bp.section_spalte and r > bp.kopfzeile
                            and section_kanon(t, rc) is not None)
        lay.daten_ab = bp.daten_ab or (sec_zeilen[0] if sec_zeilen else bp.kopfzeile + 1)
        if not sec_zeilen and zellen:
            self.warnungen.append(f"Laut Profil steht die Section in Spalte {get_column_letter(bp.section_spalte or 1)}, dort wurde "
                                  f"aber keine Gliederungsnummer gefunden (Blatt {list(self.layouts).index(blatt) + 1}).")
        if bp.kennwort_spalte:
            werte = Counter(t.strip() for (r, c), t in zellen.items() if c == bp.kennwort_spalte and r >= lay.daten_ab)
            lay.kennwort = werte.most_common(1)[0][0] if werte else ""
        lay.spec_spalte = self._spec_spalte(zellen, lay.daten_ab, {bp.section_spalte, bp.inhalt_spalte, bp.untergliederung_spalte,
                                                                    bp.kennwort_spalte}, bp.section_spalte or 0, rc)
        self._rc = rc
        self._register_abschluss(blatt, lay, zellen, rc, cfg)

    def _zahl_sections(self, blatt: str, lay: BlattLayout) -> None:
        """Section-Nummern, die als Zahl in der Excel stehen: gleich lautende sind nicht unterscheidbar (2.10 wird 2.1)."""
        if not lay.section_spalte:
            return
        nummern = [(r, self.zellen[(blatt, r, lay.section_spalte)]) for r in lay.kopf_zeilen
                   if (blatt, r, lay.section_spalte) in self.zahl_zellen]
        if not nummern:
            return
        if any("." in t for _, t in nummern):
            self.warnungen.append(f"Blatt '{blatt}': Section-Nummern stehen als Zahlen in der Excel (z. B. Zeile {nummern[0][0]}). "
                                  "Bei Zahlen fällt eine Null am Ende weg (2.10 wird zu 2.1) und die Zellen bleiben Zahlen.")
        zaehler = Counter(t for _, t in nummern)
        doppelt = {t for t, n in zaehler.items() if n > 1 and "." in t}
        if doppelt:
            lay.unsichere_sections |= doppelt
            self.warnungen.append(f"Blatt '{blatt}': Die Section-Nummer(n) {', '.join(sorted(doppelt))} kommen als Zahl mehrfach vor "
                                  "(vermutlich 2.1 und 2.10). Die Zuordnung über diese Section ist unsicher.")

    # ------------------------------------------------------------ Kandidaten
    def _tokens(self, blatt: str, zeile: int, spalte: int, text: str, ab: int, bis: int) -> tuple[list[str], array]:
        """Normformen und Positionen (Zeile, Spalte, Start, Ende je Wort) eines Textstuecks."""
        norms: list[str] = []
        pos = array("i")
        for w in woerter(text[ab:bis]):
            n = self._norm_cache.get(w.text)                    # gleiche Woerter teilen sich Normform (spart Zeit/Speicher)
            if n is None:
                n = self._norm_cache[w.text] = w.norm
            if n:
                norms.append(n)
                pos.extend((zeile, spalte, ab + w.start, ab + w.end))
        return norms, pos

    def _versteckt_art(self, blatt: str, zeile: int, spalte: int) -> str:
        if self.layouts[blatt].versteckt:
            return "blatt"
        if zeile in self.versteckte_zeilen.get(blatt, ()):
            return "zeile"
        if spalte in self.versteckte_spalten.get(blatt, ()):
            return "spalte"
        return ""

    def _kandidaten_bauen(self) -> None:
        def add(art, blatt, zeile, spalte, toks, text, start, ende, zt):
            norms, pos = toks
            if not norms:
                return
            sec = self.zeilen_section.get((blatt, zeile))
            lay = self.layouts[blatt]
            k = Kandidat(id=len(self.kandidaten), art=art, blatt=blatt, zeile=zeile, spalte=spalte,
                         norms=norms, pos=pos, zellen=self.zellen,
                         text=text, section=sec, start=start, end=ende, zelle_text=zt,
                         zelle_ist_zahl=(blatt, zeile, spalte) in self.zahl_zellen,
                         versteckt=self._versteckt_art(blatt, zeile, spalte),
                         section_unsicher=bool(sec and sec in lay.unsichere_sections),
                         kopf_spalte=(spalte in (lay.section_spalte, lay.titel_spalte) and spalte != lay.inhalt_spalte))
            self.kandidaten.append(k)
            if k.section:
                self.nach_section.setdefault(k.section, []).append(k.id)
                self.nach_blatt_section.setdefault((blatt, k.section), []).append(k.id)

        for (blatt, zeile, spalte), text in sorted(self.zellen.items()):
            lay = self.layouts[blatt]
            if lay.ignoriert:
                continue                                  # Profil: dieses Blatt wird nie abgeglichen
            if lay.register and (spalte != lay.inhalt_spalte or zeile < lay.daten_ab):
                continue                                  # Register: nur die Textspalte der Datenzeilen (ein Absatz je Zeile)
            absaetze = [(m.start(), m.end()) for m in re.finditer(r"[^\n]+", text) if m.group().strip()]
            for s, e in absaetze:
                add("absatz", blatt, zeile, spalte, self._tokens(blatt, zeile, spalte, text, s, e), text[s:e], s, e, text)
            if len(absaetze) >= 2:
                add("zelle", blatt, zeile, spalte, self._tokens(blatt, zeile, spalte, text, 0, len(text)), text, 0, len(text), text)
        for blatt, lay in self.layouts.items():
            if lay.section_spalte and not lay.kombiniert and lay.titel_spalte:
                for r in lay.kopf_zeilen:
                    a = self.zellen.get((blatt, r, lay.section_spalte))
                    b = self.zellen.get((blatt, r, lay.titel_spalte))
                    if a and b:
                        (n1, p1), (n2, p2) = (self._tokens(blatt, r, lay.section_spalte, a, 0, len(a)),
                                              self._tokens(blatt, r, lay.titel_spalte, b, 0, len(b)))
                        p1.extend(p2)
                        toks = (n1 + n2, p1)
                        add("kopf", blatt, r, lay.section_spalte, toks, f"{a} {b}", 0, len(a), a)

    # ------------------------------------------------------------ Abfragen
    def text(self, blatt: str, zeile: int, spalte: int) -> str:
        return self.zellen.get((blatt, zeile, spalte), "")

    def section_zeilen(self, blatt: str, kopf_zeile: int) -> range:
        """Zeilen, die zu einer Section gehoeren: Kopfzeile bis vor die naechste Kopfzeile."""
        lay = self.layouts[blatt]
        if lay.register:                 # Register: alle Zeilen mit demselben Gliederungswert ab der Ueberschrift
            sec = self.zeilen_section.get((blatt, kopf_zeile))
            ende = kopf_zeile + 1
            while ende <= lay.max_zeile and self.zeilen_section.get((blatt, ende)) == sec:
                ende += 1
            return range(kopf_zeile, ende)
        spaeter = [r for r in lay.kopf_zeilen if r > kopf_zeile]
        return range(kopf_zeile, (spaeter[0] if spaeter else lay.max_zeile + 1))

    def ist_zahl(self, blatt: str, zeile: int, spalte: int) -> bool:
        """True, wenn die Zelle in der Excel eine Zahl ist (kein Text)."""
        return (blatt, zeile, spalte) in self.zahl_zellen

    def beschreibung(self) -> list[dict[str, Any]]:
        return [l.beschreibung() for l in self.layouts.values()]

    # ------------------------------------------------------------ Risiken beim Einfuegen/Loeschen (fuer Welle 3)
    def _bezuege(self, blatt: str) -> list[tuple[str, int]]:
        """(Blatt, Zeile) aller Zellbezuege in Formeln des Blatts."""
        if blatt not in self._formel_bezuege:
            out: list[tuple[str, int]] = []
            for _, _, text in self._eig[blatt].formeln:
                out += self._text_bezuege(text, blatt)
            self._formel_bezuege[blatt] = out
        return self._formel_bezuege[blatt]

    @staticmethod
    def _text_bezuege(text: str, eigenes: str) -> list[tuple[str, int]]:
        out = []
        for m in _BEZUG.finditer(text):
            name = ((m.group(1) or "").replace("''", "'") or m.group(2) or eigenes).casefold()
            out.append((name, int(m.group(3))))
            if m.group(4):
                out.append((name, int(m.group(4))))
        return out

    def einfuege_risiken(self, blatt: str, zeile: int) -> list[str]:
        """Einfache deutsche Saetze: was koennte brechen, wenn an `zeile` eine Zeile eingefuegt oder geloescht wird?

        Leere Liste = unbedenklich. Aendert nichts an der Excel und am Verhalten des Tools.
        Das ist die vollstaendige (strenge) Liste, auch mit Dingen, die der Applier beim Einfuegen selbst mitzieht;
        was beim automatischen Einfuegen wirklich im Weg steht, sagt `einfuege_hindernisse`."""
        return self._einfuege_risiken(blatt, zeile, entschaerft=False)

    def einfuege_hindernisse(self, blatt: str, zeile: int, schreiber: str | None = None) -> list[str]:
        """Wie `einfuege_risiken`, aber leer, wenn beim EINFUEGEN einer Zeile an `zeile` nichts wirklich brechen kann.

        Kein Hindernis, weil der Applier es beim Einfuegen selbst mitzieht (applier._verschiebe_verbunden / _verschiebe_hoehen):
          * verbundene Zellen, die ganz unterhalb der Einfuegestelle beginnen (sie wandern mit);
          * ausgeblendete oder gruppierte Zeilen darunter (Ausblenden, Gliederungsebene und Hoehe wandern mit).
        Weiter ein Hindernis: verbundene Zellen DURCH die Einfuegestelle, ausgeblendete/gruppierte Zeilen, wenn die neue
        Zeile MITTEN in einer solchen Zeilenfolge landen wuerde (die Zeile davor und die Zeile dahinter gehoeren dazu), und alles,
        was `einfuege_risiken` sonst nennt (Formeln, bedingte Formatierung, Datenpruefung, Filter, Tabellen, benannte Bereiche,
        Hyperlinks, Bilder). Gibt es ein Hindernis, stehen die beiden Hinweise auf Unbedenkliches mit in der Liste
        (der Mensch soll alles sehen). Fuer das Loeschen von Zeilen gilt weiter `einfuege_risiken`.

        Patch-Schreiber (ADR 0007, `anwenden.schreiber: patch`, Standard; `schreiber` ueberschreibt die Einstellung): der
        Schreiber zieht Formeln, bedingte Formatierung, Datenpruefung, Filter, Druckbereich, Namen, verbundene Zellen darunter,
        Hyperlinks, Seitenumbrueche und Zeichnungs-Anker selbst mit. Hindernis ist dann nur, was src/xlsx_patch.py nicht sicher
        kann (`xlsx_patch.hindernisse`: geteilte Formel/Matrixformel ueber die Stelle, Tabellenobjekt, Pivot, INDIREKT/
        BEREICH.VERSCHIEBEN, 3D-Bezug, externe Verknuepfung, verbundene Zellen durch die Stelle, mitten in ausgeblendete
        Zeilen, Kommentare/Steuerelemente/Diagrammbezuege darunter)."""
        if schreiber is None:
            schreiber = str(config.lade("zuordnung").get("anwenden", {}).get("schreiber", "patch"))
        if schreiber != "openpyxl":
            from .xlsx_patch import hindernisse          # erst hier: das reine Lesen braucht das Schreib-Modul nicht
            return hindernisse(self.pfad, blatt, zeile)
        return self._einfuege_risiken(blatt, zeile, entschaerft=True)

    def _einfuege_risiken(self, blatt: str, zeile: int, entschaerft: bool) -> list[str]:
        eig = self._eig.get(blatt)
        if eig is None:
            return [f"Das Tabellenblatt '{blatt}' gibt es in der Excel nicht."]
        z = zeile
        risiken: list[str] = []
        harmlos: set[str] = set()         # nur im entschaerften Modus: Hinweise, die allein kein Hindernis sind
        if entschaerft:       # eine Verbindung, die erst ab der Einfuegezeile beginnt, wandert mit; nur durch die Stelle ist ein Hindernis
            durch = [b for b in eig.verbunden if b[0] < z <= b[1]]
        else:
            durch = [b for b in eig.verbunden if b[0] <= z <= b[1] and b[0] < b[1]]
        if durch:
            risiken.append("Hier laufen verbundene Zellen durch die Zeile; sie würden zerrissen oder verrutschen.")
        darunter = [b for b in eig.verbunden if (b[0] >= z if entschaerft else b[0] > z)]
        if darunter:
            risiken.append("Unterhalb sind verbundene Zellen; sie verschieben sich beim Einfügen oder Löschen nicht zuverlässig mit.")
            if entschaerft:
                harmlos.add(risiken[-1])
        eigen = blatt.casefold()
        unten = [f for f in eig.formeln if f[0] >= z]
        bezug = any(n == eigen and r >= z for anderes in self._eig for n, r in self._bezuege(anderes))
        if unten or bezug or eig.formeln_viele:
            risiken.append("Formeln werden beim Einfügen oder Löschen nicht angepasst; Bezüge auf Zeilen ab dieser Stelle "
                           "können danach auf falsche Zellen zeigen.")
        if any(b[1] >= z for b in eig.bedingt):
            risiken.append("Eine bedingte Formatierung reicht bis an oder unter diese Stelle; ihr Bereich wird nicht mitverschoben.")
        if any(b[1] >= z for b in eig.pruefung):
            risiken.append("Eine Datenüberprüfung (z. B. Auswahlliste) reicht bis an oder unter diese Stelle; ihr Bereich wird nicht mitverschoben.")
        if eig.filter and eig.filter[1] >= z:
            risiken.append("Ein Autofilter reicht bis an oder unter diese Stelle; sein Bereich wird nicht angepasst.")
        for name, r0, r1 in eig.tabellen:
            if r1 >= z:
                risiken.append(f"Die Excel-Tabelle '{name}' reicht bis an oder unter diese Stelle; ihr Bereich wächst oder schrumpft nicht mit.")
        for name, ref in eig.namen:
            if any(n == eigen and r >= z for n, r in self._text_bezuege(ref, blatt)):
                risiken.append(f"Der benannte Bereich '{name}' zeigt auf Zeilen ab dieser Stelle; er wird nicht angepasst.")
        if any(r >= z for r in eig.links):
            risiken.append("Unterhalb gibt es Hyperlinks; sie werden nicht mitverschoben.")
        gruppiert = set(eig.gruppiert)
        if entschaerft and z - 1 in gruppiert and z in gruppiert:
            risiken.append("Die neue Zeile käme mitten in eine Folge ausgeblendeter oder gruppierter Zeilen; "
                           "die Folge würde unterbrochen.")
        elif any(r >= z - 1 for r in gruppiert):
            risiken.append("In der Nähe oder darunter gibt es ausgeblendete oder gruppierte Zeilen; diese Einstellung wird nicht mitverschoben.")
            if entschaerft:
                harmlos.add(risiken[-1])
        if any(r >= z for r in eig.anker):
            risiken.append("Unterhalb sind Bilder oder Diagramme verankert; sie bleiben beim Einfügen oder Löschen an ihrer Stelle.")
        if entschaerft and all(r in harmlos for r in risiken):
            return []
        return risiken
