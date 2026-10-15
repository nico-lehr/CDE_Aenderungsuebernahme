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

Aus allen Textzellen werden "Kandidaten" gebildet (Absatz, ganze Zelle, Kopf aus
Nummer+Titel). Gegen diese Kandidaten werden die PDF-Bloecke abgeglichen.
Reine Python-Logik; schreibt nie in die Datei.
"""
from __future__ import annotations

import hashlib
import re
import weakref
import zipfile
from collections import Counter
from array import array
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText
from openpyxl.utils import coordinate_to_tuple, get_column_letter, range_boundaries
from openpyxl.worksheet.formula import ArrayFormula

from . import config
from .textutil import kopf_zerlegen, norm, section_nummer, woerter

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
                 "zelle_ist_zahl", "versteckt", "section_unsicher", "_tokens", "_norms", "_pos", "_zellen")

    def __init__(self, id: int, art: str, blatt: str, zeile: int, spalte: int, tokens: list[XTok] | None = None,
                 text: str = "", section: str | None = None, start: int = 0, end: int = 0, zelle_text: str = "",
                 zelle_ist_zahl: bool = False, versteckt: str = "", section_unsicher: bool = False, *,
                 norms: list[str] | None = None, pos: array | None = None, zellen: dict | None = None) -> None:
        self.id, self.art, self.blatt, self.zeile, self.spalte = id, art, blatt, zeile, spalte
        self.text, self.section, self.start, self.end, self.zelle_text = text, section, start, end, zelle_text
        self.zelle_ist_zahl = zelle_ist_zahl      # die Excel-Zelle ist eine Zahl (kein Text); beim Ersetzen Zahlenformat beachten
        self.versteckt = versteckt                # "" | "blatt" | "zeile" | "spalte": Stelle ist ausgeblendet
        self.section_unsicher = section_unsicher  # Section-Nummer stammt aus mehrdeutigen Zahlen (z. B. 2.1 und 2.10)
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

    def beschreibung(self) -> dict[str, Any]:
        b = lambda c: get_column_letter(c) if c else None  # noqa: E731
        return {"blatt": self.name, "section_spalte": b(self.section_spalte), "titel_spalte": b(self.titel_spalte),
                "inhalt_spalte": b(self.inhalt_spalte), "modell": self.modell,
                "sections": len(self.kopf_zeilen), "zeilen": self.max_zeile,
                "sicherheit": self.sicherheit, "sicherheit_gruende": list(self.sicherheit_gruende),
                "versteckt": self.versteckt}


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
_MODELLE: "weakref.WeakValueDictionary[tuple[str, str, str], ExcelModell]" = weakref.WeakValueDictionary()


def _datei_hash(pfad: Path) -> str:
    """SHA-1 stueckweise berechnen (grosse Dateien nie komplett in den Speicher laden)."""
    h = hashlib.sha1()
    with open(pfad, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


class ExcelModell:
    def __new__(cls, pfad: str | Path) -> "ExcelModell":
        try:
            p = Path(pfad)
            schluessel = (str(p.resolve()), _datei_hash(p),
                          repr(sorted(config.lade("zuordnung")["excel"].items())))
        except (OSError, TypeError, KeyError):
            schluessel = None
        if schluessel is not None and schluessel in _MODELLE:
            return _MODELLE[schluessel]
        self = super().__new__(cls)
        self._schluessel = schluessel
        self._fertig = False
        return self

    def __init__(self, pfad: str | Path) -> None:
        if self._fertig:                      # aus dem Zwischenspeicher: schon eingelesen
            return
        self.pfad = Path(pfad)
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
        self._lesen()
        self._fertig = True
        if self._schluessel is not None:
            _MODELLE[self._schluessel] = self

    # ------------------------------------------------------------ Lesen
    def _lesen(self) -> None:
        cfg = config.lade("zuordnung")["excel"]
        suffix = self.pfad.suffix.lower()
        if suffix not in (".xlsx", ".xlsm"):
            raise ValueError("Nur .xlsx / .xlsm werden unterstützt. Bitte die Datei in Excel als .xlsx speichern.")
        self._pruefe_paket()
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
                if isinstance(v, CellRichText):
                    n_rich += 1
                    v = str(v)
                if isinstance(v, str):
                    if v.startswith("="):
                        n_formeln += 1
                        self._formel_merken(eig, c.row, c.column, v)
                        continue
                    if v.strip():
                        self.zellen[(ws.title, c.row, c.column)] = v
                        belegt.append((c.row, c.column))
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    self.zellen[(ws.title, c.row, c.column)] = zahl_als_text(v, c.number_format)
                    self.zahl_zellen.add((ws.title, c.row, c.column))
                    belegt.append((c.row, c.column))
            lay.max_zeile, lay.max_spalte = max_z, max_s
            self._eigenschaften_lesen(ws, eig)
            self._versteckt_lesen(ws, belegt)
            if ws.merged_cells.ranges:
                self.warnungen.append(f"Blatt '{ws.title}' enthält verbundene Zellen (Zeilen einfügen/löschen verschiebt sie nur teilweise).")
        if n_formeln:
            self.warnungen.append(f"{n_formeln} Formelzellen werden nicht verändert; Zeilen einfügen/löschen passt Formeln nicht an.")
        if n_rich:
            self.warnungen.append(f"{n_rich} Zellen mit gemischter Formatierung (z. B. teilweise fett). Geänderte Zellen verlieren diese Teilformatierung.")
        je_blatt: dict[str, dict[tuple[int, int], str]] = {name: {} for name in self.layouts}
        for (b, r, c), t in self.zellen.items():
            je_blatt[b][(r, c)] = t
        for name in self.layouts:
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

    def _pruefe_paket(self) -> None:
        try:
            with zipfile.ZipFile(self.pfad) as z:
                namen = z.namelist()
        except zipfile.BadZipFile as e:
            raise ValueError("Die Excel-Datei ist beschädigt oder keine .xlsx-Datei.") from e
        if any(n.startswith(("xl/drawings/", "xl/charts/", "xl/media/")) for n in namen):
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

    def _layout_erkennen(self, blatt: str, cfg: dict, zellen: dict[tuple[int, int], str]) -> None:
        lay = self.layouts[blatt]
        if not zellen:
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
            k = Kandidat(id=len(self.kandidaten), art=art, blatt=blatt, zeile=zeile, spalte=spalte,
                         norms=norms, pos=pos, zellen=self.zellen,
                         text=text, section=sec, start=start, end=ende, zelle_text=zt,
                         zelle_ist_zahl=(blatt, zeile, spalte) in self.zahl_zellen,
                         versteckt=self._versteckt_art(blatt, zeile, spalte),
                         section_unsicher=bool(sec and sec in self.layouts[blatt].unsichere_sections))
            self.kandidaten.append(k)
            if k.section:
                self.nach_section.setdefault(k.section, []).append(k.id)
                self.nach_blatt_section.setdefault((blatt, k.section), []).append(k.id)

        for (blatt, zeile, spalte), text in sorted(self.zellen.items()):
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

        Leere Liste = unbedenklich. Aendert nichts an der Excel und am Verhalten des Tools."""
        eig = self._eig.get(blatt)
        if eig is None:
            return [f"Das Tabellenblatt '{blatt}' gibt es in der Excel nicht."]
        z = zeile
        risiken: list[str] = []
        durch = [b for b in eig.verbunden if b[0] <= z <= b[1] and b[0] < b[1]]
        if durch:
            risiken.append("Hier laufen verbundene Zellen durch die Zeile; sie würden zerrissen oder verrutschen.")
        if any(b[0] > z for b in eig.verbunden):
            risiken.append("Unterhalb sind verbundene Zellen; sie verschieben sich beim Einfügen oder Löschen nicht zuverlässig mit.")
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
        if any(r >= z - 1 for r in eig.gruppiert):
            risiken.append("In der Nähe oder darunter gibt es ausgeblendete oder gruppierte Zeilen; diese Einstellung wird nicht mitverschoben.")
        if any(r >= z for r in eig.anker):
            risiken.append("Unterhalb sind Bilder oder Diagramme verankert; sie bleiben beim Einfügen oder Löschen an ihrer Stelle.")
        return risiken
