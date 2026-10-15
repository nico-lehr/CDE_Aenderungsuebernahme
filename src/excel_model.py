"""Liest eine bestehende CDE-Excel und beschreibt sie, ohne ein festes Format vorauszusetzen.

Die Excel ist "chronologisch nach Sections aufgebaut, dahinter der Inhalt".
Wie genau, wird je Tabellenblatt aus den Daten erkannt:

  * Section-Spalte  : Spalte mit den meisten Gliederungsnummern (1, 2.1, 3.2.4 ...)
  * Titel-Spalte    : Spalte mit der Ueberschrift neben der Nummer (falls getrennt)
  * Inhalt-Spalte   : Spalte mit dem meisten Text
  * Modell          : "zeile_pro_section"  (Inhalt steht in der Zeile der Section)
                      "zeile_pro_absatz"   (jeder Absatz hat eine eigene Zeile)

Aus allen Textzellen werden "Kandidaten" gebildet (Absatz, ganze Zelle, Kopf aus
Nummer+Titel). Gegen diese Kandidaten werden die PDF-Bloecke abgeglichen.
Reine Python-Logik; schreibt nie in die Datei.
"""
from __future__ import annotations

import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText
from openpyxl.utils import get_column_letter

from . import config
from .textutil import kopf_zerlegen, norm, section_nummer, woerter


@dataclass
class XTok:
    text: str
    norm: str
    blatt: str
    zeile: int
    spalte: int
    start: int   # Zeichenposition in der Zelle
    end: int


@dataclass
class Kandidat:
    id: int
    art: str                 # absatz | zelle | kopf
    blatt: str
    zeile: int
    spalte: int
    tokens: list[XTok]
    text: str
    section: str | None
    start: int = 0           # Span in der Zelle (absatz/zelle)
    end: int = 0
    zelle_text: str = ""

    @property
    def ziel(self) -> str:
        return f"{self.blatt}!{get_column_letter(self.spalte)}{self.zeile}"

    @property
    def norms(self) -> list[str]:
        return [t.norm for t in self.tokens]


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

    def beschreibung(self) -> dict[str, Any]:
        b = lambda c: get_column_letter(c) if c else None  # noqa: E731
        return {"blatt": self.name, "section_spalte": b(self.section_spalte), "titel_spalte": b(self.titel_spalte),
                "inhalt_spalte": b(self.inhalt_spalte), "modell": self.modell,
                "sections": len(self.kopf_zeilen), "zeilen": self.max_zeile}


class ExcelModell:
    def __init__(self, pfad: str | Path) -> None:
        self.pfad = Path(pfad)
        self.warnungen: list[str] = []
        self.layouts: dict[str, BlattLayout] = {}
        self.zellen: dict[tuple[str, int, int], str] = {}
        self.zeilen_section: dict[tuple[str, int], str | None] = {}
        self.kandidaten: list[Kandidat] = []
        self.nach_section: dict[str, list[int]] = {}
        self._lesen()

    # ------------------------------------------------------------ Lesen
    def _lesen(self) -> None:
        cfg = config.lade("zuordnung")["excel"]
        suffix = self.pfad.suffix.lower()
        if suffix not in (".xlsx", ".xlsm"):
            raise ValueError("Nur .xlsx / .xlsm werden unterstützt. Bitte die Datei in Excel als .xlsx speichern.")
        self._pruefe_paket()
        wb = load_workbook(str(self.pfad), rich_text=True, keep_vba=(suffix == ".xlsm"))
        n_formeln = n_rich = n_zahl_sec = 0
        for ws in wb.worksheets:
            lay = BlattLayout(name=ws.title, max_zeile=ws.max_row, max_spalte=ws.max_column)
            self.layouts[ws.title] = lay
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if v is None:
                        continue
                    if isinstance(v, CellRichText):
                        n_rich += 1
                        v = str(v)
                    if isinstance(v, str):
                        if v.startswith("="):
                            n_formeln += 1
                            continue
                        if v.strip():
                            self.zellen[(ws.title, c.row, c.column)] = v
                    elif isinstance(v, (int, float)) and not isinstance(v, bool):
                        self.zellen[(ws.title, c.row, c.column)] = ("%g" % v) if isinstance(v, float) else str(v)
            if ws.merged_cells.ranges:
                self.warnungen.append(f"Blatt '{ws.title}' enthält verbundene Zellen (Zeilen einfügen/löschen verschiebt sie nur teilweise).")
        if n_formeln:
            self.warnungen.append(f"{n_formeln} Formelzellen werden nicht verändert; Zeilen einfügen/löschen passt Formeln nicht an.")
        if n_rich:
            self.warnungen.append(f"{n_rich} Zellen mit gemischter Formatierung (z. B. teilweise fett). Geänderte Zellen verlieren diese Teilformatierung.")
        for name in self.layouts:
            self._layout_erkennen(name, cfg)
        self._kandidaten_bauen()

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
    def _layout_erkennen(self, blatt: str, cfg: dict) -> None:
        lay = self.layouts[blatt]
        zellen = {(r, c): t for (b, r, c), t in self.zellen.items() if b == blatt}
        zaehler: Counter = Counter()
        for (r, c), t in zellen.items():
            if section_nummer(t) is not None or kopf_zerlegen(t) is not None:
                zaehler[c] += 1
        if zaehler:
            spalte, anzahl = zaehler.most_common(1)[0]
            if anzahl >= cfg["section_spalte_min"]:
                lay.section_spalte = spalte
        if lay.section_spalte:
            kombi = sum(1 for (r, c), t in zellen.items() if c == lay.section_spalte and kopf_zerlegen(t))
            nur_nr = sum(1 for (r, c), t in zellen.items() if c == lay.section_spalte and section_nummer(t))
            lay.kombiniert = kombi > nur_nr
            for (r, c), t in sorted(zellen.items()):
                if c == lay.section_spalte and (section_nummer(t) or kopf_zerlegen(t)):
                    lay.kopf_zeilen.append(r)
            if not lay.kombiniert:
                titel = Counter()
                for r in lay.kopf_zeilen:
                    for c in range(lay.section_spalte + 1, lay.max_spalte + 1):
                        t = zellen.get((r, c))
                        if t and len(t) < 160 and "\n" not in t:
                            titel[c] += 1
                            break
                if titel:
                    lay.titel_spalte = titel.most_common(1)[0][0]
        laengen: Counter = Counter()
        for (r, c), t in zellen.items():
            if c not in (lay.section_spalte, lay.titel_spalte):
                laengen[c] += len(t) if len(t) >= cfg["min_inhalt_zeichen"] else 0
        if laengen:
            lay.inhalt_spalte = laengen.most_common(1)[0][0]
        if lay.kopf_zeilen and lay.inhalt_spalte:
            mit = sum(1 for r in lay.kopf_zeilen if zellen.get((r, lay.inhalt_spalte)))
            lay.modell = "zeile_pro_section" if mit / len(lay.kopf_zeilen) >= 0.6 else "zeile_pro_absatz"
        # Section je Zeile
        aktuell: str | None = None
        kopfset = set(lay.kopf_zeilen)
        for r in range(1, lay.max_zeile + 1):
            if r in kopfset:
                t = zellen[(r, lay.section_spalte)]  # type: ignore[index]
                aktuell = section_nummer(t) or (kopf_zerlegen(t) or (None, ""))[0]
            self.zeilen_section[(blatt, r)] = aktuell

    # ------------------------------------------------------------ Kandidaten
    def _tokens(self, blatt: str, zeile: int, spalte: int, text: str, ab: int, bis: int) -> list[XTok]:
        return [XTok(w.text, w.norm, blatt, zeile, spalte, ab + w.start, ab + w.end)
                for w in woerter(text[ab:bis]) if w.norm]

    def _kandidaten_bauen(self) -> None:
        def add(art, blatt, zeile, spalte, toks, text, start, ende, zt):
            if not toks:
                return
            k = Kandidat(id=len(self.kandidaten), art=art, blatt=blatt, zeile=zeile, spalte=spalte, tokens=toks,
                         text=text, section=self.zeilen_section.get((blatt, zeile)), start=start, end=ende, zelle_text=zt)
            self.kandidaten.append(k)
            if k.section:
                self.nach_section.setdefault(k.section, []).append(k.id)

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
                        toks = (self._tokens(blatt, r, lay.section_spalte, a, 0, len(a))
                                + self._tokens(blatt, r, lay.titel_spalte, b, 0, len(b)))
                        add("kopf", blatt, r, lay.section_spalte, toks, f"{a} {b}", 0, len(a), a)

    # ------------------------------------------------------------ Abfragen
    def text(self, blatt: str, zeile: int, spalte: int) -> str:
        return self.zellen.get((blatt, zeile, spalte), "")

    def section_zeilen(self, blatt: str, kopf_zeile: int) -> range:
        """Zeilen, die zu einer Section gehoeren: Kopfzeile bis vor die naechste Kopfzeile."""
        lay = self.layouts[blatt]
        spaeter = [r for r in lay.kopf_zeilen if r > kopf_zeile]
        return range(kopf_zeile, (spaeter[0] if spaeter else lay.max_zeile + 1))

    def beschreibung(self) -> list[dict[str, Any]]:
        return [l.beschreibung() for l in self.layouts.values()]
