"""Liest ein Aenderungs-PDF und erkennt, was neu ist und was entfaellt.

Signale (alle in config/farben.yaml einstellbar):
  * Textfarbe            (blau = neu, rot = entfaellt)
  * farbige Hinterlegung (gefuellte Rechtecke hinter dem Text)
  * Durchstreichung      (duenne waagerechte Linie in Zeichenmitte)
  * PDF-Kommentare       (Highlight / StrikeOut)

Ergebnis: Woerter (PdfToken) mit Status keep/del/add/mod, zu Bloecken
(Absatz / Ueberschrift) zusammengefasst. Reine Python-Logik, kein LLM, keine UI.

Massstab (Bilanz-Prinzip): Der Leser liest entweder richtig ODER meldet die Unsicherheit sichtbar
(PdfToken.hinweis, PdfBlock.hinweise, PdfAnalyse.luecken, PdfAnalyse.warnungen). Nie still falsch.
"""
from __future__ import annotations

import colorsys
import re
import statistics
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pdfplumber

from . import config
from .models import ADD, DEL, KEEP, MOD, PdfBlock, PdfToken, luecke
from .textutil import norm, section_nummer

_RE_CID = re.compile(r"\(cid:\d+\)")
_PRIVATE_ZEICHEN = re.compile("[\ue000-\uf8ff]")
_LIGATUR = re.compile("[\ufb00-\ufb06]")
_NUR_PRIVAT = re.compile("[\ue000-\uf8ff]+")

# Hinweis-Kennung am Wort -> Satz am Block (einfaches Deutsch, ohne Dokumentinhalt)
_HINWEIS_TEXT = {
    "ligatur": "Das PDF enthält Ligaturen (z. B. ﬁ). Der Text kann in der Excel anders geschrieben sein.",
    "nicht_lesbar": "Einzelne Zeichen konnten nicht gelesen werden.",
    "unbekannte_farbe": "Einzelne Wörter haben eine Farbe, die nicht in der Legende steht; sie gelten als unverändert.",
    "linie_unklar": "Eine Linie läuft quer durch den Text (Zierlinie?). Sie wurde nicht als Durchstreichung gewertet.",
    "link": "Hyperlink: die blaue Farbe wurde nicht als „neu“ gewertet.",
    "teilweise_markiert": "Ein Wort ist nur zum Teil markiert (Linie oder Fläche deckt nicht das ganze Wort). Bitte prüfen.",
    "design_blau": "Die Überschrift ist blau wie alle Überschriften des Dokuments. Das wurde als Gestaltung gewertet, nicht als „neu“.",
    "design_flaeche": "Eine Farbfläche hinter dem Text sieht nach Gestaltung aus (Tabellenkopf, Streifen). "
                      "Sie wurde nicht als Markierung gewertet.",
}


@dataclass
class PdfAnalyse:
    pfad: str
    seiten: int
    blocks: list[PdfBlock]
    ignoriert: list[dict[str, Any]] = field(default_factory=list)
    farben: list[dict[str, Any]] = field(default_factory=list)
    warnungen: list[str] = field(default_factory=list)
    seiten_masse: dict[int, tuple[float, float]] = field(default_factory=dict)
    luecken: list[dict[str, Any]] = field(default_factory=list)   # Schema 2: models.luecke(...)
    # Bereiche, in denen die Woerter einer Luecke gezaehlt werden: (Luecke, Seite, [Rechtecke]); nur intern
    regionen: list[tuple[dict[str, Any], int, list[tuple[float, float, float, float]]]] = field(
        default_factory=list, repr=False)

    @property
    def markierte_woerter(self) -> int:
        n = sum(1 for b in self.blocks for t in b.tokens if t.status != KEEP)
        return n + sum(i["woerter"] for i in self.ignoriert)


# ---------------------------------------------------------------- Farben
def _rgb(farbe: Any) -> tuple[float, float, float] | None:
    if farbe is None or isinstance(farbe, str):
        return None
    if isinstance(farbe, (int, float)):
        farbe = (farbe,)
    try:
        c = tuple(float(x) for x in farbe)
    except (TypeError, ValueError):
        return None
    if len(c) == 1:
        return (c[0], c[0], c[0])
    if len(c) == 3:
        return c  # type: ignore[return-value]
    if len(c) == 4:
        cc, m, y, k = c
        return ((1 - cc) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
    return None


def _hex(rgb: tuple[float, float, float]) -> str:
    return "#%02x%02x%02x" % tuple(max(0, min(255, round(v * 255))) for v in rgb)


def klasse(rgb: tuple[float, float, float] | None, cfg: dict, hintergrund: bool = False) -> str | None:
    """'blau' | 'rot' | None."""
    if rgb is None:
        return None
    r, g, b = (max(0.0, min(1.0, v)) for v in rgb)
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    grenzen = cfg["hintergrund"] if hintergrund else cfg["text"]
    if s < grenzen["min_saettigung"] or v < grenzen.get("min_helligkeit", 0.25):
        return None
    grad = h * 360.0
    lo, hi = cfg["farbton"]["blau"]
    if lo <= grad <= hi:
        return "blau"
    if grad <= cfg["farbton"]["rot_bis"] or grad >= cfg["farbton"]["rot_ab"]:
        return "rot"
    return None


def _bedeutung(farbklasse: str | None, cfg: dict) -> str | None:
    if farbklasse is None:
        return None
    wert = cfg["legende"].get(farbklasse)
    return wert if wert in ("neu", "entfaellt") else None


def _unbekannt_art(rgb: tuple[float, float, float], cfg: dict) -> str:
    """Wie auffaellig ist eine Farbe, die KEINER Legende zugeordnet ist?
    '' = unauffaellig (Schwarz, Weiss, dunkles Grau) | 'bunt' | 'pastell' | 'grau' (mittleres Grau)."""
    u = cfg["unbekannt"]
    if not u["aktiv"]:
        return ""
    r, g, b = (max(0.0, min(1.0, v)) for v in rgb)
    _, s, v = colorsys.rgb_to_hsv(r, g, b)
    if v < u["min_helligkeit"]:
        return ""
    if s >= u["min_saettigung"]:
        return "bunt"
    if s >= u["pastell_min_saettigung"] and v >= u["pastell_min_helligkeit"]:
        return "pastell"
    if s < u["pastell_min_saettigung"] and u["grau_min_helligkeit"] <= v <= u["grau_max_helligkeit"]:
        return "grau"
    return ""


# ---------------------------------------------------------------- Seitendaten
@dataclass
class _Z:
    """Ein Zeichen mit allem, was zur Bewertung gebraucht wird."""
    text: str
    x0: float
    x1: float
    top: float
    bottom: float
    size: float
    font: str
    status: str = KEEP
    leer: bool = False                 # Leerraum-Zeichen (wird nur fuer Geometrie gebraucht)
    unsichtbar: bool = False           # Render-Modus 3 (unsichtbare Textebene)
    weiss: bool = False
    hex: str = ""
    farbklasse: str | None = None
    farbe_bed: str | None = None       # Bedeutung der Textfarbe ("neu" | "entfaellt" | None)
    unb: str = ""                      # Art der nicht zugeordneten Textfarbe (siehe _unbekannt_art)
    unb_flaeche: str = ""              # Farbe einer nicht zugeordneten Hinterlegung / Hervorhebung
    unscharf: bool = False             # Status stammt nur aus Geometrie (Linie, Flaeche, Kommentar)
    hinweis: set = field(default_factory=set)


@dataclass
class _Info:
    """Zusatzangaben zu einem Wort (nur fuer die Luecken-Zaehlung)."""
    unb: bool = False
    farbe: str = ""
    unlesbar: bool = False


@dataclass
class _Zeile:
    tokens: list[PdfToken]
    infos: list[_Info]
    top: float
    bottom: float
    x0: float
    x1: float
    size: float
    fett: bool
    seite: int
    unsichtbar: bool = False


@dataclass
class _SeitenDaten:
    nr: int
    breite: float
    hoehe: float
    zeichen: list[_Z] = field(default_factory=list)
    bild_anteil: float = 0.0
    bilder: list[tuple[float, float, float, float]] = field(default_factory=list)
    tabellen: list[tuple[float, float, float, float]] = field(default_factory=list)
    zu_viele_linien: bool = False
    grafik_objekte: int = 0
    rect_farben: Counter = field(default_factory=Counter)


# ---------------------------------------------------------------- Seite lesen
def _render_modi(page: Any) -> list[int] | None:
    """Text-Render-Modus je Zeichen (3 = unsichtbar). pdfplumber gibt ihn nicht heraus, deshalb wird die Seite
    hier mit einem eigenen Geraet ausgewertet und das Ergebnis als Layout der Seite abgelegt.
    None, wenn das nicht moeglich ist (dann gelten alle Zeichen als sichtbar)."""
    try:
        from pdfminer.pdfinterp import PDFPageInterpreter
        from pdfplumber.page import PDFPageAggregatorWithMarkedContent
    except Exception:       # pragma: no cover - andere pdfplumber-Version
        return None

    class _Geraet(PDFPageAggregatorWithMarkedContent):
        def __init__(self, *a: Any, **kw: Any) -> None:
            super().__init__(*a, **kw)
            self.modi: list[int] = []
            self._modus = 0

        def render_string(self, textstate: Any, seq: Any, ncs: Any, graphicstate: Any) -> None:
            self._modus = int(getattr(textstate, "render", 0) or 0)
            super().render_string(textstate, seq, ncs, graphicstate)

        def render_char(self, *a: Any, **kw: Any) -> float:   # type: ignore[override]
            adv = super().render_char(*a, **kw)
            self.modi.append(self._modus)
            return adv

    try:
        if hasattr(page, "_layout"):
            return None
        geraet = _Geraet(page.pdf.rsrcmgr, pageno=page.page_number, laparams=page.pdf.laparams)
        PDFPageInterpreter(page.pdf.rsrcmgr, geraet).process_page(page.page_obj)
        page._layout = geraet.get_result()
        return geraet.modi
    except Exception:
        return None


def _zeichenfarbe(c: dict) -> tuple[tuple[float, float, float] | None, bool]:
    """(RGB oder None, unklar). 'unklar' = Farbe aus Muster/Sonderfarbraum, die nicht umgerechnet werden kann."""
    farbe = c.get("non_stroking_color")
    if farbe is None:
        return None, False
    if isinstance(farbe, str):
        return None, True
    try:
        werte = tuple(float(x) for x in farbe)
    except (TypeError, ValueError):
        return None, True
    raum = str(c.get("ncs") or "")
    if len(werte) == 1 and raum in ("Separation", "DeviceN", "Indexed", "Pattern", "Lab"):
        return None, werte[0] > 0.05
    if len(werte) == 3 and raum == "Lab":
        return None, True
    return _rgb(werte), False


def _bild_anteil(bilder: list[tuple[float, float, float, float]], w: float, h: float) -> float:
    if not bilder or w <= 0 or h <= 0:
        return 0.0
    return min(1.0, sum((x1 - x0) * (b - t) for x0, t, x1, b in bilder) / (w * h))


def _lese_seite(page: Any, nr: int, cfg: dict) -> _SeitenDaten:
    sd = _SeitenDaten(nr=nr, breite=float(page.width), hoehe=float(page.height))
    modi = _render_modi(page)
    roh = list(page.chars)
    if modi is not None and len(modi) != len(roh):
        modi = None
    for i, c in enumerate(roh):
        text = c.get("text") or ""
        rgb, unklar = _zeichenfarbe(c)
        leer = not text.strip()
        z = _Z(text=text, x0=c["x0"], x1=c["x1"], top=c["top"], bottom=c["bottom"],
               size=float(c.get("size") or 10), font=str(c.get("fontname") or ""), leer=leer)
        z.unsichtbar = bool(modi and modi[i] in (3, 7))
        if rgb is not None:
            z.hex = _hex(rgb)
            z.weiss = min(rgb) >= 0.97
            z.farbklasse = klasse(rgb, cfg)
        elif unklar:
            z.hex = "sonderfarbe"
        if not leer and not z.unsichtbar:
            z.farbe_bed = _bedeutung(z.farbklasse, cfg)
            if z.farbklasse is None:
                z.unb = _unbekannt_art(rgb, cfg) if rgb is not None else ("sonder" if unklar else "")
        sd.zeichen.append(z)
    # Bilder
    w, h = sd.breite, sd.hoehe
    for im in page.images:
        x0, x1 = max(0.0, im["x0"]), min(w, im["x1"])
        top, bot = max(0.0, im["top"]), min(h, im["bottom"])
        if x1 > x0 and bot > top:
            sd.bilder.append((x0, top, x1, bot))
    sd.bild_anteil = _bild_anteil(sd.bilder, w, h)
    sd.grafik_objekte = len(page.curves) + len(page.rects) + len(page.lines)
    _tabellen_suchen(page, sd, cfg)
    return sd


# ---------------------------------------------------------------- Tabellen mit Rahmenlinien
def _tabellen_suchen(page: Any, sd: _SeitenDaten, cfg: dict) -> None:
    tc = cfg["tabellen"]
    lang, hoch = tc["linie_min_laenge"], tc["linie_min_hoehe"]
    wag: list[tuple[float, float, float]] = []      # (y, x0, x1)
    senk: list[tuple[float, float, float]] = []     # (x, top, bottom)
    for l in page.lines:
        dx, dy = l["x1"] - l["x0"], l["bottom"] - l["top"]
        if dy <= 1.0 and dx >= lang:
            wag.append(((l["top"] + l["bottom"]) / 2, l["x0"], l["x1"]))
        elif dx <= 1.0 and dy >= hoch:
            senk.append(((l["x0"] + l["x1"]) / 2, l["top"], l["bottom"]))
    for r in page.rects:
        bw, bh = r["x1"] - r["x0"], r["bottom"] - r["top"]
        if bh <= 1.6 and bw >= lang:
            wag.append(((r["top"] + r["bottom"]) / 2, r["x0"], r["x1"]))
        elif bw <= 1.6 and bh >= hoch:
            senk.append(((r["x0"] + r["x1"]) / 2, r["top"], r["bottom"]))
        elif r.get("stroke") and bw >= hoch and bh >= hoch:       # Zelle mit Rahmen
            wag += [(r["top"], r["x0"], r["x1"]), (r["bottom"], r["x0"], r["x1"])]
            senk += [(r["x0"], r["top"], r["bottom"]), (r["x1"], r["top"], r["bottom"])]
    if len(wag) + len(senk) > cfg["seiten"]["max_linien"]:
        sd.zu_viele_linien = True
        return
    if len(wag) < 2 or len(senk) < 2:
        return
    tol = tc["toleranz"]
    n = len(wag) + len(senk)
    eltern = list(range(n))

    def finde(i: int) -> int:
        while eltern[i] != i:
            eltern[i] = eltern[eltern[i]]
            i = eltern[i]
        return i

    for i, (y, a, b) in enumerate(wag):
        for j, (x, t, u) in enumerate(senk):
            if a - tol <= x <= b + tol and t - tol <= y <= u + tol:
                eltern[finde(i)] = finde(len(wag) + j)
    gruppen: dict[int, list[int]] = {}
    for i in range(n):
        gruppen.setdefault(finde(i), []).append(i)
    for mitglieder in gruppen.values():
        ws = [wag[i] for i in mitglieder if i < len(wag)]
        ss = [senk[i - len(wag)] for i in mitglieder if i >= len(wag)]
        zeilen_y = _verschmelze([y for y, _, _ in ws], tol)
        spalten_x = _verschmelze([x for x, _, _ in ss], tol)
        if (len(zeilen_y) - 1) * (len(spalten_x) - 1) >= tc["min_zellen"]:
            sd.tabellen.append((min(a for _, a, _ in ws), min(y for y, _, _ in ws),
                                max(b for _, _, b in ws), max(y for y, _, _ in ws)))


def _verschmelze(werte: list[float], tol: float) -> list[float]:
    out: list[float] = []
    for w in sorted(werte):
        if not out or w - out[-1] > tol:
            out.append(w)
    return out


# ---------------------------------------------------------------- Markierungen
class _Markierer:
    """Sammelt pro Seite alle Markierungs-Signale und beantwortet: Status je Zeichen.

    Ablauf: __init__ sammelt die rohen Signale, pruefe() entscheidet mit Hilfe der Zeichen, welche davon wirklich
    Markierungen sind (und welche Gestaltung: Zierlinie, Tabellenkopf, Hyperlink), status() bewertet ein Zeichen."""

    def __init__(self, page: Any, cfg: dict) -> None:
        self.cfg = cfg
        self.flaechen: list[tuple[float, float, float, float, str]] = []       # Hinterlegungen = Markierung
        self.unb_flaechen: list[tuple[float, float, float, float, str]] = []   # nicht zugeordnete Hinterlegungen
        self.design: list[tuple[float, float, float, float]] = []              # Hinterlegungen = Gestaltung
        self.striche: list[tuple[float, float, float]] = []
        self.annot: list[tuple[float, float, float, float, str]] = []
        self.unb_annot: list[tuple[float, float, float, float, str]] = []
        self.links: list[tuple[float, float, float, float]] = []
        self.rect_farben: Counter = Counter()
        self._flaechen_roh: list[tuple[float, float, float, float, str | None, str]] = []   # (Box, Bedeutung, Farbe)
        self._striche_roh: list[tuple[float, float, float, bool]] = []                      # (x0, x1, y, hell)
        w, h = float(page.width), float(page.height)
        if cfg["hintergrund"]["aktiv"]:
            hg = cfg["hintergrund"]
            for r in page.rects:
                if not r.get("fill"):
                    continue
                rgb = _rgb(r.get("non_stroking_color"))
                if rgb is None:
                    continue
                if r["width"] * r["height"] > hg["max_flaeche_anteil"] * w * h:
                    continue
                if r["height"] <= 1.6:     # duenne Rechtecke sind Linien (siehe unten)
                    continue
                k = klasse(rgb, cfg, hintergrund=True)
                self.rect_farben[(_hex(rgb), k)] += 1
                bed = _bedeutung(k, cfg)
                box = (r["x0"], r["top"], r["x1"], r["bottom"])
                if bed or (k is None and self._kraeftig_hg(rgb)):
                    self._flaechen_roh.append(box + (bed, _hex(rgb)))
        d = cfg["durchstreichung"]
        if d["aktiv"]:
            bed = _bedeutung("durchgestrichen", cfg)
            if bed:
                for l in page.lines:
                    if (l.get("linewidth", 1) or 0) <= d["max_linienstaerke_pt"] and abs(l["bottom"] - l["top"]) <= 1.2 \
                            and (l["x1"] - l["x0"]) < d["max_breite_anteil"] * w:
                        self._striche_roh.append((l["x0"], l["x1"], (l["top"] + l["bottom"]) / 2,
                                                  self._hell(l.get("stroking_color"))))
                for r in page.rects:
                    if r["height"] <= 1.6 and 2 <= r["width"] < d["max_breite_anteil"] * w:
                        self._striche_roh.append((r["x0"], r["x1"], (r["top"] + r["bottom"]) / 2,
                                                  self._hell(r.get("non_stroking_color"))))
        if cfg["annotationen"]["aktiv"]:
            self._annotationen(page, h)

    def _kraeftig_hg(self, rgb: tuple[float, float, float]) -> bool:
        hg = self.cfg["hintergrund"]
        _, s, v = colorsys.rgb_to_hsv(*(max(0.0, min(1.0, x)) for x in rgb))
        return s >= hg["min_saettigung"] and v >= hg["min_helligkeit"]

    def _hell(self, farbe: Any) -> bool:
        """Hellgraue (neutrale, helle) Linie? Das ist eine Zierlinie, keine Durchstreichung."""
        rgb = _rgb(farbe)
        if rgb is None:
            return False
        d = self.cfg["durchstreichung"]
        _, s, v = colorsys.rgb_to_hsv(*(max(0.0, min(1.0, x)) for x in rgb))
        return s <= d["hell_max_saettigung"] and v >= d["hell_min_helligkeit"]

    def _annotationen(self, page: Any, h: float) -> None:
        cfg = self.cfg
        for a in page.annots:
            data = a.get("data") or {}
            sub = data.get("Subtype")
            name = getattr(sub, "name", str(sub))
            if name == "Link":
                self.links.append((a["x0"], a["top"], a["x1"], a["bottom"]))
                continue
            if name == "StrikeOut":
                bed = _bedeutung("durchgestrichen", cfg)
                unb = None
            elif name == "Highlight":
                rgb = _rgb(data.get("C"))
                k = klasse(rgb, cfg, hintergrund=True)
                bed = _bedeutung(k, cfg)
                unb = None if (bed or k is not None) else (_hex(rgb) if rgb is not None else "-")
            else:
                continue
            if not bed and not unb:
                continue
            quad = data.get("QuadPoints")
            try:
                q = [float(x) for x in quad] if quad else []
            except (TypeError, ValueError):
                q = []
            boxen = []
            if len(q) >= 8:
                for i in range(0, len(q) - 7, 8):
                    xs, ys = q[i:i + 8:2], q[i + 1:i + 8:2]
                    boxen.append((min(xs), h - max(ys), max(xs), h - min(ys)))
            else:
                boxen.append((a["x0"], a["top"], a["x1"], a["bottom"]))
            for x0, top, x1, bot in boxen:
                if bed:
                    self.annot.append((x0, top, x1, bot, bed))
                else:
                    self.unb_annot.append((x0, top, x1, bot, unb or "-"))

    # ---- Entscheidung: Markierung oder Gestaltung? ----------------------------------------------
    def pruefe(self, zeichen: list[_Z]) -> None:
        """Entscheidet mit Hilfe der Zeichen der Seite, welche Linien und Flaechen echte Markierungen sind."""
        cfg = self.cfg
        d = cfg["durchstreichung"]
        for x0, x1, y, hell in self._striche_roh:
            treffer = [c for c in zeichen
                       if c.top + d["band"][0] * (c.bottom - c.top) <= y <= c.top + d["band"][1] * (c.bottom - c.top)
                       and min(x1, c.x1) - max(x0, c.x0) > 0]
            if not treffer:
                continue
            size = statistics.median(c.size for c in treffer)
            ueber = max(0.0, min(c.x0 for c in treffer) - x0) + max(0.0, x1 - max(c.x1 for c in treffer))
            if hell or ueber > d["max_ueberstand_em"] * size:
                # Zierlinie quer durch den Text: nicht als Durchstreichung werten, aber nicht verschweigen
                for c in treffer:
                    if not c.leer and (min(x1, c.x1) - max(x0, c.x0)) >= d["min_abdeckung"] * max(c.x1 - c.x0, 0.1):
                        c.hinweis.add("linie_unklar")
                continue
            self.striche.append((x0, x1, y))
        f = cfg["flaechen"]
        for x0, top, x1, bot, bed, hx in self._flaechen_roh:
            innen = [c for c in zeichen if x0 - 0.5 <= (c.x0 + c.x1) / 2 <= x1 + 0.5
                     and top - 0.5 <= (c.top + c.bottom) / 2 <= bot + 0.5]
            if not any(not c.leer for c in innen):
                continue
            size = statistics.median(c.size for c in innen if not c.leer)
            ueber = max(0.0, min(c.x0 for c in innen) - x0) + max(0.0, x1 - max(c.x1 for c in innen))
            besetzt = sum((c.x1 - c.x0) * (c.bottom - c.top) for c in innen) / max((x1 - x0) * (bot - top), 0.1)
            anliegend = (ueber <= f["max_ueberstand_em"] * size and besetzt >= f["min_fuellgrad"]
                         and (bot - top) <= f["max_hoehe_zeilen"] * size)
            if not anliegend:
                self.design.append((x0, top, x1, bot))      # Tabellenkopf, Baenderung, Kasten ...
            elif bed:
                self.flaechen.append((x0, top, x1, bot, bed))
            else:
                self.unb_flaechen.append((x0, top, x1, bot, hx))

    def status(self, c: _Z) -> str:
        cfg = self.cfg
        bed: set[str] = set()
        farbe_bed = c.farbe_bed
        cx, cy = (c.x0 + c.x1) / 2, (c.top + c.bottom) / 2
        if farbe_bed == "neu" and c.farbklasse == "blau" and not c.leer:
            for x0, top, x1, bot in self.links:
                if x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bot + 1:
                    farbe_bed = None                    # Hyperlink: das Blau ist Gestaltung, keine Neu-Markierung
                    c.hinweis.add("link")
                    break
        if farbe_bed:
            bed.add(farbe_bed)
        geo = False
        for x0, top, x1, bot, b in self.flaechen:
            if x0 - 0.5 <= cx <= x1 + 0.5 and top - 0.5 <= cy <= bot + 0.5:
                bed.add(b)
                geo = True
        for x0, top, x1, bot, b in self.annot:
            if x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bot + 1:
                bed.add(b)
                geo = True
        if self.striche:
            d = cfg["durchstreichung"]
            h = c.bottom - c.top
            lo, hi = c.top + d["band"][0] * h, c.top + d["band"][1] * h
            breite = max(c.x1 - c.x0, 0.1)
            for x0, x1, y in self.striche:
                if lo <= y <= hi:
                    ueber = min(x1, c.x1) - max(x0, c.x0)
                    if ueber / breite >= d["min_abdeckung"]:
                        bed.add(_bedeutung("durchgestrichen", cfg) or "entfaellt")
                        geo = True
                        break
        if "entfaellt" in bed:
            erg = DEL
        elif "neu" in bed:
            erg = ADD
        else:
            erg = KEEP
        c.unscharf = erg != KEEP and geo and farbe_bed != ("neu" if erg == ADD else "entfaellt")
        if not c.leer:
            if erg == KEEP:
                for x0, top, x1, bot in self.design:
                    if x0 - 0.5 <= cx <= x1 + 0.5 and top - 0.5 <= cy <= bot + 0.5:
                        c.hinweis.add("design_flaeche")
                        break
                for x0, top, x1, bot, hx in self.unb_flaechen:
                    if x0 - 0.5 <= cx <= x1 + 0.5 and top - 0.5 <= cy <= bot + 0.5:
                        c.unb_flaeche = hx
                        break
                else:
                    for x0, top, x1, bot, hx in self.unb_annot:
                        if x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bot + 1:
                            c.unb_flaeche = hx
                            break
        c.status = erg
        return erg


def _zeichenstatus(g: list[_Z], anteil: float) -> list[str]:
    """Status je Zeichen eines Wortes. Der Status jedes Zeichens bleibt erhalten, damit nichts verloren geht:
    'Netzleistung' (gestrichen) + 'Leistung' (neu) ohne Leerzeichen bleibt ein Wort mit altem UND neuem Text,
    'ab' + 'gesichert' (neu) behaelt das unveraenderte 'ab'.
    Nur bei Markierungen aus Geometrie (Linie, Flaeche, Kommentar) wird ein einzelnes unmarkiertes Randzeichen
    (Satzzeichen, hoechstens ein Buchstabe) dem Rest angepasst, weil Linien und Flaechen oft knapp daneben liegen."""
    n = len(g)
    st = [z.status for z in g]
    n_del = st.count(DEL)
    n_add = st.count(ADD)
    if n_del and n_add:
        return st                      # Ersetzung ohne Leerzeichen: Zeichen einzeln behalten
    marke = DEL if n_del else ADD if n_add else None
    if marke is None or (n_del + n_add) / n < anteil:
        return st
    behalten = [i for i, z in enumerate(g) if z.status == KEEP]
    buchstaben = [i for i in behalten if g[i].text[:1].isalnum()]
    knapp_daneben = (len(buchstaben) == 1 and buchstaben[0] in (0, n - 1)
                     and all(z.unscharf for z in g if z.status == marke))
    if not buchstaben or knapp_daneben:
        return [marke] * n
    return st


def _woerter(zeichen: list[_Z], seite: int, cfg: dict, grund_hex: str = "") -> tuple[list[PdfToken], list[_Info]]:
    """Zeichen einer Zeile -> Woerter (mit Statusanteilen)."""
    lay = cfg["layout"]
    zeichen = sorted(zeichen, key=lambda z: z.x0)
    gruppen: list[list[_Z]] = []
    for z in zeichen:
        if gruppen and z.x0 - gruppen[-1][-1].x1 <= lay["wort_luecke"] * z.size:
            gruppen[-1].append(z)
        else:
            gruppen.append([z])
    tokens = []
    infos = []
    for g in gruppen:
        n = len(g)
        st = _zeichenstatus(g, cfg["durchstreichung"]["wort_anteil"])
        old = "".join(z.text for z, s in zip(g, st) if s in (KEEP, DEL))
        new = "".join(z.text for z, s in zip(g, st) if s in (KEEP, ADD))
        if all(s == KEEP for s in st):
            status = KEEP
        elif all(s == DEL for s in st):
            status = DEL
        elif all(s == ADD for s in st):
            status = ADD
        else:
            status = MOD
        bbox = (min(z.x0 for z in g), min(z.top for z in g), max(z.x1 for z in g), max(z.bottom for z in g))
        info = _Info()
        unb = [z for z, s in zip(g, st) if s == KEEP and ((z.unb and z.hex != grund_hex) or z.unb_flaeche)]
        if unb and len(unb) * 2 >= n:
            info.unb = True
            z0 = unb[0]
            info.farbe = z0.hex if (z0.unb and z0.hex != grund_hex) else z0.unb_flaeche
        if _RE_CID.search(old + new) or _PRIVATE_ZEICHEN.search(old + new) or "\ufffd" in old + new:
            info.unlesbar = True
        hin: set[str] = set().union(*(z.hinweis for z in g))
        if info.unlesbar:
            hin.add("nicht_lesbar")
        if info.unb:
            hin.add("unbekannte_farbe")
        if _LIGATUR.search(old + new):
            hin.add("ligatur")
        if status == MOD and any(z.unscharf for z in g):
            hin.add("teilweise_markiert")
        tokens.append(PdfToken(old=old, new=new, status=status, page=seite, bbox=bbox, hinweis=";".join(sorted(hin))))
        infos.append(info)
    return tokens, infos


def _zeilen(zeichen: list[_Z], seite: int, cfg: dict, grund_hex: str = "") -> list[_Zeile]:
    tol = cfg["layout"]["zeilen_toleranz"]
    zeichen = [z for z in zeichen if not z.leer]
    zeichen = sorted(zeichen, key=lambda z: ((z.top + z.bottom) / 2, z.x0))
    cluster: list[list[_Z]] = []
    mitten: list[float] = []
    for z in zeichen:
        mitte = (z.top + z.bottom) / 2
        if cluster and abs(mitte - mitten[-1]) <= tol * z.size:
            cluster[-1].append(z)
            mitten[-1] = statistics.fmean((c.top + c.bottom) / 2 for c in cluster[-1])
        else:
            cluster.append([z])
            mitten.append(mitte)
    zeilen = []
    for cl in cluster:
        toks, infos = _woerter(cl, seite, cfg, grund_hex)
        if not toks:
            continue
        t0 = toks[0]
        if t0.status == KEEP and _NUR_PRIVAT.fullmatch(t0.old):      # Symbolschrift-Aufzaehlungspunkt am Zeilenanfang: kein Lesefehler
            t0.hinweis = ";".join(h for h in t0.hinweis.split(";") if h and h != "nicht_lesbar")
            infos[0].unlesbar = False
        fett_n = sum(1 for z in cl if re.search(r"bold|black|heavy|semibold", z.font, re.I))
        zeilen.append(_Zeile(
            tokens=toks, infos=infos, top=min(z.top for z in cl), bottom=max(z.bottom for z in cl),
            x0=min(z.x0 for z in cl), x1=max(z.x1 for z in cl),
            size=statistics.median(z.size for z in cl), fett=fett_n / len(cl) > 0.6, seite=seite,
            unsichtbar=sum(1 for z in cl if z.unsichtbar or z.weiss) * 2 > len(cl)))
    zeilen.sort(key=lambda z: z.top)
    return zeilen


# ---------------------------------------------------------------- Bloecke
_AUFZAEHLUNG = re.compile(r"^([•\-–·▪●○*]|\(?\d{1,2}\)|[a-z]\))$")


def _erste_form(t: PdfToken) -> str:
    return t.new or t.old


def _ist_ueberschrift(z: _Zeile, body: float, cfg: dict) -> bool:
    t = z.tokens[0]
    nr_alt = section_nummer(t.old) if t.old else None
    nr_neu = section_nummer(t.new) if t.new else None
    wenige = len(z.tokens) <= 25
    gross = z.size >= body * cfg["layout"]["ueberschrift_faktor"]
    if (nr_alt or nr_neu) and len(z.tokens) >= 2 and wenige and (z.fett or gross):
        return True
    return bool(z.fett and z.size >= body * 1.15 and wenige)


def _kopf_fuss_schluessel(z: _Zeile) -> str:
    text = " ".join(_erste_form(t) for t in z.tokens)
    return re.sub(r"\d+", "#", text.casefold())


def _zu_bloecken(alle_zeilen: list[list[_Zeile]], seiten_h: dict[int, float], cfg: dict,
                 ergebnis: PdfAnalyse) -> list[list[_Zeile]]:
    """Fasst die Zeilen zu Bloecken zusammen. Gibt je Seite die Zeilen zurueck, die KEINE Kopf-/Fusszeile sind."""
    lay = cfg["layout"]
    # Kopf-/Fusszeilen finden
    kandidaten: dict[str, set[int]] = {}
    for zl in alle_zeilen:
        for z in zl:
            h = seiten_h[z.seite]
            if z.top < lay["kopf_fuss_anteil"] * h or z.bottom > (1 - lay["kopf_fuss_anteil"]) * h:
                kandidaten.setdefault(_kopf_fuss_schluessel(z), set()).add(z.seite)
    n_seiten = len(alle_zeilen)
    muster = re.compile(r"^(seite|page)\s*#", re.I)

    def im_rand(z: _Zeile) -> bool:
        h = seiten_h[z.seite]
        return z.top < lay["kopf_fuss_anteil"] * h or z.bottom > (1 - lay["kopf_fuss_anteil"]) * h

    # Schriftgroesse des Fliesstextes (ohne Randzeilen), nur zur Pruefung der Randzeilen
    g0 = Counter()
    for zl in alle_zeilen:
        for z in zl:
            if not im_rand(z):
                g0[round(z.size * 2) / 2] += sum(len(t.old or t.new) for t in z.tokens)
    body0 = g0.most_common(1)[0][0] if g0 else 11.0
    rand_ids = _abgesetzte_randzeilen(alle_zeilen, seiten_h, body0, lay)

    def ist_kopf_fuss(z: _Zeile) -> bool:
        if id(z) in rand_ids:
            return True
        if not im_rand(z):
            return False
        key = _kopf_fuss_schluessel(z)
        return bool(muster.match(key)) or (n_seiten >= 2 and len(kandidaten.get(key, ())) >= max(2, n_seiten // 2))

    groessen = Counter()
    for zl in alle_zeilen:
        for z in zl:
            if not ist_kopf_fuss(z):
                groessen[round(z.size * 2) / 2] += sum(len(t.old or t.new) for t in z.tokens)
    body = groessen.most_common(1)[0][0] if groessen else 11.0

    sec_alt: str | None = None
    sec_neu: str | None = None
    idx = 0
    seiten_nutz: list[list[_Zeile]] = []
    abstaende: list[float] = []
    for zl in alle_zeilen:
        nutz = []
        for z in zl:
            if ist_kopf_fuss(z):
                gew = sum(1 for t in z.tokens if t.status != KEEP)
                if gew:
                    ergebnis.ignoriert.append({"seite": z.seite, "grund": "kopf_fuss", "woerter": gew,
                                               "text": " ".join(_erste_form(t) for t in z.tokens)})
                continue
            nutz.append(z)
        seiten_nutz.append(nutz)
        abstaende += [b.top - a.top for a, b in zip(nutz, nutz[1:]) if 0.9 * a.size < b.top - a.top < 3 * a.size]
    _design_ueberschriften(seiten_nutz, body, cfg, ergebnis)
    median = _zeilenabstand(abstaende, body)
    weit = median >= 1.5 * body          # grosser Zeilenabstand: Absatzende erkennt man an der kurzen Zeile
    rechts_x = sorted(z.x1 for zl in seiten_nutz for z in zl if len(z.tokens) >= 5)
    rechts = rechts_x[int(0.9 * (len(rechts_x) - 1))] if rechts_x else None
    alle_z = [z for zl in seiten_nutz for z in zl]
    if alle_z and ergebnis.seiten_masse:
        links = min(z.x0 for z in alle_z)
        breite = max(b for b, _ in ergebnis.seiten_masse.values())
        schwelle_kurz = links + 0.75 * (breite - 2 * links)   # Zeile endet deutlich vor dem rechten Rand
    else:
        schwelle_kurz = None

    def kurz(z: _Zeile) -> bool:
        return (rechts is not None and z.x1 < rechts - 3 * z.size) or (schwelle_kurz is not None and z.x1 < schwelle_kurz)
    for nutz in seiten_nutz:
        if not nutz:
            continue
        aktuell: list[_Zeile] = []
        art = "absatz"

        def abschliessen() -> None:
            nonlocal idx, sec_alt, sec_neu, aktuell, art
            if not aktuell:
                return
            toks = [t for z in aktuell for t in z.tokens]
            if art == "ueberschrift":
                t0 = toks[0]
                n_alt = section_nummer(t0.old) if t0.old else None
                n_neu = section_nummer(t0.new) if t0.new else None
                if n_alt or n_neu:
                    sec_alt, sec_neu = n_alt, n_neu
                else:   # unnummerierte Ueberschrift: Zuordnung ueber Text, Nummern bleiben
                    sec_alt = sec_neu = None
            else:
                toks = _silben(toks, aktuell, cfg)
            bbox = (min(z.x0 for z in aktuell), min(z.top for z in aktuell),
                    max(z.x1 for z in aktuell), max(z.bottom for z in aktuell))
            ergebnis.blocks.append(PdfBlock(idx=idx, page=aktuell[0].seite, kind=art, tokens=toks, bbox=bbox,
                                            section_alt=sec_alt, section_neu=sec_neu,
                                            hinweise=_block_hinweise(toks)))
            idx += 1
            aktuell = []

        vorher: _Zeile | None = None
        for z in nutz:
            kopf = _ist_ueberschrift(z, body, cfg)
            neu_start = (not aktuell or kopf or art == "ueberschrift"
                         or (vorher is not None and z.top - vorher.top > lay["absatz_faktor"] * median)
                         or (weit and vorher is not None and kurz(vorher))
                         or bool(_AUFZAEHLUNG.match(_erste_form(z.tokens[0]))))
            if neu_start:
                abschliessen()
                art = "ueberschrift" if kopf else "absatz"
            aktuell.append(z)
            vorher = z
        abschliessen()
    return seiten_nutz


def _design_ueberschriften(seiten_nutz: list[list[_Zeile]], body: float, cfg: dict, ergebnis: PdfAnalyse) -> None:
    """Sind ALLE Ueberschriften des Dokuments blau und ist sonst kaum etwas markiert, ist das Blau Gestaltung:
    die Ueberschriften gelten als unveraendert (mit Hinweis und Luecke), nicht als neu."""
    dc = cfg["ueberschriften_design"]
    koepfe = [z for nutz in seiten_nutz for z in nutz if _ist_ueberschrift(z, body, cfg)]
    if len(koepfe) < dc["min_anzahl"] or not all(t.status == ADD for z in koepfe for t in z.tokens):
        return
    kopf_ids = {id(z) for z in koepfe}
    rest = [t for nutz in seiten_nutz for z in nutz if id(z) not in kopf_ids for t in z.tokens]
    if rest and sum(1 for t in rest if t.status != KEEP) / len(rest) > dc["max_anteil_text"]:
        return
    je_seite: Counter = Counter()
    for z in koepfe:
        for t in z.tokens:
            t.old, t.status = t.new, KEEP
            t.hinweis = ";".join(sorted(set(filter(None, t.hinweis.split(";"))) | {"design_blau"}))
        je_seite[z.seite] += len(z.tokens)
    for seite, n in sorted(je_seite.items()):
        ergebnis.luecken.append(luecke(
            "sonstiges", seite, f"Seite {seite}: Alle Überschriften des Dokuments sind blau, im übrigen Text ist kaum "
                                f"etwas markiert. Das Blau wurde als Gestaltung gewertet, nicht als „neu“.", n))


def _abgesetzte_randzeilen(alle_zeilen: list[list[_Zeile]], seiten_h: dict[int, float], body: float,
                           lay: dict) -> set[int]:
    """Zeilen am oberen/unteren Seitenrand, die vom Text abgesetzt stehen und nicht groesser sind als der Fliesstext:
    Kopf-/Fusszeilen, auch wenn sie sich nicht auf anderen Seiten wiederholen (z. B. einseitiges Dokument)."""
    ids: set[int] = set()
    for zl in alle_zeilen:
        if len(zl) < 2:
            continue
        h = seiten_h[zl[0].seite]
        rand = lay["kopf_fuss_anteil"] * h

        def abgesetzt(z: _Zeile, abstand: float) -> bool:
            return abstand >= lay["kopf_fuss_abstand_em"] * z.size and z.size <= body * lay["kopf_fuss_max_schrift"]

        # oben: von oben nach unten; die unterste abgesetzte Randzeile bestimmt, wie viele Zeilen Kopf sind
        letzte = -1
        for i, z in enumerate(zl[:-1]):
            if z.bottom > rand:
                break
            if abgesetzt(z, zl[i + 1].top - z.bottom):
                letzte = i
        ids.update(id(z) for z in zl[:letzte + 1])
        # unten: von unten nach oben
        erste = len(zl)
        for i in range(len(zl) - 1, 0, -1):
            z = zl[i]
            if z.top < h - rand:
                break
            if abgesetzt(z, z.top - zl[i - 1].bottom):
                erste = i
        ids.update(id(z) for z in zl[erste:])
    return ids


def _block_hinweise(tokens: list[PdfToken]) -> list[str]:
    """Saetze zu den Hinweisen der Woerter eines Blocks (jeder Satz einmal)."""
    codes = []
    for t in tokens:
        for c in t.hinweis.split(";") if t.hinweis else ():
            if c not in codes:
                codes.append(c)
    return [_HINWEIS_TEXT[c] for c in codes if c in _HINWEIS_TEXT]


def _zeilenabstand(abstaende: list[float], body: float) -> float:
    """Typischer Zeilenabstand (dokumentweit): kleinste Gruppe (+-8 %) mit mindestens 2 Treffern."""
    if not abstaende:
        return body * 1.3
    gruppen: list[list[float]] = []
    for a in sorted(abstaende):
        if gruppen and a <= gruppen[-1][0] * 1.08:
            gruppen[-1].append(a)
        else:
            gruppen.append([a])
    stark = [g for g in gruppen if len(g) >= 2] or gruppen
    return statistics.median(stark[0])


def _silben(tokens: list[PdfToken], zeilen: list[_Zeile], cfg: dict) -> list[PdfToken]:
    """Fuegt am Zeilenende getrennte Woerter ('Kapazi-' + 'taet') wieder zusammen."""
    lay = cfg["layout"]
    if not lay["silbentrennung_zusammenfuegen"] or len(zeilen) < 2:
        return tokens
    ausnahmen = {norm(w) for w in lay["keine_silbentrennung_vor"]}
    erste_der_zeilen = {id(z.tokens[0]) for z in zeilen[1:]}
    out: list[PdfToken] = []
    for t in tokens:
        if out and id(t) in erste_der_zeilen:
            v = out[-1]
            if (v.status == KEEP and t.status == KEEP and len(v.old) > 2 and v.old.endswith("-")
                    and t.old[:1].islower() and norm(t.old) not in ausnahmen):
                bbox = (v.bbox[0], min(v.bbox[1], t.bbox[1]), max(v.bbox[2], t.bbox[2]), max(v.bbox[3], t.bbox[3]))
                hin = ";".join(sorted((set(v.hinweis.split(";")) | set(t.hinweis.split(";"))) - {""}))
                out[-1] = PdfToken(old=v.old[:-1] + t.old, new=v.new[:-1] + t.new, status=KEEP, page=v.page, bbox=bbox,
                                   hinweis=hin)
                continue
        out.append(t)
    return out


# ---------------------------------------------------------------- Luecken
def _in_box(t: PdfToken, box: tuple[float, float, float, float]) -> bool:
    cx, cy = (t.bbox[0] + t.bbox[2]) / 2, (t.bbox[1] + t.bbox[3]) / 2
    return box[0] - 2 <= cx <= box[2] + 2 and box[1] - 2 <= cy <= box[3] + 2


def _ist_marke(token: PdfToken) -> bool:
    """Ein Aufzaehlungszeichen / eine Nummer ('1.', 'a)', '•')?"""
    return bool(re.fullmatch(r"[\W_]{1,3}|\(?[0-9A-Za-z]{1,3}[.)]", _erste_form(token)))


def _spalten_regionen(zeilen: list[_Zeile], tabellen: list[tuple[float, float, float, float]],
                      cfg: dict) -> list[tuple[float, float, float, float]]:
    """Bereiche mit mehreren Textspalten nebeneinander (oder spaltigen Tabellen ohne Rahmen).

    Jede Zeile wird an grossen Luecken in Abschnitte zerlegt. Zwei Hinweise auf Spalten:
      (A) mindestens `min_zeilen_luecke` Zeilen haben an derselben Stelle eine solche Luecke (Spalten auf gleicher Hoehe);
      (B) mindestens `min_zeilen` Abschnitte beginnen am selben x und mindestens ebenso viele (im Verhaeltnis
          `min_anteil`) enden links davon, mit Abstand dazwischen und in derselben Hoehe der Seite
          (das geht auch, wenn die Zeilen der Spalten nicht auf gleicher Hoehe stehen)."""
    sp = cfg["spalten"]
    segs: list[tuple[float, float, float, float, float]] = []          # (x0, x1, top, bottom, size)
    luecken: list[tuple[int, float, float, float, float]] = []         # (Zeile, links, rechts, top, bottom)
    for n, z in enumerate(zeilen):
        mitte = (z.top + z.bottom) / 2
        if any(t[0] - 2 <= (z.x0 + z.x1) / 2 <= t[2] + 2 and t[1] - 2 <= mitte <= t[3] + 2 for t in tabellen):
            continue
        ts = sorted(z.tokens, key=lambda t: t.bbox[0])
        teil = [ts[0]]
        gruppen = []
        for t in ts[1:]:
            if t.bbox[0] - teil[-1].bbox[2] >= sp["min_luecke_em"] * z.size:
                if not (len(teil) == 1 and _ist_marke(teil[0]) and not gruppen):
                    luecken.append((n, teil[-1].bbox[2], t.bbox[0], z.top, z.bottom))
                gruppen.append(teil)
                teil = [t]
            else:
                teil.append(t)
        gruppen.append(teil)
        for g in gruppen:
            if len(g) == 1 and _ist_marke(g[0]):
                continue
            segs.append((g[0].bbox[0], g[-1].bbox[2], z.top, z.bottom, z.size))
    ergebnis: list[tuple[float, float, float, float]] = []
    # (A) mehrere Zeilen mit einer Luecke, hinter der der Text an derselben Stelle weitergeht (Spalten auf gleicher Hoehe)
    rest = list(luecken)
    while rest:
        beste: list[tuple[int, float, float, float, float]] = []
        for _, _, x, _, _ in rest:
            gleich = [g for g in rest if abs(g[2] - x) <= 2.5]
            if len({g[0] for g in gleich}) > len({g[0] for g in beste}):
                beste = gleich
        nummern = {g[0] for g in beste}
        if len(nummern) < sp["min_zeilen_luecke"]:
            break
        z_ok = [zeilen[n] for n in nummern]
        ergebnis.append((min(z.x0 for z in z_ok), min(z.top for z in z_ok),
                         max(z.x1 for z in z_ok), max(z.bottom for z in z_ok)))
        rest = [g for g in rest if g[0] not in nummern]
    # (B) zweite Spalte beginnt an derselben Stelle, linke Abschnitte enden davor
    benutzt: set[int] = set()
    starts = Counter(round(x0 / 2) * 2 for x0, _, _, _, _ in segs)
    for x, n in sorted(starts.items(), key=lambda kv: -kv[1]):
        if n < sp["min_zeilen"]:
            break
        rechts = [i for i, sg in enumerate(segs) if abs(sg[0] - x) <= 2.5 and i not in benutzt]
        if len(rechts) < sp["min_zeilen"]:
            continue
        groesse = statistics.median(segs[i][4] for i in rechts)
        abstand = sp["min_spaltenabstand_em"] * groesse
        links = [i for i, sg in enumerate(segs) if sg[1] <= x - abstand and sg[0] < x - abstand and i not in benutzt]
        if len(links) < sp["min_zeilen"]:
            continue
        r_top, r_bot = min(segs[i][2] for i in rechts), max(segs[i][3] for i in rechts)
        links = [i for i in links if segs[i][3] >= r_top and segs[i][2] <= r_bot]
        if len(links) < sp["min_zeilen"] or min(len(links), len(rechts)) < sp["min_anteil"] * max(len(links), len(rechts)):
            continue
        l_top, l_bot = min(segs[i][2] for i in links), max(segs[i][3] for i in links)
        ueber = min(r_bot, l_bot) - max(r_top, l_top)
        if ueber < sp["min_ueberlappung"] * min(r_bot - r_top, l_bot - l_top):
            continue
        alle = links + rechts
        benutzt.update(alle)
        ergebnis.append((min(segs[i][0] for i in alle), min(segs[i][2] for i in alle),
                         max(segs[i][1] for i in alle), max(segs[i][3] for i in alle)))
    return ergebnis


def _seite_luecken(sd: _SeitenDaten, zeilen_alle: list[_Zeile], nutz: list[_Zeile], cfg: dict,
                   ergebnis: PdfAnalyse, ohne_text: bool) -> bool:
    """Meldet, was auf dieser Seite nicht (sicher) gelesen werden konnte. True, wenn die Seite als Scan gilt."""
    sc = cfg["seiten"]
    lay = cfg["layout"]
    nr, w, h = sd.nr, sd.breite, sd.hoehe
    rand = lay["kopf_fuss_anteil"]
    sichtbar = sum(len(z.tokens) for z in nutz if not z.unsichtbar)
    unsichtbar = sum(len(z.tokens) for z in nutz if z.unsichtbar)
    # --- Scan-Seiten ---------------------------------------------------------------------------
    if sd.bild_anteil >= sc["bild_seite_anteil"] and sichtbar <= sc["bild_seite_max_woerter"]:
        if unsichtbar:
            text = (f"Seite {nr} ist ein Scan mit unsichtbarer Textebene (Texterkennung). Markierungen im Bild "
                    f"werden nicht gelesen.")
        else:
            text = f"Seite {nr} besteht fast nur aus einem Bild (Scan). Markierungen im Bild werden nicht gelesen."
        ergebnis.luecken.append(luecke("scan_seite", nr, text, 0))
        if not ohne_text:
            ergebnis.warnungen.append(f"Seite {nr} besteht fast nur aus einem Bild. Markierungen im Bild werden nicht gelesen.")
        return True
    if ohne_text:
        if sd.bilder or sd.grafik_objekte >= sc["grafik_min_objekte"]:
            ergebnis.luecken.append(luecke(
                "scan_seite", nr, f"Seite {nr} enthält keinen lesbaren Text, aber ein Bild oder eine Zeichnung "
                                  f"(Scan?). Änderungen darauf werden nicht gelesen.", 0))
        return True
    # --- Bilder im Text ------------------------------------------------------------------------
    for x0, t, x1, b in sd.bilder:
        flaeche = (x1 - x0) * (b - t) / (w * h)
        mitte = (t + b) / 2
        if flaeche < sc["bild_min_flaeche_anteil"] or mitte < rand * h or mitte > (1 - rand) * h:
            continue
        if sd.bild_anteil >= sc["bild_seite_anteil"]:      # Hintergrundbild einer Textseite
            continue
        ergebnis.luecken.append(luecke(
            "bild", nr, f"Seite {nr}: Ein Bild oder Diagramm liegt im Text. Markierungen im Bild werden nicht gelesen.", 0))
    if sd.zu_viele_linien:
        ergebnis.luecken.append(luecke(
            "bild", nr, f"Seite {nr} enthält eine Zeichnung mit sehr vielen Linien. Sie wird nicht gelesen.", 0))
    # --- Tabellen und Spalten --------------------------------------------------------------------
    if sd.tabellen:
        l = luecke("tabelle", nr, f"Seite {nr}: Eine Tabelle wurde zeilenweise als normaler Text gelesen. "
                                  f"Zellen können falsch zugeordnet sein.", 0)
        ergebnis.luecken.append(l)
        ergebnis.regionen.append((l, nr, list(sd.tabellen)))
    sp = _spalten_regionen(nutz, sd.tabellen, cfg)
    if sp:
        l = luecke("mehrspaltig", nr, f"Seite {nr}: Text steht in mehreren Spalten nebeneinander. "
                                      f"Die Wörter können vermischt gelesen worden sein.", 0)
        ergebnis.luecken.append(l)
        ergebnis.regionen.append((l, nr, sp))
    # --- Farbflaechen, die wie Gestaltung aussehen (Tabellenkopf, Streifen, Kasten) ----------------
    design = [t for z in nutz for t in z.tokens if "design_flaeche" in t.hinweis
              and not any(_in_box(t, box) for box in sd.tabellen)]
    if design:
        ergebnis.luecken.append(luecke(
            "sonstiges", nr, f"Seite {nr}: {len(design)} Wörter stehen auf einer Farbfläche, die nach Gestaltung aussieht "
                             f"(Kasten, Streifen). Sie wurde nicht als Markierung gewertet.", len(design)))
    # --- Farben und Zeichen ----------------------------------------------------------------------
    farb = [(i.farbe, 1) for z in nutz for i in z.infos if i.unb]
    if farb:
        farben = [f for f, _ in Counter(f for f, _ in farb).most_common(3) if f]
        zusatz = f" ({', '.join(farben)})" if farben else ""
        ergebnis.luecken.append(luecke(
            "unbekannte_farbe", nr,
            f"Seite {nr}: {len(farb)} Wörter haben eine Farbe{zusatz}, die nicht in der Legende steht. "
            f"Sie wurden als unverändert gewertet.", len(farb)))
    unles = sum(1 for z in nutz for i in z.infos if i.unlesbar)
    if unles:
        ergebnis.luecken.append(luecke(
            "nicht_lesbarer_text", nr,
            f"Seite {nr}: {unles} Wörter enthalten Zeichen, die nicht gelesen werden konnten "
            f"(Schrift ohne Zeichenzuordnung).", unles))
    return False


# ---------------------------------------------------------------- Einstieg
def analysiere_pdf(pfad: str | Path, cfg: dict | None = None) -> PdfAnalyse:
    cfg = cfg or config.lade("farben")
    pfad = Path(pfad)
    ergebnis = PdfAnalyse(pfad=str(pfad), seiten=0, blocks=[])
    seiten_daten: list[_SeitenDaten] = []
    markierer: dict[int, _Markierer] = {}
    with pdfplumber.open(str(pfad)) as pdf:
        ergebnis.seiten = len(pdf.pages)
        for nr, page in enumerate(pdf.pages, 1):
            sd = _lese_seite(page, nr, cfg)
            seiten_daten.append(sd)
            if any(not z.leer for z in sd.zeichen):
                markierer[nr] = _Markierer(page, cfg)
    # Grundfarbe des Dokuments (Farbe, in der der Grossteil des Textes steht): nie eine Markierung
    zaehler: Counter = Counter()
    for sd in seiten_daten:
        for z in sd.zeichen:
            if not z.leer and z.hex:
                zaehler[z.hex] += 1
    grund_hex = ""
    if zaehler:
        hx, n = zaehler.most_common(1)[0]
        if n / sum(zaehler.values()) >= cfg["unbekannt"]["grundfarbe_anteil"]:
            grund_hex = hx
    text_farben: Counter = Counter()
    rect_farben: Counter = Counter()
    unbekannt: Counter = Counter()
    alle_zeilen: list[list[_Zeile]] = []
    seiten_h: dict[int, float] = {}
    ohne_text: set[int] = set()
    for sd in seiten_daten:
        nr = sd.nr
        seiten_h[nr] = sd.hoehe
        ergebnis.seiten_masse[nr] = (sd.breite, sd.hoehe)
        chars = [z for z in sd.zeichen if not z.leer]
        if not chars:
            ohne_text.add(nr)
            alle_zeilen.append([])
            continue
        m = markierer[nr]
        rect_farben.update(m.rect_farben)
        m.pruefe(sd.zeichen)
        for z in sd.zeichen:
            m.status(z)
        for z in chars:
            if z.hex and z.hex != "sonderfarbe":
                text_farben[(z.hex, z.farbklasse)] += 1
                if z.unb == "bunt" and z.hex != grund_hex and z.farbklasse is None:
                    unbekannt[z.hex] += 1
        alle_zeilen.append(_zeilen(sd.zeichen, nr, cfg, grund_hex))
    nutz_pro_seite = _zu_bloecken(alle_zeilen, seiten_h, cfg, ergebnis)
    for sd, zl, nutz in zip(seiten_daten, alle_zeilen, nutz_pro_seite):
        _seite_luecken(sd, zl, nutz, cfg, ergebnis, sd.nr in ohne_text)
    for l, seite, boxen in ergebnis.regionen:
        l["woerter"] = sum(1 for b in ergebnis.blocks if b.page == seite for t in b.tokens
                           if t.status != KEEP and any(_in_box(t, box) for box in boxen))
    ergebnis.luecken.sort(key=lambda l: (l["seite"], l["art"]))
    # Farbuebersicht fuer das Dashboard (zum Kalibrieren)
    for (hx, k), n in text_farben.most_common():
        ergebnis.farben.append({"farbe": hx, "zeichen": n, "klasse": k or "-", "art": "Text",
                                "bedeutung": cfg["legende"].get(k, "-") if k else "keine Markierung"})
    for (hx, k), n in rect_farben.most_common():
        ergebnis.farben.append({"farbe": hx, "zeichen": n, "klasse": k or "-", "art": "Hinterlegung (Rechtecke)",
                                "bedeutung": cfg["legende"].get(k, "-") if k else "keine Markierung"})
    for nr in sorted(ohne_text):
        ergebnis.warnungen.append(f"Seite {nr} enthält keinen auslesbaren Text (Scan?). Sie wird nicht ausgewertet.")
    for hx, n in unbekannt.most_common(5):
        ergebnis.warnungen.append(f"Farbe {hx} ({n} Zeichen) ist nicht in der Legende (config/farben.yaml) und wird als unverändert behandelt.")
    if not any(b.hat_aenderung for b in ergebnis.blocks) and not ergebnis.ignoriert:
        ergebnis.warnungen.append("Im PDF wurden keine Änderungen erkannt. Bitte Farben im PDF prüfen (config/farben.yaml).")
    return ergebnis
