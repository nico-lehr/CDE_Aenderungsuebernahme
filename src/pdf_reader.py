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
    "gemischt_blau_strich": "Ein Wort ist blau UND durchgestrichen (widersprüchlich). Es wurde als „entfällt“ gelesen.",
    "gemischt_widerspruch": "Ein Wort ist gleichzeitig als „neu“ und als „entfällt“ markiert. Es wurde als „entfällt“ gelesen.",
    "unterstrichen": "Ein unterstrichenes Wort: eine Unterstreichung hat keine feste Bedeutung und wurde nicht ausgewertet.",
    "kommentar": "Ein PDF-Kommentar (z. B. Unterstreichung, Notiz) liegt auf dem Text. Seine Bedeutung ist unklar; "
                 "er wurde nicht ausgewertet.",
    "design_blau": "Die Überschrift ist blau wie alle Überschriften des Dokuments. Das wurde als Gestaltung gewertet, nicht als „neu“.",
    "design_flaeche": "Eine Farbfläche hinter dem Text sieht nach Gestaltung aus (Tabellenkopf, Streifen). "
                      "Sie wurde nicht als Markierung gewertet.",
    "mehrspaltig": "Dieses Wort steht in einem Bereich mit mehreren Spalten nebeneinander. Die Wörter können vermischt "
                   "gelesen worden sein (ein Satz kann zerrissen sein).",
    "tabelle": "Dieses Wort steht in einer Tabelle. Die Tabelle wurde zeilenweise als normaler Text gelesen; "
               "Zellen können falsch zugeordnet sein.",
    # Schema 3 (Leser-Treue R1-R8)
    "rand_unklar": "Eine Linie oder Fläche deckt den Rand eines Wortes nicht genau ab (z. B. ein Satzzeichen daneben). "
                   "Der Rand wurde dem Rest des Wortes zugeschlagen. Bitte prüfen.",
    "absatzgrenze_unklar": "Es ist nicht sicher, ob hier ein neuer Absatz beginnt (die Zeile davor endet etwas früher "
                           "als die anderen). Bitte prüfen.",
    "trennung_unklar": "Ein neues Wort ist am Zeilenende getrennt. Es wurde ohne Trennstrich zusammengefügt, kommt so aber "
                       "sonst nicht im Dokument vor. Bitte prüfen.",
    "hochzahl_markiert": "Eine kleine hochgestellte Marke (z. B. Fußnotenzahl) ist markiert. Sie wurde nicht als Teil des "
                         "Wortes gelesen. Bitte prüfen.",
    "linienfarbe_unklar": "Die Farbe einer Linie durch den Text konnte nicht gelesen werden. Sie wurde als Durchstreichung "
                          "gewertet. Bitte prüfen.",
}


def hinweis_text(code: str) -> str:
    """Satz in einfachem Deutsch zu einem Hinweis-Code eines Worts (leer, wenn es keinen gibt)."""
    return _HINWEIS_TEXT.get("kommentar" if code.startswith("kommentar_") else code, "")


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
    weich: bool = False                # weicher Trennstrich (Code 173), von pdfminer als Leerzeichen gelesen
    weiss: bool = False
    hex: str = ""
    farbklasse: str | None = None
    farbe_bed: str | None = None       # Bedeutung der Textfarbe ("neu" | "entfaellt" | None)
    unb: str = ""                      # Art der nicht zugeordneten Textfarbe (siehe _unbekannt_art)
    unb_flaeche: str = ""              # Farbe einer nicht zugeordneten Hinterlegung / Hervorhebung
    unscharf: bool = False             # Status stammt nur aus Geometrie (Linie, Flaeche, Kommentar)
    hinweis: set = field(default_factory=set)
    nr: int = -1                       # Position in der Zeichenfolge der Seite (Reihenfolge im Inhaltsstrom)


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
    weich_ende: bool = False          # die Zeile endet mit einem weichen Trennstrich


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
    rahmen: Any = None                 # _Rahmen (Leserichtung der Seite)
    fremd_woerter: int = 0             # Woerter in einer anderen Schreibrichtung als der Text der Seite
    fremd_farbig: bool = False         # ... davon mindestens eines in Legendenfarbe


# ---------------------------------------------------------------- Seite lesen
@dataclass(frozen=True)
class _Rahmen:
    """Bringt die Seite in Leserichtung. art: 'h' = Text steht waagerecht (nichts zu tun) | 'ab' = Text laeuft von
    oben nach unten | 'auf' = von unten nach oben | 'um' = kopfueber. (w, h) = Masse der Seite, wie pdfplumber sie meldet."""
    art: str = "h"
    w: float = 0.0
    h: float = 0.0

    def box(self, x0: float, top: float, x1: float, bot: float) -> tuple[float, float, float, float]:
        if self.art == "ab":
            return (top, self.w - x1, bot, self.w - x0)
        if self.art == "auf":
            return (self.h - bot, x0, self.h - top, x1)
        if self.art == "um":
            return (self.w - x1, self.h - bot, self.w - x0, self.h - top)
        return (x0, top, x1, bot)

    @property
    def masse(self) -> tuple[float, float]:
        return (self.h, self.w) if self.art in ("ab", "auf") else (self.w, self.h)


def _kasten(o: dict, rahmen: _Rahmen) -> tuple[float, float, float, float]:
    return rahmen.box(o["x0"], o["top"], o["x1"], o["bottom"])


def _richtung(c: dict) -> str:
    """Schreibrichtung eines Zeichens: 'h' | 'um' (kopfueber) | 'ab' | 'auf' | 'schraeg'."""
    m = c.get("matrix") or (1, 0, 0, 1, 0, 0)
    a, b = float(m[0]), float(m[1])
    if abs(b) <= 0.2 * abs(a):
        return "h" if a >= 0 else "um"
    if abs(a) <= 0.2 * abs(b):
        return "ab" if b < 0 else "auf"
    return "schraeg"


def _render_modi(page: Any) -> list[tuple[int, bool]] | None:
    """Je Zeichen: (Text-Render-Modus (3 = unsichtbar), weicher Trennstrich?). pdfplumber gibt ihn nicht heraus, deshalb wird die Seite
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
            self.modi: list[tuple[int, bool]] = []
            self._modus = 0

        def render_string(self, textstate: Any, seq: Any, ncs: Any, graphicstate: Any) -> None:
            self._modus = int(getattr(textstate, "render", 0) or 0)
            super().render_string(textstate, seq, ncs, graphicstate)

        def render_char(self, *a: Any, **kw: Any) -> float:   # type: ignore[override]
            adv = super().render_char(*a, **kw)
            weich = False
            try:
                font, cid = a[1], a[5]
                # Code 173 (weicher Trennstrich) liest pdfminer als Leerzeichen; in Schriften mit festem Code ist er das
                weich = cid == 173 and not font.is_multibyte() and not str(font.to_unichr(cid)).strip()
            except Exception:
                pass
            self.modi.append((self._modus, weich))
            return adv

        def paint_path(self, gstate: Any, *a: Any, **kw: Any) -> None:   # type: ignore[override]
            # R6: Farbraum der Linien/Flaechen merken (pdfminer gibt ihn nur fuer Zeichen heraus; pdfplumber uebernimmt
            # die Attribute ncs/scs, wenn sie am Objekt stehen)
            vorher = len(getattr(self.cur_item, "_objs", []))
            super().paint_path(gstate, *a, **kw)
            try:
                for obj in self.cur_item._objs[vorher:]:
                    obj.ncs, obj.scs = gstate.ncs, gstate.scs
            except Exception:
                pass

    try:
        if hasattr(page, "_layout"):
            return None
        geraet = _Geraet(page.pdf.rsrcmgr, pageno=page.page_number, laparams=page.pdf.laparams)
        PDFPageInterpreter(page.pdf.rsrcmgr, geraet).process_page(page.page_obj)
        page._layout = geraet.get_result()
        return geraet.modi
    except Exception:
        return None


class _Sonderfarben:
    """Rechnet Sonderfarben (Separation / DeviceN mit einer Tinte, z. B. Firmenfarben) in RGB um.

    pdfplumber nennt nur den Tintenwert (1.0 = VOLLE Farbe, nicht Weiss!), nicht den Farbraum-Namen. Deshalb werden
    die Sonderfarb-Raeume der Seite ausgewertet; ist das Ergebnis nicht eindeutig, bleibt die Farbe 'unklar'."""

    def __init__(self, page: Any) -> None:
        self.raeume: list[Any] = []
        self._cache: dict[tuple, tuple[float, float, float] | None] = {}
        try:
            from pdfminer.pdftypes import resolve1
            self._resolve = resolve1
            res = resolve1(page.page_obj.resources) or {}
            farbraeume = resolve1(res.get("ColorSpace")) or {}
            for spez in farbraeume.values():
                spez = resolve1(spez)
                if isinstance(spez, list) and len(spez) >= 4 and getattr(resolve1(spez[0]), "name", "") in ("Separation", "DeviceN"):
                    self.raeume.append(spez)
        except Exception:
            self.raeume = []

    def __call__(self, werte: tuple[float, ...]) -> tuple[float, float, float] | None:
        if werte not in self._cache:
            ergebnisse = set()
            for spez in self.raeume:
                rgb = self._rgb(spez, werte)
                if rgb is None:
                    ergebnisse.add(None)
                else:
                    ergebnisse.add(tuple(round(v, 2) for v in rgb))
            self._cache[werte] = ergebnisse.pop() if len(ergebnisse) == 1 else None
        return self._cache[werte]

    def _rgb(self, spez: list, werte: tuple[float, ...]) -> tuple[float, float, float] | None:
        try:
            r = self._resolve
            if len(werte) != (1 if getattr(r(spez[0]), "name", "") == "Separation" else len(r(spez[1]))):
                return None
            fn = r(spez[3])
            if not isinstance(fn, dict) or r(fn.get("FunctionType")) != 2 or len(werte) != 1:
                return None                                  # nur die einfache Tintenfunktion (Typ 2) wird berechnet
            c0 = [float(r(x)) for x in (r(fn.get("C0")) or [0.0])]
            c1 = [float(r(x)) for x in (r(fn.get("C1")) or [1.0])]
            n = float(r(fn.get("N")) if fn.get("N") is not None else 1.0)
            t = max(0.0, min(1.0, werte[0]))
            return _rgb(tuple(a + (t ** n) * (b - a) for a, b in zip(c0, c1)))
        except Exception:
            return None


def _zeichenfarbe(c: dict, sonder: Any = None) -> tuple[tuple[float, float, float] | None, bool]:
    """(RGB oder None, unklar). 'unklar' = Farbe aus Muster/Sonderfarbraum, die nicht umgerechnet werden kann."""
    return _farbe(c.get("non_stroking_color"), c.get("ncs"), sonder)


def _objektfarbe(o: dict, strich: bool, sonder: Any = None) -> tuple[tuple[float, float, float] | None, bool]:
    """Farbe einer Linie/Flaeche MIT Farbraum (R6), wie bei Zeichen: (RGB oder None, unklar).
    strich=True: Linienfarbe (stroking), sonst Fuellfarbe. Ohne bekannten Farbraum wie frueher (Grau/RGB/CMYK nach Anzahl)."""
    if strich:
        return _farbe(o.get("stroking_color"), o.get("scs"), sonder)
    return _farbe(o.get("non_stroking_color"), o.get("ncs"), sonder)


def _farbe(farbe: Any, raum_roh: Any, sonder: Any = None) -> tuple[tuple[float, float, float] | None, bool]:
    if farbe is None:
        return None, False
    if isinstance(farbe, str):
        return None, True
    if isinstance(farbe, (int, float)):
        farbe = (farbe,)
    try:
        werte = tuple(float(x) for x in farbe)
    except (TypeError, ValueError):
        return None, True
    raum = str(raum_roh or "")
    if len(werte) == 1 and raum in ("Separation", "DeviceN", "Indexed", "Pattern", "Lab"):
        rgb = sonder(werte) if (sonder is not None and raum in ("Separation", "DeviceN")) else None
        if rgb is not None:
            return rgb, False
        return None, werte[0] > 0.05          # Tinte 0 = kein Farbauftrag; sonst Farbe, die wir nicht umrechnen koennen
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
    # Schreibrichtung: Steht der Text (fast) ueberall quer oder kopfueber, wird die Seite in Leserichtung gebracht.
    # Zeichen in einer anderen Richtung (Randvermerk, Wasserzeichen) werden nicht gelesen, aber gemeldet.
    sonder = _Sonderfarben(page)
    richtungen = [_richtung(c) for c in roh]
    zaehl = Counter(r for r, c in zip(richtungen, roh) if (c.get("text") or "").strip())
    haupt = zaehl.most_common(1)[0][0] if zaehl else "h"
    rahmen = _Rahmen(haupt if haupt in ("ab", "auf", "um") else "h", sd.breite, sd.hoehe)
    sd.rahmen = rahmen
    sd.breite, sd.hoehe = rahmen.masse
    fremd_wort_ende = True
    for i, c in enumerate(roh):
        text = c.get("text") or ""
        leer = not text.strip()
        if haupt != "schraeg" and richtungen[i] != haupt:
            # Zeichen einer anderen Schreibrichtung: nur zaehlen
            if leer:
                fremd_wort_ende = True
            else:
                if fremd_wort_ende:
                    sd.fremd_woerter += 1
                fremd_wort_ende = False
                rgb_f, _ = _zeichenfarbe(c, sonder)
                if _bedeutung(klasse(rgb_f, cfg), cfg):
                    sd.fremd_farbig = True
            continue
        rgb, unklar = _zeichenfarbe(c, sonder)
        x0, top, x1, bot = rahmen.box(c["x0"], c["top"], c["x1"], c["bottom"])
        z = _Z(text=text, x0=x0, x1=x1, top=top, bottom=bot,
               size=float(c.get("size") or 10), font=str(c.get("fontname") or ""), leer=leer, nr=i)
        z.unsichtbar = bool(modi and modi[i][0] in (3, 7))
        z.weich = bool(modi and modi[i][1])
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
        bx0, btop, bx1, bbot = _kasten(im, rahmen)
        x0, x1 = max(0.0, bx0), min(w, bx1)
        top, bot = max(0.0, btop), min(h, bbot)
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
        x0, top, x1, bot = _kasten(l, sd.rahmen)
        dx, dy = x1 - x0, bot - top
        if dy <= 1.0 and dx >= lang:
            wag.append(((top + bot) / 2, x0, x1))
        elif dx <= 1.0 and dy >= hoch:
            senk.append(((x0 + x1) / 2, top, bot))
    for r in page.rects:
        x0, top, x1, bot = _kasten(r, sd.rahmen)
        bw, bh = x1 - x0, bot - top
        if bh <= 1.6 and bw >= lang:
            wag.append(((top + bot) / 2, x0, x1))
        elif bw <= 1.6 and bh >= hoch:
            senk.append(((x0 + x1) / 2, top, bot))
        elif r.get("stroke") and bw >= hoch and bh >= hoch:       # Zelle mit Rahmen
            wag += [(top, x0, x1), (bot, x0, x1)]
            senk += [(x0, top, bot), (x1, top, bot)]
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
# Kommentar-Typen, die Text markieren oder anmerken koennten, hier aber nicht ausgewertet werden
_ANDERE_KOMMENTARE = ("Underline", "Squiggly", "Caret", "Ink", "FreeText", "Stamp", "Square", "Circle", "Line",
                      "Polygon", "PolyLine", "Redact")


def _drehe_punkt(px: float, py: float, rot: int, w: float, h: float) -> tuple[float, float]:
    """Punkt aus dem Seitenraum (PDF-Koordinaten, y nach oben) in die angezeigte Seite (x, top) -- wie pdfplumber."""
    turns = (rot // 90) % 4
    for i in range(turns):
        comp = w if i == turns % 2 else h
        px, py = py, comp - px
    return px, h - py


class _Markierer:
    """Sammelt pro Seite alle Markierungs-Signale und beantwortet: Status je Zeichen.

    Ablauf: __init__ sammelt die rohen Signale, pruefe() entscheidet mit Hilfe der Zeichen, welche davon wirklich
    Markierungen sind (und welche Gestaltung: Zierlinie, Tabellenkopf, Hyperlink), status() bewertet ein Zeichen."""

    def __init__(self, page: Any, cfg: dict, rahmen: _Rahmen) -> None:
        self.cfg = cfg
        self.rahmen = rahmen
        self.flaechen: list[tuple[float, float, float, float, str]] = []       # Hinterlegungen = Markierung
        self.unb_flaechen: list[tuple[float, float, float, float, str]] = []   # nicht zugeordnete Hinterlegungen
        self.design: list[tuple[float, float, float, float]] = []              # Hinterlegungen = Gestaltung
        self.striche: list[tuple[float, float, float]] = []
        self.unterstriche: list[tuple[float, float, float]] = []
        self.annot: list[tuple[float, float, float, float, str]] = []
        self.unb_annot: list[tuple[float, float, float, float, str]] = []
        self.links: list[tuple[float, float, float, float]] = []
        self.andere: list[dict[str, Any]] = []          # Kommentare anderer Art: {"typ", "boxen", "treffer"}
        self.rect_farben: Counter = Counter()
        self._flaechen_roh: list[tuple[float, float, float, float, str | None, str]] = []   # (Box, Bedeutung, Farbe)
        # (x0, x1, y, hell, lang, farbe_unklar); lang = breiter als durchstreichung.max_breite_anteil der Seite
        self._linien_roh: list[tuple[float, float, float, bool, bool, bool]] = []
        self.sonder = _Sonderfarben(page)
        w, h = float(page.width), float(page.height)
        bw, bh = rahmen.masse
        if cfg["hintergrund"]["aktiv"]:
            hg = cfg["hintergrund"]
            for r in page.rects:
                if not r.get("fill"):
                    continue
                rgb, unklar = _objektfarbe(r, False, self.sonder)     # R6: Farbraum beachten (Sonderfarbe Ton 1,0 != weiss)
                if rgb is None and not unklar:
                    continue
                if r["width"] * r["height"] > hg["max_flaeche_anteil"] * w * h:
                    continue
                box = _kasten(r, rahmen)
                if box[3] - box[1] <= 1.6:     # duenne Rechtecke sind Linien (siehe unten)
                    continue
                if rgb is None:
                    # Farbe nicht berechenbar: nicht still uebergehen, sondern wie eine unbekannte Hinterlegung melden
                    self.rect_farben[("sonderfarbe", None)] += 1
                    self._flaechen_roh.append(box + (None, "sonderfarbe"))
                    continue
                k = klasse(rgb, cfg, hintergrund=True)
                self.rect_farben[(_hex(rgb), k)] += 1
                bed = _bedeutung(k, cfg)
                if bed or (k is None and self._kraeftig_hg(rgb)):
                    self._flaechen_roh.append(box + (bed, _hex(rgb)))
        d = cfg["durchstreichung"]
        if d["aktiv"]:
            # R5: keine Breitenkappe mehr fuer Durchstreichungen (die Kappe gilt nur noch fuer Unterstreichungen, siehe pruefe)
            kappe = d["max_breite_anteil"] * bw
            for l in page.lines:
                x0, top, x1, bot = _kasten(l, rahmen)
                if (l.get("linewidth", 1) or 0) <= d["max_linienstaerke_pt"] and abs(bot - top) <= 1.2 and x1 - x0 >= 2:
                    hell, unklar = self._hell(*_objektfarbe(l, True, self.sonder))
                    self._linien_roh.append((x0, x1, (top + bot) / 2, hell, x1 - x0 >= kappe, unklar))
            for r in page.rects:
                x0, top, x1, bot = _kasten(r, rahmen)
                if bot - top <= 1.6 and 2 <= x1 - x0:
                    hell, unklar = self._hell(*_objektfarbe(r, not r.get("fill"), self.sonder))
                    self._linien_roh.append((x0, x1, (top + bot) / 2, hell, x1 - x0 >= kappe, unklar))
            # flache Kurven/Pfade (z. B. als Polylinie gezeichnete Striche) sind Linienkandidaten
            for c in page.curves:
                x0, top, x1, bot = _kasten(c, rahmen)
                if bot - top <= d.get("kurve_max_hoehe_pt", 1.2) and x1 - x0 >= 2 \
                        and (c.get("linewidth", 1) or 0) <= d["max_linienstaerke_pt"]:
                    hell, unklar = self._hell(*_objektfarbe(c, bool(c.get("stroke")), self.sonder))
                    self._linien_roh.append((x0, x1, (top + bot) / 2, hell, x1 - x0 >= kappe, unklar))
        if cfg["annotationen"]["aktiv"]:
            self._annotationen(page)

    def _kraeftig_hg(self, rgb: tuple[float, float, float]) -> bool:
        hg = self.cfg["hintergrund"]
        _, s, v = colorsys.rgb_to_hsv(*(max(0.0, min(1.0, x)) for x in rgb))
        return s >= hg["min_saettigung"] and v >= hg["min_helligkeit"]

    def _hell(self, rgb: tuple[float, float, float] | None, unklar: bool = False) -> tuple[bool, bool]:
        """(hell, farbe_unklar): Hellgraue (neutrale, helle) Linie? Das ist eine Zierlinie, keine Durchstreichung.
        farbe_unklar: die Farbe (Sonderfarbraum) konnte nicht berechnet werden."""
        if rgb is None:
            return False, unklar
        d = self.cfg["durchstreichung"]
        _, s, v = colorsys.rgb_to_hsv(*(max(0.0, min(1.0, x)) for x in rgb))
        return s <= d["hell_max_saettigung"] and v >= d["hell_min_helligkeit"], False

    def _boxen(self, a: dict, data: dict, page: Any) -> list[tuple[float, float, float, float]]:
        """Bereiche eines Kommentars in Leserichtung (aus QuadPoints, sonst aus dem Rechteck)."""
        quad = data.get("QuadPoints")
        try:
            q = [float(x) for x in quad] if quad else []
        except (TypeError, ValueError):
            q = []
        boxen = []
        if len(q) >= 8:
            rot = int(getattr(page, "rotation", 0) or 0)
            w, h = float(page.width), float(page.height)
            for i in range(0, len(q) - 7, 8):
                pts = [_drehe_punkt(q[i + j], q[i + j + 1], rot, w, h) for j in range(0, 8, 2)]
                xs, ts = [p[0] for p in pts], [p[1] for p in pts]
                boxen.append(self.rahmen.box(min(xs), min(ts), max(xs), max(ts)))
        else:
            boxen.append(_kasten(a, self.rahmen))
        return boxen

    def _annotationen(self, page: Any) -> None:
        cfg = self.cfg
        for a in page.annots:
            data = a.get("data") or {}
            sub = data.get("Subtype")
            name = getattr(sub, "name", str(sub))
            if name == "Link":
                self.links.append(_kasten(a, self.rahmen))
                continue
            if name in _ANDERE_KOMMENTARE:
                self.andere.append({"typ": name, "boxen": self._boxen(a, data, page), "treffer": 0})
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
            for x0, top, x1, bot in self._boxen(a, data, page):
                if bed:
                    self.annot.append((x0, top, x1, bot, bed))
                else:
                    self.unb_annot.append((x0, top, x1, bot, unb or "-"))

    # ---- Entscheidung: Markierung oder Gestaltung? ----------------------------------------------
    def pruefe(self, zeichen: list[_Z]) -> None:
        """Entscheidet mit Hilfe der Zeichen der Seite, welche Linien und Flaechen echte Markierungen sind."""
        cfg = self.cfg
        d = cfg["durchstreichung"]
        streichen = _bedeutung("durchgestrichen", cfg) is not None
        def gedeckt(c: _Z, x0: float, x1: float) -> bool:
            return (min(x1, c.x1) - max(x0, c.x0)) >= d["min_abdeckung"] * max(c.x1 - c.x0, 0.1)

        for x0, x1, y, hell, lang, farbe_unklar in self._linien_roh:
            for art, band in (("strich", d["band"]), ("unter", d["unterstreichung_band"])):
                if art == "strich" and not streichen:
                    continue
                if art == "unter" and lang:
                    continue                    # lange Linien unter dem Text: Trennlinie/Rahmen, keine Unterstreichung
                treffer = [c for c in zeichen
                           if c.top + band[0] * (c.bottom - c.top) <= y <= c.top + band[1] * (c.bottom - c.top)
                           and min(x1, c.x1) - max(x0, c.x0) > 0]
                if not treffer:
                    continue
                if art == "strich":
                    # R5: die Linie zaehlt nur, wenn sie die Zeichen unter ihr wirklich im Durchstreich-Band kreuzt
                    # (Anteil min_kreuz_anteil), unabhaengig von ihrer Laenge
                    unter = [c for c in zeichen if not c.leer and c.top <= y <= c.bottom
                             and min(x1, c.x1) - max(x0, c.x0) > 0]
                    gekreuzt = [c for c in treffer if not c.leer]
                    if unter and len(gekreuzt) < d.get("min_kreuz_anteil", 0.6) * len(unter):
                        for c in gekreuzt:
                            if gedeckt(c, x0, x1):
                                c.hinweis.add("linie_unklar")
                        continue
                size = statistics.median(c.size for c in treffer)
                ueber = max(0.0, min(c.x0 for c in treffer) - x0) + max(0.0, x1 - max(c.x1 for c in treffer))
                if hell or ueber > d["max_ueberstand_em"] * size:
                    if art == "strich":
                        # Zierlinie quer durch den Text: nicht als Durchstreichung werten, aber nicht verschweigen
                        for c in treffer:
                            if not c.leer and gedeckt(c, x0, x1):
                                c.hinweis.add("linie_unklar")
                    continue
                if art == "strich" and farbe_unklar:
                    # R6: Linienfarbe nicht berechenbar (Sonderfarbraum): als Durchstreichung werten, aber melden
                    for c in treffer:
                        if not c.leer and gedeckt(c, x0, x1):
                            c.hinweis.add("linienfarbe_unklar")
                (self.striche if art == "strich" else self.unterstriche).append((x0, x1, y))
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

    def _linie_trifft(self, linien: list[tuple[float, float, float]], c: _Z, band: list[float]) -> bool:
        d = self.cfg["durchstreichung"]
        h = c.bottom - c.top
        lo, hi = c.top + band[0] * h, c.top + band[1] * h
        breite = max(c.x1 - c.x0, 0.1)
        return any(lo <= y <= hi and (min(x1, c.x1) - max(x0, c.x0)) / breite >= d["min_abdeckung"]
                   for x0, x1, y in linien)

    def status(self, c: _Z) -> str:
        cfg = self.cfg
        quelle: dict[str, set[str]] = {}          # Bedeutung -> woher
        farbe_bed = c.farbe_bed
        cx, cy = (c.x0 + c.x1) / 2, (c.top + c.bottom) / 2
        in_link = False
        if not c.leer:
            for x0, top, x1, bot in self.links:
                if x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bot + 1:
                    in_link = True
                    break
        if in_link and farbe_bed == "neu" and c.farbklasse == "blau":
            farbe_bed = None                        # Hyperlink: das Blau ist Gestaltung, keine Neu-Markierung
            c.hinweis.add("link")
        if farbe_bed:
            quelle.setdefault(farbe_bed, set()).add("farbe")
        for x0, top, x1, bot, b in self.flaechen:
            if x0 - 0.5 <= cx <= x1 + 0.5 and top - 0.5 <= cy <= bot + 0.5:
                quelle.setdefault(b, set()).add("flaeche")
        for x0, top, x1, bot, b in self.annot:
            if x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bot + 1:
                quelle.setdefault(b, set()).add("annot")
        if self.striche and self._linie_trifft(self.striche, c, cfg["durchstreichung"]["band"]):
            quelle.setdefault(_bedeutung("durchgestrichen", cfg) or "entfaellt", set()).add("strich")
        if "entfaellt" in quelle:
            erg = DEL
        elif "neu" in quelle:
            erg = ADD
        else:
            erg = KEEP
        gewinner = quelle.get("entfaellt" if erg == DEL else "neu", set())
        c.unscharf = erg != KEEP and "farbe" not in gewinner
        if "neu" in quelle and "entfaellt" in quelle and not c.leer:
            c.hinweis.add("gemischt_blau_strich" if "farbe" in quelle["neu"] and "strich" in quelle["entfaellt"]
                          else "gemischt_widerspruch")
        if not c.leer:
            for k in self.andere:
                if any(x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bot + 1 for x0, top, x1, bot in k["boxen"]):
                    k["treffer"] += 1
                    c.hinweis.add("kommentar_" + k["typ"].lower())
            if erg == KEEP:
                if self.unterstriche and not in_link and self._linie_trifft(self.unterstriche, c,
                                                                            cfg["durchstreichung"]["unterstreichung_band"]):
                    c.hinweis.add("unterstrichen")
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

    def meldungen(self, seite: int) -> list[str]:
        """Kommentare, die auf keinem erkannten Text liegen."""
        return [f"Seite {seite}: Ein PDF-Kommentar vom Typ {k['typ']} liegt nicht auf erkanntem Text und wurde nicht ausgewertet."
                for k in self.andere if not k["treffer"]]


def _zeichenstatus(g: list[_Z], anteil: float) -> tuple[list[str], bool]:
    """Status je Zeichen eines Wortes (fuer old/new/status des Tokens, Verhalten wie bisher) und ob die Markierung nur
    aus Geometrie stammt (Linie, Flaeche, Kommentar). Der Status jedes Zeichens bleibt erhalten, damit nichts verloren geht:
    'Netzleistung' (gestrichen) + 'Leistung' (neu) ohne Leerzeichen bleibt ein Wort mit altem UND neuem Text,
    'ab' + 'gesichert' (neu) behaelt das unveraenderte 'ab'.
    Unmarkierte Satzzeichen am markierten Wort und (nur bei Geometrie) ein einzelner unmarkierter Randbuchstabe werden dem
    Rest angepasst, weil Linien und Flaechen oft knapp daneben liegen. Die Zeichen-Laeufe (`PdfToken.teile`) uebernehmen
    diese Glaettung nur bei Geometrie (siehe _woerter)."""
    n = len(g)
    st = [z.status for z in g]
    n_del = st.count(DEL)
    n_add = st.count(ADD)
    if n_del and n_add:
        return st, False               # Ersetzung ohne Leerzeichen: Zeichen einzeln behalten
    marke = DEL if n_del else ADD if n_add else None
    geometrie = marke is not None and all(z.unscharf for z in g if z.status == marke)
    if marke is None or (n_del + n_add) / n < anteil:
        return st, geometrie
    behalten = [i for i, z in enumerate(g) if z.status == KEEP]
    buchstaben = [i for i in behalten if g[i].text[:1].isalnum()]
    knapp_daneben = len(buchstaben) == 1 and buchstaben[0] in (0, n - 1) and geometrie
    if not buchstaben or knapp_daneben:
        return [marke] * n, geometrie
    return st, geometrie


_GESCHUETZT = ("\xa0", " ", " ")      # geschuetzte Leerzeichen (auch schmal / Ziffernbreite)


def _teile(g: list[_Z], st: list[str]) -> list[tuple[str, str]]:
    """Zeichen-Laeufe [(status, text)] eines Wortes; leer, wenn das ganze Wort denselben Status hat."""
    if len(set(st)) <= 1:
        return []
    out: list[tuple[str, str]] = []
    for z, s in zip(g, st):
        if out and out[-1][0] == s:
            out[-1] = (s, out[-1][1] + z.text)
        else:
            out.append((s, z.text))
    return out


def _woerter(zeichen: list[_Z], seite: int, cfg: dict, grund_hex: str = "",
             leere: list[_Z] | None = None, hoch: list[_Z] | None = None) -> tuple[list[PdfToken], list[_Info]]:
    """Zeichen einer Zeile -> Woerter (mit Statusanteilen).

    R1: `leere` = die Leerraum-Zeichen der Zeile (Leerzeichen, geschuetztes Leerzeichen, Tab). Steht eines zwischen zwei
    Zeichen, trennt es die Woerter auch bei kleiner Luecke (Vorrang vor der Schwelle, z. B. enger Blocksatz) und bestimmt
    den Trenner, sofern ueberhaupt etwas Platz ist (layout.leerzeichen_min_luecke_em). Ohne Leerraum-Zeichen entscheidet
    die Luecke (layout.wort_luecke). (Eine Luecke ohne Zeichen trennt weiter:
    viele PDFs setzen Woerter einzeln, auch in Zeilen, die sonst Leerzeichen enthalten.)"""
    lay = cfg["layout"]
    zeichen = sorted(zeichen, key=lambda z: z.x0)
    leer = sorted((z for z in (leere or []) if not z.weich and z.text), key=lambda z: z.x0)
    mitten = [(z.x0 + z.x1) / 2 for z in leer]

    min_luecke = lay.get("leerzeichen_min_luecke_em", 0.05)

    def dazwischen(a: _Z, b: _Z) -> list[_Z]:
        # Leerzeichen-Zeichen zwischen a und b: Mitte zwischen den Mitten, es beginnt nicht im Zeichen davor (ein
        # Leerzeichen, das das Zeichen davor ueberdeckt, ist ein Phantom ohne Vorschub) und es steht auch in der
        # Lesereihenfolge des Inhaltsstroms zwischen a und b (LibreOffice zeichnet Phantome hinter dem Trennstrich)
        lo, hi = (a.x0 + a.x1) / 2, (b.x0 + b.x1) / 2
        n_lo, n_hi = min(a.nr, b.nr), max(a.nr, b.nr)
        return [z for z, m in zip(leer, mitten) if lo < m < hi and z.x0 >= a.x1 - min_luecke * a.size
                and (z.nr < 0 or n_lo < z.nr < n_hi)]

    gruppen: list[list[_Z]] = []
    trenner: list[str] = []          # Trenner HINTER der Gruppe gleichen Index
    for z in zeichen:
        if gruppen:
            a = gruppen[-1][-1]
            luecke = z.x0 - a.x1
            # Ein Leerzeichen-Zeichen zaehlt nur, wenn zwischen den Zeichen auch wirklich Platz ist
            # (layout.leerzeichen_min_luecke_em): LibreOffice setzt an weichen Trennstellen Leerzeichen ohne Vorschub
            # (Phantome mitten im Wort, teils ausser der Reihe), die das Wort nicht trennen duerfen.
            zw = dazwischen(a, z) if (leer and luecke >= min_luecke * z.size) else []
            if not zw and luecke <= lay["wort_luecke"] * z.size:
                gruppen[-1].append(z)
                continue
            texte = "".join(x.text for x in zw)
            if any(c in texte for c in _GESCHUETZT):
                trenner.append("\xa0")
            elif "\t" in texte:
                trenner.append("\t")
            else:
                trenner.append(" ")
        gruppen.append([z])
    trenner.append(" ")              # am Zeilenende: Leerzeichen (zusammengefuegte Trennungen setzen das selbst)
    tokens = []
    infos = []
    for gi, g in enumerate(gruppen):
        n = len(g)
        roh = [z.status for z in g]
        st, geometrie = _zeichenstatus(g, cfg["durchstreichung"]["wort_anteil"])
        # R2: Zeichen-Laeufe. Die Farbe ist je Zeichen eindeutig -> roher Status je Zeichen. Geglaettet wird nur bei
        # Markierungen aus Geometrie (Linie, Flaeche, Kommentar); hat die Glaettung etwas umentschieden: rand_unklar.
        teile = _teile(g, st if geometrie else roh)
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
        if geometrie and st != roh:
            hin.add("rand_unklar")
        # R3: markierte Hochzahl (Fussnoten-Marke), die zu diesem Wort gehoert: nicht im Text, aber gemeldet
        naechstes_x0 = gruppen[gi + 1][0].x0 if gi + 1 < len(gruppen) else float("inf")
        if any(h.status != KEEP and (gi == 0 or h.x0 >= g[0].x0) and h.x0 < naechstes_x0 for h in hoch or ()):
            hin.add("hochzahl_markiert")
        tokens.append(PdfToken(old=old, new=new, status=status, page=seite, bbox=bbox, hinweis=";".join(sorted(hin)),
                               trenner=trenner[gi], teile=teile))
        infos.append(info)
    return tokens, infos


def _hochzahlen(cluster: list[list[_Z]], lay: dict) -> tuple[list[list[_Z]], list[list[_Z]]]:
    """R3: Hochzahlen (Fussnoten-Marken): Zeichen deutlich kleiner als der Median der Zeile (layout.hochzahl_faktor) UND
    angehoben (Unterkante mindestens layout.hochzahl_anhebung Schriftgroessen ueber der Grundlinie der Zeile).
    Sie gehoeren nicht in den Wort-Text. Gibt (Zeilen ohne Hochzahlen, Hochzahlen je Zeile) zurueck. Eine eigene
    kleine 'Zeile' aus Hochzahlen (Mitte zu weit weg fuer dieselbe Zeile) wird der Zeile darunter zugeordnet."""
    faktor = lay.get("hochzahl_faktor", 0.75)
    anhebung = lay.get("hochzahl_anhebung", 0.15)
    max_zeichen = lay.get("hochzahl_max_zeichen", 6)
    infos = []
    for cl in cluster:
        med = statistics.median(c.size for c in cl)
        haupt = [c for c in cl if c.size >= faktor * med]
        basis = statistics.median(c.bottom for c in haupt) if haupt else None
        infos.append((med, basis, min(c.top for c in cl), min(c.x0 for c in cl), max(c.x1 for c in cl)))
    rest: list[list[_Z]] = []
    hoch: list[list[_Z]] = []
    for cl, (med, basis, _, _, _) in zip(cluster, infos):
        h = [c for c in cl if basis is not None and c.size < faktor * med and c.bottom <= basis - anhebung * med]
        bleibt = [c for c in cl if not any(c is x for x in h)]
        rest.append(bleibt if bleibt else cl)
        hoch.append(h if bleibt else [])
    # ganze kleine Cluster (nur Hochzahlen) an die Zeile darunter haengen
    for i, cl in enumerate(rest):
        if not cl or len(cl) > max_zeichen:
            continue
        gross = max(c.size for c in cl)
        for j, (med, basis, top, x0, x1) in enumerate(infos):
            if j == i or basis is None or not rest[j] or gross >= faktor * med:
                continue
            if all(top <= c.bottom <= basis - anhebung * med and x0 - med <= c.x0 and c.x1 <= x1 + med for c in cl):
                hoch[j].extend(cl)
                rest[i] = []
                break
    return rest, hoch


def _zeilen(zeichen: list[_Z], seite: int, cfg: dict, grund_hex: str = "") -> list[_Zeile]:
    lay = cfg["layout"]
    tol = lay["zeilen_toleranz"]
    weiche = [z for z in zeichen if z.weich]
    leere = [z for z in zeichen if z.leer and not z.weich and not z.unsichtbar]
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
    cluster, hoch_je = _hochzahlen(cluster, lay) if cluster else ([], [])
    zeilen = []
    for cl, hoch in zip(cluster, hoch_je):
        if not cl:
            continue
        c_top, c_bot = min(z.top for z in cl), max(z.bottom for z in cl)
        c_x0, c_x1 = min(z.x0 for z in cl), max(z.x1 for z in cl)
        leer_cl = [z for z in leere if c_top <= (z.top + z.bottom) / 2 <= c_bot and c_x0 <= (z.x0 + z.x1) / 2 <= c_x1]
        toks, infos = _woerter(cl, seite, cfg, grund_hex, leer_cl, hoch)
        if not toks:
            continue
        t0 = toks[0]
        if t0.status == KEEP and _NUR_PRIVAT.fullmatch(t0.old):      # Symbolschrift-Aufzaehlungspunkt am Zeilenanfang = Aufzaehlungszeichen
            t0.old = t0.new = "•"
            t0.teile = []
            t0.hinweis = ";".join(h for h in t0.hinweis.split(";") if h and h != "nicht_lesbar")
            infos[0].unlesbar = False
        fett_n = sum(1 for z in cl if re.search(r"bold|black|heavy|semibold", z.font, re.I))
        zeilen.append(_Zeile(
            tokens=toks, infos=infos, top=min(z.top for z in cl), bottom=max(z.bottom for z in cl),
            x0=min(z.x0 for z in cl), x1=max(z.x1 for z in cl),
            size=statistics.median(z.size for z in cl), fett=fett_n / len(cl) > 0.6, seite=seite,
            unsichtbar=sum(1 for z in cl if z.unsichtbar or z.weiss) * 2 > len(cl),
            weich_ende=any(min(c.top for c in cl) <= (w.top + w.bottom) / 2 <= max(c.bottom for c in cl)
                           and w.x0 >= max(c.x1 for c in cl) - 1 for w in weiche)))
    zeilen.sort(key=lambda z: z.top)
    return zeilen


# ---------------------------------------------------------------- Bloecke
_PUNKTE = "•·▪▫■□●○◦‣⁃◆◇➢➤►▶✓✔→*"
_RE_PUNKT = re.compile(f"^[{re.escape(_PUNKTE)}]$")
_RE_STRICH = re.compile(r"^[-–—]$")
_RE_ZAHL_PUNKT = re.compile(r"^(\d{1,2})\.$")
_RE_KLAMMER = re.compile(r"^(\(?\d{1,2}\)|\(?[A-Za-z]\)|\(?(?:i{1,3}|iv|vi{0,3}|ix|x)\))$")
_RE_ROEMISCH = re.compile(r"^(?:i{1,3}|iv|vi{0,3}|ix|x)\.$")


def _marke(text: str) -> tuple[str, int] | None:
    """Aufzaehlungszeichen am Zeilenanfang: ('punkt' | 'strich' | 'zahl' | 'klammer' | 'roemisch', Nummer) oder None."""
    if _RE_PUNKT.match(text):
        return "punkt", 0
    if _RE_STRICH.match(text):
        return "strich", 0
    m = _RE_ZAHL_PUNKT.match(text)
    if m:
        return "zahl", int(m.group(1))
    if _RE_KLAMMER.match(text):
        return "klammer", 0
    if _RE_ROEMISCH.match(text):
        return "roemisch", 0
    return None


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
    abstaende: list[float] = []           # Zeilenabstaende im Fliesstext (Zeilen in Fliesstext-Schriftgroesse)
    alle_abstaende: list[float] = []      # Ersatz, falls es keine zwei Fliesstext-Zeilen untereinander gibt
    weich_ids = {id(z.tokens[-1]) for zl in alle_zeilen for z in zl if z.weich_ende}
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
        for a, b in zip(nutz, nutz[1:]):
            if 0.9 * a.size < b.top - a.top < 3 * a.size:
                alle_abstaende.append(b.top - a.top)
                if abs(a.size - body) <= 0.5 and abs(b.size - body) <= 0.5:
                    abstaende.append(b.top - a.top)
    _design_ueberschriften(seiten_nutz, body, cfg, ergebnis)
    median = _zeilenabstand(abstaende or alle_abstaende, body)
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

    def kurz_im_text(z: _Zeile) -> bool:        # deutlich kuerzer als die langen Zeilen des Textes
        return rechts is not None and z.x1 < rechts - 3 * z.size

    def kurz(z: _Zeile) -> bool:
        return kurz_im_text(z) or (schwelle_kurz is not None and z.x1 < schwelle_kurz)

    # R4: "Haette das erste Wort der naechsten Zeile noch in die vorige gepasst?" Rechter Rand je Zeile: 90-%-Wert der
    # Zeilenenden langer Zeilen DERSELBEN Seite, die waagerecht mit ihr ueberlappen (Spalten bleiben getrennt).
    seiten_zeilen: dict[int, list[_Zeile]] = {}
    for zl in seiten_nutz:
        for z in zl:
            seiten_zeilen.setdefault(z.seite, []).append(z)
    rand_cache: dict[int, float | None] = {}

    def rechter_rand(z: _Zeile) -> float | None:
        if id(z) not in rand_cache:
            ref = sorted(l.x1 for l in seiten_zeilen.get(z.seite, []) if len(l.tokens) >= 5 and l.x0 < z.x1 and l.x1 > z.x0)
            rand_cache[id(z)] = ref[int(0.9 * (len(ref) - 1))] if len(ref) >= lay.get("absatz_passt_min_zeilen", 3) else None
        return rand_cache[id(z)]

    def passt(vorher: _Zeile, tok: PdfToken) -> str:
        """'ja' (haette sicher gepasst -> harter Umbruch) | 'nein' | 'unklar' (knapp) | '' (unbekannt: kein Rand)."""
        r = rechter_rand(vorher)
        if r is None:
            return ""
        if _erste_form(vorher.tokens[-1]).endswith("-") or vorher.weich_ende:
            return "nein"                          # Silbentrennung: die Zeile ist voll
        rest = (r - vorher.x1) - ((tok.bbox[2] - tok.bbox[0]) + lay.get("leerzeichen_em", 0.28) * vorher.size)
        if rest >= lay.get("absatz_passt_toleranz_em", 0.5) * vorher.size:
            return "ja"
        return "unklar" if rest > 0 else "nein"

    wortschatz = {norm(x) for zl in alle_zeilen for z in zl for t in z.tokens for x in (t.old, t.new) if x}
    block_zeilen: dict[int, tuple[_Zeile, _Zeile]] = {}      # id(Block) -> (erste, letzte Zeile)
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
                n_alt = next((section_nummer(t.old) for t in toks if t.old), None)      # '{-2.3-} {+2.2+} Titel'
                n_neu = next((section_nummer(t.new) for t in toks if t.new), None)
                if n_alt or n_neu:
                    sec_alt, sec_neu = n_alt, n_neu
                else:   # unnummerierte Ueberschrift: Zuordnung ueber Text, Nummern bleiben
                    sec_alt = sec_neu = None
            else:
                toks = _silben(toks, aktuell, cfg, weich_ids, wortschatz)
            bbox = (min(z.x0 for z in aktuell), min(z.top for z in aktuell),
                    max(z.x1 for z in aktuell), max(z.bottom for z in aktuell))
            ergebnis.blocks.append(PdfBlock(idx=idx, page=aktuell[0].seite, kind=art, tokens=toks, bbox=bbox,
                                            section_alt=sec_alt, section_neu=sec_neu,
                                            hinweise=_block_hinweise(toks)))
            block_zeilen[id(ergebnis.blocks[-1])] = (aktuell[0], aktuell[-1])
            idx += 1
            aktuell = []

        vorher: _Zeile | None = None
        block_marke: str | None = None        # Art des Aufzaehlungszeichens, mit dem der aktuelle Block beginnt
        letzte_nr = 0                         # letzte Nummer einer Liste "1." "2." ...
        for z in nutz:
            kopf = _ist_ueberschrift(z, body, cfg)
            # zweite Zeile einer Ueberschrift: gleiche Schrift, enger Abstand, keine eigene Nummer
            folge_kopf = (art == "ueberschrift" and vorher is not None and abs(z.size - vorher.size) < 0.5
                          and z.fett == vorher.fett and z.top - vorher.top <= 1.6 * vorher.size
                          and section_nummer(z.tokens[0].old or z.tokens[0].new) is None)
            marke = _marke(_erste_form(z.tokens[0]))
            ist_listenpunkt = False
            if marke and vorher is not None:
                m_art, m_nr = marke
                if m_art == "zahl":
                    ist_listenpunkt = m_nr == 1 or m_nr == letzte_nr + 1
                elif m_art == "strich":
                    # ein Strich am Zeilenanfang ist nur dann ein Listenpunkt, wenn davor ein Satz endet,
                    # die Zeile davor kurz war oder die Liste mit einem Strich schon laeuft
                    ende = _erste_form(vorher.tokens[-1])[-1:]
                    ist_listenpunkt = block_marke == "strich" or ende in (".", ":", ";", "!", "?") or kurz_im_text(vorher)
                else:
                    ist_listenpunkt = True
            grenze = ""
            if aktuell and art == "absatz" and not kopf and vorher is not None and not ist_listenpunkt:
                grenze = passt(vorher, z.tokens[0])
            neu_start = (not aktuell or (kopf and not folge_kopf) or (art == "ueberschrift" and not folge_kopf)
                         or (not folge_kopf and (
                             (vorher is not None and z.top - vorher.top > lay["absatz_faktor"] * median)
                             or (weit and vorher is not None and kurz(vorher))
                             or ist_listenpunkt or grenze == "ja")))
            if grenze == "unklar" and not neu_start:
                t0 = z.tokens[0]
                t0.hinweis = ";".join(sorted(set(filter(None, t0.hinweis.split(";"))) | {"absatzgrenze_unklar"}))
            if neu_start:
                abschliessen()
                art = "ueberschrift" if kopf else "absatz"
                block_marke = marke[0] if (marke and ist_listenpunkt) else None
                if kopf:
                    letzte_nr = 0
                elif marke and ist_listenpunkt and marke[0] == "zahl":
                    letzte_nr = marke[1]
            aktuell.append(z)
            vorher = z
        abschliessen()
    _seitenumbruch(ergebnis, block_zeilen, weich_ids, {norm(w) for w in lay["keine_silbentrennung_vor"]}, kurz,
                   lay["silbentrennung_min_buchstaben"], wortschatz, passt)
    return seiten_nutz


def _seitenumbruch(ergebnis: PdfAnalyse, block_zeilen: dict[int, tuple[_Zeile, _Zeile]], weich_ids: set[int],
                   ausnahmen: set[str], kurz: Any, min_buchstaben: int = 2, wortschatz: set[str] | None = None,
                   passt: Any = None) -> None:
    """Absatz ueber den Seitenumbruch: ein am Seitenende getrenntes Wort wird zusammengefuegt; setzt sich ein Absatz
    auf der naechsten Seite fort, bekommen beide Bloecke einen Hinweis (die Woerter bleiben je Seite).

    R4: Der letzte Block einer Seite bekommt `fortsetzung` (die Bloecke werden NICHT zusammengelegt, das macht der Planer):
      "sicher"   kein Satzende UND (naechste Seite beginnt klein ODER Silbentrennung verbunden) UND das erste Wort der
                 naechsten Seite haette nicht mehr in die letzte Zeile gepasst;
      "moeglich" sonst, wenn die letzte Zeile voll ist (auch mit Satzende), die Grenze knapp ist oder die naechste Seite
                 ohne Satzende klein weitergeht;
      ""         sonst (die letzte Zeile endet frueh: neuer Absatz)."""
    bl = ergebnis.blocks
    i = 0
    while i < len(bl) - 1:
        b, b2 = bl[i], bl[i + 1]
        if b.kind == b2.kind == "absatz" and b2.page > b.page and b.tokens and b2.tokens:
            letzte = block_zeilen.get(id(b), (None, None))[1]
            v, t = b.tokens[-1], b2.tokens[0]
            p = passt(letzte, t) if (passt is not None and letzte is not None) else ""
            verbunden = False
            m = _verbinden(v, t, id(v) in weich_ids, ausnahmen, min_buchstaben, wortschatz)
            if m is not None:
                verbunden = True
                b.tokens[-1] = m
                del b2.tokens[0]
                b.hinweise += [h for h in _block_hinweise([m]) if h not in b.hinweise]
                if not b2.tokens:
                    del bl[i + 1]
                    for k, rest in enumerate(bl[i + 1:], i + 1):
                        rest.idx = k
                    i += 1
                    continue
                b2.bbox = (min(x.bbox[0] for x in b2.tokens), min(x.bbox[1] for x in b2.tokens),
                           max(x.bbox[2] for x in b2.tokens), max(x.bbox[3] for x in b2.tokens))
                v, t = b.tokens[-1], b2.tokens[0]
            text_v, text_t = _erste_form(v), _erste_form(t)
            satzende = bool(re.search(r"[.!?:;)\"”»]$", text_v))
            klein = text_t[:1].islower()
            voll = p == "nein" or (p == "" and letzte is not None and not kurz(letzte))
            if klein or (not satzende and letzte is not None and not kurz(letzte)):
                for blk, satz in ((b, "Der Absatz geht auf der nächsten Seite weiter."),
                                  (b2, "Dieser Absatz ist die Fortsetzung des Absatzes von der vorigen Seite.")):
                    if satz not in blk.hinweise:
                        blk.hinweise.append(satz)
            if not satzende and (klein or verbunden) and (p == "nein" or verbunden):
                b.fortsetzung = "sicher"
            elif voll or p == "unklar" or (klein and not satzende):
                b.fortsetzung = "moeglich"
        i += 1


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
            t.old, t.status, t.teile = t.new, KEEP, []
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
    saetze: list[str] = []
    for c in codes:
        satz = _HINWEIS_TEXT.get("kommentar" if c.startswith("kommentar_") else c)
        if satz and satz not in saetze:
            saetze.append(satz)
    return saetze


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


def _abkuerzung(vor: str) -> bool:
    """Steht vor dem Zeilenend-Strich eine Abkuerzung oder Ziffer ('XML-', 'BESS-', '2-')? Dann ist er ein echter Bindestrich."""
    if not vor:
        return False
    if vor[-1].isdigit():
        return True
    teil = re.split(r"[-/]", vor)[-1]
    return len(teil) >= 2 and teil.isupper()


def _laeufe(t: PdfToken) -> list[tuple[str, str]]:
    """Zeichen-Laeufe eines Tokens (auch wenn `teile` leer ist)."""
    if t.teile:
        return list(t.teile)
    return [(t.status, t.new if t.status == ADD else t.old)]


def _verbinden(v: PdfToken, t: PdfToken, weich: bool, ausnahmen: set[str], min_buchstaben: int = 2,
               wortschatz: set[str] | None = None) -> PdfToken | None:
    """Fuegt ein am Zeilenende getrenntes Wort wieder zusammen ('Kapazi-' + 'taet'; mit '-' oder weichem Trennstrich).
    Gilt fuer unveraenderte, neue und gestrichene Woerter, wenn beide Teile denselben Status haben.
    `min_buchstaben` (config layout.silbentrennung_min_buchstaben): so viele Buchstaben muss das Stueck VOR dem Trennstrich haben.
    Der weiche Trennstrich ist unsichtbar; er zaehlt nicht mit ('Ne' + 'nnleistung' ist genauso lang wie 'Ne-' + 'nnleistung').

    R8: Grossbuchstabe nach dem Strich ('E-' + 'Mail') oder Abkuerzung/Ziffer davor ('XML-' + 'basiert') -> der Strich bleibt
    (echter Bindestrich). Kleinbuchstabe danach -> ohne Strich. Bei NEUEM Text ohne weichen Strich nur dann sicher, wenn das
    zusammengefuegte Wort sonst im Dokument vorkommt (`wortschatz`, normiert); kommt nur die Form mit Strich vor, bleibt er;
    sonst ohne Strich + Hinweis 'trennung_unklar'. `trennstelle` = Position im Wort (Zeichenfolge aller Laeufe): bei
    behaltenem Strich zeigt sie auf den Strich (er ist optional), sonst auf die Fuegestelle."""
    if v.status != t.status or v.status not in (KEEP, ADD, DEL):
        return None
    text_v = v.new if v.status == ADD else v.old
    text_t = t.new if t.status == ADD else t.old
    if not text_v or not text_t or norm(text_t) in ausnahmen:
        return None
    hinweis_neu: set[str] = set()
    if text_v.endswith("-"):
        vor = text_v[:-1]
        if not vor or not vor[-1].isalnum():
            return None
        if text_t[:1].isupper():
            mit_strich = True
        elif text_t[:1].islower():
            mit_strich = _abkuerzung(vor)
            if not mit_strich and len(vor) < min_buchstaben:
                return None
            if not mit_strich and v.status == ADD and wortschatz is not None:
                if norm(vor + text_t) not in wortschatz:
                    if norm(vor + "-" + text_t) in wortschatz:
                        mit_strich = True
                    else:
                        hinweis_neu.add("trennung_unklar")
        else:
            return None
        strich = 0 if mit_strich else 1
    elif weich and text_v[-1].isalpha() and text_t[:1].islower():
        if len(text_v) < min_buchstaben:
            return None
        mit_strich, strich = False, 0
    else:
        return None
    bbox = (v.bbox[0], min(v.bbox[1], t.bbox[1]), max(v.bbox[2], t.bbox[2]), max(v.bbox[3], t.bbox[3]))
    if v.page != t.page:
        bbox = v.bbox
    hin = ";".join(sorted((set(v.hinweis.split(";")) | set(t.hinweis.split(";")) | hinweis_neu) - {""}))
    alt = (v.old[:len(v.old) - strich] + t.old) if v.old else ""
    neu = (v.new[:len(v.new) - strich] + t.new) if v.new else ""
    # Zeichen-Laeufe zusammensetzen (Strich am Ende des ersten Teils entfernen, wenn er wegfaellt)
    lv = _laeufe(v)
    if strich:
        s_, tx = lv[-1]
        lv[-1] = (s_, tx[:-1])
        lv = [x for x in lv if x[1]]
    fuge = sum(len(tx) for _, tx in lv)
    teile: list[tuple[str, str]] = []
    for s_, tx in lv + _laeufe(t):
        if teile and teile[-1][0] == s_:
            teile[-1] = (s_, teile[-1][1] + tx)
        else:
            teile.append((s_, tx))
    return PdfToken(old=alt, new=neu, status=v.status, page=v.page, bbox=bbox, hinweis=hin, trenner=t.trenner,
                    teile=teile if len(teile) > 1 else [], trennstelle=fuge - 1 if mit_strich and text_v.endswith("-") else fuge)


def _silben(tokens: list[PdfToken], zeilen: list[_Zeile], cfg: dict, weich_ids: set[int],
            wortschatz: set[str] | None = None) -> list[PdfToken]:
    """Fuegt am Zeilenende getrennte Woerter wieder zusammen. weich_ids: Woerter, hinter denen ein weicher
    Trennstrich steht (wird um zusammengefuegte Woerter ergaenzt)."""
    lay = cfg["layout"]
    if not lay["silbentrennung_zusammenfuegen"] or len(zeilen) < 2:
        return tokens
    ausnahmen = {norm(w) for w in lay["keine_silbentrennung_vor"]}
    mindest = lay["silbentrennung_min_buchstaben"]
    erste_der_zeilen = {id(z.tokens[0]) for z in zeilen[1:]}
    out: list[PdfToken] = []
    for t in tokens:
        if out and id(t) in erste_der_zeilen:
            v = out[-1]
            m = _verbinden(v, t, id(v) in weich_ids, ausnahmen, mindest, wortschatz)
            if m is not None:
                if id(t) in weich_ids:
                    weich_ids.add(id(m))
                out[-1] = m
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
    # --- Text in anderer Schreibrichtung (Randvermerk, Wasserzeichen) ------------------------------
    if sd.fremd_woerter:
        if sd.fremd_farbig:
            ergebnis.luecken.append(luecke(
                "sonstiges", nr, f"Seite {nr}: {sd.fremd_woerter} Wörter stehen in einer anderen Schreibrichtung als der "
                                 f"übrige Text und sind farbig markiert. Sie wurden nicht gelesen.", sd.fremd_woerter))
        else:
            ergebnis.warnungen.append(
                f"Seite {nr}: {sd.fremd_woerter} Wörter stehen in einer anderen Schreibrichtung als der übrige Text "
                f"(Randvermerk, Wasserzeichen?) und wurden nicht gelesen.")
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
                markierer[nr] = _Markierer(page, cfg, sd.rahmen)
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
        ergebnis.warnungen += m.meldungen(nr)
        alle_zeilen.append(_zeilen(sd.zeichen, nr, cfg, grund_hex))
    nutz_pro_seite = _zu_bloecken(alle_zeilen, seiten_h, cfg, ergebnis)
    for sd, zl, nutz in zip(seiten_daten, alle_zeilen, nutz_pro_seite):
        _seite_luecken(sd, zl, nutz, cfg, ergebnis, sd.nr in ohne_text)
    for l, seite, boxen in ergebnis.regionen:
        l["woerter"] = sum(1 for b in ergebnis.blocks if b.page == seite for t in b.tokens
                           if t.status != KEEP and any(_in_box(t, box) for box in boxen))
        if l["art"] in ("mehrspaltig", "tabelle"):      # geaenderte Woerter im Bereich tragen den Hinweis (der Planer uebernimmt sie nie allein)
            for b in ergebnis.blocks:
                if b.page != seite:
                    continue
                gesetzt = False
                for t in b.tokens:
                    if t.status != KEEP and any(_in_box(t, box) for box in boxen):
                        t.hinweis = ";".join(sorted(set(filter(None, t.hinweis.split(";"))) | {l["art"]}))
                        gesetzt = True
                if gesetzt:
                    b.hinweise += [h for h in _block_hinweise(b.tokens) if h not in b.hinweise]
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
