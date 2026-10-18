"""Excel-Markierung der Aenderungen (ADR 0005): reine Funktionen plus ein kleiner openpyxl-Teil.

In der CDE-Excel sehen Aenderungen aus wie im PDF:
  * neuer Inhalt            = rot
  * alter Inhalt (entfaellt) = blau UND durchgestrichen
  * ersetzte Stelle         = Alt (blau, durchgestrichen) DIREKT vor Neu (rot) in derselben Zelle, dazwischen ein Leerzeichen
  * ganz gestrichene Section = die Zeile bleibt, ihr Inhalt ist blau und durchgestrichen
Eine markierte Zelle ist ein Rich-Text-Wert (CellRichText) aus Stuecken. Jedes Stueck hat ein Format:
  "normal" | "alt" | "neu".
Zwei Ansichten sind daraus ableitbar und machen die Gegenprobe stark:
  * alt_ansicht = normal + alt  -> muss genau der ORIGINAL-Text sein (nichts geht verloren)
  * neu_ansicht = normal + neu  -> muss genau der neue Absatz laut PDF sein

Die Bedeutung eines Stuecks in einer VORHANDENEN Zelle (Idempotenz, echte "nachher"-Excel) wird so gelesen:
durchgestrichen = alt (egal welche Farbe, wie im PDF), rote Schrift ohne Strich = neu, sonst normal.
"""
from __future__ import annotations

import colorsys
from dataclasses import dataclass
from typing import Any, Iterable

from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font

from . import config
from .models import Lauf

NORMAL, ALT, NEU = "normal", "alt", "neu"
FORMATE = (NORMAL, ALT, NEU)


def konfig(cfg: dict | None = None) -> dict[str, Any]:
    """Abschnitt `markierung:` aus config/zuordnung.yaml (Regel 28)."""
    if cfg and "markierung" in cfg:
        return cfg["markierung"]
    return config.lade("zuordnung").get("markierung", {})


def aktiv(cfg: dict | None = None) -> bool:
    """Werden Aenderungen farbig in der Excel markiert? (`markierung.aenderungen: farbe`; `keine` = alte Reinigungs-Logik)"""
    return str(konfig(cfg).get("aenderungen", "keine")) == "farbe"


@dataclass(frozen=True)
class Stueck:
    """Ein Textstueck einer Zelle mit Format. `basis` ist optional die Schrift des Originals an dieser Stelle (InlineFont)."""
    format: str
    text: str
    basis: Any = None

    def __post_init__(self) -> None:
        if self.format not in FORMATE:
            raise ValueError(f"unbekanntes Format {self.format!r}")


# ------------------------------------------------------------------ Zustand mit Markierung (reine Logik)
_RANG = {"": 0, " ": 1, "\t": 1, " ": 2, " ": 2, "\n": 3}


def zustand_markiert(laeufe: Iterable[Lauf], angewendet: set[str], cfg: dict | None = None) -> list[Stueck]:
    """Wie soll ein PDF-Absatz in der Excel aussehen, wenn genau die Aenderungen in `angewendet` eingetragen sind?

    Entfaellt-Laeufe (d) einer angewendeten Aenderung bleiben stehen (Format alt), Neu-Laeufe (a) werden rot (Format neu),
    nicht angewendete a-Laeufe fehlen, nicht angewendete d-Laeufe bleiben normaler Text. Folgt auf ein Alt-Stueck direkt ein
    Neu-Stueck (verschmolzene Ersetzung), steht dazwischen der konfigurierte Trenner (Vorgabe: nichts, wie im PDF)."""
    trenner = str(konfig(cfg).get("trenner", ""))
    sichtbar: list[Lauf] = []
    for l in laeufe:
        if l.art == "a" and l.aenderung not in angewendet:
            continue
        if l.art == "s" and sichtbar and sichtbar[-1].art == "s":
            if _RANG.get(l.text, 1) > _RANG.get(sichtbar[-1].text, 1):
                sichtbar[-1] = l
            continue
        sichtbar.append(l)
    while sichtbar and sichtbar[0].art == "s":
        sichtbar.pop(0)
    while sichtbar and sichtbar[-1].art == "s":
        sichtbar.pop()
    stuecke: list[Stueck] = []
    for l in sichtbar:
        if l.art == "d" and l.aenderung in angewendet:
            f = ALT
        elif l.art == "a":
            f = NEU
        else:
            f = NORMAL
        if stuecke and stuecke[-1].format == ALT and f == NEU:
            stuecke.append(Stueck(NORMAL, trenner))
        stuecke.append(Stueck(f, l.text))
    return zusammenfassen(stuecke)


def zusammenfassen(stuecke: Iterable[Stueck]) -> list[Stueck]:
    """Aufeinanderfolgende Stuecke mit gleichem Format (und gleicher Basisschrift) werden eins; leere fallen weg."""
    out: list[Stueck] = []
    for s in stuecke:
        if not s.text:
            continue
        if out and out[-1].format == s.format and out[-1].basis is s.basis:
            out[-1] = Stueck(s.format, out[-1].text + s.text, s.basis)
        else:
            out.append(s)
    return out


def ansicht(stuecke: Iterable[Stueck], welche: str) -> str:
    """'alt' = normal + alt (Original), 'neu' = normal + neu (Stand nach der Aenderung)."""
    if welche not in (ALT, NEU):
        raise ValueError("ansicht: 'alt' oder 'neu'")
    return "".join(s.text for s in stuecke if s.format in (NORMAL, welche))


# ------------------------------------------------------------------ Lesen einer Zelle
def _farbton(farbe: Any) -> tuple[float, float, float] | None:
    """(Farbton in Grad, Saettigung, Helligkeit) einer openpyxl-Farbe mit RGB; sonst None (Theme/Index)."""
    rgb = getattr(farbe, "rgb", None)
    if not isinstance(rgb, str) or len(rgb) not in (6, 8):
        return None
    try:
        r, g, b = (int(rgb[-6:][i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except ValueError:
        return None
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    return h * 360.0, s, v


def ist_rot(farbe: Any) -> bool:
    t = _farbton(farbe)
    return bool(t) and t[1] >= 0.45 and t[2] >= 0.35 and (t[0] <= 18 or t[0] >= 335)


def format_von_schrift(schrift: Any) -> str:
    """Bedeutung einer Schrift: durchgestrichen = alt (jede Farbe), rot ohne Strich = neu, sonst normal."""
    if schrift is None:
        return NORMAL
    if getattr(schrift, "strike", None):
        return ALT
    if ist_rot(getattr(schrift, "color", None)):
        return NEU
    return NORMAL


def stuecke_aus_zelle(zelle: Any) -> list[Stueck]:
    """Die Stuecke einer Zelle (Rich Text je Stueck, sonst die Zellschrift als Ganzes)."""
    v = zelle.value
    if isinstance(v, CellRichText):
        out = []
        for b in v:
            if isinstance(b, TextBlock):
                out.append(Stueck(format_von_schrift(b.font), b.text, b.font))
            else:
                out.append(Stueck(format_von_schrift(zelle.font), str(b), None))
        return zusammenfassen(out)
    if v is None:
        return []
    return [Stueck(format_von_schrift(zelle.font), v if isinstance(v, str) else str(v))]


def ansichten(wert_oder_zelle: Any) -> tuple[str, str, list[Stueck]]:
    """(alt_ansicht, neu_ansicht, stuecke) einer Zelle oder eines Zellwerts (str / CellRichText)."""
    if hasattr(wert_oder_zelle, "value") and hasattr(wert_oder_zelle, "font"):
        st = stuecke_aus_zelle(wert_oder_zelle)
    elif isinstance(wert_oder_zelle, CellRichText):
        st = zusammenfassen(Stueck(format_von_schrift(b.font), b.text, b.font) if isinstance(b, TextBlock)
                            else Stueck(NORMAL, str(b)) for b in wert_oder_zelle)
    else:
        st = [Stueck(NORMAL, "" if wert_oder_zelle is None else str(wert_oder_zelle))]
    return ansicht(st, ALT), ansicht(st, NEU), st


def hat_markierung(stuecke: Iterable[Stueck]) -> bool:
    return any(s.format != NORMAL for s in stuecke)


# ------------------------------------------------------------------ Schreiben
def _farbe(format_: str, cfg: dict | None) -> str | None:
    k = konfig(cfg)
    if format_ == ALT:
        return str(k.get("alt_farbe", "FF00B0F0"))
    if format_ == NEU:
        return str(k.get("neu_farbe", "FFFF0000"))
    return None


def inline_schrift(basis: Any, zellschrift: Any, format_: str, cfg: dict | None = None) -> InlineFont:
    """Schrift eines Stuecks: die Schrift des Originals (Name, Groesse, fett, kursiv, ...) plus Farbe/Strich je Format.
    Ohne diese Kopie faellt ein Rich-Text-Stueck auf die Standardschrift zurueck."""
    quelle = basis if basis is not None else zellschrift
    f = InlineFont(
        rFont=getattr(quelle, "rFont", None) or getattr(quelle, "name", None),
        sz=getattr(quelle, "sz", None) or getattr(quelle, "size", None),
        b=getattr(quelle, "b", None), i=getattr(quelle, "i", None),
        u=getattr(quelle, "u", None), family=getattr(quelle, "family", None),
        charset=getattr(quelle, "charset", None), scheme=getattr(quelle, "scheme", None),
        vertAlign=getattr(quelle, "vertAlign", None))
    farbe = _farbe(format_, cfg)
    if format_ == NORMAL and getattr(quelle, "color", None) is not None:
        f.color = quelle.color
    elif farbe:
        f.color = farbe
    f.strike = True if format_ == ALT else None
    return f


def schreibe(zelle: Any, stuecke: Iterable[Stueck], cfg: dict | None = None) -> None:
    """Schreibt Stuecke in eine Zelle. Alles normal -> einfacher Text; ein einziges markiertes Format -> Zellschrift
    (rot bzw. blau+durchgestrichen) und einfacher Text; Mischung -> Rich Text. Zahlenformat, Fuellung, Ausrichtung bleiben."""
    st = zusammenfassen(stuecke)
    if not st:
        zelle.value = None
        return
    formate = {s.format for s in st}
    if len(formate) == 1:
        text = "".join(s.text for s in st)
        zelle.value = text
        if formate != {NORMAL}:
            faerbe_zelle(zelle, next(iter(formate)), cfg)
        return
    zelle.value = CellRichText(*[TextBlock(inline_schrift(s.basis, zelle.font, s.format, cfg), s.text) for s in st])


def faerbe_zelle(zelle: Any, format_: str, cfg: dict | None = None) -> None:
    """Ganze Zelle: Schrift rot (neu) bzw. blau+durchgestrichen (alt), sonst unveraendert. Auch fuer Zahlen und Formeln."""
    if format_ == NORMAL:
        return
    f = zelle.font
    farbe = _farbe(format_, cfg)
    zelle.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, u=f.u, family=f.family, charset=f.charset, scheme=f.scheme,
                      vertAlign=f.vertAlign, color=farbe, strike=True if format_ == ALT else f.strike)
