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

Auftrag 003 (ADR 0005 Nachtrag, ADR 0007): `markiere_ops` baut aus dem VORHANDENEN Zellinhalt (Stuecke samt Markierungen
frueherer Runden) und den geplanten Text-Operationen die End-Stuecke: geloeschter Text bleibt als Alt-Stueck an seiner
Stelle, eingefuegter Text steht als Neu-Stueck direkt dahinter (kein Trenner, ausser das PDF hat dort Leerraum), jedes neue
Stueck bekommt die Grundschrift seiner Umgebung, vorhandene Laeufe behalten ihre Schrift unveraendert (`behalten`).
Die Positionen der Operationen beziehen sich auf den Text, den das Excel-Modell gesehen hat. Heute ist das der FLACHE Text
(`str(CellRichText)`, auch frueher durchgestrichene Laeufe; src/excel_model.py `_lesen`); liest das Modell kuenftig die
Neu-Ansicht (ADR 0006: durchgestrichene Laeufe zaehlen nicht), gilt `sicht="neu"` -- der Applier erkennt das je Zelle.
"""
from __future__ import annotations

import colorsys
from dataclasses import dataclass, field
from typing import Any, Iterable

from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font

from . import config
from .models import Lauf, Operation

NORMAL, ALT, NEU = "normal", "alt", "neu"
FORMATE = (NORMAL, ALT, NEU)
SICHT_FLACH, SICHT_NEU = "flach", "neu"      # Sicht des Excel-Modells auf eine Zelle (siehe Modul-Text)


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
    """Ein Textstueck einer Zelle mit Format. `basis` ist optional die Schrift des Originals an dieser Stelle (InlineFont).

    `behalten`: vorhandener Text -- seine Schrift bleibt genau so (auch Markierungen frueherer Runden, z. B. dunkelrot oder
    durchgestrichen); sonst bekommt ein Stueck mit Format alt/neu beim Schreiben Farbe und Strich. `aenderung`: ID der
    Aenderung, die das Stueck in diesem Lauf markiert bzw. eingefuegt hat ("" = vorhandener Text)."""
    format: str
    text: str
    basis: Any = None
    behalten: bool = False
    aenderung: str = ""

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
        v = out[-1] if out else None
        if (v is not None and v.format == s.format and v.basis is s.basis and v.behalten == s.behalten
                and v.aenderung == s.aenderung):
            out[-1] = Stueck(s.format, v.text + s.text, s.basis, s.behalten, s.aenderung)
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


def sicht_text(stuecke: Iterable[Stueck], sicht: str = SICHT_FLACH) -> str:
    """Text einer Zelle in der Sicht des Excel-Modells: flach = alles, neu = ohne durchgestrichene (alte) Stuecke."""
    return "".join(s.text for s in stuecke if sicht != SICHT_NEU or s.format != ALT)


# ------------------------------------------------------------------ Markieren nach Operationen (Auftrag 003)
@dataclass
class Markierung:
    """Ergebnis von `markiere_ops`.

    stuecke: die End-Stuecke der Zelle (zum Schreiben).
    zeichen: je Zeichen der Stuecke (gleiche Reihenfolge) die Herkunft fuer die Gegenprobe:
      (zeichen, klasse, pos, luecke, aenderung)
      klasse   = alt | neu fuer Markierungen DIESES Laufs (alt auch fuer schon durchgestrichenen Text in einer Loeschung),
                 sonst normal (auch Markierungen frueherer Runden und Leerraum zwischen Alt und Neu);
      pos      = Stelle im Original-Text der Sicht (-1 = neu eingefuegt oder in der Sicht unsichtbar);
      luecke   = bei neuen/unsichtbaren Zeichen die Stelle zwischen zwei Original-Zeichen, an der sie stehen (sonst = pos);
      aenderung = Aenderung, die das Zeichen markiert/eingefuegt hat ("" = vorhandener Text).
    neu_markiert: je Aenderung die Zahl sichtbarer Zeichen (ohne Leerraum), die DIESER Lauf neu markiert oder eingefuegt hat.
    schon_alt: je Aenderung die Zahl der Zeichen in ihren Loeschungen, die schon durchgestrichen waren (Idempotenz)."""
    stuecke: list[Stueck]
    zeichen: list[tuple[str, str, int, int, str]]
    neu_markiert: dict[str, int] = field(default_factory=dict)
    schon_alt: dict[str, int] = field(default_factory=dict)

    @property
    def text(self) -> str:
        return "".join(s.text for s in self.stuecke)


def trenner_aus_laeufen(laeufe: Iterable[Lauf]) -> dict[str, str]:
    """Leerraum, der im PDF zwischen dem entfallenden und dem neuen Teil EINER Aenderung steht ("{-12-} {+7+}" -> " ").
    Aenderungen, deren Alt und Neu im PDF aneinanderkleben ("{-alt-}|{+neu+}"), bekommen "" (wie im PDF)."""
    ll = list(laeufe)
    out: dict[str, str] = {}
    for i, l in enumerate(ll):
        if l.art != "d" or not l.aenderung:
            continue
        j, sep = i + 1, ""
        while j < len(ll) and ll[j].art == "s":
            if _RANG.get(ll[j].text, 1) > _RANG.get(sep, 0):
                sep = ll[j].text
            j += 1
        if j < len(ll) and ll[j].art == "a" and ll[j].aenderung == l.aenderung:
            out[l.aenderung] = sep
    return out


def markiere_ops(vorher: list[Stueck], ops: Iterable[Operation], sicht: str = SICHT_FLACH,
                 zwischen: dict[str, str] | None = None, cfg: dict | None = None) -> Markierung:
    """End-Stuecke einer Zelle (reine Funktion, schreibt nichts).

    vorher: Stuecke der Zelle (stuecke_aus_zelle, mit Schrift); ops: ersetzen/einfuegen/zelle_leeren dieser Zelle
    (Positionen im Text der `sicht`; abgedeckte Operationen zaehlen nicht).
      * geloeschter Text (ersetzen, zelle_leeren) bleibt stehen: Format alt (blau + durchgestrichen); schon durchgestrichener
        Text behaelt seine Schrift;
      * eingefuegter Text steht an seiner Stelle als Format neu (rot), bei einer Ersetzung DIREKT hinter dem alten Text
        (dazwischen nur der Leerraum aus dem PDF, `zwischen[aenderung]`, sonst `markierung.trenner`);
      * neue Stuecke bekommen die Grundschrift der Umgebung (Zeichen davor, am Anfang das erste Zeichen);
      * alles andere bleibt, wie es ist (Markierungen frueherer Runden eingeschlossen).
    Reihenfolge an derselben Stelle wie im Applier (`_roh_anwenden`): Ersatztext einer dort endenden Ersetzung, dann
    Einfuegungen (nach seq), dann Einfuegungen per leerer Ersetzung (nach seq)."""
    trenner = str(konfig(cfg).get("trenner", ""))
    zwischen = zwischen or {}
    orig: list[tuple[str, int, int]] = []           # (Zeichen, Stueck-Index, Position in der Sicht | -1)
    n = 0
    for i, st in enumerate(vorher):
        unsichtbar = sicht == SICHT_NEU and st.format == ALT
        for ch in st.text:
            orig.append((ch, i, -1 if unsichtbar else n))
            if not unsichtbar:
                n += 1
    loesch = [""] * n
    einf: dict[int, list[tuple[tuple, str, str, str]]] = {}
    for o in ops:
        if o.abgedeckt:
            continue
        if o.art == "zelle_leeren":
            loesch = [x or o.aenderung for x in loesch]
            continue
        if o.art not in ("ersetzen", "einfuegen"):
            continue
        if not 0 <= o.start <= max(o.start, o.ende) <= n:
            raise ValueError("Operation liegt außerhalb des Zelltexts")
        if o.art == "ersetzen":
            for p in range(o.start, o.ende):
                loesch[p] = loesch[p] or o.aenderung
            if o.text and o.ende > o.start:
                einf.setdefault(o.ende, []).append(((0, o.start, o.seq), zwischen.get(o.aenderung, trenner), o.text,
                                                    o.aenderung))
            elif o.text:
                einf.setdefault(o.start, []).append(((2, o.seq, 0), "", o.text, o.aenderung))
        elif o.text:
            einf.setdefault(o.start, []).append(((1, o.seq, 0), "", o.text, o.aenderung))
    sicht_basis = [vorher[i].basis for _, i, p in orig if p >= 0]

    def basis_an(g: int) -> Any:
        if 0 < g <= len(sicht_basis):
            return sicht_basis[g - 1]
        if sicht_basis:
            return sicht_basis[0]
        return vorher[0].basis if vorher else None

    roh: list[tuple[str, str, Any, bool, str, str, int, int]] = []   # Zeichen, Format, Basis, behalten, Aenderung, Klasse, pos, Luecke

    def einsetzen(g: int) -> None:
        for _, sep, text, aid in sorted(einf.get(g, ()), key=lambda e: e[0]):
            b = basis_an(g)
            roh.extend((ch, NORMAL, b, False, aid, NORMAL, -1, g) for ch in sep)
            roh.extend((ch, NEU, b, False, aid, NEU, -1, g) for ch in text)

    luecke = 0
    for ch, i, p in orig:
        st = vorher[i]
        if p < 0:
            roh.append((ch, st.format, st.basis, True, "", NORMAL, -1, luecke))
            continue
        einsetzen(p)
        luecke = p + 1
        aid = loesch[p]
        if aid and st.format == ALT:                 # schon durchgestrichen: Schrift bleibt, gilt als entfallen
            roh.append((ch, ALT, st.basis, True, aid, ALT, p, p))
        elif aid:
            roh.append((ch, ALT, st.basis, False, aid, ALT, p, p))
        else:
            roh.append((ch, st.format, st.basis, True, "", NORMAL, p, p))
    einsetzen(n)
    stuecke = zusammenfassen(Stueck(f, ch, b, beh, aid) for ch, f, b, beh, aid, _, _, _ in roh)
    zeichen = [(ch, kl, p, lu, aid) for ch, _, _, _, aid, kl, p, lu in roh]
    neu_markiert: dict[str, int] = {}
    schon_alt: dict[str, int] = {}
    for ch, _, _, beh, aid, kl, _, _ in roh:
        if not aid:
            continue
        if not beh and kl != NORMAL and not ch.isspace():
            neu_markiert[aid] = neu_markiert.get(aid, 0) + 1
        elif beh and kl == ALT:
            schon_alt[aid] = schon_alt.get(aid, 0) + 1
    return Markierung(stuecke, zeichen, neu_markiert, schon_alt)


def klassen_stuecke(zeichen: Iterable[tuple[str, str, int, int, str]]) -> list[list[str]]:
    """[[klasse, text], ...] aus `Markierung.zeichen` (fuer die Anzeige, Aenderung.stuecke, Schema 4)."""
    out: list[list[str]] = []
    for ch, kl, *_ in zeichen:
        if out and out[-1][0] == kl:
            out[-1][1] += ch
        else:
            out.append([kl, ch])
    return out


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


def _schrift_behalten(basis: Any, zellschrift: Any) -> InlineFont:
    """Schrift eines vorhandenen Stuecks unveraendert (auch Farbe und Strich frueherer Runden)."""
    if isinstance(basis, InlineFont):
        return basis
    q = basis if basis is not None else zellschrift
    return InlineFont(rFont=getattr(q, "rFont", None) or getattr(q, "name", None), sz=getattr(q, "sz", None),
                      b=getattr(q, "b", None), i=getattr(q, "i", None), u=getattr(q, "u", None),
                      strike=getattr(q, "strike", None), color=getattr(q, "color", None),
                      family=getattr(q, "family", None), charset=getattr(q, "charset", None),
                      scheme=getattr(q, "scheme", None), vertAlign=getattr(q, "vertAlign", None))


def schreibe(zelle: Any, stuecke: Iterable[Stueck], cfg: dict | None = None) -> None:
    """Schreibt Stuecke in eine Zelle. Alles normal -> einfacher Text; ein einziges markiertes Format -> Zellschrift
    (rot bzw. blau+durchgestrichen) und einfacher Text; Mischung -> Rich Text. Zahlenformat, Fuellung, Ausrichtung bleiben.
    Stuecke mit eigener Schrift (`basis`, z. B. vorhandene Laeufe) werden immer Rich Text; `behalten` = Schrift unveraendert."""
    st = zusammenfassen(stuecke)
    if not st:
        zelle.value = None
        return
    formate = {s.format for s in st}
    if len(formate) == 1 and all(s.basis is None and not s.behalten for s in st):
        text = "".join(s.text for s in st)
        zelle.value = text
        if formate != {NORMAL}:
            faerbe_zelle(zelle, next(iter(formate)), cfg)
        return
    if all(s.behalten and s.basis is None for s in st):
        zelle.value = "".join(s.text for s in st)
        return
    zelle.value = CellRichText(*[TextBlock(_schrift_behalten(s.basis, zelle.font) if s.behalten
                                           else inline_schrift(s.basis, zelle.font, s.format, cfg), s.text) for s in st])


def faerbe_zelle(zelle: Any, format_: str, cfg: dict | None = None) -> None:
    """Ganze Zelle: Schrift rot (neu) bzw. blau+durchgestrichen (alt), sonst unveraendert. Auch fuer Zahlen und Formeln."""
    if format_ == NORMAL:
        return
    f = zelle.font
    farbe = _farbe(format_, cfg)
    zelle.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, u=f.u, family=f.family, charset=f.charset, scheme=f.scheme,
                      vertAlign=f.vertAlign, color=farbe, strike=True if format_ == ALT else f.strike)
