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
import re
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
        # normale Stuecke mit gleicher Schrift sehen gleich aus (ob vorhanden oder als Leerraum neu): ein Lauf
        if (v is not None and v.format == s.format and v.basis is s.basis
                and ((v.behalten == s.behalten and v.aenderung == s.aenderung) or s.format == NORMAL)):
            out[-1] = Stueck(s.format, v.text + s.text, s.basis, v.behalten and s.behalten, v.aenderung or s.aenderung)
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
                 zwischen: dict[str, str] | None = None, cfg: dict | None = None,
                 bereiche: list[tuple[int, int, list[tuple[str, str, int, int, str]]]] | None = None) -> Markierung:
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
    Einfuegungen (nach seq), dann Einfuegungen per leerer Ersetzung (nach seq).
    bereiche: [(start, ende, eintraege)] -- Absatz-Bereiche, deren Markierung aus dem PDF projiziert wurde (`projiziere`;
    Positionen relativ zu start). Dort gilt die Projektion statt der Operationen; eine Operation in einem Bereich ist ein
    Widerspruch (ValueError, der Aufrufer nimmt dann nur die Operationen)."""
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
    orig_text = [ch for ch, _, p in orig if p >= 0]
    loesch = [""] * n
    einf: dict[int, list[tuple[tuple, str, str, str]]] = {}
    schreibweise: dict[str, str] = {}
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
            s, e, text = o.start, o.ende, o.text
            if text and e > s:                   # was vorn/hinten gleich bleibt (an Wortgrenzen), wird nicht markiert
                vorn, hinten = _gemeinsam("".join(orig_text[s:e]), text, schreibweise, cfg)
                s, e, text = s + vorn, e - hinten, text[vorn:len(text) - hinten]
            if text.strip() == "" and "".join(orig_text[s:e]).strip() == "":
                if text:                         # reiner Leerraum: nie markieren (eingefuegt wird er normal)
                    einf.setdefault(s, []).append(((2, o.seq, 0), "", text, o.aenderung))
                continue
            for p in range(s, e):
                loesch[p] = loesch[p] or o.aenderung
            if text and e > s:
                einf.setdefault(e, []).append(((0, s, o.seq), zwischen.get(o.aenderung, trenner), text, o.aenderung))
            elif text:
                einf.setdefault(s, []).append(((2, o.seq, 0), "", text, o.aenderung))
        elif o.text:
            einf.setdefault(o.start, []).append(((1, o.seq, 0), "", o.text, o.aenderung))
    p = 0
    while p < n:                                 # Leerraum am Rand eines gestrichenen Stuecks wird nicht gestrichen
        if not loesch[p]:
            p += 1
            continue
        q = p
        while q < n and loesch[q] == loesch[p]:
            q += 1
        a, b = p, q
        while a < b and orig_text[a].isspace():
            loesch[a] = ""
            a += 1
        while b > a and orig_text[b - 1].isspace():
            loesch[b - 1] = ""
            b -= 1
        p = q
    sichtbar: list[int] = []                         # Stueck-Index je sichtbarem Zeichen
    unsichtbar_bei: dict[int, list[tuple[str, int]]] = {}
    for ch, i, p in orig:
        if p < 0:
            unsichtbar_bei.setdefault(len(sichtbar), []).append((ch, i))
        else:
            sichtbar.append(i)
    bereich_ab: dict[int, tuple[int, list]] = {}
    ende_vorher = -1
    for a, e, ent in sorted(bereiche or [], key=lambda b: b[0]):
        if not 0 <= a <= e <= n or a < ende_vorher or a in bereich_ab:
            raise ValueError("Bereiche passen nicht zum Zelltext")
        if any(loesch[p] for p in range(a, e)) or any(a < g < e for g in einf):
            raise ValueError("Operation in einem projizierten Bereich")
        bereich_ab[a] = (e, ent)
        ende_vorher = e

    def basis_an(g: int) -> Any:
        if 0 < g <= len(sichtbar):
            return vorher[sichtbar[g - 1]].basis
        if sichtbar:
            return vorher[sichtbar[0]].basis
        return vorher[0].basis if vorher else None

    roh: list[tuple[str, str, Any, bool, str, str, int, int]] = []   # Zeichen, Format, Basis, behalten, Aenderung, Klasse, pos, Luecke

    def einsetzen(g: int) -> None:
        for _, sep, text, aid in sorted(einf.get(g, ()), key=lambda e: e[0]):
            b = basis_an(g)
            a0, a1 = len(text) - len(text.lstrip()), len(text.rstrip())     # Leerraum am Rand ist nicht markiert
            roh.extend((ch, NORMAL, b, False, aid, NORMAL, -1, g) for ch in sep + text[:a0])
            roh.extend((ch, NEU, b, False, aid, NEU, -1, g) for ch in text[a0:a1])
            roh.extend((ch, NORMAL, b, False, aid, NORMAL, -1, g) for ch in text[max(a0, a1):])

    def unsichtbar_an(g: int) -> None:
        for ch, i in unsichtbar_bei.get(g, ()):
            roh.append((ch, vorher[i].format, vorher[i].basis, True, "", NORMAL, -1, g))

    def zeichen_an(p: int, aid: str) -> None:
        st = vorher[sichtbar[p]]
        ch = orig_text[p]
        if aid and st.format == ALT:                 # schon durchgestrichen: Schrift bleibt, gilt als entfallen
            roh.append((ch, ALT, st.basis, True, aid, ALT, p, p))
        elif aid:
            roh.append((ch, ALT, st.basis, False, aid, ALT, p, p))
        else:
            roh.append((ch, st.format, st.basis, True, "", NORMAL, p, p))

    g = 0
    while g <= n:
        unsichtbar_an(g)
        einsetzen(g)
        if g in bereich_ab:
            e, ent = bereich_ab[g]
            naechste = g + 1
            for ch, kl, r, rl, aid in ent:
                q = g + (r if r >= 0 else rl)
                while naechste < e and naechste <= q:
                    unsichtbar_an(naechste)
                    naechste += 1
                if r >= 0:
                    if not 0 <= g + r < e or orig_text[g + r] != ch:
                        raise ValueError("Projektion passt nicht zum Zelltext")
                    zeichen_an(g + r, aid if kl == ALT else "")
                else:
                    roh.append((ch, NEU if kl == NEU else NORMAL, basis_an(q), False, aid, kl, -1, q))
            while naechste < e:
                unsichtbar_an(naechste)
                naechste += 1
            if e > g:
                g = e
                continue
        if g < n:
            zeichen_an(g, loesch[g])
        g += 1
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


def _gemeinsam(alt: str, neu: str, cache: dict[str, str], cfg: dict | None) -> tuple[int, int]:
    """Wie viele Zeichen vorn und hinten bei einer Ersetzung gleich bleiben und deshalb nicht markiert werden. Leerraum gilt
    als gleich, Schreibvarianten auch (Normalform: U+2011 = "-", typografische Anfuehrungszeichen ...). Gekuerzt wird nur an
    Leerraum, damit ein Wort immer ganz markiert wird wie im PDF ("Wert" -> "Werte", "2.3" -> "2.2": ganzes Wort)."""
    from .gegenprobe import normalform

    def form(c: str) -> str:
        if c.isspace():
            return " "
        if c not in cache:
            cache[c] = normalform(c, cfg)[0] or c
        return cache[c]

    def grenze(s: str, k: int) -> bool:
        return k <= 0 or k >= len(s) or s[k].isspace() or s[k - 1].isspace()

    k = 0
    while k < len(alt) and k < len(neu) and form(alt[k]) == form(neu[k]):
        k += 1
    if not (grenze(alt, k) and grenze(neu, k)):
        k = max((i + 1 for i in range(k) if alt[i].isspace()), default=0)
    m = 0
    while m < len(alt) - k and m < len(neu) - k and form(alt[-1 - m]) == form(neu[-1 - m]):
        m += 1
    if not (grenze(alt, len(alt) - m) and grenze(neu, len(neu) - m)):
        m = max((len(alt) - i for i in range(len(alt) - m, len(alt)) if alt[i].isspace()), default=0)
    return k, m


# ------------------------------------------------------------------ Markierung aus dem PDF projizieren (Nacharbeit Auftrag 003)
def _soll_zeichen(laeufe: Iterable[Lauf], aktiv: set[str], trenner: str = "") -> list[list[Any]]:
    """Soll eines PDF-Absatzes je Zeichen: [zeichen, klasse, optional, aenderung] (wie zustand_markiert).
    Ein Trenner (Leerraum zwischen zwei Woertern) ist im PDF nicht markiert und bleibt normal -- ausser er liegt zwischen
    zwei Zeichen derselben markierten Klasse ("{-ISO 44556-}", "{+neuer Text+}"): dann gehoert er zur Markierung."""
    sichtbar: list[Lauf] = []
    for l in laeufe:
        if l.art == "a" and l.aenderung not in aktiv:
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
    out: list[list[Any]] = []
    vorige = ""
    for l in sichtbar:
        k = ALT if (l.art == "d" and l.aenderung in aktiv) else (NEU if l.art == "a" else NORMAL)
        if vorige == ALT and k == NEU:
            out.extend([c, NORMAL, False, ""] for c in trenner)
        out.extend([c, k, l.optional, l.aenderung if k != NORMAL else ""] for c in l.text)
        vorige = k
    fest = [i for i, z in enumerate(out) if not z[0].isspace()]
    behalten = _alter_trenner_zwischen_neu(out, fest)
    for a, b in zip(fest, fest[1:]):
        if b - a > 1 and out[a][1] == out[b][1] != NORMAL and a not in behalten:
            for k in range(a + 1, b):
                out[k][1], out[k][3] = out[a][1], out[a][3]
    return out


def _alter_trenner_zwischen_neu(out: list[list[Any]], fest: list[int]) -> set[int]:
    """Welcher unmarkierte Trenner zwischen zwei neuen Woertern ist der EINZIGE Leerraum des alten Textes an dieser Stelle?

    Beispiel Listennummer: "{-c.-}{+d.+}" (alte Nummer gestrichen, neue direkt dahinter), dann das unmarkierte Polster der Excel
    (Tab, geschuetzte Leerzeichen), dann ein neues erstes Wort "{+Besondere +}", dann der alte Text. Das Polster steht zwischen
    zwei neuen Stuecken, gehoert aber zum alten Text (zustand() behaelt es auch): es bleibt normal und darf nicht zur Markierung
    werden, sonst fehlt dem alten Text der Leerraum und die Zuordnung zur Excel scheitert. Rueckgabe: Anfangsindizes `a` der
    Luecken (a, b) zwischen zwei festen Zeichen, die nicht umgefaerbt werden. Gilt nur, wenn vor UND hinter dem neuen Block
    altes Zeichen steht und dort kein anderer Leerraum des alten Textes liegt; bei mehreren Luecken gewinnt wie in zustand() der
    staerkste Trenner, bei Gleichstand der erste."""
    behalten: set[int] = set()
    k = 0
    while k < len(fest):
        if out[fest[k]][1] != NEU:
            k += 1
            continue
        k0 = k
        while k + 1 < len(fest) and out[fest[k + 1]][1] == NEU:
            k += 1
        links, rechts = (fest[k0 - 1] if k0 > 0 else None), (fest[k + 1] if k + 1 < len(fest) else None)
        k += 1
        if links is None or rechts is None:
            continue
        aussen = list(range(links + 1, fest[k0])) + list(range(fest[k - 1] + 1, rechts))
        if any(out[i][1] != NEU for i in aussen):
            continue                                    # der alte Text hat dort schon Leerraum
        luecken = [(a, b) for a, b in zip(fest[k0:k], fest[k0 + 1:k])
                   if b - a > 1 and all(out[i][1] == NORMAL for i in range(a + 1, b))]
        if luecken:
            a, b = max(luecken, key=lambda ab: (_RANG.get("".join(out[i][0] for i in range(ab[0] + 1, ab[1])), 1), -ab[0]))
            behalten.add(a)
    return behalten


_NUMMER = re.compile(r"^\(?[\w.]{1,6}[.)]$")        # Listennummer: "3.", "d.", "2.1.", "a)", "(iv)"


def _nummer_ersetzt(soll: list[list[Any]], x_: int | None, b0: int, b1: int, r_text: list[str], lo: int, hi: int) -> bool:
    """Ersetzt der neue Text eine Listennummer, die in der Excel von Polster (geschuetzte Leerzeichen, Tab, mehrfacher
    Leerraum) gefolgt wird, und beginnt er selbst mit einer Nummer und hat mehr Text dahinter?"""
    if x_ is None or soll[x_][1] != ALT or not lo < hi or not all(r_text[y].isspace() for y in range(lo, hi)):
        return False
    if hi - lo < 2 and r_text[lo] == " ":
        return False                              # ein einzelnes Leerzeichen ist kein Polster
    a = x_
    while a > 0 and soll[a - 1][1] == ALT:
        a -= 1
    alt = "".join(z[0] for z in soll[a:x_ + 1]).strip()
    text = "".join(soll[s][0] for s in range(b0, b1 + 1))
    teile = text.split(None, 1)
    return bool(_NUMMER.match(alt) and not text[:1].isspace() and len(teile) == 2 and _NUMMER.match(teile[0]))


def projiziere(texte: list[str], laeufe: Iterable[Lauf], aktiv: set[str],
               cfg: dict | None = None) -> list[list[tuple[str, str, int, int, str]]] | None:
    """Markierung eines PDF-Absatzes auf die Excel-Zeichen uebertragen (reine Funktion).

    texte: der Absatz-Bereich je Zelle in der Sicht des Excel-Modells (mehrere Zellen werden wie in der Gegenprobe mit einem
    Leerzeichen verbunden). Jedes Excel-Zeichen bekommt ueber die Normalform-Zuordnung die Klasse seines PDF-Gegenstuecks
    (normal | alt) und bleibt selbst unveraendert -- auch Excel-eigene Zeichen wie U+2011 oder doppelter Leerraum. Neue
    PDF-Laeufe werden an ihrer Stelle eingefuegt, mit genau dem Leerraum, den das PDF zeigt, ohne vorhandenen Excel-Leerraum
    zu verdoppeln; dieser Leerraum ist nicht markiert.
    Rueckgabe je Zelle: Eintraege (zeichen, klasse, pos | -1, luecke, aenderung) in Ausgabereihenfolge, Positionen relativ
    zum Bereich; None = keine sichere Zuordnung (Text passt nicht, der Aufrufer bleibt bei den Operationen)."""
    from .gegenprobe import normalform
    if not texte:
        return None
    soll = _soll_zeichen(laeufe, aktiv, str(konfig(cfg).get("trenner", "")))
    alt_idx = [i for i, z in enumerate(soll) if z[1] != NEU]
    r_text: list[str] = []
    grenzen: list[tuple[int, int]] = []            # (Anfang, Ende) je Zelle in r_text
    for k, t in enumerate(texte):
        if k:
            r_text.append(" ")                     # gedachtes Leerzeichen zwischen zwei Zellen
        grenzen.append((len(r_text), len(r_text) + len(t)))
        r_text.extend(t)
    ns, ms = normalform("".join(soll[i][0] for i in alt_idx), cfg)
    nr, mr = normalform("".join(r_text), cfg)
    abb: dict[int, int] = {}                       # Soll-Index -> Excel-Index
    i = j = 0
    while i < len(ns) and j < len(nr):
        si = alt_idx[ms[i]]
        if ns[i] == nr[j]:
            abb.setdefault(si, mr[j])
            i += 1
            j += 1
        elif soll[si][2] or (ns[i] == " " and i + 1 < len(ns) and soll[alt_idx[ms[i + 1]]][2]):
            i += 1
        else:
            return None
    while i < len(ns) and (ns[i] == " " or soll[alt_idx[ms[i]]][2]):
        i += 1
    if i < len(ns) or j < len(nr):
        return None
    n = len(r_text)
    zelle_von = [-1] * n
    for k, (a, e) in enumerate(grenzen):
        for x in range(a, e):
            zelle_von[x] = k
    klasse: list[str | None] = [None] * n
    aid: list[str] = [""] * n
    for si, rj in abb.items():
        klasse[rj], aid[rj] = soll[si][1], soll[si][3]

    def nachbar(x: int, schritt: int) -> int | None:
        x += schritt
        while 0 <= x < n and (r_text[x].isspace() or klasse[x] is None):
            x += schritt
        return x if 0 <= x < n else None

    for x in range(n):                             # Excel-eigene Zeichen ohne PDF-Gegenstueck: innen alt, sonst normal
        if klasse[x] is None and not r_text[x].isspace():
            li, re_ = nachbar(x, -1), nachbar(x, 1)
            alt_beide = li is not None and re_ is not None and klasse[li] == klasse[re_] == ALT
            klasse[x], aid[x] = (ALT, aid[li]) if alt_beide else (NORMAL, "")
    x = 0
    while x < n:                                   # Leerraum: wie sein PDF-Gegenstueck, sonst alt nur zwischen zwei alten
        if not r_text[x].isspace():
            x += 1
            continue
        e = x
        while e < n and r_text[e].isspace():
            e += 1
        gesetzt = [y for y in range(x, e) if klasse[y] is not None]
        if gesetzt:
            k_, a_ = klasse[gesetzt[0]], aid[gesetzt[0]]
        else:
            li, re_ = nachbar(x, -1), nachbar(e - 1, 1)
            alt_beide = li is not None and re_ is not None and klasse[li] == klasse[re_] == ALT
            k_, a_ = (ALT, aid[li]) if alt_beide else (NORMAL, "")
        for y in range(x, e):
            klasse[y], aid[y] = k_, a_
        x = e
    einf: dict[int, list[list[tuple[str, str, str]]]] = {}

    def fest(si: int) -> bool:
        return si in abb and not soll[si][0].isspace()

    k = 0
    while k < len(soll):
        if soll[k][1] != NEU:
            k += 1
            continue
        b0 = k
        while k < len(soll) and soll[k][1] == NEU:
            k += 1
        b1 = k - 1
        x_ = next((s for s in range(b0 - 1, -1, -1) if fest(s)), None)
        y_ = next((s for s in range(b1 + 1, len(soll)) if fest(s)), None)
        vor = "".join(soll[s][0] for s in range(x_ + 1 if x_ is not None else 0, b0) if soll[s][0].isspace())
        nach = "".join(soll[s][0] for s in range(b1 + 1, y_ if y_ is not None else len(soll)) if soll[s][0].isspace())
        lo = abb[x_] + 1 if x_ is not None else 0
        hi = abb[y_] if y_ is not None else n
        if lo > hi:
            return None
        if _nummer_ersetzt(soll, x_, b0, b1, r_text, lo, hi):
            # Die Listennummer selbst wird ersetzt ("{-d.-}{+f. Besondere +}", Excel: "d." + Polster/Tab + "Ausfuehrung"):
            # die neue Nummer steht direkt hinter der alten, der uebrige neue Text HINTER dem Polster (nicht an der Nummer)
            text = "".join(soll[s][0] for s in range(b0, b1 + 1))
            k1 = len(text.split(None, 1)[0])
            k2 = len(text) - len(text[k1:].lstrip())
            a_ = soll[b0][3]
            rest = [(soll[s][0], NEU, soll[s][3] or a_) for s in range(b0 + k2, b1 + 1)]
            while nach and rest and rest[-1][0].isspace():
                rest.pop()                         # ein unmarkierter Trenner dahinter ersetzt das markierte Leerzeichen
            einf.setdefault(lo, []).append([(soll[s][0], NEU, soll[s][3] or a_) for s in range(b0, b0 + k1)])
            # der Excel-Leerraum der Nummer steht VOR dem neuen Text; was das PDF hinter dem neuen Text unmarkiert zeigt, fehlt dort
            einf.setdefault(hi, []).append(rest + [(c, NORMAL, a_) for c in nach])
            continue
        if not vor:
            p = lo                                 # klebt am Zeichen davor
        elif not nach:
            p = hi                                 # klebt am Zeichen dahinter
        elif lo < hi and zelle_von[lo] < 0:
            p = lo                                 # Zellgrenze: der neue Text gehoert ans Ende der Zelle davor
        else:
            # Leerraum auf beiden Seiten: der neue Text kommt HINTER den ganzen Excel-Leerraum, vor das naechste behaltene
            # Zeichen ("3)" + viele NBSP + "Allgemeine" -> "3)<NBSP...> {+Weitere +}Allgemeine"), nie mitten hinein
            p = lo
            while p < hi and r_text[p].isspace():
                p += 1
        if x_ is None or any(c.isspace() for c in r_text[lo:p]):
            vor = ""                               # Excel hat den Leerraum davor schon
        if y_ is None or any(c.isspace() for c in r_text[p:hi]):
            nach = ""
        a_ = soll[b0][3]
        blk = [(soll[s][0], NEU, soll[s][3] or a_) for s in range(b0, b1 + 1)]
        # markierter Leerraum am Rand des neuen Textes (LESEN: Teil des a-Laufs) ist rot -- ausser die Excel hat dort schon
        # Leerraum (oder er steht ohnehin unmarkiert davor/dahinter): ein vorhandenes Leerzeichen wird nie gedoppelt
        while blk and blk[0][0].isspace() and (vor or (p > 0 and r_text[p - 1].isspace())):
            blk.pop(0)
        while blk and blk[-1][0].isspace() and (nach or (p < n and r_text[p].isspace())):
            blk.pop()
        if not blk:
            continue
        einf.setdefault(p, []).append([(c, NORMAL, a_) for c in vor] + blk + [(c, NORMAL, a_) for c in nach])

    def ort(p: int) -> tuple[int, int]:
        for z, (a, e) in enumerate(grenzen):
            if a <= p <= e:
                return z, p - a
        return len(grenzen) - 1, grenzen[-1][1] - grenzen[-1][0]

    out: list[list[tuple[str, str, int, int, str]]] = [[] for _ in texte]
    for p in range(n + 1):
        for block in einf.get(p, ()):
            z, lu = ort(p)
            out[z].extend((c, kl, -1, lu, a_) for c, kl, a_ in block)
        if p < n and zelle_von[p] >= 0:
            z, rp = ort(p)
            kl = klasse[p] or NORMAL
            out[z].append((r_text[p], kl, rp, rp, aid[p] if kl != NORMAL else ""))
    return out


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
