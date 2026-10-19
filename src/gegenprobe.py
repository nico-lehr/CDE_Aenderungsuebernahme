"""Gegenprobe ("doppelte Buchfuehrung", ADR 0004): reine Funktionen, keine Excel-/PDF-Bibliothek.

Ein PDF-Absatz ist eine Liste von Laeufen (models.Lauf): k = bleibt, d = entfaellt, a = neu, s = Trenner.
Daraus laesst sich fuer JEDE Menge angewendeter Aenderungen der Absatztext bauen (`zustand`):
  * zustand(laeufe, set())        = der alte Absatz (so muss er VORHER in der Excel stehen, Pruefung A1)
  * zustand(laeufe, alle IDs)     = der neue Absatz (so verlangt ihn das PDF)
  * zustand(laeufe, angewendet)   = so muss er NACH dem Speichern in der Excel stehen (Pruefung A3)

Verglichen wird in einer "Normalform" (streng, aber fair): Gross/Klein und Satzzeichen zaehlen, Leerraum-Varianten,
Anfuehrungszeichen-/Strich-Varianten, weiche Trennzeichen und Hochzahlen nicht (Tabellen in config/zuordnung.yaml
`normalform:`). Keine Dokumentinhalte in Rueckgaben fuer Berichte: `Vergleich.art` ist nur eine Kategorie (Regel 16).
"""
from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field
from typing import Any, Iterable

from . import config
from .models import Absatz, Lauf, Operation

# Rang beim Zusammenlegen von Trennern (der "staerkste" bleibt): geklebt < Leerzeichen/Tab < geschuetzt < Zeilenumbruch
_RANG = {"": 0, " ": 1, "\t": 1, " ": 2, " ": 2, "\n": 3}


def _cfg(cfg: dict | None) -> dict:
    if cfg and "normalform" in cfg:
        return cfg["normalform"]
    if cfg and "entfernen" in cfg:                     # direkt der Abschnitt
        return cfg
    return config.lade("zuordnung").get("normalform", {})


# ---------------------------------------------------------------- Zustand eines Absatzes
def _gefiltert(laeufe: Iterable[Lauf], aktiv: set[str]) -> list[Lauf]:
    """Laeufe, die im Zustand `aktiv` sichtbar sind; Trenner an entfernten Stuecken werden zusammengelegt."""
    sichtbar: list[Lauf] = []
    for l in laeufe:
        if l.art == "d" and l.aenderung in aktiv:
            continue
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
    return sichtbar


def zustand_mit_optional(laeufe: Iterable[Lauf], aktiv: set[str] | None = None) -> tuple[str, frozenset[int]]:
    """Wie `zustand`, dazu die Positionen der optionalen Zeichen (duerfen in der Excel fehlen)."""
    teile: list[str] = []
    optional: set[int] = set()
    pos = 0
    for l in _gefiltert(laeufe, set(aktiv or ())):
        if l.optional:
            optional.update(range(pos, pos + len(l.text)))
        teile.append(l.text)
        pos += len(l.text)
    return "".join(teile), frozenset(optional)


def zustand(laeufe: Iterable[Lauf], aktiv: set[str] | None = None) -> str:
    """Text des Absatzes, wenn genau die Aenderungen in `aktiv` angewendet sind.

    Beispiele: "Anlage␣{-insgesamt-}." -> "Anlage ." (Normalform "Anlage."); "darf␣{-12-}␣{+7+}␣ms" -> "darf 7 ms";
    "{-Zudem-}␣gilt" -> "gilt"; "x␣{-alt-}{+neu+}␣y" -> "x neu y"."""
    return zustand_mit_optional(laeufe, aktiv)[0]


# ---------------------------------------------------------------- Normalform
def normalform(text: str, cfg: dict | None = None) -> tuple[str, list[int]]:
    """Vergleichsform und Abbildung: Zeichen i der Normalform stammt aus Zeichen abbildung[i] des Rohtexts."""
    c = _cfg(cfg)
    entfernen = set(c.get("entfernen", []))
    ersetzen = dict(c.get("ersetzen", {}))
    leerraum = set(c.get("leerraum", [])) | {" "}
    vor = set(c.get("kein_leerzeichen_vor", ""))
    nach = set(c.get("kein_leerzeichen_nach", ""))
    listen = tuple(c.get("listenzeichen", []))

    roh: list[tuple[str, int]] = []
    for i, z in enumerate(text):
        if z in entfernen:
            continue
        if z in leerraum or z.isspace():
            roh.append((" ", i))
            continue
        z = ersetzen.get(z, z)
        for t in unicodedata.normalize("NFKC", z):
            if unicodedata.combining(t) and roh and roh[-1][0] != " ":
                roh[-1] = (unicodedata.normalize("NFC", roh[-1][0] + t), roh[-1][1])
                continue
            roh.append((ersetzen.get(t, t), i))
    # Leerraum zusammenlegen, am Rand und vor/nach Klammern und Satzzeichen weg
    out: list[tuple[str, int]] = []
    for z, i in roh:
        if z == " ":
            if not out or out[-1][0] == " " or out[-1][0] in nach:
                continue
            out.append((z, i))
            continue
        if z in vor and out and out[-1][0] == " ":
            out.pop()
        out.append((z, i))
    while out and out[-1][0] == " ":
        out.pop()
    # optionales Listenzeichen am Anfang ("• Text" == "Text")
    if len(out) >= 2 and out[0][0] in listen and out[1][0] == " ":
        out = out[2:]
    return "".join(z for z, _ in out), [i for _, i in out]


# ---------------------------------------------------------------- Vergleich
@dataclass
class Vergleich:
    gleich: bool
    abbildung: dict[int, int] = field(default_factory=dict)   # Rohindex PDF-Text -> Rohindex Excel-Text (gleiche Zeichen)
    art: str = ""        # erste Abweichung: "" | leerraum | satzzeichen | gross_klein | ziffer | buchstabe | pdf_laenger | excel_laenger


def _art(a: str, b: str) -> str:
    if a == " " or b == " ":
        return "leerraum"
    if a.casefold() == b.casefold():
        return "gross_klein"
    if a.isdigit() or b.isdigit():
        return "ziffer"
    if not a.isalnum() or not b.isalnum():
        return "satzzeichen"
    return "buchstabe"


def vergleiche(pdf: str, excel: str, optional: frozenset[int] = frozenset(), cfg: dict | None = None) -> Vergleich:
    """Ist der PDF-Text (mit optionalen Zeichen) in Normalform gleich dem Excel-Text?"""
    np_, pm = normalform(pdf, cfg)
    nx, xm = normalform(excel, cfg)
    i = j = 0
    abb: dict[int, int] = {}
    while i < len(np_) and j < len(nx):
        if np_[i] == nx[j]:
            abb[pm[i]] = xm[j]
            i += 1
            j += 1
        elif pm[i] in optional:
            i += 1
        elif np_[i] == " " and i + 1 < len(np_) and pm[i + 1] in optional:   # "Ver- sicherung": Leerzeichen am optionalen Strich
            i += 1
        else:
            return Vergleich(False, abb, _art(np_[i], nx[j]))
    while i < len(np_) and (pm[i] in optional or np_[i] == " "):
        i += 1
    if i < len(np_):
        return Vergleich(False, abb, "pdf_laenger")
    if j < len(nx):
        return Vergleich(False, abb, "excel_laenger")
    return Vergleich(True, abb, "")


def gleich(pdf: str, excel: str, optional: frozenset[int] = frozenset(), cfg: dict | None = None) -> bool:
    return vergleiche(pdf, excel, optional, cfg).gleich


# ---------------------------------------------------------------- A3: Pruefung nach dem Speichern
def _delta(op: Operation) -> int:
    return len(op.text) - (op.ende - op.start)


def neue_spanne(start: int, ende: int, ops: Iterable[Operation], eigene: set[str]) -> tuple[int, int]:
    """Lage des Bereichs [start, ende) des Originaltexts, nachdem `ops` angewendet wurden.

    Eigene Operationen (Aenderung in `eigene`) liegen im Bereich und aendern seine Laenge; fremde liegen ausserhalb
    und verschieben ihn nur, wenn sie davor liegen."""
    s, e = start, ende
    for op in ops:
        if op.art not in ("ersetzen", "einfuegen"):
            continue
        d = _delta(op)
        if op.aenderung in eigene:
            e += d
            if op.start < start:                           # beginnt vor dem Bereich (z. B. Zeilenumbruch davor mit weg):
                s -= start - op.start                      # der Bereich beginnt dann an der Operation
        elif op.ende <= start and not (op.start == op.ende == start):
            s += d
            e += d
        elif op.start == op.ende == start:                 # fremde Einfuegung genau am Anfang gilt als davor
            s += d
            e += d
    # Eine eigene Loeschung, die ueber den Bereich hinausreicht (ganzer Absatz samt Zeilenumbruch davor/danach),
    # macht ihn leer -- nie kuerzer als leer
    return s, max(s, e)


def pruefe_absatz(absatz: Absatz, zellen: dict[tuple[int, int], tuple[str, str, list[Operation]]],
                  angewendet: set[str], cfg: dict | None = None) -> str | None:
    """A3: Steht der Absatz nach dem Speichern so in der Excel, wie das PDF es fuer die angewendeten Aenderungen verlangt?

    zellen: (zeile, spalte) -> (Originaltext, gespeicherter Text, angewendete Operationen in dieser Zelle).
    Rueckgabe None = in Ordnung, sonst ein Code: gegenprobe_text | gegenprobe_umbruch | gegenprobe_aussen | gegenprobe_zelle."""
    eigene = absatz.aenderungen
    stuecke: list[str] = []
    for zeile, spalte, start, ende in absatz.zellen:
        daten = zellen.get((zeile, spalte))
        if daten is None:
            return "gegenprobe_zelle"
        orig, gesp, ops = daten
        s, e = neue_spanne(start, ende, ops, eigene)
        if not (0 <= s <= e <= len(gesp)):
            return "gegenprobe_aussen"
        stueck = gesp[s:e]
        if stueck.count("\n") > orig[start:ende].count("\n"):
            return "gegenprobe_umbruch"
        if not [o for o in ops if o.aenderung not in eigene]:     # nur eigene Aenderungen in dieser Zelle: Rest bytegleich
            vorn, hinten = orig[:start], orig[ende:]
            if gesp[:s] != vorn or gesp[e:] != hinten:
                # Ein ganz entfernter Absatz nimmt einen angrenzenden Zeilenumbruch mit -- das ist erlaubt
                if not (stueck == "" and (gesp == vorn.rstrip("\n") + hinten or gesp == vorn + hinten.lstrip("\n")
                                          or gesp == (vorn + hinten).strip("\n"))):
                    return "gegenprobe_aussen"
        stuecke.append(stueck)
    soll, optional = zustand_mit_optional(absatz.laeufe, set(angewendet) & eigene)
    if not gleich(soll, " ".join(stuecke), optional, cfg):
        return "gegenprobe_text"
    return None


def absaetze_aus_liste(daten: Any) -> list[Absatz] | None:
    """plan.json `absaetze` lesen; fehlt das Feld (Schema <= 2), None (dann gilt die alte Nachkontrolle)."""
    if not isinstance(daten, list):
        return None
    return [Absatz.aus_dict(d) for d in daten]


# ---------------------------------------------------------------- A3 mit Markierung (ADR 0005 Nachtrag, ADR 0007)
# Mit markierung.aenderungen = farbe bleibt der alte Text stehen (blau + durchgestrichen) und der neue steht rot dahinter.
# Geprueft wird je Absatz gegen das PDF: Alt-Ansicht (normal + alt) = alter Absatz, Neu-Ansicht (normal + neu) = Absatz laut
# PDF fuer die angewendeten Aenderungen, Formatklasse je Zeichen = Soll, und kein Zeichen des Originals geht verloren.
_N, _A, _U = "normal", "alt", "neu"


def zustand_markiert_zeichen(laeufe: Iterable[Lauf], aktiv: set[str], trenner: str = "") -> list[tuple[str, str, bool]]:
    """Soll eines Absatzes mit Markierung je Zeichen: (zeichen, klasse, optional). Wie markierung.zustand_markiert:
    entfallende Laeufe angewendeter Aenderungen bleiben (alt), neue Laeufe sind neu, nicht angewendete neue fehlen."""
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
    out: list[tuple[str, str, bool]] = []
    vorige = ""
    for l in sichtbar:
        k = _A if (l.art == "d" and l.aenderung in aktiv) else (_U if l.art == "a" else _N)
        if vorige == _A and k == _U:
            out.extend((c, _N, False) for c in trenner)
        out.extend((c, k, l.optional) for c in l.text)
        vorige = k
    return out


def _klassen_gleich(soll: list[tuple[str, str, bool]], ist: list[tuple[str, str]], cfg: dict | None) -> bool:
    """Gleiche Zeichen (Normalform, ohne Leerraum) mit gleicher Formatklasse; optionale Soll-Zeichen duerfen fehlen."""
    ns, ms = normalform("".join(c for c, _, _ in soll), cfg)
    ni, mi = normalform("".join(c for c, _ in ist), cfg)
    a = [(c, soll[i][1], soll[i][2]) for c, i in zip(ns, ms) if c != " "]
    b = [(c, ist[i][1]) for c, i in zip(ni, mi) if c != " "]
    i = j = 0
    while i < len(a) and j < len(b):
        if a[i][0] == b[j][0] and a[i][1] == b[j][1]:
            i += 1
            j += 1
        elif a[i][2]:
            i += 1
        else:
            return False
    while i < len(a) and a[i][2]:
        i += 1
    return i == len(a) and j == len(b)


def pruefe_absatz_markiert(absatz: Absatz, zellen: dict[tuple[int, int], tuple[str, list[tuple[str, str, int, int, str]]]],
                           angewendet: set[str], cfg: dict | None = None, trenner: str = "",
                           klassen: bool = True) -> str | None:
    """A3 mit Markierung: Steht der Absatz nach dem Schreiben so in der Excel, wie das PDF es verlangt?

    zellen: (zeile, spalte) -> (Originaltext in der Sicht des Excel-Modells, Zeichen der Zelle nach dem Schreiben als
    (zeichen, klasse, pos, luecke, aenderung) -- siehe markierung.Markierung). Rueckgabe None = in Ordnung, sonst ein Code:
    gegenprobe_zelle | gegenprobe_verlust (ein Zeichen des Originals fehlt) | gegenprobe_aussen (eigene Markierung ausserhalb
    des Absatzes) | gegenprobe_alt | gegenprobe_text | gegenprobe_umbruch | gegenprobe_format.
    klassen=False: ohne den Vergleich der Formatklassen (Pruefung des Plans: stimmt der Text, den die Operationen ergeben?)."""
    eigene = absatz.aenderungen
    aktiv = set(angewendet) & eigene
    ist: list[tuple[str, str]] = []
    alt_umbrueche = 0
    for k, (zeile, spalte, start, ende) in enumerate(absatz.zellen):
        daten = zellen.get((zeile, spalte))
        if daten is None:
            return "gegenprobe_zelle"
        vorher, zeichen = daten
        if "".join(z[0] for z in zeichen if z[2] >= 0) != vorher or [z[2] for z in zeichen if z[2] >= 0] != list(range(len(vorher))):
            return "gegenprobe_verlust"
        if not 0 <= start <= ende <= len(vorher):
            return "gegenprobe_aussen"
        alt_umbrueche += vorher[start:ende].count("\n")
        region: list[tuple[str, str]] = []
        for ch, kl, pos, lu, aid in zeichen:
            if pos >= 0:
                drin = start <= pos < ende
                if not drin and aid and aid in eigene and kl != _N and not ch.isspace():
                    return "gegenprobe_aussen"
            elif not aid:
                continue                # ohne Aenderung und ohne Position: aeltere Markierung ausserhalb der Ansicht
            else:
                drin = start < lu < ende or (lu in (start, ende) and bool(aid) and aid in eigene)
                if not drin and aid and aid in eigene and not ch.isspace():
                    return "gegenprobe_aussen"
            if drin:
                region.append((ch, kl))
        if k:
            ist.append((" ", _N))
        ist.extend(region)
    alt_ist = "".join(c for c, kl in ist if kl != _U)
    neu_ist = "".join(c for c, kl in ist if kl != _A)
    alt_soll, opt_alt = zustand_mit_optional(absatz.laeufe, set())
    neu_soll, opt_neu = zustand_mit_optional(absatz.laeufe, aktiv)
    if not gleich(alt_soll, alt_ist, opt_alt, cfg):
        return "gegenprobe_alt"
    if neu_ist.count("\n") > max(alt_umbrueche, neu_soll.count("\n")):
        return "gegenprobe_umbruch"
    if not gleich(neu_soll, neu_ist, opt_neu, cfg):
        return "gegenprobe_text"
    if klassen and not _klassen_gleich(zustand_markiert_zeichen(absatz.laeufe, aktiv, trenner), ist, cfg):
        return "gegenprobe_format"
    return None
