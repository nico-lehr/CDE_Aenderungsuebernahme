"""Absatz-Abgleich "wie ein sorgfaeltiger Mensch" (ADR 0004, Bausteine B und D).

Ein Mensch liest den ganzen Absatz alt (bleibt + entfaellt) und neu (bleibt + neu), sucht die Stelle in der Excel in
Dokument-Reihenfolge, traegt nur dann ohne Nachfrage ein, wenn der alte Absatz dort GENAU so steht, und uebertraegt die
Aenderungen Zeichen fuer Zeichen; was die Excel selbst an Zeichen hat (geschuetztes Leerzeichen, Tabulator ...), bleibt.

Bausteine (reine Logik, schreibt nie in eine Datei):
  * `laeufe_aus_tokens`  PDF-Woerter -> Laeufe (k bleibt, d entfaellt, a neu, s Trenner) mit einer ID je Aenderung
  * `spuren`             PDF-Bloecke -> logische Absaetze (ueber Seitenumbrueche verbunden)
  * `ExaktIndex`         Normalform jedes Excel-Absatzes / jeder Ueberschrift -> Fundstellen (auch Teilstuecke)
  * `reissverschluss`    laengste aufsteigende Kette (PDF-Reihenfolge = Excel-Reihenfolge) -> Status je Absatz
  * `ops_fuer_spur`      A1 (alter Absatz steht exakt da?) + B (je Aenderung eine Operation `ersetzen`)
Keine Dokumentinhalte in Codes/Gruenden fuer Berichte (Regel 16).
"""
from __future__ import annotations

import bisect
import re
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Any, Iterable, Sequence

from . import config
from .gegenprobe import _RANG, Vergleich, normalform, vergleiche, zustand, zustand_mit_optional
from .models import ADD, DEL, KEEP, Lauf, Operation, PdfBlock, PdfToken
from .textutil import kopf_zerlegen, norm

if TYPE_CHECKING:                      # nur fuer Typangaben
    from .excel_model import ExcelModell, Kandidat

# Status je Absatz nach dem Reissverschluss
EINDEUTIG, REIHENFOLGE, MEHRDEUTIG, WIDERSPRUCH, KEINER = "eindeutig", "durch_reihenfolge", "mehrdeutig", "widerspruch", "keiner"
AUTOMATISCH_ERLAUBT = (EINDEUTIG, REIHENFOLGE)

_ART = {KEEP: "k", DEL: "d", ADD: "a"}


def _zuordnung_cfg(cfg: dict | None) -> dict:
    return cfg if cfg and ("planer" in cfg or "matcher" in cfg) else config.lade("zuordnung")


# =============================================================================================== 1. Laeufe
def _wort_teile(t: PdfToken) -> list[tuple[str, str]]:
    """(art, text) eines Worts: Zeichen-Laeufe vom PDF-Leser oder das ganze Wort."""
    if t.teile:
        return [(_ART.get(st, "k"), tx) for st, tx in t.teile if tx]
    if t.status == KEEP:
        return [("k", t.old or t.new)]
    if t.status == DEL:
        return [("d", t.old)]
    if t.status == ADD:
        return [("a", t.new)]
    out = []                                      # MOD ohne Zeichen-Laeufe: alt weg, neu hin (aneinandergeklebt)
    if t.old:
        out.append(("d", t.old))
    if t.new:
        out.append(("a", t.new))
    return out


def _mit_trennstrich(t: PdfToken, teile: list[tuple[str, str]]) -> list[tuple[str, str, bool]]:
    """Ein Trennstrich an der Zeilenend-Trennung (`trennstelle`) wird ein eigener, optionaler Lauf."""
    out = [(a, x, False) for a, x in teile]
    if t.trennstelle < 0 or len(teile) != 1:
        return out
    art, text = teile[0]
    p = t.trennstelle
    if 0 <= p < len(text) and text[p] == "-":
        out = [(art, text[:p], False), (art, "-", True), (art, text[p + 1:], False)]
        out = [x for x in out if x[1]]
    return out


@dataclass
class _Aufbau:
    laeufe: list[Lauf]
    lauf_token: list[int]                 # Token-Index je Lauf (bei Trennern: das Wort davor)
    gruppen: list[list[int]]              # Lauf-Indizes je Aenderung (Gruppe); ID = praefix + Nummer (ab 1)
    gruppen_token: list[list[int]]        # Token-Indizes je Gruppe
    zaehl_token: list[list[int]]          # Token, die in dieser Gruppe als markierte Woerter zaehlen (jedes genau einmal)


def _aufbau(tokens: Sequence[PdfToken], praefix: str = "G") -> _Aufbau:
    roh: list[tuple[str, str, bool, int]] = []
    for i, t in enumerate(tokens):
        for art, text, opt in _mit_trennstrich(t, _wort_teile(t)):
            roh.append((art, text, opt, i))
        if i + 1 < len(tokens) and t.trenner != "":
            roh.append(("s", t.trenner or " ", False, i))
    # Gruppen: maximale Folge von d/a-Laeufen (Trenner dazwischen unterbrechen nicht, ein k-Lauf schon)
    gruppe_von: list[int] = []
    gruppen: list[list[int]] = []
    offen = False
    for k, (art, _, _, _) in enumerate(roh):
        if art in ("d", "a"):
            if not offen:
                gruppen.append([])
                offen = True
            gruppen[-1].append(k)
            gruppe_von.append(len(gruppen) - 1)
        else:
            if art == "k":
                offen = False
            gruppe_von.append(-1)
    laeufe = [Lauf(art, text, f"{praefix}{gruppe_von[k] + 1}" if gruppe_von[k] >= 0 else "", opt)
              for k, (art, text, opt, _) in enumerate(roh)]
    lauf_token = [i for _, _, _, i in roh]
    gruppen_token = [sorted({lauf_token[k] for k in g}) for g in gruppen]
    # Jedes markierte Wort zaehlt genau einmal (Bilanz): in der ersten Gruppe, in der es vorkommt
    zaehl: list[list[int]] = [[] for _ in gruppen]
    gezaehlt: set[int] = set()
    for gi, toks in enumerate(gruppen_token):
        for i in toks:
            if i not in gezaehlt and tokens[i].status != KEEP:
                zaehl[gi].append(i)
                gezaehlt.add(i)
    if gruppen:
        for i, t in enumerate(tokens):               # markiertes Wort ohne d/a-Lauf (sollte nicht vorkommen): naechste Gruppe
            if t.status != KEEP and i not in gezaehlt:
                gi = min(range(len(gruppen)), key=lambda g: min(abs(i - j) for j in gruppen_token[g]))
                zaehl[gi].append(i)
                gezaehlt.add(i)
    return _Aufbau(laeufe, lauf_token, gruppen, gruppen_token, zaehl)


def laeufe_aus_tokens(tokens: Sequence[PdfToken], praefix: str = "G") -> list[Lauf]:
    """PDF-Woerter -> Laeufe. Je Aenderung (maximale Folge nicht-k-Laeufe) eine ID `praefix + Nummer` (ab 1).

    Hat ein Wort Zeichen-Laeufe (`teile`), wird jeder Teil ein Lauf ohne Trenner dazwischen (blau "insgesamt" + schwarzer
    Punkt -> a"insgesamt" k"."); sonst das ganze Wort (mod -> d(alt) + a(neu) aneinandergeklebt). Zwischen zwei Woertern
    steht ein Trenner-Lauf mit `token.trenner` ("" = geklebt, dann kein Trenner)."""
    return _aufbau(tokens, praefix).laeufe


def umbenennen(laeufe: Iterable[Lauf], ids: dict[str, str]) -> list[Lauf]:
    return [replace(l, aenderung=ids.get(l.aenderung, l.aenderung)) if l.aenderung else l for l in laeufe]


# =============================================================================================== 2. Spuren
@dataclass
class Spur:
    """Ein logischer PDF-Absatz (ein oder mehrere Bloecke, z. B. ueber einen Seitenumbruch)."""
    nr: int
    bloecke: list[PdfBlock]
    codes: list[str] = field(default_factory=list)         # Auffaelligkeiten des ganzen Absatzes (z. B. fortsetzung_unklar)
    _a: _Aufbau | None = field(default=None, repr=False)

    @property
    def tokens(self) -> list[PdfToken]:
        return [t for b in self.bloecke for t in b.tokens]

    @property
    def aufbau(self) -> _Aufbau:
        if self._a is None:
            self._a = _aufbau(self.tokens, f"S{self.nr}G")
        return self._a

    @property
    def laeufe(self) -> list[Lauf]:
        return self.aufbau.laeufe

    @property
    def kind(self) -> str:
        return self.bloecke[0].kind

    @property
    def erster(self) -> PdfBlock:
        return self.bloecke[0]

    @property
    def alt(self) -> str:
        return zustand(self.laeufe, set())

    @property
    def neu(self) -> str:
        return zustand(self.laeufe, {l.aenderung for l in self.laeufe if l.aenderung})

    @property
    def hat_aenderung(self) -> bool:
        return any(b.hat_aenderung for b in self.bloecke)

    @property
    def ganz_neu(self) -> bool:
        return all(b.ganz_neu for b in self.bloecke)

    @property
    def ganz_geloescht(self) -> bool:
        return all(b.ganz_geloescht for b in self.bloecke)

    def block_von_token(self, i: int) -> PdfBlock:
        for b in self.bloecke:
            if i < len(b.tokens):
                return b
            i -= len(b.tokens)
        return self.bloecke[-1]


def _fortsetzung_raten(bl: list[PdfBlock], i: int, satzende: str) -> str:
    """Ohne Angabe des PDF-Lesers: endet ein Absatz unten auf der Seite ohne Satzende und beginnt oben auf der naechsten Seite
    ein Absatz, KOENNTE es derselbe sein ("moeglich"; entschieden wird ueber A1 bzw. nie automatisch)."""
    b = bl[i]
    if i + 1 >= len(bl):
        return ""
    n = bl[i + 1]
    if b.kind != "absatz" or n.kind != "absatz" or n.page != b.page + 1 or not b.tokens:
        return ""
    if any(x.page == b.page for x in bl[i + 1:]) or any(x.page == n.page for x in bl[:i + 1]):
        return ""
    letzt = b.tokens[-1].new or b.tokens[-1].old
    if not letzt or letzt.rstrip("\"'“”»«)]")[-1:] in satzende:
        return ""
    return "moeglich"


def spuren(pdf: Any, index: "ExaktIndex | None" = None, cfg: dict | None = None) -> list[Spur]:
    """PDF-Bloecke -> logische Absaetze.

    `fortsetzung == "sicher"` (PDF-Leser): mit dem naechsten Block verbinden. "moeglich": beide Lesarten per A1 gegen die
    Excel testen (verbunden nur, wenn der verbundene alte Absatz exakt in der Excel steht). Ist der naechste Block ganz neu,
    kann A1 nichts entscheiden -> Code `fortsetzung_unklar` (nie automatisch)."""
    pl = _zuordnung_cfg(cfg).get("planer", {})
    bl: list[PdfBlock] = list(pdf.blocks)
    raten = bool(pl.get("fortsetzung_raten", True)) and not any(b.fortsetzung for b in bl)
    satzende = str(pl.get("satzende_zeichen", ".!?:;"))
    fort = [(b.fortsetzung or (_fortsetzung_raten(bl, i, satzende) if raten else "")) for i, b in enumerate(bl)]
    out: list[Spur] = []
    unklar_naechster = False
    i = 0
    while i < len(bl):
        gruppe = [bl[i]]
        codes = ["fortsetzung_unklar"] if unklar_naechster else []
        unklar_naechster = False
        j = i
        while j + 1 < len(bl) and gruppe[-1].kind == "absatz" and bl[j + 1].kind == "absatz" and fort[j]:
            nxt = bl[j + 1]
            if fort[j] == "sicher":
                gruppe.append(nxt)
                j += 1
                continue
            # "moeglich"
            if nxt.ganz_neu or gruppe[-1].ganz_neu:
                # A1 kann nicht entscheiden (kein alter Text): beide Teile nie automatisch
                unklar_naechster = True
                if "fortsetzung_unklar" not in codes:
                    codes.append("fortsetzung_unklar")
                break
            if index is not None:
                probe = Spur(-1, gruppe + [nxt])
                if probe.alt and index.finde(probe.alt, zustand_mit_optional(probe.laeufe, set())[1]):
                    gruppe.append(nxt)
                    j += 1
                    continue
            break
        out.append(Spur(len(out) + 1, gruppe, codes))
        i = j + 1
    return out


# =============================================================================================== 2b. Label (ADR 0006)
_SEC_FORM = re.compile(r"^\d{1,3}(?:\.\d{1,3}){0,6}\.?$")
_KAPITEL_WORT = re.compile(r"^(?:PART|TEIL)$", re.I)
_SUB_FORM = re.compile(r"^[A-Z]{1,2}\.$")
_LISTE_FORM = re.compile(r"^(?:\d{1,3}|[a-z]{1,3}|[A-Z]{1,3})[.)]$")


@dataclass(frozen=True)
class Label:
    """Gliederungs-Label am Anfang eines PDF-Absatzes (ADR 0006):
    art  "" (keins) | liste ("1.", "a)" -- steht im Register IM Text) | sub ("A." -- im Register in der Sub-Section-Spalte)
         | section ("1.1") | part ("PART 1") -- beide im Register in der Section-Spalte.
    alt/neu: Label im alten/neuen Stand (section/part: wie PdfBlock.section_alt/neu, z. B. "Part 1"); "" = gab es nicht.
    n: Zahl der Laeufe des Labels samt Trenner dahinter (der Text beginnt bei laeufe[n])."""
    art: str = ""
    alt: str = ""
    neu: str = ""
    n: int = 0


def label_der_spur(sp: Spur) -> Label:
    """Label eines PDF-Absatzes: Listennummer bzw. Sub-Section-Buchstabe als erstes Wort eines Absatzes, Section-Nummer oder
    Kapitelzeile ("PART 1") am Anfang einer Ueberschrift. Ohne Text dahinter gibt es kein Label."""
    toks = sp.tokens
    if len(toks) < 2:
        return Label()
    if sp.kind == "ueberschrift":
        if not (sp.erster.section_alt or sp.erster.section_neu):
            return Label()
        k = 0
        while k < min(len(toks) - 1, 3):
            formen = [x for x in (toks[k].old, toks[k].new) if x]
            if formen and all(_SEC_FORM.match(x) or _KAPITEL_WORT.match(x) for x in formen):
                k += 1
            else:
                break
        if k == 0:
            return Label()
        art = "part" if any(_KAPITEL_WORT.match(x) for t in toks[:k] for x in (t.old, t.new) if x) else "section"
        alt = (sp.erster.section_alt or "") if any(t.old for t in toks[:k]) else ""
        neu = (sp.erster.section_neu or "") if any(t.new for t in toks[:k]) else ""
    else:
        t = toks[0]
        formen = [x for x in (t.old, t.new) if x]
        if not formen:
            return Label()
        if all(_SUB_FORM.match(x) for x in formen):
            art = "sub"
        elif all(_LISTE_FORM.match(x) for x in formen):
            art = "liste"
        else:
            return Label()
        k, alt, neu = 1, t.old, t.new
    lt = sp.aufbau.lauf_token
    n = 0
    while n < len(lt) and lt[n] < k:
        n += 1
    return Label(art, alt, neu, n)


# =============================================================================================== 3. ExaktIndex
@dataclass(frozen=True)
class Fundstelle:
    """Eine Stelle in der Excel (Originalzustand), an der ein Text in Normalform genau steht.

    zellen: ((zeile, spalte, start, ende), ...) -- mehrere Zellen (Nummer + Titel) werden mit " " verbunden."""
    blatt: str
    zellen: tuple[tuple[int, int, int, int], ...]
    art: str                       # absatz (ganze Zeile einer Zelle) | zeilen (mehrere ganze Zeilen) | zelle (ganze mehrzeilige
                                   # Zelle) | teil (Stueck eines laengeren Excel-Absatzes) | kopf (Nummer + Titel)
    schluessel: tuple
    text: str
    versteckt: str = ""
    gliederung: bool = True        # das Blatt hat eine erkannte Gliederung (Section-Spalte)
    kopf_spalte: bool = False      # liegt in der Nummern-/Titelspalte (nicht im Fliesstext)
    section: str | None = None
    section_unsicher: bool = False
    zeilen_ende: bool = True       # die Stelle endet am Ende einer Zeile der Zelle
    zeilen_anfang: bool = True
    end_schluessel: tuple = ()     # Ordnungs-Schluessel des Endes (letzte Zelle, Ende)

    @property
    def zeile(self) -> int:
        return self.zellen[0][0]

    @property
    def spalte(self) -> int:
        return self.zellen[0][1]

    @property
    def start(self) -> int:
        return self.zellen[0][2]

    @property
    def ende(self) -> int:
        return self.zellen[-1][3]

    @property
    def ziel(self) -> str:
        from openpyxl.utils import get_column_letter
        return f"{self.blatt}!{get_column_letter(self.spalte)}{self.zeile}"


def _rang(lay: Any, spalte: int) -> int:
    if lay is None:
        return 3
    if spalte == lay.section_spalte:
        return 0
    if spalte == lay.titel_spalte:
        return 1
    if spalte == lay.inhalt_spalte:
        return 2
    return 3


def ordnungs_schluessel(xl: "ExcelModell", blatt: str, zeile: int, spalte: int, start: int) -> tuple:
    """Ordnung in Dokument-Reihenfolge: (Blatt-Index, Zeile, Spaltenrang [Section, Titel, Inhalt, Rest], Spalte, Start)."""
    nr = list(xl.layouts).index(blatt) if blatt in xl.layouts else 999
    return (nr, zeile, _rang(xl.layouts.get(blatt), spalte), spalte, start)


def _enthaelt(g: Fundstelle, f: Fundstelle) -> bool:
    """Liegt f ganz in g (gleiche Zelle)? Bei gleicher Spanne gewinnt die erste nach Art (absatz vor teil)."""
    if g.blatt != f.blatt or len(f.zellen) != 1:
        return False
    z, s, a, e = f.zellen[0]
    for z2, s2, a2, e2 in g.zellen:
        if (z2, s2) == (z, s) and a2 <= a and e <= e2 and (a2, e2) != (a, e):
            return True
    return False


def _woerter_norm(text: str) -> tuple[str, ...]:
    return tuple(n for n in (norm(w) for w in text.split()) if n)


class ExaktIndex:
    """Normalform jedes Excel-Absatzes (Zeile einer Zelle, getrennt an "\\n"), jeder ganzen mehrzeiligen Zelle und jeder
    Ueberschrift ("Nr Titel") -> Fundstellen. Teil-Treffer (PDF-Absatz = zusammenhaengendes Stueck eines laengeren
    Excel-Absatzes an Wortgrenzen, z. B. Seitenumbruch) werden nur in Zellen gesucht, die das seltenste Wort enthalten."""

    def __init__(self, xl: "ExcelModell", cfg: dict | None = None, wort_index: dict[str, list[int]] | None = None) -> None:
        self.xl = xl
        c = _zuordnung_cfg(cfg)
        self.nf_cfg = c.get("normalform")
        self.teil_min = int(c.get("matcher", {}).get("teil_min_woerter", 3))
        self._blatt_nr = {name: i for i, name in enumerate(xl.layouts)}
        self._nach_norms: dict[tuple[str, ...], list["Kandidat"]] = {}
        for k in xl.kandidaten:
            self._nach_norms.setdefault(tuple(k.norms), []).append(k)
        if wort_index is None:
            wort_index = {}
            for k in xl.kandidaten:
                for n in set(k.norms):
                    wort_index.setdefault(n, []).append(k.id)
        self.wort = wort_index
        self._nf: dict[tuple[str, int, int], tuple[str, list[int]]] = {}
        self._zell_kand: dict[tuple[str, int, int], "Kandidat"] = {}
        for k in xl.kandidaten:
            if k.art != "kopf":
                self._zell_kand.setdefault((k.blatt, k.zeile, k.spalte), k)
        self._cache: dict[tuple, list[Fundstelle]] = {}

    # ---------------------------------------------------------------- Bausteine
    def schluessel(self, blatt: str, zeile: int, spalte: int, start: int) -> tuple:
        return (self._blatt_nr.get(blatt, 999), zeile, _rang(self.xl.layouts.get(blatt), spalte), spalte, start)

    def _gliederung(self, blatt: str) -> bool:
        lay = self.xl.layouts.get(blatt)
        return bool(lay and lay.section_spalte)

    def _fund(self, blatt: str, zellen: list[tuple[int, int, int, int]], art: str, k: "Kandidat | None") -> Fundstelle:
        texte = [self.xl.text(blatt, z, s)[a:e] for z, s, a, e in zellen]
        z0, s0, a0, _ = zellen[0]
        zl, sl, _, el = zellen[-1]
        voll0 = self.xl.text(blatt, z0, s0)
        volll = self.xl.text(blatt, zl, sl)
        anfang = a0 == 0 or voll0[a0 - 1] == "\n" or not voll0[:a0].strip()
        ende = el >= len(volll) or volll[el] == "\n" or not volll[el:].strip()
        lay = self.xl.layouts.get(blatt)
        kopf_spalte = art == "kopf" or bool(lay and s0 in (lay.section_spalte, lay.titel_spalte) and s0 != lay.inhalt_spalte)
        return Fundstelle(blatt=blatt, zellen=tuple(zellen), art=art, schluessel=self.schluessel(blatt, z0, s0, a0),
                          text=" ".join(texte), versteckt=(k.versteckt if k is not None else ""),
                          gliederung=self._gliederung(blatt), kopf_spalte=kopf_spalte,
                          section=(k.section if k is not None else self.xl.zeilen_section.get((blatt, z0))),
                          section_unsicher=bool(k.section_unsicher) if k is not None else False,
                          zeilen_ende=ende, zeilen_anfang=anfang, end_schluessel=self.schluessel(blatt, zl, sl, el))

    def _kopf_zellen(self, k: "Kandidat") -> list[tuple[int, int, int, int]]:
        zellen = list(dict.fromkeys((t.zeile, t.spalte) for t in k.tokens))
        return [(z, s, 0, len(self.xl.text(k.blatt, z, s))) for z, s in zellen]

    def _nf_zelle(self, blatt: str, zeile: int, spalte: int) -> tuple[str, list[int]]:
        key = (blatt, zeile, spalte)
        if key not in self._nf:
            self._nf[key] = normalform(self.xl.text(blatt, zeile, spalte), self.nf_cfg)
        return self._nf[key]

    # ---------------------------------------------------------------- Suche
    def finde(self, text: str, optional: frozenset[int] = frozenset(), teil: bool = True) -> list[Fundstelle]:
        """Alle Stellen, an denen `text` (Normalform, optionale Zeichen duerfen fehlen) genau steht."""
        if not text.strip():
            return []
        key = (text, optional, teil)
        if key in self._cache:
            return self._cache[key]
        norms = _woerter_norm(text)
        out: dict[tuple, Fundstelle] = {}
        for k in self._nach_norms.get(norms, []) if norms else []:
            if self.xl.layouts[k.blatt].ignoriert:
                continue
            if not vergleiche(text, k.text, optional, self.nf_cfg).gleich:
                continue
            if k.art == "kopf":
                zellen = self._kopf_zellen(k)
            else:
                zellen = [(k.zeile, k.spalte, k.start, k.end)]
            f = self._fund(k.blatt, zellen, "kopf" if k.art == "kopf" else ("zelle" if k.art == "zelle" else "absatz"), k)
            out[(f.blatt, f.zellen)] = f
        if teil and len(norms) >= self.teil_min:
            for f in self._teil_treffer(text, optional, norms):
                out.setdefault((f.blatt, f.zellen), f)
        werte = list(out.values())
        # dieselbe Stelle zweimal (z. B. mit und ohne Aufzaehlungszeichen am Anfang): nur die groessere zaehlt
        werte = [f for f in werte if not any(g is not f and _enthaelt(g, f) for g in werte)]
        erg = sorted(werte, key=lambda f: f.schluessel)
        self._cache[key] = erg
        return erg

    def _teil_treffer(self, text: str, optional: frozenset[int], norms: tuple[str, ...]) -> list[Fundstelle]:
        selten = min(set(norms), key=lambda n: len(self.wort.get(n, ())))
        zellen: dict[tuple[str, int, int], "Kandidat"] = {}
        for kid in self.wort.get(selten, ()):
            k = self.xl.kandidaten[kid]
            if k.art == "kopf" or self.xl.layouts[k.blatt].ignoriert:
                continue
            zellen.setdefault((k.blatt, k.zeile, k.spalte), k)
        if not zellen:
            return []
        pnf, pabb = normalform(text, self.nf_cfg)
        if optional:                                   # optionale Zeichen (Trennstrich) fuer die Suche weglassen
            pnf = "".join(z for z, i in zip(pnf, pabb) if i not in optional)
        pnf = pnf.strip()
        if not pnf:
            return []
        out: list[Fundstelle] = []
        for (blatt, zeile, spalte), k in zellen.items():
            nf, abb = self._nf_zelle(blatt, zeile, spalte)
            roh = self.xl.text(blatt, zeile, spalte)
            p = nf.find(pnf)
            while p >= 0:
                q = p + len(pnf)
                if (p == 0 or nf[p - 1] == " ") and (q == len(nf) or nf[q] == " "):
                    s, e = abb[p], abb[q - 1] + 1
                    if vergleiche(text, roh[s:e], optional, self.nf_cfg).gleich:
                        art = "zelle" if (s == 0 and e == len(roh.rstrip())) and "\n" in roh[s:e] else "teil"
                        f = self._fund(blatt, [(zeile, spalte, s, e)], art, self._zell_kand.get((blatt, zeile, spalte), k))
                        if f.zeilen_anfang and f.zeilen_ende and art == "teil":
                            f = replace(f, art="absatz" if "\n" not in roh[s:e] else "zeilen")
                        out.append(f)
                p = nf.find(pnf, p + 1)
        return out

    # ---------------------------------------------------------------- Zeilen (fuer die Pruefung "nichts Unzugeordnetes dazwischen")
    def verwaltete_zeilen(self, blatt: str) -> list[tuple[tuple, tuple[int, int, int, int]]]:
        """Sortiert: (Schluessel, (zeile, spalte, start, ende)) aller Textzeilen in Section-/Titel-/Inhaltsspalte."""
        if not hasattr(self, "_vz"):
            self._vz: dict[str, list] = {}
        if blatt not in self._vz:
            lay = self.xl.layouts.get(blatt)
            spalten = {c for c in ((lay.section_spalte, lay.titel_spalte, lay.inhalt_spalte) if lay else ()) if c}
            if lay is not None and lay.register:       # Register: ein Absatz je Zeile in der Textspalte (die Section steht ueberall)
                spalten = {lay.inhalt_spalte}
            out = []
            for (b, z, s), t in self.xl.zellen.items():
                if b != blatt or s not in spalten or (lay is not None and lay.register and z < lay.daten_ab):
                    continue
                for m in re.finditer(r"[^\n]+", t):
                    if m.group().strip():
                        out.append((self.schluessel(b, z, s, m.start()), (z, s, m.start(), m.end())))
            out.sort()
            self._vz[blatt] = out
        return self._vz[blatt]


# =============================================================================================== 4. Reissverschluss
@dataclass
class Zuordnung:
    status: str                                    # EINDEUTIG | REIHENFOLGE | MEHRDEUTIG | WIDERSPRUCH | KEINER
    fund: Fundstelle | None = None                 # gewaehlte Stelle (bei MEHRDEUTIG/WIDERSPRUCH nur ein Vorschlag)
    art: str = "alt"                               # an der Stelle steht der alte ("alt") oder schon der neue ("neu") Absatz
    kandidaten: list[tuple[Fundstelle, str]] = field(default_factory=list)
    fenster: tuple[tuple | None, tuple | None] = (None, None)   # Kettennachbarn (Excel-Schluessel), fuer den Ersatz-Matcher
    auf_kette: bool = False

    @property
    def automatisch(self) -> bool:
        return self.status in AUTOMATISCH_ERLAUBT


class _MaxBaum:
    """Fenwick-Baum fuer Praefix-Maxima (Gewicht, Element)."""

    def __init__(self, n: int) -> None:
        self.n = n
        self.b: list[tuple[int, int]] = [(0, -1)] * (n + 1)

    def setze(self, i: int, wert: tuple[int, int]) -> None:
        i += 1
        while i <= self.n:
            if wert[0] > self.b[i][0]:
                self.b[i] = wert
            i += i & -i

    def max_vor(self, i: int) -> tuple[int, int]:      # Maximum ueber Positionen < i
        best = (0, -1)
        while i > 0:
            if self.b[i][0] > best[0]:
                best = self.b[i]
            i -= i & -i
        return best


def reissverschluss(kandidaten: list[list[tuple[Fundstelle, str]]], gewicht_alt: int = 4, gewicht_neu: int = 2) -> list[Zuordnung]:
    """Laengste (gewichtete) aufsteigende Kette ueber (Absatz, Fundstelle)-Paare: PDF und Excel in derselben Reihenfolge.

    kandidaten[i] = Fundstellen des i-ten PDF-Absatzes, je mit "alt" (alter Absatz steht dort) oder "neu" (neuer Absatz steht
    schon dort). Status je Absatz:
      eindeutig          eine Fundstelle, auf der Kette                         -> automatisch erlaubt
      durch_reihenfolge  mehrere, genau eine zwischen den benachbarten Ankern (davor: eindeutig oder schon ueber die
                         Reihenfolge zugeordnet; danach: eindeutig) -> automatisch erlaubt
      mehrdeutig         mehrere zwischen den Ankern                            -> nie automatisch
      widerspruch        keine Fundstelle passt in die Reihenfolge              -> nie automatisch
      keiner             keine Fundstelle                                       -> nie automatisch"""
    elemente: list[tuple[int, Fundstelle, str]] = []
    for i, cands in enumerate(kandidaten):
        for f, art in sorted(cands, key=lambda x: x[0].schluessel, reverse=True):
            elemente.append((i, f, art))
    schluessel = sorted({f.schluessel for _, f, _ in elemente})
    pos = {s: k for k, s in enumerate(schluessel)}
    baum = _MaxBaum(len(schluessel))
    dp: list[int] = []
    vor: list[int] = []
    for e, (i, f, art) in enumerate(elemente):
        w = (gewicht_alt if art == "alt" else gewicht_neu) + (1 if len(kandidaten[i]) == 1 else 0)
        best, prev = baum.max_vor(pos[f.schluessel])
        dp.append(best + w)
        vor.append(prev)
        baum.setze(pos[f.schluessel], (best + w, e))
    gewaehlt: dict[int, tuple[Fundstelle, str]] = {}
    if dp:
        e = max(range(len(dp)), key=lambda k: dp[k])
        while e >= 0:
            i, f, art = elemente[e]
            gewaehlt[i] = (f, art)
            e = vor[e]
    n = len(kandidaten)
    out = [Zuordnung(KEINER, kandidaten=list(kandidaten[i])) for i in range(n)]
    for i in range(n):
        if len(kandidaten[i]) == 1 and i in gewaehlt:
            out[i].status, (out[i].fund, out[i].art), out[i].auf_kette = EINDEUTIG, gewaehlt[i], True
    anker = [i for i in range(n) if out[i].status == EINDEUTIG]
    links: Fundstelle | None = None              # letzte schon sicher zugeordnete Stelle (eindeutig oder ueber die Reihenfolge)
    for i in range(n):
        z = out[i]
        if z.status == EINDEUTIG:
            links = z.fund
            continue
        if not kandidaten[i]:
            continue
        k = bisect.bisect_right(anker, i)
        lo = links.end_schluessel if links is not None else None
        hi = out[anker[k]].fund.schluessel if k < len(anker) else None          # type: ignore[union-attr]
        dazwischen = [(f, a) for f, a in kandidaten[i] if (lo is None or f.schluessel > lo) and (hi is None or f.schluessel < hi)]
        z.auf_kette = i in gewaehlt
        if len(kandidaten[i]) == 1:
            z.status = WIDERSPRUCH
            z.fund, z.art = kandidaten[i][0]
        elif len(dazwischen) == 1:
            z.status = REIHENFOLGE
            z.fund, z.art = dazwischen[0]
            links = z.fund
        elif dazwischen:
            z.status = MEHRDEUTIG
            wahl = gewaehlt.get(i) if gewaehlt.get(i) in dazwischen else None
            if wahl is None:
                sichtbar = [x for x in dazwischen if not x[0].versteckt]
                wahl = (sichtbar or dazwischen)[0]
            z.fund, z.art = wahl
        else:
            z.status = WIDERSPRUCH
            wahl = gewaehlt.get(i)
            if wahl is None:
                sichtbar = [x for x in kandidaten[i] if not x[0].versteckt]
                wahl = (sichtbar or kandidaten[i])[0]
            z.fund, z.art = wahl
    # Fenster je Absatz: naechste Kettenglieder (automatisch erlaubte Zuordnungen) davor/danach
    kette = [i for i in range(n) if out[i].automatisch]
    for i in range(n):
        k = bisect.bisect_left(kette, i)
        vorher = [j for j in kette[:k] if j != i]
        nachher = [j for j in kette[k:] if j != i]
        lo = out[vorher[-1]].fund.end_schluessel if vorher else None              # type: ignore[union-attr]
        hi = out[nachher[0]].fund.schluessel if nachher else None                 # type: ignore[union-attr]
        out[i].fenster = (lo, hi)
    return out


# =============================================================================================== 5. A1 + B
@dataclass
class SpurOps:
    """Ergebnis von `ops_fuer_spur`."""
    vergleich: Vergleich
    ops: dict[str, list[Operation]] = field(default_factory=dict)       # Aenderungs-ID -> Operationen
    probleme: dict[str, list[str]] = field(default_factory=dict)        # Aenderungs-ID -> Codes (dann hoechstens "mittel")
    excel_text: str = ""


def _sichtbar(laeufe: list[Lauf], aktiv: set[str]) -> list[tuple[int, Lauf]]:
    """Wie gegenprobe._gefiltert, aber mit dem Index jedes Laufs (und ohne die Raender zu kuerzen)."""
    out: list[tuple[int, Lauf]] = []
    for k, l in enumerate(laeufe):
        if l.art == "d" and l.aenderung in aktiv:
            continue
        if l.art == "a" and l.aenderung not in aktiv:
            continue
        if l.art == "s" and out and out[-1][1].art == "s":
            if _RANG.get(l.text, 1) > _RANG.get(out[-1][1].text, 1):
                out[-1] = (k, l)
            continue
        out.append((k, l))
    return out


def _positionen(laeufe: list[Lauf], aktiv: set[str]) -> tuple[str, dict[int, tuple[int, int]]]:
    """Text wie `zustand(laeufe, aktiv)` und die Lage jedes sichtbaren Laufs darin."""
    vis = _sichtbar(laeufe, aktiv)
    while vis and vis[0][1].art == "s":
        vis.pop(0)
    while vis and vis[-1][1].art == "s":
        vis.pop()
    text: list[str] = []
    lage: dict[int, tuple[int, int]] = {}
    p = 0
    for k, l in vis:
        lage[k] = (p, p + len(l.text))
        text.append(l.text)
        p += len(l.text)
    return "".join(text), lage


def _ws(z: str) -> bool:
    return z.isspace()


def _naht_teile(teil: list[Lauf]) -> tuple[str, str, str]:
    """(Trenner vorn, Kern, Trenner hinten) einer Laufliste."""
    vorn = hinten = ""
    if teil and teil[0].art == "s":
        vorn = teil[0].text
        teil = teil[1:]
    if teil and teil[-1].art == "s":
        hinten = teil[-1].text
        teil = teil[:-1]
    return vorn, "".join(l.text for l in teil), hinten


def ops_fuer_spur(laeufe: list[Lauf], fund: Fundstelle, zelltext: Any, cfg: dict | None = None,
                  seq: int = 0) -> SpurOps:
    """A1 + B fuer einen Absatz an einer Excel-Stelle.

    A1: `vergleiche(zustand(laeufe, {}), Excel-Text)` -- nur wenn gleich, gibt es Operationen.
    B : je Aenderung EINE Operation `ersetzen` zwischen den Excel-Positionen der benachbarten k-Stuecke; die Zeichen der
        Excel ausserhalb (auch NBSP/Tab an den Naehten) bleiben, Satzzeichen bleiben richtig angeklebt.
    `zelltext(blatt, zeile, spalte)` liefert den Original-Zelltext."""
    c = _zuordnung_cfg(cfg)
    nf = c.get("normalform", {})
    kein_vor = set(nf.get("kein_leerzeichen_vor", ".,;:!?)]}"))
    kein_nach = set(nf.get("kein_leerzeichen_nach", "([{"))
    alt, lage = _positionen(laeufe, set())
    _, optional = zustand_mit_optional(laeufe, set())
    # Excel-Text der Stelle (mehrere Zellen mit " " verbunden) und Rueckabbildung
    segmente: list[tuple[int, int, int, int, int]] = []        # (Offset in X, zeile, spalte, start, ende)
    teile: list[str] = []
    off = 0
    for z, s, a, e in fund.zellen:
        t = zelltext(fund.blatt, z, s)[a:e]
        segmente.append((off, z, s, a, e))
        teile.append(t)
        off += len(t) + 1
    X = " ".join(teile)
    v = vergleiche(alt, X, optional, nf or None)
    erg = SpurOps(v, excel_text=X)
    if not v.gleich:
        return erg
    abb = v.abbildung
    ids = list(dict.fromkeys(l.aenderung for l in laeufe if l.aenderung))
    for gid in ids:
        idx = [k for k, l in enumerate(laeufe) if l.aenderung == gid]
        i0, i1 = idx[0], idx[-1]
        probleme: list[str] = []
        # Nachbar-k-Stuecke mit abgebildeten Zeichen suchen
        iL = next((k for k in range(i0 - 1, -1, -1) if laeufe[k].art == "k"), None)
        iR = next((k for k in range(i1 + 1, len(laeufe)) if laeufe[k].art == "k"), None)

        def abgebildet(k: int, letztes: bool) -> int | None:
            a, e = lage[k]
            rng = range(e - 1, a - 1, -1) if letztes else range(a, e)
            for p in rng:
                if p in abb:
                    return abb[p] + (1 if letztes else 0)
            return None
        while iL is not None and abgebildet(iL, True) is None:
            probleme.append("naht_unklar")
            iL = next((k for k in range(iL - 1, -1, -1) if laeufe[k].art == "k"), None)
        while iR is not None and abgebildet(iR, False) is None:
            probleme.append("naht_unklar")
            iR = next((k for k in range(iR + 1, len(laeufe)) if laeufe[k].art == "k"), None)
        xa = abgebildet(iL, True) if iL is not None else 0
        xb = abgebildet(iR, False) if iR is not None else len(X)
        if xa is None or xb is None or xb < xa:
            erg.probleme[gid] = ["naht_unklar"]
            continue
        # Fein (Auftrag 003): je zusammenhaengender Aenderung (Teilgruppe zwischen zwei Trennern) eine eigene Operation; die
        # Zeichen der Excel zwischen zwei Teilgruppen (geschuetztes Leerzeichen, Tab ...) bleiben unangetastet.
        teile = _teilgruppen(laeufe, i0, i1)
        stuecke: list[tuple[int, int, str]] = []
        for j, (s0, s1) in enumerate(teile):
            st = _teil_op(laeufe, X, abgebildet, j, teile, iL, iR, xa, xb, kein_vor, kein_nach)
            if st is None:
                stuecke = []
                break
            stuecke.append(st)
        if not stuecke:
            st = _gruppen_op(laeufe, X, iL, iR, xa, xb, gid, kein_vor, kein_nach)
            if st is None:
                erg.probleme[gid] = ["leer"]
                continue
            stuecke = [st]
        ops: list[Operation] = []
        for ra, rb, neu in stuecke:
            if "\n" in X[ra:rb]:
                probleme.append("excel_umbruch")
            ops += _auf_zellen(ra, rb, neu, X, segmente, fund, gid, seq, probleme)
        if ops and not ("mehrere_zellen" in probleme and len(stuecke) > 1):
            erg.ops[gid] = ops
        if probleme:
            erg.probleme[gid] = list(dict.fromkeys(probleme))
    return erg


def _teilgruppen(laeufe: list[Lauf], i0: int, i1: int) -> list[tuple[int, int]]:
    """Teilgruppen einer Aenderung (Lauf-Indizes [von, bis]): zusammenhaengende d/a-Laeufe ohne Trenner dazwischen. Getrennt
    wird an einem Trenner nur, wenn auf BEIDEN Seiten alter Text steht und eine Seite auch neuen hat ("{-x-}{+y+} {-z-}{+w+}"):
    dann liegt zwischen zwei Stellen der Excel ein Zeichen der Excel, das bleiben soll. "{-12-} {+7+}" bleibt eine Ersetzung,
    "{-a-} {-b-}" eine Loeschung, reine Einfuegungen haengen an ihrem Nachbarn."""
    cluster: list[tuple[int, int, frozenset]] = []
    k = i0
    while k <= i1:
        if laeufe[k].art not in ("d", "a"):
            k += 1
            continue
        a = k
        while k + 1 <= i1 and laeufe[k + 1].art in ("d", "a"):
            k += 1
        cluster.append((a, k, frozenset(laeufe[x].art for x in range(a, k + 1))))
        k += 1
    out: list[list[tuple[int, int, frozenset]]] = []
    for c in cluster:
        if out:
            arten = frozenset().union(*(x[2] for x in out[-1]))
            if not ("d" in arten and "d" in c[2] and ("a" in arten or "a" in c[2])):
                out[-1].append(c)
                continue
        out.append([c])
    return [(g[0][0], g[-1][1]) for g in out]


def _gruppen_op(laeufe: list[Lauf], X: str, iL: int | None, iR: int | None, xa: int, xb: int, gid: str,
                kein_vor: set, kein_nach: set) -> tuple[int, int, str] | None:
    """Eine Operation fuer die ganze Aenderung zwischen den Nachbar-k-Stuecken (Bereich [ra, rb) und neuer Text)."""
    mitte = laeufe[(iL + 1 if iL is not None else 0):(iR if iR is not None else len(laeufe))]
    return _zone_op(mitte, gid, X, xa, xb, iL is None, iR is None,
                    iR is not None and laeufe[iR].text[:1] in kein_vor, iL is not None and laeufe[iL].text[-1:] in kein_nach)


def _zone_op(mitte: list[Lauf], gid: str, X: str, xa: int, xb: int, rand_l: bool, rand_r: bool, vor_satzzeichen: bool,
             nach_klammer: bool) -> tuple[int, int, str] | None:
    """Ersatz fuer den ganzen Excel-Bereich [xa, xb) zwischen zwei Ankern (frueheres Verfahren, Naehte der Excel)."""
    alt_teil = [l for _, l in _sichtbar(mitte, set())]
    neu_teil = [l for _, l in _sichtbar(mitte, {gid})]
    lo, kern_alt, to = _naht_teile(alt_teil)
    ln, kern_neu, tn = _naht_teile(neu_teil)
    # am Absatzrand gibt es keine Naht; Satzzeichen kleben am Nachbarn
    if rand_l:
        lo = ln = ""
    if rand_r:
        to = tn = ""
    if vor_satzzeichen:
        tn = ""
    if nach_klammer:
        ln = ""
    mx = X[xa:xb]
    lx = mx[:len(mx) - len(mx.lstrip())]
    tx = mx[len(mx.rstrip()):] if mx.strip() else ""
    if kern_alt and kern_neu:                        # ersetzen: nur den Kern, die Naehte der Excel bleiben
        return xa + len(lx), xb - len(tx), kern_neu
    if kern_neu:                                     # einfuegen: eine Naht der Excel bleibt
        w = mx
        if not w:
            neu = ln + kern_neu + tn
        elif ln and tn and len(w) >= 2 and w.isspace():      # zwei Naht-Zeichen der Excel (" " + NBSP): je eins nach aussen
            neu = w[:1] + kern_neu + w[1:]
        elif ln and tn:
            neu = w + kern_neu + tn
        elif ln:
            neu = w + kern_neu
        elif tn:
            neu = kern_neu + w
        else:
            neu = w + kern_neu
        return xa, xb, neu
    if kern_alt:                                     # loeschen: hoechstens eine Naht bleibt
        if not (ln or tn) or rand_l or rand_r or vor_satzzeichen or nach_klammer:   # Rand / Satzzeichen: keine Naht
            return xa, xb, ""
        return xa, xb, lx or tx or (ln or tn)
    return None


def _teil_op(laeufe: list[Lauf], X: str, abgebildet: Any, j: int, teile: list[tuple[int, int]], iL: int | None,
             iR: int | None, xa0: int, xb0: int, kein_vor: set, kein_nach: set) -> tuple[int, int, str] | None:
    """Operation einer Teilgruppe: ersetzt bzw. loescht genau ihren alten Text (die Excel-Stelle der d-Laeufe) oder fuegt
    ohne Bereich ein; Leerraum der Excel neben dem alten Text bleibt (ausser eine Loeschung braucht eine Naht weniger).
    Anker sind die k-Stuecke um die Aenderung und der alte Text der Nachbar-Teilgruppen. None = nicht fein bestimmbar."""
    s0, s1 = teile[j]
    gid = next(l.aenderung for l in laeufe[s0:s1 + 1] if l.aenderung)
    # Anker: alter Text der vorigen/naechsten Teilgruppe, sonst die k-Stuecke um die ganze Aenderung
    xa, geteilt_l = xa0, False
    for k in range(s0 - 1, (iL if iL is not None else -1), -1):
        if laeufe[k].art == "d":
            p = abgebildet(k, True)
            if p is None:
                return None
            xa, geteilt_l = p, True
            break
    xb, geteilt_r = xb0, False
    for k in range(s1 + 1, (iR if iR is not None else len(laeufe))):
        if laeufe[k].art == "d":
            p = abgebildet(k, False)
            if p is None:
                return None
            xb, geteilt_r = p, True
            break
    if xb < xa:
        return None
    p0 = teile[j - 1][1] + 1 if j > 0 else (iL + 1 if iL is not None else 0)
    p1 = teile[j + 1][0] if j + 1 < len(teile) else (iR if iR is not None else len(laeufe))
    mitte = laeufe[p0:p1]
    rand_l = iL is None and not geteilt_l
    rand_r = iR is None and not geteilt_r
    vor_satz = not geteilt_r and iR is not None and laeufe[iR].text[:1] in kein_vor
    nach_kl = not geteilt_l and iL is not None and laeufe[iL].text[-1:] in kein_nach
    zone = _zone_op(mitte, gid, X, xa, xb, rand_l, rand_r, vor_satz, nach_kl)
    if zone is None:
        return None
    za, zb, neu_zone = zone
    d_laeufe = [k for k in range(s0, s1 + 1) if laeufe[k].art == "d"]
    if d_laeufe:
        ka, kb = abgebildet(d_laeufe[0], False), abgebildet(d_laeufe[-1], True)
        if ka is None or kb is None or not (xa <= ka <= kb <= xb):
            return None
        if any(laeufe[k].art == "a" for k in range(s0, s1 + 1)):
            return ka, kb, neu_zone                    # ersetzen: genau der alte Text
        lx, tx = X[xa:ka], X[kb:xb]
        if neu_zone == "":                             # Loeschung ohne Naht (Rand, Satzzeichen): Leerraum nur auf der eigenen Seite
            return (ka if geteilt_l else xa), (kb if geteilt_r else xb), ""
        if neu_zone == lx and not geteilt_r:           # linke Naht der Excel bleibt
            return ka, xb, ""
        if neu_zone == tx and not geteilt_l:           # rechte Naht der Excel bleibt
            return xa, kb, ""
        if lx or tx:                                   # beide Seiten gehoeren Nachbarn: nur den alten Text
            return ka, kb, ""
        return za, zb, neu_zone
    # reine Einfuegung: ohne Bereich, der Leerraum der Excel bleibt. Nie genau am Anfang des alten Texts der naechsten
    # Teilgruppe (gleiche Stelle wie deren Ersetzung: die Reihenfolge waere nicht sicher) -- dann den Bereich ersetzen.
    w = X[xa:xb]
    if not w:
        return None if geteilt_r else (xa, xa, neu_zone)
    if not geteilt_r and neu_zone.startswith(w):
        return xb, xb, neu_zone[len(w):]
    if neu_zone.endswith(w):
        return xa, xa, neu_zone[:len(neu_zone) - len(w)]
    return za, zb, neu_zone


def _auf_zellen(ra: int, rb: int, neu: str, X: str, segmente: list[tuple[int, int, int, int, int]], fund: Fundstelle,
                gid: str, seq: int, probleme: list[str]) -> list[Operation]:
    """Bereich [ra, rb) im verbundenen Text -> Operation(en) je Zelle. Der kuenstliche Trenner " " zwischen zwei Zellen
    (Nummer + Titel) gehoert keiner Zelle."""
    def seg_von(p: int) -> int:
        for j in range(len(segmente) - 1, -1, -1):
            if segmente[j][0] <= p:
                return j
        return 0

    def op(j: int, a: int, b: int, text: str) -> Operation:
        o, z, s, za, _ = segmente[j]
        return Operation("ersetzen", fund.blatt, z, s, za + a - o, za + b - o, text, seq, aenderung=gid)

    ja = seg_von(ra)
    jb = seg_von(max(rb - 1, ra))
    o_a, _, _, a0, e0 = segmente[ja]
    ende_a = o_a + (e0 - a0)
    if ja == jb and rb <= ende_a:
        return [op(ja, ra, rb, neu)]
    if ra == ende_a and ja + 1 < len(segmente):          # beginnt am Zelltrenner
        j2 = ja + 1
        if rb == ra:                                     # Einfuegen genau am Trenner: an den Anfang der naechsten Zelle
            return [op(j2, segmente[j2][0], segmente[j2][0], neu.lstrip())]
        if neu[:1].isspace() or neu == "":
            o2, _, _, a2, e2 = segmente[j2]
            if rb <= o2 + (e2 - a2):
                return [op(j2, o2, rb, neu[1:] if neu else "")]
    if ja == jb and rb == ende_a + 1 and ja + 1 < len(segmente):   # endet am Zelltrenner
        if neu[-1:].isspace() or neu == "":
            return [op(ja, ra, ende_a, neu[:-1] if neu else "")]
    # Ueber zwei Zellen (Nummer + Titel): nur wenn sich der neue Text sauber in Nummer und Titel teilen laesst
    if fund.art == "kopf" and len(segmente) == 2 and ja == 0 and jb == 1:
        kz = kopf_zerlegen(neu.strip()) if neu.strip() else None
        o2 = segmente[1][0]
        if (kz or not neu.strip()) and ra == 0:
            nr, titel = kz if kz else ("", "")
            return [op(0, 0, ende_a, nr), op(1, o2, rb, titel + (" " if rb < len(X) and X[rb - 1:rb] == " " else "")
                                              if titel else "")]
    probleme.append("mehrere_zellen")
    return []


def ganz_loeschen(fund: Fundstelle, zelltext: Any, gid: str, seq: int = 0,
                  belegt: dict[tuple[int, int], list[tuple[int, int]]] | None = None
                  ) -> tuple[list[Operation], list[tuple[int, int, int, int]]]:
    """Operationen, die einen ganzen Absatz an der Stelle entfernen, und die Zellbereiche, die dabei verschwinden.

    Ganze Zelle -> Zelle leer (ersetzen durch ""); Zeile in einer mehrzeiligen Zelle -> samt dem Zeilenumbruch davor (sonst
    dem danach); Stueck einer Zeile -> samt einem angrenzenden Leerzeichen. `belegt` (Zelle -> Bereiche): schon von anderen
    Loeschungen genommene Zeichen; ein Umbruch wird nie zweimal genommen (keine ueberlappenden Operationen)."""
    ops: list[Operation] = []
    bereiche: list[tuple[int, int, int, int]] = []
    for z, s, a, e in fund.zellen:
        t = zelltext(fund.blatt, z, s)
        frei = lambda p: not any(x <= p < y for x, y in (belegt or {}).get((z, s), []))     # noqa: E731
        if not (t[:a].strip() or t[e:].strip()):
            a2, e2 = 0, len(t)
        elif fund.art in ("absatz", "zeilen", "zelle") or (not t[:a].strip(" \t\xa0") or t[:a].endswith("\n")) and \
                (not t[e:].strip(" \t\xa0") or t[e:].startswith("\n")):
            a2, e2 = a, e
            vorn = len(t[:a].rstrip(" \t\xa0").rstrip("\n"))
            rest = t[e:]
            hinten = e + (len(rest) - len(rest.lstrip(" \t\xa0\n")))
            if t[:a].rstrip(" \t\xa0").endswith("\n") and t[:a].strip() and all(frei(p) for p in range(vorn, a)):
                a2 = vorn
            elif hinten > e and t[e:].strip() and all(frei(p) for p in range(e, hinten)):
                e2 = hinten
        else:
            a2, e2 = a, e
            if e2 < len(t) and t[e2] in " \t\xa0":
                while e2 < len(t) and t[e2] in " \t\xa0":
                    e2 += 1
            else:
                while a2 > 0 and t[a2 - 1] in " \t\xa0":
                    a2 -= 1
        ops.append(Operation("ersetzen", fund.blatt, z, s, a2, e2, "", seq, aenderung=gid))
        bereiche.append((z, s, a2, e2))
    return ops, bereiche
