"""Plant, was in die Excel eingetragen wird -- ohne die Excel zu veraendern.

Eingabe: PDF-Analyse (Woerter mit Status) + Excel-Modell.
Ausgabe: Liste von Aenderungen. Jede Aenderung ist entweder
  * uebernommen  (Konfidenz hoch: wird automatisch eingetragen),
  * pruefen      (Vorschlag steht fest, ein Mensch bestaetigt),
  * nicht_uebernommen (mit Grund). Nichts geht still verloren.

Alles Regeln, keine KI. Alle Operationen beziehen sich auf den Originalzustand
der Excel; die Reihenfolge beim Eintragen regelt der Applier.
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Any, Iterable, Iterator

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter

from . import config
from .excel_model import Kandidat, ExcelModell, zahl_als_text
from .abgleich import (ExaktIndex, Fundstelle, Spur, Zuordnung, _woerter_norm, ganz_loeschen, ops_fuer_spur,
                       reissverschluss, spuren, umbenennen)
from .gegenprobe import normalform, vergleiche, zustand, zustand_mit_optional
from .matcher import Matcher, Treffer
from .models import (ADD, DEL, HOCH, KEEP, MITTEL, MOD, NICHT, NIEDRIG, PRUEFEN, UEBERNOMMEN,
                     Absatz, Aenderung, Lauf, Operation, PdfBlock, PdfToken)
from .pdf_reader import PdfAnalyse, hinweis_text
from .textutil import kopf_zerlegen, norm, section_nummer

STATUS_VON_STUFE = {HOCH: UEBERNOMMEN, MITTEL: PRUEFEN}


def _planer_cfg(cfg: dict | None) -> dict:
    """Abschnitt `planer:` der Konfiguration (Schwellen und Zeichenlisten, Regel 28)."""
    if cfg and "planer" in cfg:
        return cfg["planer"]
    return config.lade("zuordnung").get("planer", {})


def _trennung_ohne_ausnahme(muster: "re.Pattern[str]", text: str, ausnahmen: set[str]) -> bool:
    """Gibt es im Text einen Treffer des Trennmusters, dessen Rest-Wort keine erlaubte Ausnahme ist ("Ein- und Ausgang")?"""
    for m in muster.finditer(text):
        rest = (m.groupdict().get("rest") or "").casefold()
        if rest not in ausnahmen:
            return True
    return False


def neutext_probleme(texte: Iterable[str], cfg: dict | None = None) -> dict[str, Any]:
    """Prueft Neutext auf Dinge, die in der Excel Schaden anrichten oder auf kaputtes PDF-Lesen hindeuten
    (Ligaturen, Zeichencode-Reste, nicht zusammengefuegte Silbentrennung, Tabellen-Spaltentrenner; Muster in config planer:).

    Rueckgabe: steuerzeichen (nicht schreibbar), verdaechtig (Liste einfacher Hinweise -> nie automatisch),
    formel (Text beginnt mit '=' und wuerde von Excel als Formel gelesen)."""
    pl = _planer_cfg(cfg)
    zeichen = "".join(pl.get("verdaechtige_zeichen", []))
    muster = [re.compile(m) for m in pl.get("verdaechtige_muster", [])]
    trennung = [re.compile(m) for m in pl.get("trennmuster", [])]
    ausnahmen = {str(w).casefold() for w in pl.get("trennung_ausnahmen", [])}
    tabelle = [re.compile(m, re.MULTILINE) for m in pl.get("tabellen_muster", [])]
    tab_min = max(int(pl.get("tabellen_min_treffer", 1)), 1)
    out: dict[str, Any] = {"steuerzeichen": False, "verdaechtig": [], "formel": False}
    for t in texte:
        if not isinstance(t, str) or not t:
            continue
        if ILLEGAL_CHARACTERS_RE.search(t):
            out["steuerzeichen"] = True
        if zeichen and any(z in t for z in zeichen):
            h = "Der neue Text enthält Zeichen (z. B. Ligaturen), die beim Lesen des PDF verfälscht sein können. Bitte mit dem PDF vergleichen."
            if h not in out["verdaechtig"]:
                out["verdaechtig"].append(h)
        if any(m.search(t) for m in muster):
            h = "Der neue Text enthält Zeichencode-Reste aus dem PDF (z. B. „(cid:12)“). Die PDF-Schrift wurde nicht sauber gelesen."
            if h not in out["verdaechtig"]:
                out["verdaechtig"].append(h)
        if any(_trennung_ohne_ausnahme(m, t, ausnahmen) for m in trennung):
            h = ("Der neue Text enthält ein Trennmuster wie „Wort- Rest“ (Silbentrennung am Zeilenende). Das Wort könnte "
                 "zerrissen sein. Bitte mit dem PDF vergleichen.")
            if h not in out["verdaechtig"]:
                out["verdaechtig"].append(h)
        if tabelle and sum(len(m.findall(t)) for m in tabelle) >= tab_min:
            h = ("Der neue Text enthält senkrechte Striche als Spaltentrenner (Text aus einer Tabelle?). Tabellen werden nicht "
                 "gelesen. Bitte mit dem PDF vergleichen.")
            if h not in out["verdaechtig"]:
                out["verdaechtig"].append(h)
        if t.startswith("="):
            out["formel"] = True
    return out


_ZAHL_TEXT = re.compile(r"-?\d+(\.\d+)?")


def zahl_aus_text(text: str, zahlenformat: str | None = None) -> int | float | None:
    """Zahl, die genau diesen Text als Anzeige hat -- sonst None (dann bliebe nur Text, das Zahlenformat ginge verloren).

    Ohne `zahlenformat` wird nur geprueft, ob der Text eine einfache Zahl ist (Planer, kennt das Format nicht);
    mit Format muss die Anzeige der Zahl (siehe `zahl_als_text`) buchstabengleich sein (Applier). Hoechstens 15 Ziffern."""
    t = text.strip()
    if not _ZAHL_TEXT.fullmatch(t) or sum(c.isdigit() for c in t) > 15:
        return None
    wert: int | float = float(t) if "." in t else int(t)
    if zahlenformat is not None and zahl_als_text(wert, zahlenformat) != t:
        return None
    return wert


def _wort_gleich(a: str, b: str) -> bool:
    """Zwei Woerter gelten als gleich: gleiche Normform; Woerter ohne Normform ('•') muessen buchstabengleich sein."""
    na, nb = norm(a), norm(b)
    return na == nb if (na or nb) else a == b


def _woerter_gleich(excel: list[str], pdf: list[str]) -> bool:
    return len(excel) == len(pdf) and bool(pdf) and all(_wort_gleich(x, y) for x, y in zip(excel, pdf))


@dataclass
class Plan:
    aenderungen: list[Aenderung]
    layout: list[dict]
    warnungen: list[str] = field(default_factory=list)
    ignoriert: list[dict] = field(default_factory=list)
    markierte_woerter: int = 0
    luecken: list[dict] = field(default_factory=list)   # Schema 2: Dinge, die nicht gelesen werden konnten
    absaetze: list[Absatz] = field(default_factory=list)   # Schema 3: PDF-Absaetze fuer die Gegenprobe (ADR 0004)


def _bbox(tokens: list[PdfToken]) -> tuple[float, float, float, float]:
    return (min(t.bbox[0] for t in tokens), min(t.bbox[1] for t in tokens),
            max(t.bbox[2] for t in tokens), max(t.bbox[3] for t in tokens))


@dataclass
class _Ort:
    """Ankerstelle fuer einen NEUEN Absatz bzw. eine neue Section (aus der Kette oder, unsicher, aus dem alten Matcher)."""
    blatt: str
    zeile: int
    spalte: int
    art: str                 # kopf | absatz | zelle | teil
    start: int
    end: int
    text: str
    score: float
    stufe: str
    davor: bool              # der Anker steht VOR dem neuen Block (sonst danach)
    gruende: list[str] = field(default_factory=list)
    codes: list[str] = field(default_factory=list)

    @property
    def ziel(self) -> str:
        return f"{self.blatt}!{get_column_letter(self.spalte)}{self.zeile}"


# Saetze zu Hinweis-Codes, die der PDF-Leser (noch) nicht erklaert (einfaches Deutsch, ohne Dokumentinhalt)
_CODE_SATZ = {
    "fortsetzung_unklar": "Der Absatz geht vielleicht über einen Seitenumbruch weiter (oder ist ein neuer Absatz). "
                          "Das ist nicht sicher; bitte prüfen.",
    "absatzgrenze_unklar": "Die Grenze zwischen zwei Absätzen ist im PDF nicht sicher erkennbar.",
    "trennung_unklar": "Eine Silbentrennung am Zeilenende konnte nicht sicher aufgelöst werden.",
    "rand_unklar": "Die Markierung endet nicht sauber an einer Wortgrenze.",
    "farbe_uneinig": "Zwei Lesewege sehen die Farbe dieser Stelle verschieden.",
    "tinte_unerklaert": "Im Seitenbild gibt es Farbe, die keiner Markierung zugeordnet werden konnte.",
}

# Gruende (Codes) in einfachem Deutsch, wenn eine Aenderung nicht automatisch eingetragen wird
_GRUND = {
    "mehrdeutig": "Der alte Absatz steht mehrmals in der Excel; die Reihenfolge entscheidet nicht, welche Stelle gemeint ist.",
    "widerspruch": "Der alte Absatz steht in der Excel an einer Stelle, die nicht zur Reihenfolge im PDF passt.",
    "excel_weicht_ab": "Der alte Absatz aus dem PDF steht so nicht in der Excel; die Stelle ist nur ähnlich.",
    "excel_umbruch": "In der Excel steht an der Stelle ein Zeilenumbruch.",
    "naht_unklar": "Die genaue Stelle im Excel-Absatz konnte nicht sicher bestimmt werden.",
    "mehrere_zellen": "Die Änderung betrifft mehrere Excel-Zellen.",
    "versteckt": "Die Stelle in der Excel ist ausgeblendet.",
    "kopfspalte": "Die Stelle liegt in der Nummern- oder Titelspalte und nicht im Fließtext.",
    "layout_unsicher": "Der Aufbau des Blatts wurde nur unsicher erkannt.",
    "section_unpassend": "Die Section passt nicht zur gefundenen Stelle.",
    "kurz": "Sehr kurzer Text, die Stelle ist nicht sicher.",
    "anker_teil": "Der PDF-Text deckt nur ein Stück des längeren Excel-Absatzes ab (Seitenumbruch?); der neue Absatz könnte "
                  "mitten im Satz landen.",
    "anker_luecke": "Zwischen der Ankerstelle und dem nächsten zugeordneten Absatz steht in der Excel Text, der keinem PDF-Absatz "
                    "zugeordnet ist.",
    "anker_unsicher": "Der Absatz davor konnte in der Excel nicht sicher zugeordnet werden; die Einfügestelle ist nur ungefähr bestimmbar.",
}

_A1_SATZ = {
    "leerraum": "Leerzeichen", "satzzeichen": "ein Satzzeichen", "gross_klein": "Groß-/Kleinschreibung",
    "ziffer": "eine Ziffer", "buchstabe": "ein Buchstabe", "pdf_laenger": "im PDF steht mehr Text",
    "excel_laenger": "in der Excel steht mehr Text",
}


def _in(a: Fundstelle, b: Fundstelle) -> bool:
    """Liegt Fundstelle a ganz in b (gleiche Zelle)?"""
    if a.blatt != b.blatt or len(a.zellen) != 1:
        return False
    z, s, x, y = a.zellen[0]
    return any((z2, s2) == (z, s) and x2 <= x and y <= y2 for z2, s2, x2, y2 in b.zellen)


class _Planer:
    def __init__(self, pdf: PdfAnalyse, xl: ExcelModell, cfg: dict) -> None:
        self.pdf, self.xl, self.cfg = pdf, xl, cfg
        self.matcher = Matcher(xl, cfg)
        self.out: list[Aenderung] = []
        self.absaetze: list[Absatz] = []
        self._tr: dict[tuple, Treffer | None] = {}
        self.verbraucht: set[int] = set()   # Spuren, die schon als Teil einer neuen Section behandelt wurden
        self._ag: list[str] = []                    # Gruende, warum die Ankerstelle des aktuellen neuen Blocks unsicher ist
        self._risiko_cache: dict[tuple[str, int], list[str]] = {}               # (Blatt, Einfuegezeile) -> Risiken
        self._sperre: dict[str, list[str]] = {}                                 # Aenderung -> Sperr-Codes des Absatzes
        self._fenster: tuple | None = None          # Reissverschluss-Fenster fuer den Ersatz-Matcher
        self._deckel: list[str] | None = None       # Ersatz-Matcher aktiv: hoechstens "mittel", mit diesen Codes
        self._spur_von: dict[str, int] = {}         # Aenderung -> Spur-Nummer (Index)
        self._geloescht_bis: dict[tuple[str, int, int], list[tuple[int, int]]] = {}   # Zelle -> geloeschte Bereiche ganzer Absaetze
        self._ist: str = ""                         # Excel-Text der Ersatz-Stelle (fuer ist_absatz)
        self._soll: str = ""

    # ------------------------------------------------------------ Helfer
    def treffer(self, b: PdfBlock) -> Treffer | None:
        key = (b.idx, self._fenster)
        if key not in self._tr:
            a = [(i, norm(t.old)) for i, t in enumerate(b.tokens) if t.old and norm(t.old)]
            self._tr[key] = self.matcher.finde(a, b.section_alt, ist_kopf=(b.kind == "ueberschrift"), fenster=self._fenster)
        return self._tr[key]

    def _basis(self, b: PdfBlock, typ: str, toks: list[PdfToken], alt: str, neu: str) -> Aenderung:
        ae = Aenderung(id=f"A{len(self.out) + 1:04d}", typ=typ, seite=toks[0].page, bbox=_bbox(toks),
                       section=b.section_neu or b.section_alt, alt=alt, neu=neu,
                       woerter=sum(1 for t in toks if t.status != KEEP), block_idx=b.idx,
                       kontext=(b.neu_text or b.alt_text)[:400])
        self.out.append(ae)
        if self._aktuelle_spur is not None:
            self._spur_von[ae.id] = self._aktuelle_spur
            self._sperre[ae.id] = list(self._spur_sperre)
            self._hinweise_nur(ae)
        ae.ist_absatz = self._ist
        ae.soll_absatz = self._soll
        return ae

    _a1_satz: str = ""
    _aktuelle_spur: int | None = None
    _spur_sperre: list[str] = []
    _spur_hinweise: list[str] = []

    def _hinweise_nur(self, ae: Aenderung) -> None:
        for c in self._spur_hinweise:
            satz = hinweis_text(c)
            if satz:
                self._hinweis(ae, satz)

    def _codes_des_absatzes(self, sp: Spur, mit: Iterable[Spur] = ()) -> tuple[list[str], list[str]]:
        """(Sperr-Codes, reine Hinweis-Codes) ueber ALLE Woerter des Absatzes (nicht nur die geaenderten).
        Sperr-Codes stehen in config planer.unsichere_hinweise (Praefix "kommentar" gilt fuer kommentar_*)."""
        pl = _planer_cfg(self.cfg)
        liste = [str(c) for c in pl.get("unsichere_hinweise", [])]
        nur = {str(c) for c in pl.get("hinweis_codes", [])}
        sperre: list[str] = []
        hinweise: list[str] = []
        for s in [sp, *mit]:
            codes = list(s.codes)
            for t in s.tokens:
                codes += [c for c in t.hinweis.split(";") if c]
            for c in codes:
                if any(c == x or c.startswith(x + "_") for x in liste):
                    if c not in sperre:
                        sperre.append(c)
                elif (c in nur or c.split("_")[0] in nur) and c not in hinweise:
                    hinweise.append(c)
        return sperre, hinweise

    def _setze(self, ae: Aenderung, stufe: str, grund: str = "", codes: Iterable[str] = ()) -> None:
        codes = list(codes)
        # Neutext pruefen: Steuerzeichen koennen nicht geschrieben werden, Ligaturen/(cid:N) sind nie "sicher"
        texte = [ae.neu] + [o.text for o in ae.ops] + [t for o in ae.ops for t in o.zellen.values()]
        p = neutext_probleme(texte, self.cfg)
        if p["steuerzeichen"]:
            self._nicht(ae, "Der neue Text enthält Steuerzeichen und kann nicht in die Excel geschrieben werden.")
            self._code(ae, "steuerzeichen")
            return
        if self._deckel is not None:
            if stufe == HOCH:
                stufe = MITTEL
            codes += self._deckel
            if self._deckel and "excel_weicht_ab" in self._deckel:
                grund = (grund + "; " if grund else "") + _GRUND["excel_weicht_ab"]
                if self._a1_satz:
                    self._hinweis(ae, self._a1_satz)
        sperre = self._sperre.get(ae.id, [])
        if sperre:
            if stufe == HOCH:
                stufe = MITTEL
            for c in sperre:
                satz = hinweis_text(c) or _CODE_SATZ.get(c, "")
                if satz:
                    self._hinweis(ae, satz)
                codes.append("sperre:" + c)
            grund = (grund + "; " if grund else "") + ("Im Absatz gibt es eine Stelle, die nicht sicher gelesen werden konnte "
                                                       "(siehe Hinweise). Der ganze Absatz wird nicht automatisch eingetragen.")
        if p["verdaechtig"]:
            if stufe == HOCH:
                stufe = MITTEL
            for h in p["verdaechtig"]:
                self._hinweis(ae, h)
            codes.append("neutext_verdaechtig")
            grund = (grund + "; " if grund else "") + "Der neue Text könnte aus dem PDF verfälscht gelesen worden sein."
        if p["formel"] and any(o.art in ("ersetzen", "einfuegen", "zeile_einfuegen") for o in ae.ops):
            self._hinweis(ae, "Der neue Text beginnt mit „=“. Er wird als Text eingetragen, nicht als Formel.")
        for op in ae.ops:                # Zahl-Zelle: der neue Text muss wieder eine Zahl sein, sonst geht das Zahlenformat verloren
            if op.art in ("ersetzen", "einfuegen") and self.xl.ist_zahl(op.blatt, op.zeile, op.spalte):
                alt_z = self.xl.text(op.blatt, op.zeile, op.spalte)
                ergebnis = alt_z[:op.start] + op.text + alt_z[op.ende:]
                if ergebnis.strip() and zahl_aus_text(ergebnis) is None:
                    if stufe == HOCH:
                        stufe = MITTEL
                    self._hinweis(ae, "Die Zelle in der Excel ist eine Zahl, der neue Text ist keine. Das Zahlenformat geht verloren.")
                    grund = (grund + "; " if grund else "") + "Zahlenformat geht verloren (die Zelle ist eine Zahl, der neue Text nicht)."
                    codes.append("zahlformat")
                    break
        risiken = self._zeilen_risiken(ae)
        if risiken:                      # Zeile einfuegen verschiebt Dinge, die openpyxl nicht mitzieht: nie automatisch
            if stufe == HOCH:
                stufe = MITTEL
            for r in risiken:
                self._hinweis(ae, r)
            codes.append("zeilen_risiko")
            grund = (grund + "; " if grund else "") + ("Beim Einfügen der neuen Zeile könnten Dinge in der Excel verrutschen "
                                                       "oder brechen (siehe Hinweise). Bitte vor dem Eintragen prüfen.")
        ae.konfidenz = stufe
        ae.status = STATUS_VON_STUFE.get(stufe, NICHT)
        for c in codes:
            self._code(ae, c)
        if grund:
            ae.grund = (ae.grund + "; " if ae.grund else "") + grund

    @staticmethod
    def _code(ae: Aenderung, code: str) -> None:
        if code and code not in ae.pruefcodes:
            ae.pruefcodes.append(code)

    def _zeilen_risiken(self, ae: Aenderung) -> list[str]:
        """Hindernisse (einfache Saetze) aller Zeilen-Einfuegungen einer Aenderung. Leer = unbedenklich.
        Was der Applier selbst mitzieht (verbundene Zellen unterhalb, ausgeblendete Zeilen darunter), zaehlt nicht."""
        out: list[str] = []
        for op in ae.ops:
            if op.art != "zeile_einfuegen" or op.abgedeckt:
                continue
            pos = op.zeile if op.vor else op.zeile + 1
            if (op.blatt, pos) not in self._risiko_cache:
                self._risiko_cache[(op.blatt, pos)] = self.xl.einfuege_hindernisse(op.blatt, pos)
            out += [r for r in self._risiko_cache[(op.blatt, pos)] if r not in out]
        return out

    @staticmethod
    def _hinweis(ae: Aenderung, text: str) -> None:
        if text not in ae.hinweise:
            ae.hinweise.append(text)

    def _nicht(self, ae: Aenderung, grund: str, ziel: str = "") -> None:
        ae.status, ae.konfidenz, ae.ops = NICHT, NIEDRIG, []
        ae.grund = (ae.grund + "; " if ae.grund else "") + grund
        if ziel:
            ae.ziel = ziel

    def _kein_treffer(self, ae: Aenderung, b: PdfBlock) -> None:
        tr = self.treffer(b)
        if tr is None:
            self._nicht(ae, "Der Absatz wurde in der Excel nicht gefunden.")
            self._code(ae, "nicht_gefunden")
        else:
            self._nicht(ae, f"Zu unsichere Stelle in der Excel (Übereinstimmung {tr.score:.0%}).", tr.kandidat.ziel)
            self._code(ae, "zu_unsicher")
            ae.excel_text = tr.kandidat.text[:300]
        ae.score = tr.score if tr else 0.0

    def _ziel(self, xtok) -> str:
        return f"{xtok.blatt}!{get_column_letter(xtok.spalte)}{xtok.zeile}"

    @staticmethod
    def _gruppen(b: PdfBlock) -> Iterator[tuple[int, int]]:
        i, n = 0, len(b.tokens)
        while i < n:
            if b.tokens[i].status == KEEP:
                i += 1
                continue
            j = i
            while j < n and b.tokens[j].status != KEEP:
                j += 1
            yield i, j
            i = j

    # ------------------------------------------------------------ Hauptschleife
    def lauf(self) -> None:
        self.idx = ExaktIndex(self.xl, self.cfg, self.matcher.index)
        self.spuren = spuren(self.pdf, self.idx, self.cfg)
        self.zuo = self._zuordnen()
        self._zugeordnet = self._zugeordnete_zeilen()
        for i, sp in enumerate(self.spuren):
            if i in self.verbraucht or not sp.hat_aenderung:
                continue
            self._spur_start(i)
            if sp.ganz_neu:
                self._neu(i)
            elif sp.ganz_geloescht:
                self._geloescht_spur(i)
            else:
                self._text_spur(i)
            self._spur_ende()
        self._loeschungen_koppeln()
        self._leere_zellen_naht()

    def _leere_zellen_naht(self) -> None:
        """Wird der ganze Inhalt einer Zelle geloescht und an ihrem Rand ein neuer Absatz eingesetzt, braucht der neue Absatz
        keinen Zeilenumbruch zum (verschwundenen) Nachbarn."""
        leer = {(o.blatt, o.zeile, o.spalte) for a in self.out for o in a.ops
                if o.art == "ersetzen" and o.text == "" and o.start == 0 and o.ende == len(self.xl.text(o.blatt, o.zeile, o.spalte))
                and o.ende > 0}
        for a in self.out:
            if a.typ not in ("absatz_neu",):
                continue
            for o in a.ops:
                if o.art == "ersetzen" and o.start == o.ende and (o.blatt, o.zeile, o.spalte) in leer:
                    o.text = o.text.strip("\n")

    def _spur_start(self, i: int, mit: Iterable[Spur] = ()) -> None:
        self._aktuelle_spur = i
        self._spur_sperre, self._spur_hinweise = self._codes_des_absatzes(self.spuren[i], mit)
        self._fenster = self.zuo[i].fenster
        self._deckel = None
        self._a1_satz = ""
        self._ist, self._soll = "", ""

    def _spur_ende(self) -> None:
        self._aktuelle_spur = None
        self._spur_sperre, self._spur_hinweise = [], []
        self._fenster, self._deckel = None, None
        self._ist, self._soll = "", ""

    # ------------------------------------------------------------ Reissverschluss
    def _zuordnen(self) -> list[Zuordnung]:
        nf = self.cfg.get("normalform")
        alle: list[list[tuple[Fundstelle, str]]] = []
        for sp in self.spuren:
            cands: list[tuple[Fundstelle, str]] = []
            ids = {l.aenderung for l in sp.laeufe if l.aenderung}
            alt, opt_alt = zustand_mit_optional(sp.laeufe, set())
            neu, opt_neu = zustand_mit_optional(sp.laeufe, ids)
            orte: set[tuple] = set()
            if alt.strip():
                for f in self.idx.finde(alt, opt_alt):
                    cands.append((f, "alt"))
                    orte.add((f.blatt, f.zellen))
            if ids and neu.strip() and normalform(neu, nf)[0] != normalform(alt, nf)[0]:
                for f in self.idx.finde(neu, opt_neu):
                    if (f.blatt, f.zellen) in orte:
                        continue
                    # Ein Stueck des alten Absatzes ist kein "schon eingetragen" (am Ende gestrichenes Wort); umgekehrt ist der
                    # alte Absatz als Stueck des schon eingetragenen neuen (am Ende angehaengter Satz) keine Stelle zum Aendern.
                    if any(_in(f, g) for g, art in cands if art == "alt"):
                        continue
                    cands = [(g, art) for g, art in cands if not (art == "alt" and _in(g, f))]
                    cands.append((f, "neu"))
            alle.append(self._filtern(sp, cands))
        return reissverschluss(alle)

    @staticmethod
    def _filtern(sp: Spur, cands: list[tuple[Fundstelle, str]]) -> list[tuple[Fundstelle, str]]:
        """Fundstellen, die zaehlen: Ueberschriften nur fuer Ueberschriften; Blaetter ohne Gliederung (z. B. "Historie") nur, wenn
        es im gegliederten Blatt keine gibt; Nummern-/Titelspalte fuer Fliesstext nur, wenn es sonst keine gibt; sichtbare
        Stellen schlagen versteckte; ganze Absaetze schlagen Teilstuecke."""
        if sp.kind != "ueberschrift":
            cands = [c for c in cands if c[0].art != "kopf"]
            if any(not c[0].kopf_spalte for c in cands):
                cands = [c for c in cands if not c[0].kopf_spalte]
        # ein ganzer Excel-Absatz schlaegt Stuecke laengerer Absaetze (gleicher Satz zufaellig mitten in einem anderen Absatz)
        if any(c[0].art != "teil" for c in cands):
            cands = [c for c in cands if c[0].art != "teil"]
        if any(c[0].gliederung for c in cands):
            cands = [c for c in cands if c[0].gliederung]
        if any(not c[0].versteckt for c in cands):
            cands = [c for c in cands if not c[0].versteckt]
        return cands

    def _zugeordnete_zeilen(self) -> set[tuple[str, int, int, int]]:
        """(Blatt, Zeile, Spalte, Zeilenanfang) aller Excel-Textzeilen, die ein PDF-Absatz sicher (auf der Kette) belegt."""
        out: set[tuple[str, int, int, int]] = set()
        for z in self.zuo:
            if not z.automatisch or z.fund is None:
                continue
            for zeile, spalte, a, e in z.fund.zellen:
                t = self.xl.text(z.fund.blatt, zeile, spalte)
                for m in re.finditer(r"[^\n]+", t):
                    if m.start() < max(e, a + 1) and m.end() > a:
                        out.add((z.fund.blatt, zeile, spalte, m.start()))
        return out

    def _unzugeordnet_dazwischen(self, blatt: str, lo: tuple, hi: tuple | None) -> bool:
        for key, (z, s, a, _) in self.idx.verwaltete_zeilen(blatt):
            if key <= lo:
                continue
            if hi is not None and key >= hi:
                break
            if (blatt, z, s, a) not in self._zugeordnet:
                return True
        return False

    def _stufe_an_stelle(self, sp: Spur, fund: Fundstelle, z: Zuordnung, gruende: list[str], codes: list[str],
                         text_aenderung: bool = True) -> str:
        """Sicherheit einer In-Zellen-Aenderung an einer exakten Fundstelle (A1 ist erfuellt)."""
        stufe = HOCH
        if not z.automatisch:
            stufe = MITTEL
            codes.append(z.status)
            gruende.append(_GRUND.get(z.status, ""))
        if fund.versteckt:
            stufe = MITTEL
            codes.append("versteckt")
            gruende.append({"blatt": f"das Blatt '{fund.blatt}' ist ausgeblendet", "zeile": "die Zeile ist ausgeblendet",
                            "spalte": "die Spalte ist ausgeblendet"}.get(fund.versteckt, _GRUND["versteckt"]))
        if sp.kind != "ueberschrift" and fund.kopf_spalte:
            stufe = MITTEL
            codes.append("kopfspalte")
            gruende.append(_GRUND["kopfspalte"])
        lay = self.xl.layouts.get(fund.blatt)
        # Blatt-Aufbau-Sperre (ADR 0004, 7): gilt nicht fuer sichtbare In-Zellen-Aenderungen im Fliesstext an einer eindeutigen
        # (auch ueber die Reihenfolge eindeutigen) exakten Stelle; fuer Nummer/Titel und neue Zeilen bleibt sie.
        if lay is not None and lay.sicherheit != "hoch" and (fund.kopf_spalte or not text_aenderung or not z.automatisch):
            stufe = MITTEL
            codes.append("layout_unsicher")
            gruende.append(f"der Aufbau des Blatts '{fund.blatt}' wurde nur unsicher erkannt")
        sec = sp.erster.section_alt
        if sec and fund.section and not fund.section_unsicher and fund.section != sec and \
                self.cfg["matcher"].get("section_unpassend_stufe", MITTEL) != HOCH:
            stufe = MITTEL
            codes.append("section_unpassend")
            gruende.append("Section passt nicht zur gefundenen Stelle")
        if fund.section_unsicher and fund.kopf_spalte:
            stufe = MITTEL
            codes.append("section_unsicher")
            gruende.append("die Section-Nummer steht als Zahl in der Excel und ist nicht eindeutig (2.1 oder 2.10?)")
        if len(_woerter_norm(sp.alt)) <= self.cfg["schwellen"]["kurz_tokens"] and z.status != "eindeutig" and \
                not (sec and fund.section == sec):
            stufe = MITTEL
            codes.append("kurz")
            gruende.append(_GRUND["kurz"])
        return stufe

    # ------------------------------------------------------------ Aenderungen je Absatz (exakt)
    def _spur_aenderungen(self, sp: Spur) -> tuple[list[Aenderung], list[Lauf]]:
        """Je Gruppe eine Aenderung (IDs in Reihenfolge); dazu die Laeufe mit den echten IDs."""
        a = sp.aufbau
        aes: list[Aenderung] = []
        ids: dict[str, str] = {}
        for gi, runs in enumerate(a.gruppen):
            gid = sp.laeufe[runs[0]].aenderung
            seg = sp.laeufe[runs[0]:runs[-1] + 1]
            alt = zustand(seg, set())
            neu = zustand(seg, {gid})
            toks = [sp.tokens[i] for i in a.gruppen_token[gi]]
            seite = toks[0].page
            b = sp.block_von_token(a.gruppen_token[gi][0])
            typ = "text_ersetzt" if alt and neu else "text_eingefuegt" if neu else "text_geloescht"
            ae = Aenderung(id=f"A{len(self.out) + 1:04d}", typ=typ, seite=seite,
                           bbox=_bbox([t for t in toks if t.page == seite]), section=b.section_neu or b.section_alt,
                           alt=alt, neu=neu, woerter=len(a.zaehl_token[gi]), block_idx=b.idx,
                           kontext=(sp.neu or sp.alt)[:400])
            self.out.append(ae)
            self._spur_von[ae.id] = self._aktuelle_spur if self._aktuelle_spur is not None else -1
            self._sperre[ae.id] = list(self._spur_sperre)
            self._hinweise_nur(ae)
            ids[gid] = ae.id
            aes.append(ae)
        laeufe = umbenennen(sp.laeufe, ids)
        soll = zustand(laeufe, set(ids.values()))
        for ae in aes:
            ae.soll_absatz = soll
        return aes, laeufe

    def _absatz(self, blatt: str, zellen: Iterable[tuple[int, int, int, int]], laeufe: list[Lauf]) -> None:
        self.absaetze.append(Absatz(id=f"P{len(self.absaetze) + 1:04d}", blatt=blatt, zellen=[tuple(z) for z in zellen],
                                    laeufe=laeufe))

    def _text_spur(self, i: int) -> None:
        sp, z = self.spuren[i], self.zuo[i]
        if z.fund is not None and z.art == "neu" and z.automatisch:
            self._schon_eingetragen(sp, z.fund)
            return
        if z.fund is not None and z.art == "alt" and z.status in ("eindeutig", "durch_reihenfolge", "mehrdeutig"):
            self._exakt(i, z.fund)
            return
        # keiner / widerspruch: alter Matcher im Fenster zwischen den Ankern, hoechstens "mittel"
        if self._ersatz_text(i):
            return
        if z.fund is not None and z.art == "alt":            # Widerspruch: die exakte Stelle als Vorschlag
            self._exakt(i, z.fund)
            return
        self._deckel = ["excel_weicht_ab"] if z.status == "keiner" else [z.status]
        for b in sp.bloecke:
            for i0, i1 in self._gruppen(b):
                self._text_gruppe(b, i0, i1)

    def _schon_eingetragen(self, sp: Spur, fund: Fundstelle) -> None:
        aes, laeufe = self._spur_aenderungen(sp)
        for ae in aes:
            ae.ist_absatz, ae.excel_text, ae.score = fund.text, fund.text[:300], 1.0
            self._nicht(ae, "Der neue Text steht schon so in der Excel (die Änderung ist offenbar schon eingetragen).", fund.ziel)
            self._code(ae, "schon_eingetragen")

    def _exakt(self, i: int, fund: Fundstelle) -> None:
        sp, z = self.spuren[i], self.zuo[i]
        aes, laeufe = self._spur_aenderungen(sp)
        seq = sp.erster.idx * 1000
        so = ops_fuer_spur(laeufe, fund, self.xl.text, self.cfg, seq)
        if not so.vergleich.gleich:                       # sollte nicht vorkommen (Fundstelle ist exakt)
            for ae in aes:
                self._nicht(ae, "Der alte Absatz steht so nicht in der Excel. Bitte von Hand prüfen.", fund.ziel)
                self._code(ae, "a1_" + (so.vergleich.art or "abweichung"))
            return
        self._absatz(fund.blatt, fund.zellen, laeufe)
        for ae in aes:
            ae.ist_absatz, ae.excel_text, ae.score, ae.ziel = fund.text, fund.text[:300], 1.0, fund.ziel
            gruende: list[str] = []
            codes: list[str] = []
            stufe = self._stufe_an_stelle(sp, fund, z, gruende, codes)
            ops = so.ops.get(ae.id, [])
            for c in so.probleme.get(ae.id, []):
                stufe = MITTEL
                codes.append(c)
                gruende.append(_GRUND.get(c, ""))
            if not ops:
                self._nicht(ae, "Die Stelle der Änderung im Excel-Absatz konnte nicht sicher bestimmt werden. Bitte von Hand "
                                "prüfen.", fund.ziel)
                for c in codes or ["naht_unklar"]:
                    self._code(ae, c)
                continue
            for o in ops:
                o.seq = seq + sp.aufbau.gruppen_token[aes.index(ae)][0]
            ae.ops = ops
            ae.ziel = f"{fund.blatt}!{get_column_letter(ops[0].spalte)}{ops[0].zeile}"
            self._setze(ae, stufe, "; ".join(g for g in dict.fromkeys(gruende) if g), codes)

    def _a1_codes(self, alt: str, excel: str) -> list[str]:
        v = vergleiche(alt, excel, frozenset(), self.cfg.get("normalform"))
        return ["a1_" + v.art] if v.art else []

    def _ersatz_text(self, i: int) -> bool:
        """keiner/widerspruch: ein Vorschlag ueber den alten Matcher im Fenster. True, wenn er etwas geplant hat."""
        sp, z = self.spuren[i], self.zuo[i]
        schw = self.cfg["schwellen"]
        trs = [self.treffer(b) for b in sp.bloecke]
        if not any(tr is not None and tr.score >= schw["mittel"] for tr in trs):
            return False
        for b, tr in zip(sp.bloecke, trs):
            basis = ["excel_weicht_ab"] if z.status == "keiner" else [z.status, "excel_weicht_ab"]
            if tr is not None:
                a1 = self._a1_codes(b.alt_text if len(sp.bloecke) > 1 else sp.alt, tr.kandidat.text)
                art = a1[0][3:] if a1 else ""
                self._a1_satz = ""
                self._ist = tr.kandidat.text
                self._soll = sp.neu
                self._deckel = basis + a1
                if art:
                    self._a1_satz = f"Der alte Absatz aus dem PDF steht so nicht in der Excel (Abweichung: {_A1_SATZ.get(art, art)})."
            else:
                self._deckel = basis
            if b.hat_aenderung:
                if b.ganz_geloescht:
                    self._geloescht(b)
                else:
                    for i0, i1 in self._gruppen(b):
                        self._text_gruppe(b, i0, i1)
        return True

    # ------------------------------------------------------------ Geloeschte Absaetze
    def _geloescht_spur(self, i: int) -> None:
        sp, z = self.spuren[i], self.zuo[i]
        if z.fund is not None and z.art == "alt" and (z.status in ("eindeutig", "durch_reihenfolge", "mehrdeutig")
                                                       or (z.status == "widerspruch" and not self._ersatz_moeglich(i))):
            fund = z.fund
            aes, laeufe = self._spur_aenderungen(sp)
            ist_kopf = sp.kind == "ueberschrift"
            for ae in aes:
                ae.typ = "section_geloescht" if ist_kopf else "absatz_geloescht"
            if len(aes) != 1:                                  # ein ganz gestrichener Absatz ist genau eine Gruppe
                for ae in aes:
                    self._nicht(ae, "Der gestrichene Absatz konnte nicht als Ganzes zugeordnet werden. Bitte von Hand prüfen.",
                                fund.ziel)
                return
            ae = aes[0]
            ae.alt = sp.alt
            belegt = {(z0, s0): r for (b0, z0, s0), r in self._geloescht_bis.items() if b0 == fund.blatt}
            ops, bereiche = ganz_loeschen(fund, self.xl.text, ae.id, sp.erster.idx * 1000, belegt)
            # Die Operationen nehmen den Zeilenumbruch selbst mit (kein Aufraeumen leerer Zeilen durch den Applier noetig,
            # sonst stimmen die Stellen der anderen Absaetze derselben Zelle fuer die Gegenprobe A3 nicht mehr).
            for z0, s0, a0, e0 in bereiche:
                self._geloescht_bis[(fund.blatt, z0, s0)] = self._geloescht_bis.get((fund.blatt, z0, s0), []) + [(a0, e0)]
            self._absatz(fund.blatt, bereiche, laeufe)
            ae.ist_absatz, ae.excel_text, ae.score, ae.ziel, ae.ops = fund.text, fund.text[:300], 1.0, fund.ziel, ops
            gruende: list[str] = []
            codes: list[str] = []
            stufe = self._stufe_an_stelle(sp, fund, z, gruende, codes, text_aenderung=not ist_kopf)
            if fund.art == "teil":
                # der gestrichene Absatz ist nur ein Stueck eines laengeren Excel-Absatzes: der Rest bleibt stehen -- ob das so
                # gemeint ist, entscheidet ein Mensch
                stufe = MITTEL
                codes.append("teil_absatz")
                gruende.append("In der Excel steht im selben Absatz mehr Text, der im PDF nicht gestrichen ist. Nur der gestrichene "
                               "Text wird entfernt, der Rest bleibt stehen.")
                self._hinweis(ae, gruende[-1])
            self._setze(ae, stufe, "; ".join(g for g in dict.fromkeys(gruende) if g), codes)
            return
        self._deckel = ["excel_weicht_ab"] if z.status == "keiner" else [z.status, "excel_weicht_ab"]
        for b in sp.bloecke:
            tr = self.treffer(b)
            self._ist = tr.kandidat.text if tr is not None else ""
            self._deckel = (["excel_weicht_ab"] if z.status == "keiner" else [z.status, "excel_weicht_ab"]) + \
                (self._a1_codes(b.alt_text, tr.kandidat.text) if tr is not None else [])
            self._geloescht(b)

    def _ersatz_moeglich(self, i: int) -> bool:
        schw = self.cfg["schwellen"]
        return any((tr := self.treffer(b)) is not None and tr.score >= schw["mittel"] for b in self.spuren[i].bloecke)

    def _loeschungen_koppeln(self) -> None:
        """Eine gestrichene Section und der gestrichene Inhalt darunter sind eine Gruppe (`gruppe`): der Status gilt fuer alle
        gemeinsam.

        Sonst bliebe eine verwaiste Zeile zurueck (Nummer und Titel weg, alter Inhalt steht noch da) oder ein Teil des
        Inhalts verschwindet, der Rest nicht. Ist ein Mitglied der Gruppe nicht sicher (pruefen) oder nicht eintragbar,
        bleiben alle anderen auf pruefen. Die Vorschlaege (Operationen) bleiben stehen."""
        je_spur: dict[int, list[Aenderung]] = {}
        for ae in self.out:
            if ae.typ in ("absatz_geloescht", "section_geloescht"):
                je_spur.setdefault(self._spur_von.get(ae.id, -1), []).append(ae)
        for kopf in [ae for ae in self.out if ae.typ == "section_geloescht"]:
            i = self._spur_von.get(kopf.id, -1)
            if i < 0:
                continue
            folge: list[Aenderung] = []
            k = i + 1
            while k < len(self.spuren) and self.spuren[k].ganz_geloescht and self.spuren[k].kind == "absatz":
                folge += [a for a in je_spur.get(k, []) if a.typ == "absatz_geloescht"]
                k += 1
            if not folge:
                continue
            gruppe = [kopf] + folge
            for m in gruppe:
                m.gruppe = kopf.id
            if all(m.status == UEBERNOMMEN for m in gruppe):
                continue
            for m in gruppe:
                if m.status == UEBERNOMMEN:
                    self._herabstufen(m, "Gehört zu einer gestrichenen Section, deren Löschung nicht ganz sicher ist: Section und "
                                         "gestrichener Inhalt werden nur zusammen gelöscht, sonst bliebe ein Rest stehen "
                                         "(z. B. eine Zeile mit altem Inhalt ohne Nummer und Titel).",
                                      "Section und gestrichener Inhalt gehören zusammen. Bitte alles zusammen prüfen.")
                    self._code(m, "gruppe")


    def _herabstufen(self, ae: Aenderung, grund: str, hinweis: str) -> None:
        """Aenderung von 'uebernommen' auf 'pruefen' (der Vorschlag mit seinen Operationen bleibt stehen)."""
        ae.status, ae.konfidenz = PRUEFEN, MITTEL
        ae.grund = (ae.grund + "; " if ae.grund else "") + grund
        self._hinweis(ae, hinweis)

    # ------------------------------------------------------------ Text in Absatz/Ueberschrift
    def _text_gruppe(self, b: PdfBlock, i0: int, i1: int) -> None:
        toks = b.tokens[i0:i1]
        alt = " ".join(t.old for t in toks if t.old)
        neu = " ".join(t.new for t in toks if t.new)
        typ = "text_ersetzt" if alt and neu else "text_eingefuegt" if neu else "text_geloescht"
        ae = self._basis(b, typ, toks, alt, neu)
        tr = self.treffer(b)
        schw = self.cfg["schwellen"]
        if tr is None or tr.score < schw["mittel"]:
            self._kein_treffer(ae, b)
            return
        cand = tr.kandidat
        ae.score = tr.score
        ae.ziel = cand.ziel
        ae.excel_text = cand.text[:300]
        stufe = tr.stufe
        gruende = list(tr.gruende)
        alt_idx = [i for i in range(i0, i1) if b.tokens[i].old]
        d_idx = [i for i in alt_idx if norm(b.tokens[i].old)]
        leer_idx = [i for i in alt_idx if not norm(b.tokens[i].old)]     # alte Woerter ohne Normform ("•", ".")
        seq = b.idx * 1000 + i0
        neu_w = [t.new for t in toks if t.new]

        if d_idx:
            fehlt = [i for i in d_idx if i not in tr.zuordnung]
            if fehlt:
                if self._schon_da(b, i0, i1, tr, neu_w):
                    self._nicht(ae, "Der neue Text steht schon in der Excel (die Änderung ist offenbar schon eingetragen).", cand.ziel)
                elif not neu_w and len(fehlt) == len(d_idx):
                    self._nicht(ae, "Der gestrichene Text steht nicht mehr in der Excel (vielleicht schon entfernt, vielleicht "
                                    "anders formuliert). Bitte von Hand prüfen.", cand.ziel)
                else:
                    self._nicht(ae, "Der alte Text steht in der Excel anders (abweichende Formulierung). "
                                    "Bitte von Hand prüfen.", cand.ziel)
                return
            js = sorted(tr.zuordnung[i] for i in d_idx)
            xt = [cand.tokens[j] for j in js]
            zellen = {(t.blatt, t.zeile, t.spalte) for t in xt}
            if len(zellen) > 1:
                if not self._kopf_mehrzellig(ae, b, i0, i1, cand, xt, neu, seq, stufe):
                    self._nicht(ae, "Die Änderung betrifft mehrere Excel-Zellen.", cand.ziel)
                return
            if js[-1] - js[0] + 1 != len(js):
                stufe = MITTEL
                gruende.append("in der Excel stehen zusätzliche Wörter zwischen dem alten Text")
            blatt, zeile, spalte = xt[0].blatt, xt[0].zeile, xt[0].spalte
            zt = self.xl.text(blatt, zeile, spalte)
            start, ende = xt[0].start, xt[-1].end
            vorn = [b.tokens[i].old for i in leer_idx if i < d_idx[0]]
            hinten = [b.tokens[i].old for i in leer_idx if i > d_idx[-1]]
            if vorn or hinten:       # gestrichene Aufzaehlungszeichen/Satzzeichen am Rand gehoeren zum Bereich
                rand = self._rand_leer(zt, start, ende, vorn, hinten, cand)
                if rand is None:
                    self._nicht(ae, "Ein gestrichenes Zeichen ohne Buchstaben (z. B. Aufzählungszeichen) wurde in der Excel "
                                    "nicht sicher gefunden. Bitte von Hand prüfen.", cand.ziel)
                    return
                start, ende = rand
            if neu:
                op = Operation("ersetzen", blatt, zeile, spalte, start, ende, neu, seq, aenderung=ae.id)
            else:
                s, e = self._luecke(zt, start, ende)
                op = Operation("ersetzen", blatt, zeile, spalte, s, e, "", seq, aenderung=ae.id)
            ae.ops = [op]
            ae.ziel = self._ziel(xt[0])
        elif leer_idx:
            # Nur Woerter ohne Normform ("•", "."): ihre Stelle aus dem Rohtext zwischen zwei sicheren Nachbarwoertern bestimmen
            span = self._leer_span(b, i0, i1, tr, cand)
            if span is None:
                self._nicht(ae, "Ein gestrichenes Zeichen ohne Buchstaben (z. B. Aufzählungszeichen) wurde in der Excel "
                                "nicht sicher gefunden. Bitte von Hand prüfen.", cand.ziel)
                return
            blatt, zeile, spalte, start, ende = span
            zt = self.xl.text(blatt, zeile, spalte)
            if neu:
                op = Operation("ersetzen", blatt, zeile, spalte, start, ende, neu, seq, aenderung=ae.id)
            else:
                s, e = self._luecke(zt, start, ende)
                op = Operation("ersetzen", blatt, zeile, spalte, s, e, "", seq, aenderung=ae.id)
            ae.ops = [op]
            ae.ziel = f"{blatt}!{get_column_letter(spalte)}{zeile}"
        else:
            # reine Einfuegung: nach dem naechsten vorangehenden, zugeordneten Wort
            vor = next((i for i in range(i0 - 1, -1, -1) if i in tr.zuordnung), None)
            if vor is not None:
                xt = cand.tokens[tr.zuordnung[vor]]
                if vor != i0 - 1:
                    stufe = MITTEL
                    gruende.append("Einfügestelle nur ungefähr bestimmbar")
                op = Operation("einfuegen", xt.blatt, xt.zeile, xt.spalte, xt.end, xt.end, " " + neu, seq, aenderung=ae.id)
            else:
                nach = next((i for i in range(i1, len(b.tokens)) if i in tr.zuordnung), None)
                if nach is None:
                    self._nicht(ae, "Keine feste Stelle für die Einfügung gefunden.", cand.ziel)
                    return
                xt = cand.tokens[tr.zuordnung[nach]]
                op = Operation("einfuegen", xt.blatt, xt.zeile, xt.spalte, xt.start, xt.start, neu + " ", seq, aenderung=ae.id)
            if self._schon_da(b, i0, i1, tr, neu_w):
                self._nicht(ae, "Der neue Text steht schon in der Excel (die Änderung ist offenbar schon eingetragen).",
                            self._ziel(xt))
                return
            ae.ops = [op]
            ae.ziel = self._ziel(xt)
        self._setze(ae, stufe, "; ".join(gruende))

    # ------------------------------------------------------------ Hilfen fuer Woerter ohne Normform und Idempotenz
    def _schon_da(self, b: PdfBlock, i0: int, i1: int, tr: Treffer, neu_w: list[str]) -> bool:
        """Steht der neue Text schon direkt an der Einfuegestelle? (zweiter Lauf auf dem eigenen Ergebnis)"""
        if not neu_w:
            return False
        cand = tr.kandidat
        zuo = tr.zuordnung
        vor = next((i for i in range(i0 - 1, -1, -1) if i in zuo), None)
        if vor is not None and cand.art != "kopf":
            xt = cand.tokens[zuo[vor]]
            rest = re.findall(r"\S+", self.xl.text(xt.blatt, xt.zeile, xt.spalte)[xt.end:cand.end])
            if _woerter_gleich(rest[:len(neu_w)], neu_w):
                return True
        nach = next((i for i in range(i1, len(b.tokens)) if i in zuo), None)
        if nach is not None and cand.art != "kopf":
            xt = cand.tokens[zuo[nach]]
            davor = re.findall(r"\S+", self.xl.text(xt.blatt, xt.zeile, xt.spalte)[cand.start:xt.start])
            if _woerter_gleich(davor[-len(neu_w):], neu_w):
                return True
        return False

    def _rand_leer(self, zt: str, s: int, e: int, vorn: list[str], hinten: list[str], cand: Kandidat) -> tuple[int, int] | None:
        """Erweitert den Bereich [s, e) um direkt angrenzende Woerter ohne Normform, falls sie wirklich dort stehen."""
        lo = cand.start if cand.art != "kopf" else 0
        hi = cand.end if cand.art != "kopf" else len(zt)
        for w in reversed(vorn):
            m = re.search(r"(\S+)\s*$", zt[:s])
            if not m or m.group(1) != w or m.start(1) < lo:
                return None
            s = m.start(1)
        for w in hinten:
            m = re.match(r"\s*(\S+)", zt[e:])
            if not m or m.group(1) != w or e + m.end() > hi:
                return None
            e += m.end()
        return s, e

    def _leer_span(self, b: PdfBlock, i0: int, i1: int, tr: Treffer, cand: Kandidat) -> tuple[str, int, int, int, int] | None:
        """Stelle (Blatt, Zeile, Spalte, Start, Ende) der alten Woerter ohne Normform einer Gruppe.

        Der Rohtext zwischen den beiden naechsten sicher zugeordneten Nachbarwoertern muss genau die alten Woerter des
        PDF enthalten -- sonst gibt es keine Stelle (None)."""
        if cand.art == "kopf":
            return None
        zuo = tr.zuordnung
        vor = next((i for i in range(i0 - 1, -1, -1) if i in zuo), None)
        nach = next((i for i in range(i1, len(b.tokens)) if i in zuo), None)
        zt = cand.zelle_text
        s = cand.tokens[zuo[vor]].end if vor is not None else cand.start
        e = cand.tokens[zuo[nach]].start if nach is not None else cand.end
        if e < s:
            return None
        von = vor + 1 if vor is not None else 0
        bis = nach if nach is not None else len(b.tokens)
        idx = [i for i in range(von, bis) if b.tokens[i].old]
        excel = [(m.group(), s + m.start(), s + m.end()) for m in re.finditer(r"\S+", zt[s:e])]
        if [w for w, _, _ in excel] != [b.tokens[i].old for i in idx]:
            return None
        ks = [k for k, i in enumerate(idx) if i0 <= i < i1]
        if not ks:
            return None
        return cand.blatt, cand.zeile, cand.spalte, excel[ks[0]][1], excel[ks[-1]][2]

    @staticmethod
    def _luecke(text: str, s: int, e: int) -> tuple[int, int]:
        """Beim Loeschen eines Wortbereichs ein angrenzendes Leerzeichen mitnehmen."""
        if s > 0 and text[s - 1] == " ":
            while s > 0 and text[s - 1] == " ":
                s -= 1
        else:
            while e < len(text) and text[e] == " ":
                e += 1
        return s, e

    def _kopf_mehrzellig(self, ae, b, i0, i1, cand, xt, neu, seq, stufe) -> bool:
        """Ueberschrift, bei der Nummer UND Titel (zwei Zellen) geaendert werden."""
        if cand.art != "kopf":
            return False
        zellen = []
        for t in xt:
            if (t.blatt, t.zeile, t.spalte) not in zellen:
                zellen.append((t.blatt, t.zeile, t.spalte))
        if len(zellen) != 2:
            return False
        teile = neu.split()
        nr = teile[0] if teile and section_nummer(teile[0]) else ""
        titel = " ".join(teile[1:] if nr else teile)
        ops = []
        for (blatt, zeile, spalte), text in zip(zellen, (nr, titel)):
            toks = [t for t in xt if (t.blatt, t.zeile, t.spalte) == (blatt, zeile, spalte)]
            ops.append(Operation("ersetzen", blatt, zeile, spalte, toks[0].start, toks[-1].end, text, seq, aenderung=ae.id))
        ae.ops = ops
        ae.ziel = self._ziel(xt[0])
        self._setze(ae, MITTEL if stufe != HOCH else HOCH)
        return True

    # ------------------------------------------------------------ Geloeschte Bloecke
    def _geloescht(self, b: PdfBlock) -> None:
        ist_kopf = b.kind == "ueberschrift"
        toks = b.tokens
        ae = self._basis(b, "section_geloescht" if ist_kopf else "absatz_geloescht", toks,
                         b.alt_text, "")
        tr = self.treffer(b)
        schw = self.cfg["schwellen"]
        if tr is None or tr.score < schw["mittel"]:
            self._kein_treffer(ae, b)
            return
        cand = tr.kandidat
        ae.score, ae.ziel, ae.excel_text = tr.score, cand.ziel, cand.text[:300]
        stufe = tr.stufe
        n_alt = sum(1 for t in toks if t.old and norm(t.old))
        js = sorted(set(tr.zuordnung.values()))
        if len(js) / max(n_alt, 1) < schw["hoch"]:
            stufe = MITTEL
            ae.grund = "Der Text in der Excel stimmt nur teilweise mit dem gestrichenen Text überein."
        seq = b.idx * 1000
        alle = self._alle_gestrichen(cand, js, toks)
        if cand.art == "kopf" and alle:
            zellen = []
            for t in cand.tokens:
                z = (t.blatt, t.zeile, t.spalte)
                if z not in zellen:
                    zellen.append(z)
            ae.ops = [Operation("zelle_leeren", bl, ze, sp, seq=seq, aenderung=ae.id) for bl, ze, sp in zellen]
        else:
            ae.ops = self._loesch_ops(ae, cand, js, seq, alle)
            # Nicht alles gestrichen: nur wenn genau ein Stueck aus ganzen Zeilen entfernt wird, ist es sicher
            if not alle and not (len(ae.ops) == 1 and ae.ops[0].absatz_weg):
                stufe = MITTEL
                self._hinweis(ae, "In der Excel steht in derselben Zelle mehr Text, der im PDF nicht gestrichen ist. "
                                  "Nur die gestrichenen Wörter werden entfernt, der Rest bleibt stehen.")
                ae.grund = (ae.grund + "; " if ae.grund else "") + "Nur ein Teil des Excel-Textes ist im PDF gestrichen."
        self._setze(ae, stufe)

    def _alle_gestrichen(self, cand: Kandidat, js: list[int], toks: list[PdfToken]) -> bool:
        """Sind ALLE Woerter der Excel-Stelle im PDF gestrichen? (nur dann darf eine Zelle/ein Absatz ganz entfernt werden)"""
        if len(set(js)) != len(cand.tokens):
            return False
        # Woerter ohne Normform ("•") stehen nicht in cand.tokens, muessen aber auch mitgestrichen sein
        if cand.art == "kopf":
            texte = [self.xl.text(*z) for z in dict.fromkeys((t.blatt, t.zeile, t.spalte) for t in cand.tokens)]
        else:
            texte = [cand.zelle_text[cand.start:cand.end]]
        leer = Counter(w for t in texte for w in t.split() if not norm(w))
        gestrichen = Counter(t.old for t in toks if t.old and not norm(t.old))
        return not (leer - gestrichen)

    def _loesch_ops(self, ae: Aenderung, cand: Kandidat, js: list[int], seq: int, alle: bool) -> list[Operation]:
        zt = cand.zelle_text
        if alle and cand.art in ("absatz", "zelle"):
            rest = (zt[:cand.start] + zt[cand.end:]).strip()
            if not rest:
                return [Operation("zelle_leeren", cand.blatt, cand.zeile, cand.spalte, seq=seq, aenderung=ae.id)]
            return [Operation("ersetzen", cand.blatt, cand.zeile, cand.spalte, cand.start, cand.end, "", seq,
                              aenderung=ae.id, absatz_weg=True)]
        # gezielt nur die gestrichenen Woerter entfernen: je zusammenhaengendem Stueck (gleiche Zelle, nur Leerraum dazwischen)
        stuecke: list[list[int]] = []
        for j in sorted(set(js)):
            if stuecke:
                p, t = cand.tokens[stuecke[-1][-1]], cand.tokens[j]
                text = self.xl.text(t.blatt, t.zeile, t.spalte)
                if (p.blatt, p.zeile, p.spalte) == (t.blatt, t.zeile, t.spalte) and j == stuecke[-1][-1] + 1 \
                        and not text[p.end:t.start].strip():
                    stuecke[-1].append(j)
                    continue
            stuecke.append([j])
        ops: list[Operation] = []
        for k, st in enumerate(stuecke):
            t0, t1 = cand.tokens[st[0]], cand.tokens[st[-1]]
            text = self.xl.text(t0.blatt, t0.zeile, t0.spalte)
            s0, e0 = t0.start, t1.end
            davor, danach = text[:s0].rstrip(" "), text[e0:].lstrip(" ")
            ganze_zeilen = (davor == "" or davor.endswith("\n")) and (danach == "" or danach.startswith("\n"))
            if ganze_zeilen:
                ops.append(Operation("ersetzen", t0.blatt, t0.zeile, t0.spalte, s0, e0, "", seq + k, aenderung=ae.id, absatz_weg=True))
            else:
                s, e = self._luecke(text, s0, e0)
                ops.append(Operation("ersetzen", t0.blatt, t0.zeile, t0.spalte, s, e, "", seq + k, aenderung=ae.id))
        return ops

    # ------------------------------------------------------------ Neue Bloecke
    def _ort_aus_fund(self, i: int, fund: Fundstelle, z: Zuordnung, direkt: bool, davor: bool) -> _Ort:
        """Ankerstelle aus einer exakten Fundstelle. "hoch" nur, wenn der Anker direkt davor steht, sicher auf der Kette liegt,
        sein Ende das Ende einer Excel-Zeile ist und dazwischen in den verwalteten Spalten nichts Unzugeordnetes liegt."""
        stufe = HOCH
        gruende: list[str] = []
        codes: list[str] = []
        if not z.automatisch:
            stufe = MITTEL
            codes.append("anker_unsicher")
            gruende.append(_GRUND.get(z.status, _GRUND["anker_unsicher"]))
        if not direkt:
            stufe = MITTEL
            codes.append("anker_unsicher")
            gruende.append("der Absatz davor konnte in der Excel nicht zugeordnet werden, die Einfügestelle ist nur ungefähr bestimmbar")
        if davor and not fund.zeilen_ende:
            stufe = MITTEL
            codes.append("anker_teil")
            gruende.append("der PDF-Text deckt nur den Anfang oder ein Stück des längeren Excel-Absatzes ab (Seitenumbruch?); "
                           "neuer Absatz könnte mitten im Satz landen")
        if not davor and not fund.zeilen_anfang:
            stufe = MITTEL
            codes.append("anker_teil")
            gruende.append(_GRUND["anker_teil"])
        if davor:
            hi = self.zuo[i].fenster[1]
            if hi is not None and hi[0] != fund.end_schluessel[0]:
                hi = None
            if self._unzugeordnet_dazwischen(fund.blatt, fund.end_schluessel, hi):
                stufe = MITTEL
                codes.append("anker_luecke")
                gruende.append("zwischen der Ankerstelle und dem nächsten zugeordneten Absatz steht in der Excel Text, der keinem "
                               "PDF-Absatz zugeordnet ist")
        if fund.versteckt:
            stufe = MITTEL
            codes.append("versteckt")
            gruende.append(_GRUND["versteckt"][:-1].lower())
        lay = self.xl.layouts.get(fund.blatt)
        if lay is not None and lay.sicherheit != "hoch":
            stufe = MITTEL
            codes.append("layout_unsicher")
            gruende.append(f"der Aufbau des Blatts '{fund.blatt}' wurde nur unsicher erkannt")
        z0, s0, a0, _ = fund.zellen[0]
        zl, sl, _, el = fund.zellen[-1]
        return _Ort(fund.blatt, z0, s0, fund.art, a0, el, fund.text, 1.0, stufe, davor, gruende, list(dict.fromkeys(codes)))

    def _ort_aus_treffer(self, tr: Treffer, davor: bool) -> _Ort:
        k = tr.kandidat
        if k.art != "kopf" and tr.zuordnung:
            ende = k.tokens[max(tr.zuordnung.values())].end
        else:
            ende = k.end
        gruende = list(tr.gruende) + [g for g in tr.anker_gruende if g not in tr.gruende]
        gruende.append("der Absatz davor steht so nicht in der Excel, die Einfügestelle ist nur ungefähr bestimmbar")
        return _Ort(k.blatt, k.zeile, k.spalte, k.art, k.start, ende, k.text, tr.score, MITTEL, davor, gruende,
                    ["anker_unsicher"])

    def _anker(self, i: int) -> _Ort | None:
        """Naechster Kettenanker vor dem neuen Block (sonst danach). Ohne Kette: alter Matcher (hoechstens "mittel")."""
        schw = self.cfg["schwellen"]
        uebersprungen = 0
        for k in range(i - 1, -1, -1):
            sp, z = self.spuren[k], self.zuo[k]
            if z.fund is not None and (z.automatisch or not sp.ganz_neu):
                return self._ort_aus_fund(i, z.fund, z, uebersprungen == 0, True)
            if sp.ganz_neu:
                continue
            self._fenster = z.fenster
            tr = self.treffer(sp.bloecke[-1])
            if tr is not None and tr.score >= schw["mittel"]:
                return self._ort_aus_treffer(tr, True)
            uebersprungen += 1
        for k in range(i + 1, len(self.spuren)):
            sp, z = self.spuren[k], self.zuo[k]
            if z.fund is not None and (z.automatisch or not sp.ganz_neu):
                o = self._ort_aus_fund(i, z.fund, z, False, False)
                return o
            if sp.ganz_neu:
                continue
            self._fenster = z.fenster
            tr = self.treffer(sp.bloecke[0])
            if tr is not None and tr.score >= schw["mittel"]:
                return self._ort_aus_treffer(tr, False)
        return None

    def _basis_neu(self, sp: Spur, typ: str) -> Aenderung:
        toks = sp.tokens
        seite = toks[0].page
        ae = Aenderung(id=f"A{len(self.out) + 1:04d}", typ=typ, seite=seite, bbox=_bbox([t for t in toks if t.page == seite]),
                       section=sp.erster.section_neu or sp.erster.section_alt, alt="", neu=sp.neu,
                       woerter=sum(1 for t in toks if t.status != KEEP), block_idx=sp.erster.idx, kontext=sp.neu[:400])
        self.out.append(ae)
        self._spur_von[ae.id] = self._aktuelle_spur if self._aktuelle_spur is not None else -1
        self._sperre[ae.id] = list(self._spur_sperre)
        self._hinweise_nur(ae)
        ae.soll_absatz = sp.neu
        return ae

    def _neu(self, i: int) -> None:
        sp = self.spuren[i]
        ist_kopf = sp.kind == "ueberschrift"
        koerper: list[int] = []
        if ist_kopf:
            k = i + 1
            while k < len(self.spuren) and self.spuren[k].ganz_neu and self.spuren[k].kind == "absatz":
                koerper.append(k)
                k += 1
        # die Folgeabsaetze stehen in derselben Excel-Zeile: ihr Zweifel gilt auch fuer die Section (und umgekehrt)
        self._spur_start(i, [self.spuren[k] for k in koerper])
        ae = self._basis_neu(sp, "section_neu" if ist_kopf else "absatz_neu")
        if koerper:
            ae.gruppe = ae.id
        z = self.zuo[i]
        if z.fund is not None and z.art == "neu" and z.automatisch:          # steht schon an genau dieser Stelle (Idempotenz)
            schon = z.fund
            if ist_kopf:
                grund = f"Die Section steht schon in der Excel ({schon.ziel}); die Änderung ist offenbar schon eingetragen."
                self._nicht(ae, grund, schon.ziel)
                self._code(ae, "schon_eingetragen")
                for k in koerper:
                    self.verbraucht.add(k)
                    kae = self._basis_neu(self.spuren[k], "absatz_neu")
                    kae.gruppe = ae.id
                    self._nicht(kae, "Die zugehörige neue Section steht schon in der Excel.", schon.ziel)
                    self._code(kae, "schon_eingetragen")
                return
            if len([w for w in sp.neu.split() if norm(w)]) >= _planer_cfg(self.cfg).get("doppelt_min_woerter", 3):
                self._nicht(ae, f"Der Absatz steht schon in der Excel ({schon.ziel}); die Änderung ist offenbar schon eingetragen.",
                            schon.ziel)
                self._code(ae, "schon_eingetragen")
                ae.excel_text = schon.text[:300]
                return
            self._hinweis(ae, f"Ein sehr kurzer Absatz mit gleichem Text steht schon in der Excel ({schon.ziel}). "
                              "Bitte prüfen, ob er doppelt wäre.")
            self._code(ae, "vielleicht_doppelt")
            doppelt = True
        else:
            doppelt = False
        ort = self._anker(i)
        if ort is None:
            self._nicht(ae, "Keine Ankerstelle in der Excel gefunden (kein benachbarter Text zuordenbar).")
            self._code(ae, "kein_anker")
            for k in koerper:
                self.verbraucht.add(k)
                kae = self._basis_neu(self.spuren[k], "absatz_neu")
                kae.gruppe = ae.id
                self._nicht(kae, "Gehört zu einer neuen Section ohne Ankerstelle.")
                self._code(kae, "kein_anker")
            return
        stufe = ort.stufe if not doppelt else MITTEL
        self._ag = list(dict.fromkeys(g for g in ort.gruende if g)) if stufe != HOCH else []
        self._ag_codes = list(ort.codes) if stufe != HOCH else []
        if stufe != HOCH:
            self._anker_grund(ae, self._ag)
        ae.score = ort.score
        ae.excel_text = ort.text[:200]
        lay = self.xl.layouts[ort.blatt]
        if lay.inhalt_spalte is None and not (ist_kopf and lay.section_spalte):
            self._nicht(ae, "Die Inhaltsspalte der Excel konnte nicht erkannt werden.", ort.ziel)
            self._code(ae, "keine_inhaltsspalte")
            return
        z_cfg = self.cfg["einfuegen"]
        seq = sp.erster.idx * 1000
        self._nachbar_spalte = self._naechste_textspalte(i, ort.blatt) if ort.art == "kopf" else None
        if ist_kopf:
            self._neue_section(ae, sp, [self.spuren[k] for k in koerper], koerper, ort, lay, stufe, ort.davor, seq)
        else:
            modus = z_cfg["neuer_absatz"]
            if modus == "auto":
                modus = "zelle" if lay.modell == "zeile_pro_section" else "zeile"
            if modus == "zelle":
                self._absatz_in_zelle(ae, sp, ort, lay, stufe, seq)
            else:
                self._absatz_als_zeile(ae, sp, ort, lay, stufe, ort.davor, seq)

    _ag_codes: list[str] = []

    def _anker_grund(self, ae: Aenderung, gruende: list[str]) -> None:
        """Gruende, warum die Ankerstelle fuer einen neuen Absatz unsicher ist: in `grund` und als Hinweis."""
        for g in gruende:
            self._hinweis(ae, g[:1].upper() + g[1:] + ".")
        if gruende:
            ae.grund = (ae.grund + "; " if ae.grund else "") + "Die Stelle für den neuen Absatz ist unsicher: " + "; ".join(gruende)

    def _absatz_stil(self, cand: Any, lay) -> int:
        """Zeile, deren Format eine neue Absatz-Zeile bekommt: eine Absatzzeile, nie die Ueberschriftszeile."""
        kopf = set(lay.kopf_zeilen)
        if cand.zeile not in kopf or not lay.inhalt_spalte:
            return cand.zeile
        for r in self.xl.section_zeilen(cand.blatt, cand.zeile):
            if r not in kopf and self.xl.text(cand.blatt, r, lay.inhalt_spalte).strip():
                return r
        for r in range(cand.zeile + 1, lay.max_zeile + 1):
            if r not in kopf and self.xl.text(cand.blatt, r, lay.inhalt_spalte).strip():
                return r
        return cand.zeile

    def _neue_section(self, ae, sp: Spur, koerper: list[Spur], koerper_nr: list[int], cand: _Ort, lay, stufe, davor, seq) -> None:
        text = sp.neu
        nr_titel = kopf_zerlegen(text)
        nr, titel = nr_titel if nr_titel else ("", text)
        zellen: dict[int, str] = {}
        if lay.section_spalte is None:
            zellen[lay.inhalt_spalte] = text            # type: ignore[index]
        elif lay.kombiniert:
            zellen[lay.section_spalte] = text
        elif lay.titel_spalte:
            zellen[lay.section_spalte] = nr
            zellen[lay.titel_spalte] = titel
        else:
            zellen[lay.section_spalte] = text
        modus_zeile = lay.modell == "zeile_pro_section" and lay.inhalt_spalte is not None
        if modus_zeile and koerper:
            zellen[lay.inhalt_spalte] = "\n".join(kb.neu for kb in koerper)       # type: ignore[index]
        kopf_zeilen = [r for r in lay.kopf_zeilen if r <= cand.zeile]
        stil = kopf_zeilen[-1] if kopf_zeilen else cand.zeile
        op = Operation("zeile_einfuegen", cand.blatt, cand.zeile, 0, text="", seq=seq, zellen=zellen,
                       stil_von=stil, vor=not davor, aenderung=ae.id)
        ae.ops = [op]
        ae.ziel = f"{cand.blatt}!Zeile {cand.zeile + (0 if not davor else 1)} (neu)"
        self._setze(ae, stufe, codes=self._neu_codes(stufe))
        folge = seq
        for k, kb in zip(koerper_nr, koerper):
            self.verbraucht.add(k)
            kae = self._basis_neu(kb, "absatz_neu")
            kae.gruppe = ae.id
            kae.score = ae.score
            kae.excel_text = ""
            if ae.status == NICHT:            # die Section selbst ging nicht (z. B. Steuerzeichen) -> Folgeabsatz auch nicht
                self._nicht(kae, "Gehört zu einer neuen Section, die nicht eingetragen werden kann.")
            elif modus_zeile:
                for h in ae.hinweise:                 # Risiken und Ankerzweifel der Section gelten auch fuer den Absatz in ihrer Zeile
                    self._hinweis(kae, h)
                # Der Folgeabsatz steht in der Zelle der neuen Section-Zeile. Eine "abgedeckte" Markierungs-Operation
                # (ohne Zellen) bindet ihn an die Section: Status und Freigabe gelten fuer beide gemeinsam (applier).
                kae.ops = [Operation("zeile_einfuegen", cand.blatt, cand.zeile, lay.inhalt_spalte or 0, seq=seq,
                                     abgedeckt=True, vor=not davor, aenderung=kae.id)]
                kae.ziel = ae.ziel
                self._setze(kae, stufe, "als Teil der neuen Section eingetragen", codes=self._neu_codes(stufe))
                kae.konfidenz, kae.status = ae.konfidenz, ae.status
                for c in ae.pruefcodes:
                    self._code(kae, c)
                if ae.status == PRUEFEN and ae.grund:
                    kae.grund = (kae.grund + "; " if kae.grund else "") + ae.grund
                self._hinweis(kae, "Dieser Absatz gehört zur neuen Section und wird nur zusammen mit ihr eingetragen.")
            else:
                folge += 1
                kop = Operation("zeile_einfuegen", cand.blatt, cand.zeile, 0, text="", seq=folge,
                                zellen={lay.inhalt_spalte: kb.neu}, stil_von=self._absatz_stil(cand, lay),   # type: ignore[dict-item]
                                vor=not davor, aenderung=kae.id)
                kae.ops = [kop]
                kae.ziel = ae.ziel
                self._anker_grund(kae, self._ag)
                self._setze(kae, stufe, codes=self._neu_codes(stufe))

    def _neu_codes(self, stufe: str) -> list[str]:
        return list(self._ag_codes) if stufe != HOCH else []

    _nachbar_spalte: int | None = None

    def _text_spalte(self, cand: _Ort, lay) -> int | None:
        """Spalte fuer einen neuen Absatz: wie ein Mensch neben den Nachbarabsatz (Ankerstelle im Fliesstext; ist der Anker eine
        Ueberschrift, die Spalte des naechsten zugeordneten Absatzes), sonst die erkannte Inhaltsspalte."""
        if cand.art in ("absatz", "zeilen", "zelle", "teil") and cand.spalte not in (lay.section_spalte, lay.titel_spalte):
            return cand.spalte
        if self._nachbar_spalte is not None:
            return self._nachbar_spalte
        return lay.inhalt_spalte

    def _naechste_textspalte(self, i: int, blatt: str) -> int | None:
        for k in range(i + 1, len(self.spuren)):
            z = self.zuo[k]
            if self.spuren[k].kind == "ueberschrift":
                return None
            if z.automatisch and z.fund is not None and z.fund.blatt == blatt and not z.fund.kopf_spalte:
                return z.fund.spalte
        return None

    def _absatz_in_zelle(self, ae, sp: Spur, cand: _Ort, lay, stufe, seq) -> None:
        col = self._text_spalte(cand, lay)
        zeile = cand.zeile
        ist_inhalt = cand.spalte == col and cand.art in ("absatz", "zeilen", "zelle", "teil")
        zt = self.xl.text(cand.blatt, zeile, col)            # type: ignore[arg-type]
        if ist_inhalt and cand.davor:
            pos, text = cand.end, "\n" + sp.neu
            for a0, e0 in self._geloescht_bis.get((cand.blatt, zeile, col), []):    # type: ignore[arg-type]
                if a0 < pos < e0:            # der Anker (erste Zeile) wird samt folgendem Umbruch geloescht: dahinter einsetzen
                    pos, text = e0, sp.neu + "\n"
        elif ist_inhalt:
            pos, text = cand.start, sp.neu + "\n"
        elif cand.art == "kopf" or cand.spalte in (lay.section_spalte, lay.titel_spalte):
            pos, text = 0, sp.neu + ("\n" if zt.strip() else "")
        else:
            pos, text = len(zt), ("\n" if zt.strip() else "") + sp.neu
        # "ersetzen" mit leerer Spanne (wie die Text-Aenderungen): bei gleicher Stelle entscheidet `seq` die Reihenfolge, so steht
        # ein neuer Absatz HINTER einem Satz, der im PDF davor ans Absatzende angehaengt wird.
        ae.ops = [Operation("ersetzen", cand.blatt, zeile, col, pos, pos, text, seq, aenderung=ae.id)]    # type: ignore[arg-type]
        ae.ziel = f"{cand.blatt}!{get_column_letter(col)}{zeile}"                                         # type: ignore[arg-type]
        self._setze(ae, stufe, codes=self._neu_codes(stufe))

    def _absatz_als_zeile(self, ae, sp: Spur, cand: _Ort, lay, stufe, davor, seq) -> None:
        if lay.inhalt_spalte is None:
            self._nicht(ae, "Die Inhaltsspalte der Excel konnte nicht erkannt werden.", cand.ziel)
            return
        op = Operation("zeile_einfuegen", cand.blatt, cand.zeile, 0, text="", seq=seq,
                       zellen={self._text_spalte(cand, lay): sp.neu}, stil_von=self._absatz_stil(cand, lay), vor=not davor,
                       aenderung=ae.id)
        ae.ops = [op]
        ae.ziel = f"{cand.blatt}!Zeile {cand.zeile + (1 if davor else 0)} (neu)"
        self._setze(ae, stufe, codes=self._neu_codes(stufe))


def plane(pdf: PdfAnalyse, xl: ExcelModell, cfg: dict | None = None) -> Plan:
    cfg = cfg or config.lade("zuordnung")
    p = _Planer(pdf, xl, cfg)
    p.lauf()
    for ae in p.out:                     # Absicherung: an der Stelle muss beim Eintragen derselbe Text stehen
        for op in ae.ops:
            if op.art == "ersetzen":
                op.erwartet = xl.text(op.blatt, op.zeile, op.spalte)[op.start:op.ende]
            elif op.art == "zelle_leeren":       # der Applier leert nur, wenn die Zelle noch genau so aussieht
                op.erwartet = xl.text(op.blatt, op.zeile, op.spalte)
    warn = list(pdf.warnungen) + list(xl.warnungen)
    plan = Plan(aenderungen=p.out, layout=xl.beschreibung(), warnungen=warn,
                ignoriert=pdf.ignoriert, markierte_woerter=pdf.markierte_woerter,
                luecken=list(pdf.luecken), absaetze=p.absaetze)
    return plan
