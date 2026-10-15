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
from .matcher import Matcher, Treffer
from .models import (ADD, DEL, HOCH, KEEP, MITTEL, MOD, NICHT, NIEDRIG, PRUEFEN, UEBERNOMMEN,
                     Aenderung, Operation, PdfBlock, PdfToken)
from .pdf_reader import PdfAnalyse
from .textutil import kopf_zerlegen, norm, section_nummer

STATUS_VON_STUFE = {HOCH: UEBERNOMMEN, MITTEL: PRUEFEN}


def _planer_cfg(cfg: dict | None) -> dict:
    """Abschnitt `planer:` der Konfiguration (Schwellen und Zeichenlisten, Regel 28)."""
    if cfg and "planer" in cfg:
        return cfg["planer"]
    return config.lade("zuordnung").get("planer", {})


def neutext_probleme(texte: Iterable[str], cfg: dict | None = None) -> dict[str, Any]:
    """Prueft Neutext auf Dinge, die in der Excel Schaden anrichten oder auf kaputtes PDF-Lesen hindeuten.

    Rueckgabe: steuerzeichen (nicht schreibbar), verdaechtig (Liste einfacher Hinweise -> nie automatisch),
    formel (Text beginnt mit '=' und wuerde von Excel als Formel gelesen)."""
    pl = _planer_cfg(cfg)
    zeichen = "".join(pl.get("verdaechtige_zeichen", []))
    muster = [re.compile(m) for m in pl.get("verdaechtige_muster", [])]
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


def _bbox(tokens: list[PdfToken]) -> tuple[float, float, float, float]:
    return (min(t.bbox[0] for t in tokens), min(t.bbox[1] for t in tokens),
            max(t.bbox[2] for t in tokens), max(t.bbox[3] for t in tokens))


class _Planer:
    def __init__(self, pdf: PdfAnalyse, xl: ExcelModell, cfg: dict) -> None:
        self.pdf, self.xl, self.cfg = pdf, xl, cfg
        self.matcher = Matcher(xl, cfg)
        self.out: list[Aenderung] = []
        self._tr: dict[int, Treffer | None] = {}
        self.verbraucht: set[int] = set()   # Bloecke, die schon als Teil einer neuen Section behandelt wurden
        self._nach_norms: dict[tuple[str, ...], list[Kandidat]] | None = None   # Normform -> Kandidaten (Idempotenz)
        self._ag: list[str] = []                    # Gruende, warum die Ankerstelle des aktuellen neuen Blocks unsicher ist
        self._risiko_cache: dict[tuple[str, int], list[str]] = {}               # (Blatt, Einfuegezeile) -> Risiken

    # ------------------------------------------------------------ Helfer
    def treffer(self, b: PdfBlock) -> Treffer | None:
        if b.idx not in self._tr:
            a = [(i, norm(t.old)) for i, t in enumerate(b.tokens) if t.old and norm(t.old)]
            self._tr[b.idx] = self.matcher.finde(a, b.section_alt, ist_kopf=(b.kind == "ueberschrift"))
        return self._tr[b.idx]

    def _basis(self, b: PdfBlock, typ: str, toks: list[PdfToken], alt: str, neu: str) -> Aenderung:
        ae = Aenderung(id=f"A{len(self.out) + 1:04d}", typ=typ, seite=toks[0].page, bbox=_bbox(toks),
                       section=b.section_neu or b.section_alt, alt=alt, neu=neu,
                       woerter=sum(1 for t in toks if t.status != KEEP), block_idx=b.idx,
                       kontext=(b.neu_text or b.alt_text)[:400])
        self.out.append(ae)
        return ae

    def _setze(self, ae: Aenderung, stufe: str, grund: str = "") -> None:
        # Neutext pruefen: Steuerzeichen koennen nicht geschrieben werden, Ligaturen/(cid:N) sind nie "sicher"
        texte = [ae.neu] + [o.text for o in ae.ops] + [t for o in ae.ops for t in o.zellen.values()]
        p = neutext_probleme(texte, self.cfg)
        if p["steuerzeichen"]:
            self._nicht(ae, "Der neue Text enthält Steuerzeichen und kann nicht in die Excel geschrieben werden.")
            return
        if p["verdaechtig"]:
            if stufe == HOCH:
                stufe = MITTEL
            for h in p["verdaechtig"]:
                self._hinweis(ae, h)
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
                    break
        risiken = self._zeilen_risiken(ae)
        if risiken:                      # Zeile einfuegen verschiebt Dinge, die openpyxl nicht mitzieht: nie automatisch
            if stufe == HOCH:
                stufe = MITTEL
            for r in risiken:
                self._hinweis(ae, r)
            grund = (grund + "; " if grund else "") + ("Beim Einfügen der neuen Zeile könnten Dinge in der Excel verrutschen "
                                                       "oder brechen (siehe Hinweise). Bitte vor dem Eintragen prüfen.")
        ae.konfidenz = stufe
        ae.status = STATUS_VON_STUFE.get(stufe, NICHT)
        if grund:
            ae.grund = (ae.grund + "; " if ae.grund else "") + grund

    def _zeilen_risiken(self, ae: Aenderung) -> list[str]:
        """Risiken (einfache Saetze) aller Zeilen-Einfuegungen einer Aenderung. Leer = unbedenklich."""
        out: list[str] = []
        for op in ae.ops:
            if op.art != "zeile_einfuegen" or op.abgedeckt:
                continue
            pos = op.zeile if op.vor else op.zeile + 1
            if (op.blatt, pos) not in self._risiko_cache:
                self._risiko_cache[(op.blatt, pos)] = self.xl.einfuege_risiken(op.blatt, pos)
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
        else:
            self._nicht(ae, f"Zu unsichere Stelle in der Excel (Übereinstimmung {tr.score:.0%}).", tr.kandidat.ziel)
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
        for b in self.pdf.blocks:
            if b.idx in self.verbraucht or not b.hat_aenderung:
                continue
            if b.ganz_neu:
                self._neu(b)
            elif b.ganz_geloescht:
                self._geloescht(b)
            else:
                for i0, i1 in self._gruppen(b):
                    self._text_gruppe(b, i0, i1)

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
    def _anker(self, b: PdfBlock) -> tuple[Treffer, bool, bool] | None:
        """(Treffer, direkt, davor): naechster zugeordneter Block vor b (sonst nach b)."""
        schw = self.cfg["schwellen"]
        uebersprungen = 0
        for k in range(b.idx - 1, -1, -1):
            blk = self.pdf.blocks[k]
            if not blk.alt_tokens:
                continue
            tr = self.treffer(blk)
            if tr is not None and tr.score >= schw["mittel"]:
                return tr, uebersprungen == 0, True
            uebersprungen += 1
        for k in range(b.idx + 1, len(self.pdf.blocks)):
            blk = self.pdf.blocks[k]
            if not blk.alt_tokens:
                continue
            tr = self.treffer(blk)
            if tr is not None and tr.score >= schw["mittel"]:
                return tr, False, False
        return None

    def _neu(self, b: PdfBlock) -> None:
        ist_kopf = b.kind == "ueberschrift"
        koerper: list[PdfBlock] = []
        if ist_kopf:
            k = b.idx + 1
            while k < len(self.pdf.blocks) and self.pdf.blocks[k].ganz_neu and self.pdf.blocks[k].kind == "absatz":
                koerper.append(self.pdf.blocks[k])
                k += 1
        ae = self._basis(b, "section_neu" if ist_kopf else "absatz_neu", b.tokens, "", b.neu_text)
        anker = self._anker(b)
        if anker is None:
            self._nicht(ae, "Keine Ankerstelle in der Excel gefunden (kein benachbarter Text zuordenbar).")
            for kb in koerper:
                self.verbraucht.add(kb.idx)
                self._nicht(self._basis(kb, "absatz_neu", kb.tokens, "", kb.neu_text), "Gehört zu einer neuen Section ohne Ankerstelle.")
            return
        tr, direkt, davor = anker
        cand = tr.kandidat
        # Fuer einen NEUEN Absatz zaehlt die Anker-Sicherheit (nicht die Sicherheit fuers Ersetzen): deckt der PDF-Text nur
        # den Anfang eines laengeren Excel-Absatzes ab, wuerde der neue Absatz mitten im Satz landen.
        stufe = tr.anker_stufe if direkt else MITTEL
        if stufe != HOCH:
            stufe = MITTEL
        self._ag = []
        if stufe != HOCH:
            if direkt:
                self._ag = list(tr.gruende) + [g for g in tr.anker_gruende if g not in tr.gruende]
            else:
                self._ag = ["der Absatz davor konnte in der Excel nicht zugeordnet werden, die Einfügestelle ist nur ungefähr bestimmbar"]
            self._anker_grund(ae, self._ag)
        ae.score = tr.score
        ae.excel_text = cand.text[:200]
        lay = self.xl.layouts[cand.blatt]
        if lay.inhalt_spalte is None and not (ist_kopf and lay.section_spalte):
            self._nicht(ae, "Die Inhaltsspalte der Excel konnte nicht erkannt werden.", cand.ziel)
            return
        z = self.cfg["einfuegen"]
        seq = b.idx * 1000
        schon = self._vorhanden(b, ist_kopf, cand)
        if schon is not None:
            if ist_kopf:
                grund = f"Die Section steht schon in der Excel ({schon.ziel}); die Änderung ist offenbar schon eingetragen."
                self._nicht(ae, grund, schon.ziel)
                for kb in koerper:
                    self.verbraucht.add(kb.idx)
                    self._nicht(self._basis(kb, "absatz_neu", kb.tokens, "", kb.neu_text),
                                "Die zugehörige neue Section steht schon in der Excel.", schon.ziel)
                return
            if len([w for w in b.neu_text.split() if norm(w)]) >= _planer_cfg(self.cfg).get("doppelt_min_woerter", 3):
                self._nicht(ae, f"Der Absatz steht schon in der Excel ({schon.ziel}); die Änderung ist offenbar schon eingetragen.",
                            schon.ziel)
                ae.excel_text = schon.text[:300]
                return
            self._hinweis(ae, f"Ein sehr kurzer Absatz mit gleichem Text steht schon in der Excel ({schon.ziel}). Bitte prüfen, ob er doppelt wäre.")
            stufe = MITTEL
        if ist_kopf:
            self._neue_section(ae, b, koerper, cand, lay, stufe, davor, seq)
        else:
            modus = z["neuer_absatz"]
            if modus == "auto":
                modus = "zelle" if lay.modell == "zeile_pro_section" else "zeile"
            if modus == "zelle":
                self._absatz_in_zelle(ae, b, cand, lay, stufe, seq, tr)
            else:
                self._absatz_als_zeile(ae, b, cand, lay, stufe, davor, seq)

    def _anker_grund(self, ae: Aenderung, gruende: list[str]) -> None:
        """Gruende, warum die Ankerstelle fuer einen neuen Absatz unsicher ist: in `grund` und als Hinweis."""
        for g in gruende:
            self._hinweis(ae, g[:1].upper() + g[1:] + ".")
        if gruende:
            ae.grund = (ae.grund + "; " if ae.grund else "") + "Die Stelle für den neuen Absatz ist unsicher: " + "; ".join(gruende)

    def _vorhanden(self, b: PdfBlock, ist_kopf: bool, cand: Kandidat) -> Kandidat | None:
        """Steht ein Absatz/eine Ueberschrift mit genau diesem neuen Text schon in der Excel? (Idempotenz)"""
        norms = tuple(n for n in (norm(w) for w in b.neu_text.split()) if n)
        if not norms:
            return None
        if self._nach_norms is None:
            self._nach_norms = {}
            for k in self.xl.kandidaten:
                if k.art in ("kopf", "absatz"):
                    self._nach_norms.setdefault(tuple(k.norms), []).append(k)
        sections = {x for x in (cand.section, b.section_neu, b.section_alt) if x}
        for k in self._nach_norms.get(norms, []):
            if k.blatt != cand.blatt:
                continue
            if ist_kopf or (k.art == "absatz" and (not sections or k.section in sections)):
                return k
        return None

    def _absatz_stil(self, cand: Kandidat, lay) -> int:
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

    def _neue_section(self, ae, b, koerper, cand, lay, stufe, davor, seq) -> None:
        text = b.neu_text
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
            zellen[lay.inhalt_spalte] = "\n".join(kb.neu_text for kb in koerper)       # type: ignore[index]
        kopf_zeilen = [r for r in lay.kopf_zeilen if r <= cand.zeile]
        stil = kopf_zeilen[-1] if kopf_zeilen else cand.zeile
        op = Operation("zeile_einfuegen", cand.blatt, cand.zeile, 0, text="", seq=seq, zellen=zellen,
                       stil_von=stil, vor=not davor, aenderung=ae.id)
        ae.ops = [op]
        ae.ziel = f"{cand.blatt}!Zeile {cand.zeile + (0 if not davor else 1)} (neu)"
        self._setze(ae, stufe)
        folge = seq
        for kb in koerper:
            self.verbraucht.add(kb.idx)
            kae = self._basis(kb, "absatz_neu", kb.tokens, "", kb.neu_text)
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
                self._setze(kae, stufe, "als Teil der neuen Section eingetragen")
                kae.konfidenz, kae.status = ae.konfidenz, ae.status
                if ae.status == PRUEFEN and ae.grund:
                    kae.grund = (kae.grund + "; " if kae.grund else "") + ae.grund
                self._hinweis(kae, "Dieser Absatz gehört zur neuen Section und wird nur zusammen mit ihr eingetragen.")
            else:
                folge += 1
                kop = Operation("zeile_einfuegen", cand.blatt, cand.zeile, 0, text="", seq=folge,
                                zellen={lay.inhalt_spalte: kb.neu_text}, stil_von=self._absatz_stil(cand, lay),   # type: ignore[dict-item]
                                vor=not davor, aenderung=kae.id)
                kae.ops = [kop]
                kae.ziel = ae.ziel
                self._anker_grund(kae, self._ag)
                self._setze(kae, stufe)

    def _absatz_in_zelle(self, ae, b, cand, lay, stufe, seq, tr) -> None:
        col = lay.inhalt_spalte
        zeile = cand.zeile
        ist_inhalt = cand.spalte == col and cand.art in ("absatz", "zelle")
        zt = self.xl.text(cand.blatt, zeile, col)            # type: ignore[arg-type]
        if ist_inhalt:
            letzte = max(tr.zuordnung.values()) if tr.zuordnung else len(cand.tokens) - 1
            pos, text = cand.tokens[letzte].end, "\n" + b.neu_text
        elif cand.art == "kopf" or cand.spalte in (lay.section_spalte, lay.titel_spalte):
            pos, text = 0, b.neu_text + ("\n" if zt.strip() else "")
        else:
            pos, text = len(zt), ("\n" if zt.strip() else "") + b.neu_text
        ae.ops = [Operation("einfuegen", cand.blatt, zeile, col, pos, pos, text, seq, aenderung=ae.id)]   # type: ignore[arg-type]
        ae.ziel = f"{cand.blatt}!{get_column_letter(col)}{zeile}"                                         # type: ignore[arg-type]
        self._setze(ae, stufe)

    def _absatz_als_zeile(self, ae, b, cand, lay, stufe, davor, seq) -> None:
        if lay.inhalt_spalte is None:
            self._nicht(ae, "Die Inhaltsspalte der Excel konnte nicht erkannt werden.", cand.ziel)
            return
        op = Operation("zeile_einfuegen", cand.blatt, cand.zeile, 0, text="", seq=seq,
                       zellen={lay.inhalt_spalte: b.neu_text}, stil_von=self._absatz_stil(cand, lay), vor=not davor,
                       aenderung=ae.id)
        ae.ops = [op]
        ae.ziel = f"{cand.blatt}!Zeile {cand.zeile + (1 if davor else 0)} (neu)"
        self._setze(ae, stufe)


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
                luecken=list(pdf.luecken))
    return plan
