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

from dataclasses import dataclass, field
from typing import Iterator

from openpyxl.utils import get_column_letter

from . import config
from .excel_model import Kandidat, ExcelModell
from .matcher import Matcher, Treffer
from .models import (ADD, DEL, HOCH, KEEP, MITTEL, MOD, NICHT, NIEDRIG, PRUEFEN, UEBERNOMMEN,
                     Aenderung, Operation, PdfBlock, PdfToken)
from .pdf_reader import PdfAnalyse
from .textutil import kopf_zerlegen, norm, section_nummer

STATUS_VON_STUFE = {HOCH: UEBERNOMMEN, MITTEL: PRUEFEN}


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
        ae.konfidenz = stufe
        ae.status = STATUS_VON_STUFE.get(stufe, NICHT)
        if grund:
            ae.grund = (ae.grund + "; " if ae.grund else "") + grund

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
        d_idx = [i for i in range(i0, i1) if b.tokens[i].old and norm(b.tokens[i].old)]
        seq = b.idx * 1000 + i0

        if d_idx:
            fehlt = [i for i in d_idx if i not in tr.zuordnung]
            if fehlt:
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
            if neu:
                op = Operation("ersetzen", blatt, zeile, spalte, start, ende, neu, seq, aenderung=ae.id)
            else:
                s, e = self._luecke(zt, start, ende)
                op = Operation("ersetzen", blatt, zeile, spalte, s, e, "", seq, aenderung=ae.id)
            ae.ops = [op]
            ae.ziel = self._ziel(xt[0])
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
            ae.ops = [op]
            ae.ziel = self._ziel(xt)
        self._setze(ae, stufe, "; ".join(gruende))

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
        if cand.art == "kopf":
            zellen = []
            for t in cand.tokens:
                z = (t.blatt, t.zeile, t.spalte)
                if z not in zellen:
                    zellen.append(z)
            ae.ops = [Operation("zelle_leeren", bl, ze, sp, seq=seq, aenderung=ae.id) for bl, ze, sp in zellen]
        else:
            ae.ops = [self._loesch_op(ae, cand, js, seq)]
        self._setze(ae, stufe)

    def _loesch_op(self, ae: Aenderung, cand: Kandidat, js: list[int], seq: int) -> Operation:
        zt = cand.zelle_text
        voll = bool(js) and js[0] == 0 and js[-1] == len(cand.tokens) - 1
        if voll and cand.art in ("absatz", "zelle"):
            rest = (zt[:cand.start] + zt[cand.end:]).strip()
            if not rest:
                return Operation("zelle_leeren", cand.blatt, cand.zeile, cand.spalte, seq=seq, aenderung=ae.id)
            return Operation("ersetzen", cand.blatt, cand.zeile, cand.spalte, cand.start, cand.end, "", seq,
                             aenderung=ae.id, absatz_weg=True)
        s0, e0 = cand.tokens[js[0]].start, cand.tokens[js[-1]].end
        davor, danach = zt[:s0].rstrip(" "), zt[e0:].lstrip(" ")
        ganze_zeilen = (davor == "" or davor.endswith("\n")) and (danach == "" or danach.startswith("\n"))
        if ganze_zeilen:
            return Operation("ersetzen", cand.blatt, cand.zeile, cand.spalte, s0, e0, "", seq, aenderung=ae.id, absatz_weg=True)
        s, e = self._luecke(zt, s0, e0)
        return Operation("ersetzen", cand.blatt, cand.zeile, cand.spalte, s, e, "", seq, aenderung=ae.id)

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
        stufe = tr.stufe if direkt else MITTEL
        ae.score = tr.score
        ae.excel_text = cand.text[:200]
        lay = self.xl.layouts[cand.blatt]
        if lay.inhalt_spalte is None and not (ist_kopf and lay.section_spalte):
            self._nicht(ae, "Die Inhaltsspalte der Excel konnte nicht erkannt werden.", cand.ziel)
            return
        z = self.cfg["einfuegen"]
        seq = b.idx * 1000
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
            if modus_zeile:
                kae.ops = []
                kae.ziel = ae.ziel
                self._setze(kae, stufe, "als Teil der neuen Section eingetragen")
                kae.status = ae.status
            else:
                folge += 1
                kop = Operation("zeile_einfuegen", cand.blatt, cand.zeile, 0, text="", seq=folge,
                                zellen={lay.inhalt_spalte: kb.neu_text}, stil_von=cand.zeile,   # type: ignore[dict-item]
                                vor=not davor, aenderung=kae.id)
                kae.ops = [kop]
                kae.ziel = ae.ziel
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
                       zellen={lay.inhalt_spalte: b.neu_text}, stil_von=cand.zeile, vor=not davor, aenderung=ae.id)
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
    warn = list(pdf.warnungen) + list(xl.warnungen)
    plan = Plan(aenderungen=p.out, layout=xl.beschreibung(), warnungen=warn,
                ignoriert=pdf.ignoriert, markierte_woerter=pdf.markierte_woerter,
                luecken=list(pdf.luecken))
    return plan
