"""Findet fuer einen PDF-Block die passende Stelle in der Excel (rein regelbasiert).

Ablauf: Woerter normalisieren -> mit einem Wort-Index wenige Kandidaten vorwaehlen
-> jeden Kandidaten mit SequenceMatcher Wort fuer Wort ausrichten -> besten waehlen.

Bewertung:  score = 0.9 * Abdeckung + 0.1 * Dice
  Abdeckung = Anteil der PDF-Woerter, die im Kandidaten in gleicher Reihenfolge stehen
  Dice      = bevorzugt knappe Kandidaten (Absatz statt ganzer Zelle)
"""
from __future__ import annotations

import math
from collections import Counter
from dataclasses import dataclass, field
from difflib import SequenceMatcher

from .excel_model import ExcelModell, Kandidat
from .models import HOCH, MITTEL


@dataclass
class Treffer:
    kandidat: Kandidat
    score: float
    abdeckung: float
    zuordnung: dict[int, int]            # Block-Token-Index -> Kandidaten-Token-Index
    zweiter: float = 0.0
    eindeutig: bool = True
    section_passt: bool | None = None
    stufe: str = MITTEL
    gruende: list[str] = field(default_factory=list)


class Matcher:
    def __init__(self, xl: ExcelModell, cfg: dict) -> None:
        self.xl = xl
        self.cfg = cfg
        self.index: dict[str, list[int]] = {}
        for k in xl.kandidaten:
            for n in set(k.norms):
                self.index.setdefault(n, []).append(k.id)
        self.n = max(len(xl.kandidaten), 1)
        self._cache: dict[tuple, Treffer | None] = {}

    def _idf(self, wort: str) -> float:
        return math.log(1 + self.n / max(len(self.index.get(wort, ())), 1))

    def finde(self, a: list[tuple[int, str]], section: str | None, ist_kopf: bool = False) -> Treffer | None:
        """a = [(Block-Token-Index, Normform)], leere Normformen sind schon entfernt."""
        if not a:
            return None
        key = (tuple(a), section, ist_kopf)
        if key in self._cache:
            return self._cache[key]
        ergebnis = self._finde(a, section, ist_kopf)
        self._cache[key] = ergebnis
        return ergebnis

    def _finde(self, a: list[tuple[int, str]], section: str | None, ist_kopf: bool) -> Treffer | None:
        schw = self.cfg["schwellen"]
        normen = [n for _, n in a]
        punkte: Counter = Counter()
        for n in set(normen):
            w = self._idf(n)
            for kid in self.index.get(n, ()):
                punkte[kid] += w
        ids = [kid for kid, _ in punkte.most_common(self.cfg["kandidaten_pro_block"])]
        if section:
            ids += [i for i in self.xl.nach_section.get(section, [])[:400] if i not in ids]
        if not ids:
            return None
        bewertet: list[tuple[float, float, Kandidat, dict[int, int], float]] = []
        for kid in ids:
            k = self.xl.kandidaten[kid]
            if k.art == "kopf" and not ist_kopf:
                continue
            sm = SequenceMatcher(None, normen, k.norms, autojunk=False)
            zuordnung: dict[int, int] = {}
            m = 0
            for i, j, size in sm.get_matching_blocks():
                for d in range(size):
                    zuordnung[a[i + d][0]] = j + d
                m += size
            if m == 0:
                continue
            cov = m / len(normen)
            dice = 2 * m / (len(normen) + len(k.norms))
            score = 0.9 * cov + 0.1 * dice
            rang = score + (0.03 if section and k.section == section else 0.0)
            bewertet.append((rang, score, k, zuordnung, cov))
        if not bewertet:
            return None
        bewertet.sort(key=lambda t: -t[0])
        rang, score, k, zuordnung, cov = bewertet[0]
        # zweitbester Treffer in einer ANDEREN Zelle (Absatz und ganze Zelle sind keine Konkurrenz)
        zweiter, zweiter_k = 0.0, None
        for r, s, kk, _, _ in bewertet[1:]:
            if (kk.blatt, kk.zeile, kk.spalte) != (k.blatt, k.zeile, k.spalte):
                zweiter, zweiter_k = r, kk
                break
        eindeutig = (rang - zweiter) >= schw["eindeutig_abstand"]
        passt = None if not (section and k.section) else (k.section == section)
        # Gleich gute Treffer in verschiedenen Sections: die passende Section entscheidet
        if not eindeutig and passt and zweiter_k is not None and zweiter_k.section != section and (score - zweiter) >= -1e-9:
            eindeutig = True
        t = Treffer(kandidat=k, score=score, abdeckung=cov, zuordnung=zuordnung, zweiter=zweiter,
                    eindeutig=eindeutig, section_passt=passt)
        gruende = t.gruende
        if score < schw["mittel"]:
            return t if score >= 0.4 else None
        stufe = HOCH if score >= schw["hoch"] else MITTEL
        if score >= schw["hoch"] and not eindeutig:
            stufe = MITTEL
            gruende.append("mehrere ähnlich gute Stellen in der Excel")
        if len(normen) <= schw["kurz_tokens"] and stufe == HOCH and not (passt or eindeutig):
            stufe = MITTEL
            gruende.append("sehr kurzer Text, Stelle nicht sicher")
        if len(normen) <= schw["kurz_tokens"] and stufe == HOCH and passt is False and not eindeutig:
            stufe = MITTEL
        t.stufe = stufe
        return t
