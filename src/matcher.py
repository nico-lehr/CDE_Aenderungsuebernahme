"""Findet fuer einen PDF-Block die passende Stelle in der Excel (rein regelbasiert).

Ablauf: Woerter normalisieren -> mit einem Wort-Index wenige Kandidaten vorwaehlen
-> jeden Kandidaten mit SequenceMatcher Wort fuer Wort ausrichten -> besten waehlen.

Bewertung:  score = gewicht_abdeckung * Abdeckung + gewicht_dice * Dice   (Gewichte: config matcher.*)
  Abdeckung = Anteil der PDF-Woerter, die im Kandidaten in gleicher Reihenfolge stehen
  Dice      = bevorzugt knappe Kandidaten (Absatz statt ganzer Zelle)

Sicherheit ("hoch" = automatisch, "mittel" = Mensch prueft). Hoechstens "mittel" gilt, wenn
  * es mehrere gleich gute Stellen gibt (auch zwei gleiche Absaetze in DERSELBEN Zelle),
  * der PDF-Text nur einen Teil (Anfang) eines laengeren Excel-Absatzes abdeckt,
  * die Section des PDF-Blocks (alte Nummer) und die Section der gefundenen Excel-Stelle beide bekannt sind und
    nicht zusammenpassen (Text steht zufaellig auch in einer anderen Section; Config matcher.section_unpassend_stufe),
  * der Aufbau des Blatts unsicher erkannt wurde, die Stelle ausgeblendet ist oder
    die Section-Nummer aus mehrdeutigen Zahlen stammt.
"""
from __future__ import annotations

import math
from collections import Counter
from dataclasses import dataclass, field
from difflib import SequenceMatcher

from .abgleich import _rang
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
    excel_abdeckung: float = 1.0         # Anteil der Woerter des Kandidaten, die der PDF-Text abdeckt
    anker_stufe: str = MITTEL            # Sicherheit, wenn an diese Stelle ein NEUER Absatz/eine neue Section gesetzt wird
    anker_gruende: list[str] = field(default_factory=list)   # Gruende, falls anker_stufe schlechter als stufe ist


def _konkurrent(k: Kandidat, kk: Kandidat) -> bool:
    """Ist kk eine echte Konkurrenz zu k? Andere Zelle: ja. Dieselbe Zelle: nur ein anderer Absatz
    (ein Absatz und die ganze Zelle sind keine Konkurrenz)."""
    if (kk.blatt, kk.zeile, kk.spalte) != (k.blatt, k.zeile, k.spalte):
        return True
    return k.art == "absatz" and kk.art == "absatz" and (k.start, k.end) != (kk.start, kk.end)


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
        self._nach_section: dict[str, list[list[int]]] = {}      # Section -> je Blatt eine Liste von Kandidaten
        for (_, sec), ids in xl.nach_blatt_section.items():
            self._nach_section.setdefault(sec, []).append(ids)

    def _idf(self, wort: str) -> float:
        return math.log(1 + self.n / max(len(self.index.get(wort, ())), 1))

    def schluessel(self, k: Kandidat) -> tuple:
        """Ordnung des Kandidaten in Dokument-Reihenfolge (wie abgleich.ordnungs_schluessel)."""
        if not hasattr(self, "_blatt_nr"):
            self._blatt_nr = {b: i for i, b in enumerate(self.xl.layouts)}
        return (self._blatt_nr.get(k.blatt, 999), k.zeile, _rang(self.xl.layouts.get(k.blatt), k.spalte), k.spalte, k.start)

    def finde(self, a: list[tuple[int, str]], section: str | None, ist_kopf: bool = False,
              fenster: tuple[tuple | None, tuple | None] | None = None) -> Treffer | None:
        """a = [(Block-Token-Index, Normform)], leere Normformen sind schon entfernt.

        fenster = (unten, oben): nur Kandidaten, deren Ordnungs-Schluessel echt dazwischen liegt (Reissverschluss: zwischen
        den Ankern davor und danach). None = ueberall."""
        if not a:
            return None
        if fenster is not None and fenster == (None, None):
            fenster = None
        key = (tuple(a), section, ist_kopf, fenster)
        if key in self._cache:
            return self._cache[key]
        ergebnis = self._finde(a, section, ist_kopf, fenster)
        self._cache[key] = ergebnis
        return ergebnis

    def _im_fenster(self, k: Kandidat, fenster: tuple[tuple | None, tuple | None] | None) -> bool:
        if fenster is None:
            return True
        s = self.schluessel(k)
        lo, hi = fenster
        return (lo is None or s > lo) and (hi is None or s < hi)

    def _finde(self, a: list[tuple[int, str]], section: str | None, ist_kopf: bool,
               fenster: tuple[tuple | None, tuple | None] | None = None) -> Treffer | None:
        schw = self.cfg["schwellen"]
        mc = self.cfg["matcher"]
        normen = [n for _, n in a]
        punkte: Counter = Counter()
        for n in set(normen):
            w = self._idf(n)
            for kid in self.index.get(n, ()):
                punkte[kid] += w
        ids = [kid for kid, _ in punkte.most_common(self.cfg["kandidaten_pro_block"])]
        if section:
            schon = set(ids)
            for blatt_ids in self._nach_section.get(section, []):      # je Blatt getrennt gezaehlt
                for i in blatt_ids[:mc["max_kandidaten_section"]]:
                    if i not in schon:
                        ids.append(i)
                        schon.add(i)
        if not ids:
            return None
        bewertet: list[tuple[float, float, Kandidat, dict[int, int], float, float]] = []
        for kid in ids:
            k = self.xl.kandidaten[kid]
            if k.art == "kopf" and not ist_kopf:
                continue
            if not self._im_fenster(k, fenster):
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
            score = mc["gewicht_abdeckung"] * cov + mc["gewicht_dice"] * dice
            bonus = mc["section_bonus"] if section and k.section == section and not k.section_unsicher else 0.0
            bewertet.append((score + bonus, score, k, zuordnung, cov, m / len(k.norms)))
        if not bewertet:
            return None
        bewertet.sort(key=lambda t: -t[0])
        rang, score, k, zuordnung, cov, ecov = bewertet[0]
        # zweitbester Treffer an einer ANDEREN Stelle (andere Zelle oder anderer Absatz derselben Zelle)
        zweiter, zweiter_k = 0.0, None
        for r, s, kk, _, _, _ in bewertet[1:]:
            if _konkurrent(k, kk):
                zweiter, zweiter_k = r, kk
                break
        eindeutig = (rang - zweiter) >= schw["eindeutig_abstand"]
        passt = None if not (section and k.section) or k.section_unsicher else (k.section == section)
        # Gleich gute Treffer in verschiedenen Sections DESSELBEN Blatts: die passende Section entscheidet
        if (not eindeutig and passt and zweiter_k is not None and zweiter_k.blatt == k.blatt
                and zweiter_k.section != section and not zweiter_k.section_unsicher and (score - zweiter) >= -1e-9):
            eindeutig = True
        t = Treffer(kandidat=k, score=score, abdeckung=cov, zuordnung=zuordnung, zweiter=zweiter,
                    eindeutig=eindeutig, section_passt=passt, excel_abdeckung=ecov)
        gruende = t.gruende
        if score < schw["mittel"]:
            return t if score >= mc["untergrenze"] else None
        stufe = HOCH if score >= schw["hoch"] else MITTEL
        if score >= schw["hoch"] and not eindeutig:
            stufe = MITTEL
            if zweiter_k is not None and (zweiter_k.blatt, zweiter_k.zeile, zweiter_k.spalte) == (k.blatt, k.zeile, k.spalte):
                gruende.append("derselbe oder ein sehr ähnlicher Absatz steht mehrmals in derselben Zelle")
            elif zweiter_k is not None and zweiter_k.blatt != k.blatt:
                gruende.append(f"ähnlich gute Stelle auch im Blatt '{zweiter_k.blatt}'")
            else:
                gruende.append("mehrere ähnlich gute Stellen in der Excel")
        if len(normen) <= schw["kurz_tokens"] and stufe == HOCH and not (passt or eindeutig):
            stufe = MITTEL
            gruende.append("sehr kurzer Text, Stelle nicht sicher")
        if len(normen) <= schw["kurz_tokens"] and stufe == HOCH and passt is False and not eindeutig:
            stufe = MITTEL
        if stufe == HOCH and passt is False:
            # PDF-Section (die ALTE Nummer, daher bleibt Umnummerierung unberuehrt) und Excel-Section sind beide bekannt
            # und verschieden: der Text steht zufaellig auch woanders. Nie automatisch, ein Mensch schaut hin.
            stufe = mc.get("section_unpassend_stufe", MITTEL)
            if stufe != HOCH:
                gruende.append("Section passt nicht zur gefundenen Stelle")
        if stufe == HOCH and k.section_unsicher:
            stufe = MITTEL
            gruende.append("die Section-Nummer steht als Zahl in der Excel und ist nicht eindeutig (2.1 oder 2.10?)")
        if stufe == HOCH:
            lay = self.xl.layouts.get(k.blatt)
            if lay is not None and lay.sicherheit != "hoch":
                stufe = MITTEL
                gruende.append(f"der Aufbau des Blatts '{k.blatt}' wurde nur unsicher erkannt")
        if stufe == HOCH and k.versteckt:
            stufe = MITTEL
            gruende.append({"blatt": f"das Blatt '{k.blatt}' ist ausgeblendet",
                            "zeile": "die Zeile ist ausgeblendet",
                            "spalte": "die Spalte ist ausgeblendet"}[k.versteckt])
        if not ist_kopf and k.kopf_spalte:
            # Fliesstext (oder eine Listennummer wie "4.") trifft die Nummern-/Titelspalte einer Ueberschrift: dort darf nie
            # automatisch geaendert werden (sonst steht der Text in der Nummernspalte); ein Mensch entscheidet.
            stufe = MITTEL
            gruende.append("die Zuordnung läuft nur über die Section-Nummer oder den Titel, die Stelle liegt in der "
                           "Nummern- oder Titelspalte und nicht im Fließtext")
        t.stufe = stufe
        # Anfangs-Treffer: der PDF-Text deckt nur einen Teil des Excel-Absatzes ab (Seitenumbruch). Ein Wort in diesem
        # Teil zu aendern ist sicher (stufe bleibt), aber einen NEUEN Absatz hier anzusetzen nicht (anker_stufe).
        t.anker_stufe = stufe
        offen = len(k.norms) - round(ecov * len(k.norms))
        if stufe == HOCH and k.art != "kopf" and ecov < mc["excel_abdeckung_min"] and offen >= mc["rest_min_woerter"]:
            t.anker_stufe = MITTEL
            t.anker_gruende.append("der PDF-Text deckt nur den Anfang oder ein Stück des längeren Excel-Absatzes ab "
                                   f"(Rest: {offen} Wörter, Seitenumbruch?); neuer Absatz könnte mitten im Satz landen")
        return t
