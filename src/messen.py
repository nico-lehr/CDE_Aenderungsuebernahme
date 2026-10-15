"""Messwerkzeug: python -m src.cli messen --korpus entwicklung|test [--paar NAME] [--ordner PFAD] [--mit-vorschlaegen]

Ein Paar = aenderungen.pdf + vorher.xlsx + nachher.xlsx (Format: docs/paar_format.md).
Das Tool laeuft auf PDF + vorher; das Ergebnis wird mit "nachher" verglichen.
Das Soll kommt NUR aus vorher <-> nachher (Zellvergleich, Zeilenausrichtung), nie aus dem PDF-Leser.

Regel 16: Konsole und Bericht enthalten NIE Text aus Dokumenten -- nur Paar-Namen, Blattnummer,
Zeile, Spalte, Aenderungs-ID, Zaehler und Fehlerart. Texte werden nur im Speicher verglichen.
Regel 29: Berichte sind append-only (NNN_<korpus>), es wird nie etwas ueberschrieben.

Aufbau
  1. Dateien lesen und Ablauf je Paar (ohne runs-Ordner, im Temp-Ordner)
  2. Soll/Ist: Zeilen und Zellen ausrichten (difflib), Einheiten bilden
  3. Kennzahlen, stille Fehler, Fehlerart-Kaskade
  4. Bericht (json + md) und Konsole
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import shutil
import sys
import tempfile
import time
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText
from openpyxl.utils import column_index_from_string, get_column_letter

from . import config
from .applier import uebernehmen
from .excel_model import ExcelModell
from .models import ADD, DEL, MOD, NICHT, PRUEFEN, SCHEMA_VERSION, UEBERNOMMEN
from .pdf_reader import analysiere_pdf
from .planner import plane
from .protokoll import bilanz

BERICHT_FORMAT = 1
FEHLERARTEN = ("pdf_falsch_gelesen", "layout_falsch", "zuordnung_falsch", "plan_falsch", "bilanz_kaputt", "unklar")
STILLE_ARTEN = ("uebernommen_aber_falsch", "kollateral", "bilanz", "nachkontrolle_ok_aber_falsch",
                "keine_wirkung", "luecke_nicht_gemeldet")
ZEILEN_AEHNLICH = 0.6       # ab hier gelten zwei Zeilen als "dieselbe Zeile, veraendert"
BLOCK_MAX = 4000            # groessere Austausch-Bloecke werden nur nach Position gepaart
DATEIEN = ("aenderungen.pdf", "vorher.xlsx", "nachher.xlsx")

_UNSICHTBAR = dict.fromkeys(map(ord, "­​‌‍﻿"), None)


# ======================================================================================
# 1. Dateien lesen
# ======================================================================================
def _k(text: str) -> str:
    """Vergleichsform einer Zelle: Unicode-normalisiert, ohne unsichtbare Zeichen, Leerraum zusammengefasst."""
    return " ".join(unicodedata.normalize("NFKC", text).translate(_UNSICHTBAR).split())


def _zelltext(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, CellRichText):
        return str(v)
    if isinstance(v, str):
        return v
    if isinstance(v, float):
        return format(v, ".15g")
    return str(v)


def _lese_excel(pfad: Path) -> list[list[list[str]]]:
    """Alle Blaetter als Zeilen mit Zelltexten (Formeln als Text)."""
    wb = load_workbook(str(pfad), rich_text=True)
    try:
        return [[[_zelltext(v) for v in row] for row in ws.iter_rows(values_only=True)] for ws in wb.worksheets]
    finally:
        wb.close()


def _blattnamen(pfad: Path) -> list[str]:
    wb = load_workbook(str(pfad), read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _sha(pfad: Path) -> str:
    h = hashlib.sha256()
    with open(pfad, "rb") as f:
        for stueck in iter(lambda: f.read(1 << 20), b""):
            h.update(stueck)
    return h.hexdigest()


@dataclass
class Zeile:
    nr: int                       # Zeilennummer im Blatt (1 = erste)
    zellen: tuple[str, ...]       # Rohtexte (hinten ohne leere Zellen)
    keys: tuple[str, ...]         # Vergleichsformen


def _zeilen(rows: list[list[str]]) -> list[Zeile]:
    """Nur nicht-leere Zeilen (leere Zeilen zaehlen nicht; die Nummern bleiben die des Originals)."""
    out: list[Zeile] = []
    for i, row in enumerate(rows, 1):
        keys = [_k(c) for c in row]
        while keys and not keys[-1]:
            keys.pop()
        if keys:
            out.append(Zeile(i, tuple(row[:len(keys)]), tuple(keys)))
    return out


# ======================================================================================
# 2. Ausrichten: Diff vorher <-> X (X = nachher oder Ergebnis)
# ======================================================================================
@dataclass
class Diff:
    zellen: dict[tuple[int, int], str] = field(default_factory=dict)   # (Zeile, Spalte) -> neuer Rohtext ("" = leer)
    weg: set[int] = field(default_factory=set)                          # vorher-Zeilen ohne Gegenstueck
    neue: list[tuple[int, tuple[str, ...], tuple[str, ...]]] = field(default_factory=list)   # (Anker, Rohtexte, Keys)
    # Anker = Nummer der letzten behaltenen vorher-Zeile VOR der neuen Zeile (0 = ganz oben).
    # "Behalten" heisst: nicht weg UND nicht ausgeleert (eine Zeile, in der nur noch unverwaltete Spalten wie
    # Bemerkung stehen, zaehlt wie eine geloeschte Zeile -- im Soll genauso wie im Ergebnis).
    zuordnung: dict[int, int] = field(default_factory=dict)             # vorher-Zeile -> Zeile in X (alle gepaarten)
    geleert: set[int] = field(default_factory=set)                      # gepaarte vorher-Zeilen, die in X ausgeleert sind


def _woerter_zaehler(text: str) -> Counter:
    return Counter(text.split())


def _enthalten(klein: str, gross: str) -> bool:
    """Steht der kuerzere Text (mindestens drei Woerter) im Wesentlichen im laengeren? (Teilstring oder Woerter-Teilmenge)"""
    if len(klein.split()) < 3 or len(klein) >= len(gross):
        return False
    if klein in gross:
        return True
    return not (_woerter_zaehler(klein) - _woerter_zaehler(gross))


def _zell_aehnlich(x: str, y: str) -> float:
    """Aehnlichkeit zweier Zelltexte (0..1). Ein gekuerzter oder verlaengerter Text (der eine steckt im anderen) zaehlt
    als dieselbe Zelle, veraendert -- so bleibt eine Zeile erkennbar, wenn ein Absatz hinzukommt oder wegfaellt."""
    if x == y:
        return 1.0
    if not x or not y:
        return 0.0
    kurz, lang = (x, y) if len(x) <= len(y) else (y, x)
    best = 0.8 if _enthalten(kurz, lang) else 0.0
    sm = SequenceMatcher(None, x, y, autojunk=False)
    if sm.real_quick_ratio() >= 0.5 and sm.quick_ratio() >= 0.5:
        r = sm.ratio()
        if r >= 0.5:
            best = max(best, r)
    return best


def _aehnlich(a: tuple[str, ...], b: tuple[str, ...]) -> float:
    """Aehnlichkeit zweier Zeilen (0..1): Anteil gleicher bzw. aehnlicher Zellen. Eine Zeile, die nur "ausgeleert" wurde
    (die belegten Zellen der einen Zeile stehen -- gleich oder gekuerzt -- in der anderen), zaehlt als dieselbe Zeile."""
    a_ne = {c: v for c, v in enumerate(a) if v}
    b_ne = {c: v for c, v in enumerate(b) if v}
    teil = 0.0
    for klein, gross in ((a_ne, b_ne), (b_ne, a_ne)):
        if klein and klein.keys() <= gross.keys():
            werte = [_zell_aehnlich(v, gross[c]) for c, v in klein.items()]
            if all(w >= 0.5 for w in werte):
                teil = max(teil, sum(werte) / len(werte))
    gewicht = summe = 0.0
    for c in range(max(len(a), len(b))):
        x = a[c] if c < len(a) else ""
        y = b[c] if c < len(b) else ""
        if not x and not y:
            continue
        gewicht += 1.0                               # jede belegte Zelle zaehlt gleich (lange Texte erdruecken Nummer/Titel nicht)
        summe += _zell_aehnlich(x, y)
    return max(teil, summe / gewicht if gewicht else 0.0)


def _paare(va: list[Zeile], xb: list[Zeile]) -> list[tuple[int, int]]:
    """Ordnungstreue Zuordnung zwischen zwei Zeilenbloecken (Austausch-Block), beste Gesamtaehnlichkeit."""
    n, m = len(va), len(xb)
    if not n or not m:
        return []
    if n * m > BLOCK_MAX:
        return [(i, i) for i in range(min(n, m)) if _aehnlich(va[i].keys, xb[i].keys) >= ZEILEN_AEHNLICH]
    s = [[_aehnlich(va[i].keys, xb[j].keys) for j in range(m)] for i in range(n)]
    best = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            b = max(best[i - 1][j], best[i][j - 1])
            if s[i - 1][j - 1] >= ZEILEN_AEHNLICH:
                b = max(b, best[i - 1][j - 1] + s[i - 1][j - 1])
            best[i][j] = b
    out: list[tuple[int, int]] = []
    i, j = n, m
    while i > 0 and j > 0:
        if s[i - 1][j - 1] >= ZEILEN_AEHNLICH and best[i][j] == best[i - 1][j - 1] + s[i - 1][j - 1]:
            out.append((i - 1, j - 1))
            i, j = i - 1, j - 1
        elif best[i][j] == best[i - 1][j]:
            i -= 1
        else:
            j -= 1
    return out[::-1]


def _diff(v: list[Zeile], x: list[Zeile], leer: Callable[[Zeile], bool] | None = None) -> Diff:
    """Zellen-, Zeilen-Unterschiede vorher -> X. `leer(zeile)` sagt, ob eine Zeile in den verwalteten Spalten nichts mehr
    enthaelt (dann ist sie als Anker so gut wie geloescht)."""
    d = Diff()
    sm = SequenceMatcher(None, [z.keys for z in v], [z.keys for z in x], autojunk=False)
    weg_idx: set[int] = set()
    leer_idx: set[int] = set()
    roh_neu: list[tuple[int, Zeile]] = []

    def weg(i: int) -> None:
        z = v[i]
        weg_idx.add(i)
        d.weg.add(z.nr)
        for c, k in enumerate(z.keys, 1):
            if k:
                d.zellen[(z.nr, c)] = ""

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                d.zuordnung[v[i1 + k].nr] = x[j1 + k].nr
            continue
        paare = _paare(v[i1:i2], x[j1:j2]) if tag == "replace" else []
        zu_x = {j1 + b: i1 + a for a, b in paare}
        zeiger = i1
        for j in range(j1, j2):
            if j in zu_x:
                i = zu_x[j]
                while zeiger < i:
                    weg(zeiger)
                    zeiger += 1
                a, b = v[i], x[j]
                d.zuordnung[a.nr] = b.nr
                if leer is not None and leer(b):
                    leer_idx.add(i)
                    d.geleert.add(a.nr)
                for c in range(max(len(a.keys), len(b.keys))):
                    ka = a.keys[c] if c < len(a.keys) else ""
                    kb = b.keys[c] if c < len(b.keys) else ""
                    if ka != kb:
                        d.zellen[(a.nr, c + 1)] = b.zellen[c] if c < len(b.zellen) else ""
                zeiger = i + 1
            else:
                roh_neu.append((zeiger - 1, x[j]))
        while zeiger < i2:
            weg(zeiger)
            zeiger += 1
    for vi, z in roh_neu:
        while vi >= 0 and (vi in weg_idx or vi in leer_idx):
            vi -= 1
        d.neue.append((v[vi].nr if vi >= 0 else 0, z.zellen, z.keys))
    return d


def _neu_abgleich(soll: list, ist: list) -> tuple[dict[int, int], dict[int, bool]]:
    """Ordnet neue Zeilen des Ist den neuen Zeilen des Soll zu.
    -> (soll_index -> ist_index, soll_index -> Ort stimmt)."""
    paar: dict[int, int] = {}
    ort: dict[int, bool] = {}
    frei = list(range(len(ist)))
    for si, (anker, _, keys) in enumerate(soll):                       # 1. gleicher Anker, gleicher Inhalt
        for ii in frei:
            if ist[ii][0] == anker and ist[ii][2] == keys:
                paar[si], ort[si] = ii, True
                frei.remove(ii)
                break
    for si, (anker, _, keys) in enumerate(soll):                       # 2. gleicher Inhalt, anderer Ort
        if si in paar:
            continue
        for ii in frei:
            if ist[ii][2] == keys:
                paar[si], ort[si] = ii, False
                frei.remove(ii)
                break
    for si, (anker, _, keys) in enumerate(soll):                       # 3. gleicher Anker, anderer Inhalt
        if si in paar:
            continue
        for ii in frei:
            if ist[ii][0] == anker:
                paar[si], ort[si] = ii, True
                frei.remove(ii)
                break
    return paar, ort


# ======================================================================================
# 3. Einheiten und Auswertung
# ======================================================================================
def _wort_liste(text: str) -> list[str]:
    from .textutil import norm
    return [w for w in (norm(x) for x in text.split()) if w]


def _wort_diff(alt: str, neu: str) -> tuple[list[str], list[str]]:
    """-> (entfallene Woerter, neue Woerter) als normalisierte Woerter."""
    a, b = _wort_liste(alt), _wort_liste(neu)
    entf: list[str] = []
    dazu: list[str] = []
    for tag, i1, i2, j1, j2 in SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag in ("replace", "delete"):
            entf += a[i1:i2]
        if tag in ("replace", "insert"):
            dazu += b[j1:j2]
    return entf, dazu


@dataclass(frozen=True)
class Ziel:
    bi: int
    zeile: int | None
    spalte: int | None


@dataclass
class Fehler:
    """Ein Eintrag im Bericht (nur Positionen, nie Text)."""
    bi: int
    zeile: int
    spalte: int
    aenderungen: list[str] = field(default_factory=list)
    art: str = ""
    stumm: bool = False
    stille_arten: set[str] = field(default_factory=set)
    quelle: str = "soll"      # soll = Soll-Zelle verfehlt | auto = Ergebnis falsch | paar = ganzes Paar


_ZIEL_ZELLE = re.compile(r"^(.+)!([A-Z]+)(\d+)$")
_ZIEL_NEU = re.compile(r"^(.+)!Zeile (\d+) \(neu\)$")


class _Auswertung:
    """Vergleicht Ergebnis (E_auto) mit nachher und sammelt alles, was der Bericht braucht."""

    def __init__(self, v_blaetter, n_blaetter, e_blaetter, e_alle_blaetter, namen: list[str], aenderungen, plan,
                 pdf_analyse, bil_auto, bil_plan, idempotent: bool, luecken_erwartet: list[dict] | None) -> None:
        self.namen = namen
        self.bi_von_name = {n: i for i, n in enumerate(namen)}
        self.V = [_zeilen(b) for b in v_blaetter]
        self.N = [_zeilen(b) for b in n_blaetter] + [[] for _ in range(max(0, len(self.V) - len(n_blaetter)))]
        self.E = [_zeilen(b) for b in e_blaetter] + [[] for _ in range(max(0, len(self.V) - len(e_blaetter)))]
        self.vtext = [{z.nr: z.zellen for z in blatt} for blatt in self.V]
        self.aend = aenderungen
        self.plan = plan
        self._inhaltsspalte: dict[int, int | None] = {}
        self.S = [_diff(self.V[i], self.N[i]) for i in range(len(self.V))]     # 1. Lauf: nur fuer die Inhaltsspalte
        self.S = [_diff(self.V[i], self.N[i], self._leer_fn(i)) for i in range(len(self.V))]
        self.I = [_diff(self.V[i], self.E[i], self._leer_fn(i)) for i in range(len(self.V))]
        self.bil_auto, self.bil_plan = bil_auto, bil_plan
        self.idempotent = idempotent
        self.fehler: dict[Any, Fehler] = {}
        self.hinweise: list[str] = []
        if len(n_blaetter) != len(v_blaetter):
            self.hinweise.append("blattanzahl_vorher_nachher_verschieden")
        if len(e_blaetter) != len(v_blaetter):
            self.hinweise.append("blattanzahl_ergebnis_verschieden")

        # PDF-erkannte Woerter (nur im Speicher)
        self.pdf_neu: Counter[str] = Counter()
        self.pdf_alt: Counter[str] = Counter()
        for b in pdf_analyse.blocks:
            for t in b.tokens:
                if t.status in (ADD, MOD) and t.new:
                    self.pdf_neu.update(_wort_liste(t.new))
                if t.status in (DEL, MOD) and t.old:
                    self.pdf_alt.update(_wort_liste(t.old))
        self.ae_woerter = {a.id: (Counter(_wort_liste(a.neu)), Counter(_wort_liste(a.alt))) for a in aenderungen}
        self.ziele_von = {a.id: self._ziele(a) for a in aenderungen}

        self._bauen()
        self.luecken_erwartet = luecken_erwartet
        self.e_alle = (self._ist_liste([_diff(self.V[i], _zeilen(e_alle_blaetter[i]) if i < len(e_alle_blaetter) else [])
                                        for i in range(len(self.V))])
                       if e_alle_blaetter is not None else None)

    # ------------------------------------------------------------------ Hilfen
    def _ziele(self, ae) -> list[Ziel]:
        out: list[Ziel] = []
        for op in ae.ops:
            bi = self.bi_von_name.get(op.blatt)
            if bi is None:
                continue
            if op.art in ("ersetzen", "einfuegen", "zelle_leeren"):
                out.append(Ziel(bi, op.zeile, op.spalte))
            elif op.art == "zeile_einfuegen":
                out.append(Ziel(bi, None, None))
            else:
                out.append(Ziel(bi, op.zeile, None))
        if not out and ae.ziel:
            m = _ZIEL_ZELLE.match(ae.ziel)
            if m and m.group(1) in self.bi_von_name:
                out.append(Ziel(self.bi_von_name[m.group(1)], int(m.group(3)), column_index_from_string(m.group(2))))
            else:
                m = _ZIEL_NEU.match(ae.ziel)
                if m and m.group(1) in self.bi_von_name:
                    out.append(Ziel(self.bi_von_name[m.group(1)], None, None))
        return out

    def _vorher(self, bi: int, zeile: int, spalte: int) -> str:
        z = self.vtext[bi].get(zeile, ())
        return z[spalte - 1] if spalte - 1 < len(z) else ""

    def _erwartet(self, bi: int, zeile: int, spalte: int) -> str:
        """Inhalt der Zelle laut nachher."""
        return self.S[bi].zellen.get((zeile, spalte), self._vorher(bi, zeile, spalte))

    def _tatsaechlich(self, bi: int, zeile: int, spalte: int) -> str:
        return self.I[bi].zellen.get((zeile, spalte), self._vorher(bi, zeile, spalte))

    # ------------------------------------------------------------------ Einheiten bauen
    def _bauen(self) -> None:
        # Soll-Einheiten: (Schluessel, bi, zeile, spalte, soll_key, art)
        self.einheiten: list[dict[str, Any]] = []
        self.soll_woerter: dict[Any, tuple[list[str], list[str]]] = {}
        self.neu_abgleich: list[tuple[dict[int, int], dict[int, bool]]] = []
        for bi, s in enumerate(self.S):
            paar, ort = _neu_abgleich(s.neue, self.I[bi].neue)
            self.neu_abgleich.append((paar, ort))
            for (z, c), inhalt in sorted(s.zellen.items()):
                key = ("z", bi, z, c)
                ist_da = (z, c) in self.I[bi].zellen
                ist = self.I[bi].zellen.get((z, c))
                erreicht = ist_da and _k(ist) == _k(inhalt)
                v = self._vorher(bi, z, c)
                self.einheiten.append({"key": key, "bi": bi, "zeile": z, "spalte": c, "soll": inhalt, "ist_da": ist_da,
                                       "ist": ist, "erreicht": erreicht, "art": "weg" if z in s.weg else "zelle",
                                       "vorher": v, "verwaltet": self._verwaltet(bi, c)})
                self.soll_woerter[key] = _wort_diff(v, inhalt)
            for si, (anker, zellen, keys) in enumerate(s.neue):
                for c, k in enumerate(keys, 1):
                    if not k:
                        continue
                    key = ("n", bi, si, c)
                    ii = paar.get(si)
                    ist = None
                    erreicht = False
                    if ii is not None:
                        ik = self.I[bi].neue[ii][2]
                        ist = self.I[bi].neue[ii][1][c - 1] if c - 1 < len(self.I[bi].neue[ii][1]) else ""
                        erreicht = ort[si] and c - 1 < len(ik) and ik[c - 1] == k
                    self.einheiten.append({"key": key, "bi": bi, "zeile": anker, "spalte": c, "soll": zellen[c - 1],
                                           "ist_da": ii is not None, "ist": ist, "erreicht": erreicht, "art": "neu",
                                           "vorher": "", "verwaltet": self._verwaltet(bi, c)})
                    self.soll_woerter[key] = ([], _wort_liste(zellen[c - 1]))

        # Ist-Aenderungen (Auto): jede veraenderte Zelle bzw. neue Zeile
        self.ist = self._ist_liste(self.I)
        self.ist_woerter: dict[Any, tuple[Counter, Counter]] = {}
        for it in self.ist:
            if it["art"] == "neu":
                self.ist_woerter[it["key"]] = (Counter(), Counter(_wort_liste(it["inhalt"])))
            else:
                entf, dazu = _wort_diff(self._vorher(it["bi"], it["zeile"], it["spalte"]), it["inhalt"])
                self.ist_woerter[it["key"]] = (Counter(entf), Counter(dazu))

    def _ist_liste(self, diffs: list[Diff]) -> list[dict[str, Any]]:
        """Alle Zellen, die im Ergebnis anders sind als vorher, mit der Angabe, ob sie laut nachher richtig sind."""
        out: list[dict[str, Any]] = []
        for bi, d in enumerate(diffs):
            s = self.S[bi]
            for (z, c), inhalt in sorted(d.zellen.items()):
                soll = s.zellen.get((z, c))
                richtig = soll is not None and _k(soll) == _k(inhalt)
                out.append({"key": ("z", bi, z, c), "bi": bi, "zeile": z, "spalte": c, "inhalt": inhalt,
                            "richtig": richtig, "art": "weg" if z in d.weg else "zelle"})
            paar, ort = _neu_abgleich(s.neue, d.neue)
            rev = {ii: si for si, ii in paar.items()}
            for ii, (anker, zellen, keys) in enumerate(d.neue):
                si = rev.get(ii)
                for c, k in enumerate(keys, 1):
                    if not k:
                        continue
                    richtig = si is not None and ort[si] and c - 1 < len(s.neue[si][2]) and s.neue[si][2][c - 1] == k
                    out.append({"key": ("n", bi, anker, keys, c), "bi": bi, "zeile": anker, "spalte": c,
                                "inhalt": zellen[c - 1], "richtig": richtig, "art": "neu", "ist_index": ii})
        return out

    # ------------------------------------------------------------------ Fehler eintragen
    def _fehler(self, key: Any, bi: int, zeile: int, spalte: int, quelle: str) -> Fehler:
        f = self.fehler.get(key)
        if f is None:
            f = self.fehler[key] = Fehler(bi, zeile, spalte, quelle=quelle)
        return f

    # ------------------------------------------------------------------ Zuordnung Aenderung <-> Zelle
    def _zugehoerig(self, W: Counter, R: Counter) -> list:
        out = []
        for ae in self.aend:
            an, aa = self.ae_woerter[ae.id]
            gesamt = sum(an.values()) + sum(aa.values())
            ueber = sum((an & W).values()) + sum((aa & R).values())
            if gesamt and ueber and ueber / gesamt >= 0.4:
                out.append(ae)
        return out

    @staticmethod
    def _deckung(unit_dazu: Counter, unit_entf: Counter, ist_entf: Counter, ist_dazu: Counter) -> float:
        """Wie viel vom Soll-Inhalt einer Einheit steckt in einer Ist-Aenderung? Massgeblich sind die NEUEN Woerter;
        gibt es keine (reine Streichung), zaehlen die entfallenen."""
        if unit_dazu:
            return sum((unit_dazu & ist_dazu).values()) / sum(unit_dazu.values())
        if unit_entf:
            return sum((unit_entf & ist_entf).values()) / sum(unit_entf.values())
        return 0.0

    def _landung(self, W: Counter, R: Counter, ausser: Any):
        """Wo im Ergebnis sind die Woerter gelandet (veraenderte Zelle, die den Soll-Inhalt enthaelt)?"""
        best, best_w = None, (0.0, 0)
        for it in self.ist:
            if it["key"] == ausser:
                continue
            entf, dazu = self.ist_woerter[it["key"]]
            d = self._deckung(W, R, entf, dazu)
            wert = (d, -(sum(dazu.values()) + sum(entf.values())))
            if d >= 0.6 and wert > best_w:
                best, best_w = it, wert
        return (best["bi"], best["zeile"], best["spalte"]) if best else None

    def _soll_ort(self, ist_dazu: Counter, ist_entf: Counter):
        """Welche Soll-Einheit passt zu dieser veraenderten Ist-Zelle (deren Soll-Inhalt steckt in ihr)?"""
        best, best_w = None, (0.0, 0)
        for e in self.einheiten:
            entf, dazu = self.soll_woerter[e["key"]]
            d = self._deckung(Counter(dazu), Counter(entf), ist_entf, ist_dazu)
            wert = (d, -(len(dazu) + len(entf)))
            if d >= 0.6 and wert > best_w:
                best, best_w = e, wert
        return (best["bi"], best["zeile"], best["spalte"]) if best else None

    def _leer_fn(self, bi: int) -> Callable[[Zeile], bool]:
        """Eine Zeile ist "leer", wenn in den verwalteten Spalten nichts steht (nur Bemerkung/Prio o. ae. uebrig)."""
        def leer(z: Zeile) -> bool:
            return not any(k for c, k in enumerate(z.keys, 1) if k and self._verwaltet(bi, c))
        return leer

    def _soll_inhaltsspalte(self, bi: int) -> int | None:
        if bi not in self._inhaltsspalte:
            self._inhaltsspalte[bi] = self._soll_inhaltsspalte_berechnen(bi)
        return self._inhaltsspalte[bi]

    def _soll_inhaltsspalte_berechnen(self, bi: int) -> int | None:
        """Spalte mit den meisten veraenderten Zeichen laut Soll (Inhaltsspalte im Soll)."""
        gewicht: Counter[int] = Counter()
        s = self.S[bi]
        for (z, c), inhalt in s.zellen.items():
            gewicht[c] += len(inhalt) + len(self._vorher(bi, z, c))
        for _anker, zellen, _keys in s.neue:
            for c, t in enumerate(zellen, 1):
                gewicht[c] += len(t)
        return gewicht.most_common(1)[0][0] if gewicht else None

    def _layout_von(self, bi: int) -> dict | None:
        name = self.namen[bi] if bi < len(self.namen) else ""
        return next((l for l in self.plan.layout if l.get("blatt") == name), None)

    def _verwaltet(self, bi: int, spalte: int) -> bool:
        """Soll die Spalte vom Tool bearbeitet werden? Nur wenn das erkannte Layout zum Soll passt (Inhaltsspalte stimmt),
        gelten die uebrigen Spalten (z.B. Bemerkung) als "nicht verwaltet" und zaehlen nicht als Soll-Zelle.
        Bei unklarem oder abweichendem Layout zaehlt jede Spalte (nichts wird versteckt)."""
        lay = self._layout_von(bi)
        if not lay or not lay.get("inhalt_spalte"):
            return True
        if column_index_from_string(lay["inhalt_spalte"]) != self._soll_inhaltsspalte(bi):
            return True
        spalten = {column_index_from_string(lay[k]) for k in ("section_spalte", "titel_spalte", "inhalt_spalte") if lay.get(k)}
        return spalte in spalten

    def _layout_abweichung(self, bi: int, spalte: int) -> bool:
        """Erkannte Inhaltsspalte weicht von der Soll-Inhaltsspalte ab, und die Zelle liegt in der Soll-Inhaltsspalte."""
        lay = self._layout_von(bi)
        if lay is None:
            return False
        soll_sp = self._soll_inhaltsspalte(bi)
        if soll_sp is None or spalte != soll_sp:
            return False
        inh = lay.get("inhalt_spalte")
        return (not inh) or column_index_from_string(inh) != soll_sp

    # ------------------------------------------------------------------ Fehlerart-Kaskade
    def _art(self, bi: int, zeile: int, spalte: int, W: Counter, R: Counter, e_text: str | None, n_text: str | None,
             in_e_geaendert: bool, ausser: Any, ist_neu: bool = False) -> str:
        if (W - self.pdf_neu) or (R - self.pdf_alt):
            return "pdf_falsch_gelesen"
        zu = self._zugehoerig(W, R)
        ziele = [z for ae in zu for z in self.ziele_von[ae.id]]
        landung = self._landung(W, R, ausser)
        # Layout: falsches Blatt/falsche Spalte
        if self._layout_abweichung(bi, spalte):
            return "layout_falsch"
        if landung and (landung[0] != bi or landung[2] != spalte):
            return "layout_falsch"
        if ziele and not any(z.bi == bi and z.spalte in (None, spalte) for z in ziele):
            return "layout_falsch"
        # Zuordnung: richtige Spalte, falsche Zeile/falscher Absatz
        if landung and landung[1] != zeile:
            return "zuordnung_falsch"
        if ziele and not any(z.bi == bi and z.spalte in (None, spalte) and z.zeile in (None, zeile) for z in ziele):
            return "zuordnung_falsch"
        if e_text is not None and n_text is not None and _k(e_text) != _k(n_text) \
                and sorted(_wort_liste(e_text)) == sorted(_wort_liste(n_text)):
            return "zuordnung_falsch"          # dieselben Woerter, aber an anderer Stelle (Absatz vertauscht)
        # Plan: Ziel stimmt, Ergebnis nicht
        if in_e_geaendert or any(z.bi == bi and z.zeile in (None, zeile) and z.spalte in (None, spalte) for z in ziele):
            return "plan_falsch"
        return "unklar"

    def _art_fremd(self, it: dict[str, Any], ae_ids: list[str]) -> str:
        """Fehlerart fuer eine Zelle im Ergebnis, die laut Soll gar nicht (so) veraendert sein sollte."""
        entf, dazu = self.ist_woerter[it["key"]]
        ort = self._soll_ort(dazu, entf)      # (Ist-Woerter: neu, entfallen)
        if ort is None:
            return "pdf_falsch_gelesen" if ae_ids else "unklar"
        if ort[0] != it["bi"] or ort[2] != it["spalte"]:
            return "layout_falsch"
        if ort[1] != it["zeile"]:
            return "zuordnung_falsch"
        return "plan_falsch"

    # ------------------------------------------------------------------ Hauptauswertung
    def auswerten(self) -> dict[str, Any]:
        # --- Ziele der "uebernommen"-Aenderungen (nach dem Anwenden) -------------------------------
        ziel_zellen: dict[tuple[int, int, int], list[str]] = {}
        neu_ops: list[tuple[str, int, dict[int, str], str]] = []     # (Aenderung, bi, Zellen, nachkontrolle)
        uebernommen = [a for a in self.aend if a.status == UEBERNOMMEN]
        for ae in uebernommen:
            for op in ae.ops:
                bi = self.bi_von_name.get(op.blatt)
                if bi is None:
                    continue
                if op.art in ("ersetzen", "einfuegen", "zelle_leeren"):
                    ziel_zellen.setdefault((bi, op.zeile, op.spalte), []).append(ae.id)
                elif op.art == "zeile_einfuegen":
                    neu_ops.append((ae.id, bi, dict(op.zellen), ae.nachkontrolle))

        def stumm(key, bi, zeile, spalte, art, ids):
            f = self._fehler(key, bi, zeile, spalte, "auto")
            f.stumm = True
            f.stille_arten.add(art)
            for i in ids:
                if i not in f.aenderungen:
                    f.aenderungen.append(i)

        # (1)(4)(5) je Ziel-Zelle der Aenderungen
        for (bi, z, c), ids in sorted(ziel_zellen.items()):
            erw, tat = self._erwartet(bi, z, c), self._tatsaechlich(bi, z, c)
            key = ("z", bi, z, c)
            if _k(tat) != _k(erw):
                stumm(key, bi, z, c, "uebernommen_aber_falsch", ids)
                if any(next(a for a in uebernommen if a.id == i).nachkontrolle == "ok" for i in ids):
                    stumm(key, bi, z, c, "nachkontrolle_ok_aber_falsch", ids)
            if _k(tat) == _k(self._vorher(bi, z, c)):
                stumm(key, bi, z, c, "keine_wirkung", ids)

        # neue Zeilen: gibt es im Ergebnis eine Zeile mit dem Inhalt der Operation, und ist sie laut Soll richtig?
        ist_neu_zellen = [it for it in self.ist if it["art"] == "neu"]
        erklaert_neu: set[Any] = set()
        for aid, bi, zellen, nk in neu_ops:
            gesucht = {c: _k(t) for c, t in zellen.items() if _k(t)}
            treffer = [it for it in ist_neu_zellen if it["bi"] == bi and it["key"][3] and
                       all(c - 1 < len(it["key"][3]) and it["key"][3][c - 1] == t for c, t in gesucht.items())]
            ref = (bi, min(zellen) if zellen else 0)
            if not gesucht:
                continue
            if not treffer:
                key = ("nop", bi, aid)
                stumm(key, bi, 0, ref[1], "keine_wirkung", [aid])
                if nk == "ok":
                    stumm(key, bi, 0, ref[1], "nachkontrolle_ok_aber_falsch", [aid])
                continue
            for it in treffer:
                erklaert_neu.add(it["key"])
                if not it["richtig"]:
                    stumm(it["key"], bi, it["zeile"], it["spalte"], "uebernommen_aber_falsch", [aid])
                    if nk == "ok":
                        stumm(it["key"], bi, it["zeile"], it["spalte"], "nachkontrolle_ok_aber_falsch", [aid])

        # (2) Kollateral: veraenderte Zelle, laut Soll falsch, und keine Aenderung zielt darauf
        zeilen_mit_ziel = {(bi, z) for (bi, z, _c) in ziel_zellen}
        kollateral = 0
        kollateral_keys: set[Any] = set()
        for it in self.ist:
            if it["richtig"]:
                continue
            key = it["key"]
            ids = ziel_zellen.get((it["bi"], it["zeile"], it["spalte"]), []) if it["art"] != "neu" else []
            if it["art"] == "neu":
                if key in erklaert_neu:
                    continue
            elif ids or (it["art"] == "weg" and (it["bi"], it["zeile"]) in zeilen_mit_ziel):
                if not ids:      # Zeile wurde entfernt/geleert, weil die Aenderung eine andere Zelle der Zeile traf
                    ids = next(i for (b, z, c), i in ziel_zellen.items() if b == it["bi"] and z == it["zeile"])
                stumm(key, it["bi"], it["zeile"], it["spalte"], "uebernommen_aber_falsch", ids)
                continue
            stumm(key, it["bi"], it["zeile"], it["spalte"], "kollateral", [])
            kollateral += 1
            kollateral_keys.add(key)

        # --- Soll-Zellen: erreicht / Grund ----------------------------------------------------
        zaehl = [e for e in self.einheiten if e["verwaltet"]]
        unverwaltet_n = len(self.einheiten) - len(zaehl)
        soll_n = len(zaehl)
        erreicht = sum(1 for e in zaehl if e["erreicht"])
        pruefen_n = nicht_n = ohne_n = 0
        for e in zaehl:
            if e["erreicht"]:
                continue
            entf, dazu = self.soll_woerter[e["key"]]
            W, R = Counter(dazu), Counter(entf)
            zu = self._zugehoerig(W, R)
            ziel_hier = [a for a in self.aend if any(
                z.bi == e["bi"] and z.zeile == e["zeile"] and z.spalte in (None, e["spalte"]) for z in self.ziele_von[a.id])]
            verwandt = {a.id: a for a in zu + ziel_hier}
            stati = {a.status for a in verwandt.values()}
            key = e["key"]
            if UEBERNOMMEN not in stati and PRUEFEN in stati:
                pruefen_n += 1
                continue
            if UEBERNOMMEN not in stati and NICHT in stati:
                nicht_n += 1
                continue
            if not verwandt:
                ohne_n += 1
            f = self._fehler(key, e["bi"], e["zeile"], e["spalte"], "soll")
            for i in sorted(verwandt):
                if i not in f.aenderungen:
                    f.aenderungen.append(i)
            if e["art"] == "neu":
                art = self._art(e["bi"], e["zeile"], e["spalte"], W, R, None, None, e["ist_da"], None, True)
            else:
                en = self._erwartet(e["bi"], e["zeile"], e["spalte"])
                et = self._tatsaechlich(e["bi"], e["zeile"], e["spalte"])
                art = self._art(e["bi"], e["zeile"], e["spalte"], W, R, et, en, e["ist_da"], key)
            f.art = f.art or art

        # Fehlerart der Eintraege, die nur als "Ergebnis falsch" entstanden sind
        for key, f in self.fehler.items():
            if f.art:
                continue
            it = next((i for i in self.ist if i["key"] == key), None)
            if it is not None and "kollateral" in f.stille_arten:
                f.art = self._art_fremd(it, [a.id for a in self.aend if self._trifft(a, it)])
            elif key[0] == "z":
                bi, z, c = key[1:]
                e = next((x for x in self.einheiten if x["key"] == key), None)
                if e is not None:
                    entf, dazu = self.soll_woerter[key]
                    f.art = self._art(bi, z, c, Counter(dazu), Counter(entf), self._tatsaechlich(bi, z, c),
                                      self._erwartet(bi, z, c), any(i["key"] == key for i in self.ist), key)
                elif it is not None:
                    f.art = self._art_fremd(it, f.aenderungen)
                else:
                    f.art = "plan_falsch"
            elif it is not None:
                f.art = self._art_fremd(it, f.aenderungen)
            else:
                f.art = "plan_falsch"

        # (3) Bilanz
        bilanz_ok = bool(self.bil_auto["stimmt"]) and bool(self.bil_plan["stimmt"])
        if not bilanz_ok:
            f = self._fehler(("paar", "bilanz"), 0, 0, 0, "paar")
            f.stumm, f.art = True, "bilanz_kaputt"
            f.stille_arten.add("bilanz")
        # Erwartete Luecken muessen gemeldet sein
        luecke_fehlt = 0
        if self.luecken_erwartet:
            gemeldet = Counter(l["art"] for l in self.plan.luecken)
            for art, n in Counter(l.get("art") for l in self.luecken_erwartet).items():
                if gemeldet.get(art, 0) < n:
                    luecke_fehlt += n - gemeldet.get(art, 0)
            if luecke_fehlt:
                f = self._fehler(("paar", "luecke"), 0, 0, 0, "paar")
                f.stumm, f.art = True, "bilanz_kaputt"
                f.stille_arten.add("luecke_nicht_gemeldet")

        # --- Treffergenauigkeit ---------------------------------------------------------------
        auto_gesamt = len(self.ist)
        auto_richtig = sum(1 for i in self.ist if i["richtig"])
        v_gesamt = v_richtig = None
        if self.e_alle is not None:
            vorhanden = {(i["key"], _k(i["inhalt"])) for i in self.ist}
            neu = [i for i in self.e_alle if (i["key"], _k(i["inhalt"])) not in vorhanden]
            v_gesamt = len(neu)
            v_richtig = sum(1 for i in neu if i["richtig"])

        # --- Lese-Luecke (Woerter, die laut Soll im PDF stehen muessten, aber nicht erkannt wurden) ----
        soll_woerter_n = sum(len(self.soll_woerter[e["key"]][0]) + len(self.soll_woerter[e["key"]][1]) for e in zaehl)
        fehlend = Counter()
        for e in zaehl:
            entf, dazu = self.soll_woerter[e["key"]]
            fehlend.update((Counter(dazu) - self.pdf_neu) + (Counter(entf) - self.pdf_alt))
        lese_luecke = sum(fehlend.values())
        zaehlung = max(0, soll_woerter_n - self.bil_plan["woerter_im_pdf"])

        # --- Zusammenfassen -------------------------------------------------------------------
        stille = [f for f in self.fehler.values() if f.stumm]
        art_zaehler = Counter()
        for f in stille:
            for a in f.stille_arten:
                art_zaehler[a] += 1
        nicht_ohne_grund = sum(1 for a in self.aend if a.status == NICHT and not a.grund)
        return {
            "bilanz_stimmt": bilanz_ok,
            "vollstaendig": bool(self.bil_auto.get("vollstaendig")),
            "stille_fehler": len(stille),
            "stille_fehler_nach_art": dict(art_zaehler),
            "soll_zellen": soll_n,
            "erreicht_auto": erreicht,
            "pruefen_zellen": pruefen_n,
            "nicht_uebernommen_zellen": nicht_n,
            "ohne_aenderung_zellen": ohne_n,
            "soll_zellen_unverwaltet": unverwaltet_n,
            "auto_richtig": auto_richtig,
            "auto_gesamt": auto_gesamt,
            "vorschlaege_richtig": v_richtig,
            "vorschlaege_gesamt": v_gesamt,
            "kollateral": kollateral,
            "idempotent": self.idempotent,
            "nicht_uebernommen_ohne_grund": nicht_ohne_grund,
            "soll_woerter": soll_woerter_n,
            "lese_luecke_woerter": lese_luecke,
            "lese_luecke_zaehlung": zaehlung,
            "luecken_nicht_gemeldet": luecke_fehlt,
            "luecken_gemeldet": len(self.plan.luecken),
            "aenderungen": len(self.aend),
            "bilanz": {k: v for k, v in self.bil_auto.items() if isinstance(v, (int, bool))},
            "aenderungen_nach_status": {s: sum(1 for a in self.aend if a.status == s) for s in (UEBERNOMMEN, PRUEFEN, NICHT)},
            "hinweise": self.hinweise,
            "fehler": [
                {"blatt": f.bi + 1 if f.quelle != "paar" else 0, "zeile": f.zeile, "spalte": f.spalte,
                 "aenderung": f.aenderungen[0] if f.aenderungen else "", "art": f.art, "stumm": f.stumm}
                for f in sorted(self.fehler.values(), key=lambda f: (f.bi, f.zeile, f.spalte, f.art))
            ],
        }

    def _trifft(self, ae, it: dict[str, Any]) -> bool:
        return any(z.bi == it["bi"] and z.zeile in (None, it["zeile"]) and z.spalte in (None, it["spalte"])
                   for z in self.ziele_von[ae.id] if ae.status == UEBERNOMMEN)


# ======================================================================================
# 4. Ablauf je Paar
# ======================================================================================
def _lade_erwartet(ordner: Path) -> dict[str, Any] | None:
    p = ordner / "erwartet.json"
    if not p.is_file():
        return None
    try:
        daten = json.loads(p.read_text(encoding="utf-8"))
        return daten if isinstance(daten, dict) else None
    except (OSError, ValueError):
        return None


def messe_paar(ordner: Path, mit_vorschlaegen: bool = False, ergebnis_kopie: Path | None = None) -> dict[str, Any]:
    """Misst ein Paar. Gibt die Kennzahlen (ohne Dokumentinhalt) zurueck; wirft bei kaputter Eingabe.
    ergebnis_kopie: nur fuer Tests -- kopiert die Ergebnis-Excel (nur Automatik) dorthin."""
    pdf, vorher, nachher = (ordner / n for n in DATEIEN)
    for p in (pdf, vorher, nachher):
        if not p.is_file():
            raise FileNotFoundError(p.name)
    erwartet = _lade_erwartet(ordner)
    hashes_vorher = {p.name: _sha(p) for p in (pdf, vorher, nachher)}
    t0 = time.monotonic()
    with tempfile.TemporaryDirectory(prefix="cde_messen_") as tmp_name:
        tmp = Path(tmp_name)
        # A: nur Automatik
        pa = analysiere_pdf(pdf)
        xl = ExcelModell(vorher)
        plan = plane(pa, xl)
        bil_plan = bilanz(plan.aenderungen, plan.markierte_woerter, plan.ignoriert, plan.luecken)
        aend_auto = copy.deepcopy(plan.aenderungen)
        e_auto = tmp / "ergebnis_auto.xlsx"
        uebernehmen(aend_auto, vorher, e_auto, None)
        bil_auto = bilanz(aend_auto, plan.markierte_woerter, plan.ignoriert, plan.luecken)
        # B: alle "pruefen" freigegeben (Qualitaet der Vorschlaege)
        e_alle = None
        if mit_vorschlaegen:
            aend_alle = copy.deepcopy(plan.aenderungen)
            e_alle = tmp / "ergebnis_alle.xlsx"
            uebernehmen(aend_alle, vorher, e_alle, {a.id for a in aend_alle if a.status == PRUEFEN})
        # Idempotenz: Ergebnis nochmals als "vorher" mit demselben PDF
        e_auto2 = tmp / "ergebnis_auto2.xlsx"
        plan2 = plane(analysiere_pdf(pdf), ExcelModell(e_auto))
        uebernehmen(plan2.aenderungen, e_auto, e_auto2, None)
        idempotent = _lese_excel(e_auto2) == _lese_excel(e_auto)

        if ergebnis_kopie is not None:
            shutil.copy2(e_auto, ergebnis_kopie)
        v_bl, n_bl, e_bl = _lese_excel(vorher), _lese_excel(nachher), _lese_excel(e_auto)
        e_alle_bl = _lese_excel(e_alle) if e_alle else None
        auswertung = _Auswertung(v_bl, n_bl, e_bl, e_alle_bl, _blattnamen(vorher), aend_auto, plan, pa, bil_auto,
                                 bil_plan, idempotent, (erwartet or {}).get("luecken"))
        res = auswertung.auswerten()

    res["original_unveraendert"] = all(_sha(p) == hashes_vorher[p.name] for p in (pdf, vorher, nachher))
    res["dauer_sekunden"] = round(time.monotonic() - t0, 2)
    if erwartet:
        res["stoerer"] = [s for s in erwartet.get("stoerer", []) if isinstance(s, str)]
        res["profil"] = erwartet.get("profil", "") if isinstance(erwartet.get("profil", ""), str) else ""
        wm = erwartet.get("woerter_markiert")
        if isinstance(wm, int):
            res["lese_luecke_erwartet"] = max(0, wm - bil_plan["woerter_im_pdf"])
    return res


def _fehlgeschlagen(name: str, exc: BaseException) -> dict[str, Any]:
    """Paar konnte nicht gemessen werden: Fehlerart "unklar", Grund nur als Name der Ausnahme-Klasse."""
    return {"name": name, "abgestuerzt": type(exc).__name__, "bilanz_stimmt": None, "vollstaendig": None,
            "stille_fehler": 0, "stille_fehler_nach_art": {}, "soll_zellen": 0, "erreicht_auto": 0, "pruefen_zellen": 0,
            "nicht_uebernommen_zellen": 0, "ohne_aenderung_zellen": 0, "soll_zellen_unverwaltet": 0, "auto_richtig": 0,
            "auto_gesamt": 0, "vorschlaege_richtig": None, "vorschlaege_gesamt": None, "kollateral": 0,
            "idempotent": None, "nicht_uebernommen_ohne_grund": 0, "soll_woerter": 0, "lese_luecke_woerter": 0,
            "lese_luecke_zaehlung": 0, "luecken_nicht_gemeldet": 0, "luecken_gemeldet": 0, "aenderungen": 0,
            "aenderungen_nach_status": {}, "hinweise": [], "original_unveraendert": None,
            "fehler": [{"blatt": 0, "zeile": 0, "spalte": 0, "aenderung": "", "art": "unklar", "stumm": False,
                        "grund": type(exc).__name__}]}


# ======================================================================================
# 5. Bericht
# ======================================================================================
def _anteil(zaehler: float, nenner: float) -> float | None:
    return round(zaehler / nenner, 4) if nenner else None


def _ampel(ok: bool | None, leer: str = "nicht gemessen") -> str:
    return leer if ok is None else ("ok" if ok else "NICHT ok")


def fasse_zusammen(paare: list[dict[str, Any]], korpus: str, generator_version: str, vorschlaege: bool) -> dict[str, Any]:
    gemessen = [p for p in paare if "abgestuerzt" not in p]

    def s(k: str) -> int:
        return sum(p.get(k) or 0 for p in paare)

    fehlerarten = {a: 0 for a in FEHLERARTEN}
    for p in paare:
        for f in p["fehler"]:
            fehlerarten[f["art"]] = fehlerarten.get(f["art"], 0) + 1
    stille_art = Counter()
    for p in paare:
        stille_art.update(p.get("stille_fehler_nach_art", {}))
    soll = s("soll_zellen")
    gesamt = {
        "paare": len(paare), "paare_gemessen": len(gemessen), "paare_abgestuerzt": len(paare) - len(gemessen),
        "soll_zellen": soll,
        "bilanz_stimmig": sum(1 for p in gemessen if p["bilanz_stimmt"]),
        "vollstaendig": sum(1 for p in gemessen if p["vollstaendig"]),
        "stille_fehler": s("stille_fehler"),
        "stille_fehler_nach_art": {a: stille_art.get(a, 0) for a in STILLE_ARTEN},
        "auto_richtig": s("auto_richtig"), "auto_gesamt": s("auto_gesamt"),
        "vorschlaege_gemessen": vorschlaege,
        "vorschlaege_richtig": s("vorschlaege_richtig") if vorschlaege else None,
        "vorschlaege_gesamt": s("vorschlaege_gesamt") if vorschlaege else None,
        "treffergenauigkeit_auto": _anteil(s("auto_richtig"), s("auto_gesamt")),
        "vorschlags_genauigkeit": _anteil(s("vorschlaege_richtig"), s("vorschlaege_gesamt")) if vorschlaege else None,
        "abdeckung_ohne_nachfrage": _anteil(s("erreicht_auto"), soll),
        "anteil_pruefen": _anteil(s("pruefen_zellen"), soll),
        "anteil_nicht_uebernommen": _anteil(s("nicht_uebernommen_zellen"), soll),
        "ohne_aenderung_zellen": s("ohne_aenderung_zellen"),
        "soll_zellen_unverwaltet": s("soll_zellen_unverwaltet"),
        "nicht_uebernommen_ohne_grund": s("nicht_uebernommen_ohne_grund"),
        "soll_woerter": s("soll_woerter"),
        "lese_luecke_woerter": s("lese_luecke_woerter"),
        "lese_luecke_anteil": _anteil(s("lese_luecke_woerter"), s("soll_woerter")),
        "idempotent": sum(1 for p in gemessen if p["idempotent"]),
        "kollateral": s("kollateral"),
        "original_unveraendert": sum(1 for p in gemessen if p["original_unveraendert"]),
        "luecken_nicht_gemeldet": s("luecken_nicht_gemeldet"),
    }
    je_stoerer: dict[str, dict[str, Any]] = {}
    for p in paare:
        for st in p.get("stoerer", []):
            e = je_stoerer.setdefault(st, {"paare": 0, "stille_fehler": 0, "soll_zellen": 0, "erreicht_auto": 0})
            e["paare"] += 1
            e["stille_fehler"] += p["stille_fehler"]
            e["soll_zellen"] += p["soll_zellen"]
            e["erreicht_auto"] += p["erreicht_auto"]
    for e in je_stoerer.values():
        e["abdeckung_ohne_nachfrage"] = _anteil(e.pop("erreicht_auto"), e.pop("soll_zellen"))
    n = len(paare)                       # alle Paare
    m = len(gemessen)                    # davon messbar
    voll = n > 0 and not gesamt["paare_abgestuerzt"]
    standard = [p for p in gemessen if p.get("profil") == "standard"]
    standard_abdeckung = _anteil(sum(p["erreicht_auto"] for p in standard), sum(p["soll_zellen"] for p in standard))
    kriterien = [
        {"nr": 1, "name": "Bilanz stimmig", "wert": f"{gesamt['bilanz_stimmig']} von {n} Paaren", "ziel": "alle Paare",
         "ampel": _ampel(voll and gesamt["bilanz_stimmig"] == n and not gesamt["luecken_nicht_gemeldet"])},
        {"nr": 2, "name": "Stille Fehler", "wert": str(gesamt["stille_fehler"]), "ziel": "0",
         "ampel": _ampel(gesamt["stille_fehler"] == 0)},
        {"nr": 3, "name": "Treffergenauigkeit automatisch",
         "wert": _prozent(gesamt["treffergenauigkeit_auto"]), "ziel": "100 %",
         "ampel": _ampel(None if not gesamt["auto_gesamt"] else gesamt["auto_richtig"] == gesamt["auto_gesamt"])},
        {"nr": 3, "name": "Genauigkeit der Vorschläge (nach Freigabe)", "wert": _prozent(gesamt["vorschlags_genauigkeit"]),
         "ziel": "mindestens 95 %",
         "ampel": _ampel(None if gesamt["vorschlags_genauigkeit"] is None else gesamt["vorschlags_genauigkeit"] >= 0.95,
                         "keine Vorschläge" if vorschlaege else "nicht gemessen")},
        {"nr": 4, "name": "Abdeckung ohne Nachfrage", "wert": _prozent(gesamt["abdeckung_ohne_nachfrage"]),
         "ziel": "mindestens 70 % (Standard-Aufbauten 90 %)",
         "ampel": _ampel(None if gesamt["abdeckung_ohne_nachfrage"] is None else
                         gesamt["abdeckung_ohne_nachfrage"] >= 0.7 and (standard_abdeckung is None or standard_abdeckung >= 0.9))},
        {"nr": 4, "name": "Nicht-Übernahmen mit Grund", "wert": f"{gesamt['nicht_uebernommen_ohne_grund']} ohne Grund",
         "ziel": "0 ohne Grund", "ampel": _ampel(gesamt["nicht_uebernommen_ohne_grund"] == 0)},
        {"nr": 5, "name": "Lese-Lücke", "wert": _prozent(gesamt["lese_luecke_anteil"]), "ziel": "höchstens 2 %",
         "ampel": _ampel(None if gesamt["lese_luecke_anteil"] is None else gesamt["lese_luecke_anteil"] <= 0.02)},
        {"nr": 6, "name": "Zweiter Lauf ändert nichts (Idempotenz)", "wert": f"{gesamt['idempotent']} von {m}",
         "ziel": "alle Paare", "ampel": _ampel(None if not m else gesamt["idempotent"] == m)},
        {"nr": 6, "name": "Kollateralschaden", "wert": str(gesamt["kollateral"]), "ziel": "0",
         "ampel": _ampel(gesamt["kollateral"] == 0)},
        {"nr": 6, "name": "Originale unverändert", "wert": f"{gesamt['original_unveraendert']} von {m}", "ziel": "alle Paare",
         "ampel": _ampel(None if not m else gesamt["original_unveraendert"] == m)},
    ]
    return {"format": BERICHT_FORMAT, "korpus": korpus, "generator_version": generator_version,
            "schema_version": SCHEMA_VERSION, "gesamt": gesamt, "fehlerarten": fehlerarten, "je_stoerer": je_stoerer,
            "kriterien": kriterien, "paare": paare}


def _prozent(x: float | None) -> str:
    return "-" if x is None else f"{x * 100:.1f} %".replace(".", ",")


def _md(b: dict[str, Any]) -> str:
    g = b["gesamt"]
    z = [f"# Messbericht {b['korpus']}", "",
         f"Paare: {g['paare']} (davon nicht messbar: {g['paare_abgestuerzt']}) · Soll-Zellen: {g['soll_zellen']} · "
         f"Schema {b['schema_version']} · Generator-Version: {b['generator_version'] or 'unbekannt'}", "",
         "Dieser Bericht enthält keine Texte aus Dokumenten, nur Zahlen, Positionen und Fehlerarten.", "",
         "## Kriterien (docs/ziel_v1.md)", "", "| Nr | Kriterium | Wert | Ziel | Ergebnis |", "|---|---|---|---|---|"]
    for k in b["kriterien"]:
        z.append(f"| {k['nr']} | {k['name']} | {k['wert']} | {k['ziel']} | {k['ampel']} |")
    z += ["", "## Zahlen", "",
          f"- Automatisch richtig: {g['auto_richtig']} von {g['auto_gesamt']}",
          f"- Anteil „zu prüfen“: {_prozent(g['anteil_pruefen'])} · Anteil „nicht übernommen“: {_prozent(g['anteil_nicht_uebernommen'])}"
          f" · Soll-Zellen ohne passende Änderung: {g['ohne_aenderung_zellen']}",
          f"- Soll-Zellen in nicht verwalteten Spalten (nicht mitgezählt): {g['soll_zellen_unverwaltet']}",
          f"- Bilanz „vollständig“: {g['vollstaendig']} von {g['paare']} Paaren",
          f"- Soll-Wörter: {g['soll_woerter']} · nicht erkannt (Lese-Lücke): {g['lese_luecke_woerter']}",
          "", "## Stille Fehler nach Art", ""]
    z += [f"- {a}: {n}" for a, n in g["stille_fehler_nach_art"].items()]
    z += ["", "## Fehlerarten", ""] + [f"- {a}: {n}" for a, n in b["fehlerarten"].items()]
    if b["je_stoerer"]:
        z += ["", "## Je Störer-Familie", "", "| Störer | Paare | Stille Fehler | Abdeckung |", "|---|---|---|---|"]
        for name, e in sorted(b["je_stoerer"].items()):
            z.append(f"| {name} | {e['paare']} | {e['stille_fehler']} | {_prozent(e['abdeckung_ohne_nachfrage'])} |")
    z += ["", "## Paare", "", "| Paar | Soll-Zellen | Abdeckung | Stille Fehler | Bilanz | Fehler |", "|---|---|---|---|---|---|"]
    for p in b["paare"]:
        ab = _prozent(_anteil(p["erreicht_auto"], p["soll_zellen"]))
        bil = "nicht messbar" if p["bilanz_stimmt"] is None else ("stimmt" if p["bilanz_stimmt"] else "NICHT")
        fe = "; ".join(" ".join(x for x in (f"Blatt {f['blatt']} Z{f['zeile']} S{f['spalte']}", f["aenderung"], f["art"],
                                            f"({f['grund']})" if f.get("grund") else "") if x) for f in p["fehler"])
        z.append(f"| {p['name']} | {p['soll_zellen']} | {ab} | {p['stille_fehler']} | {bil} | {fe or '-'} |")
    return "\n".join(z) + "\n"


def schreibe_bericht(bericht: dict[str, Any], korpus: str) -> Path:
    """Legt einen NEUEN Ordner NNN_<korpus> an (append-only) und schreibt bericht.json + bericht.md."""
    basis = config.pfad("messung_dir")
    basis.mkdir(parents=True, exist_ok=True)
    nummern = [int(m.group(1)) for d in basis.iterdir() if (m := re.match(r"^(\d{3,})_", d.name))]
    nr = max(nummern, default=0) + 1
    while True:
        ziel = basis / f"{nr:03d}_{korpus}"
        try:
            ziel.mkdir(exist_ok=False)
            break
        except FileExistsError:
            nr += 1
    with open(ziel / "bericht.json", "x", encoding="utf-8") as f:
        json.dump(bericht, f, ensure_ascii=False, indent=1)
    with open(ziel / "bericht.md", "x", encoding="utf-8") as f:
        f.write(_md(bericht))
    return ziel


def _konsole(b: dict[str, Any], ziel: Path, vorschlaege: bool) -> None:
    g = b["gesamt"]
    print(f"Messung {b['korpus']}: {g['paare']} Paare, {g['soll_zellen']} Soll-Zellen")
    if g["paare_abgestuerzt"]:
        print(f"Achtung: {g['paare_abgestuerzt']} Paar(e) konnten nicht gemessen werden (Fehlerart: unklar).")
    print("")
    breite = max(len(k["name"]) for k in b["kriterien"])
    for k in b["kriterien"]:
        print(f"  [{k['nr']}] {k['name']:<{breite}}  {k['wert']:<16} Ziel: {k['ziel']:<40} {k['ampel']}")
    print("")
    print(f"  Automatisch richtig: {g['auto_richtig']} von {g['auto_gesamt']}")
    print(f"  Anteil zu prüfen: {_prozent(g['anteil_pruefen'])}  |  nicht übernommen: {_prozent(g['anteil_nicht_uebernommen'])}"
          f"  |  ohne passende Änderung: {g['ohne_aenderung_zellen']}")
    print("  Fehlerarten: " + ", ".join(f"{a} {n}" for a, n in b["fehlerarten"].items()))
    if g["stille_fehler"]:
        print("")
        print(f"ACHTUNG: {g['stille_fehler']} stille(r) Fehler. Das Tool hat etwas als erledigt gemeldet, das nicht stimmt.")
        print("  Einzelheiten stehen im Bericht (Paar-Name, Blatt, Zeile, Spalte, Änderungs-ID).")
    if not vorschlaege:
        print("")
        print("Hinweis: Die Qualität der Vorschläge wurde nicht gemessen (Option --mit-vorschlaegen).")
    if b["korpus"] == "test":
        print("Hinweis (Regel 11): Den Test-Korpus nie zum Einstellen des Tools benutzen.")
    print("")
    print(f"Bericht: {ziel / 'bericht.md'}")


# ======================================================================================
# 6. Einstieg
# ======================================================================================
def _lade_korpus_info(eltern: Path) -> str:
    p = eltern / "KORPUS.json"
    try:
        return str(json.loads(p.read_text(encoding="utf-8")).get("generator_version", "")) if p.is_file() else ""
    except (OSError, ValueError, AttributeError):
        return ""


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="cde messen", description="Misst das Tool an Paaren (PDF + Excel vorher + Excel nachher).")
    ap.add_argument("--korpus", choices=["entwicklung", "test"], default="entwicklung")
    ap.add_argument("--paar", help="nur dieses Paar (Ordnername)")
    ap.add_argument("--ordner", help="anderer Eltern-Ordner mit Paaren (statt paare_dir/<korpus>)")
    ap.add_argument("--mit-vorschlaegen", action="store_true", help="auch die Qualität der Vorschläge (nach Freigabe) messen")
    try:
        a = ap.parse_args(argv)
    except SystemExit as e:                       # argparse beendet selbst; wir geben den Code zurueck
        return 2 if e.code else 0
    eltern = Path(a.ordner) if a.ordner else config.pfad("paare_dir") / a.korpus
    if not eltern.is_dir():
        print("Der Ordner mit den Paaren wurde nicht gefunden. Bitte Paare dort ablegen (siehe docs/paar_format.md).")
        return 2
    if a.paar:
        kandidaten = [eltern / a.paar] if (eltern / a.paar).is_dir() and Path(a.paar).name == a.paar else []
    else:
        kandidaten = sorted(d for d in eltern.iterdir() if d.is_dir() and not d.name.startswith("."))
    if not kandidaten:
        print("Es wurden keine Paare gefunden. Jedes Paar ist ein Ordner mit aenderungen.pdf, vorher.xlsx und nachher.xlsx.")
        return 2
    if a.korpus == "test":
        print("Hinweis (Regel 11): Der Test-Korpus wird nur zur Endmessung benutzt, nie zum Einstellen.")
    paare: list[dict[str, Any]] = []
    for d in kandidaten:
        try:
            p = messe_paar(d, a.mit_vorschlaegen)
        except Exception as exc:                  # ein kaputtes Paar bricht die Messung nicht ab
            p = _fehlgeschlagen(d.name, exc)
        p["name"] = d.name
        paare.append(p)
        print(f"  Paar {d.name}: {'nicht messbar' if 'abgestuerzt' in p else 'gemessen'}", flush=True)
    bericht = fasse_zusammen(paare, a.korpus, _lade_korpus_info(eltern), a.mit_vorschlaegen)
    ziel = schreibe_bericht(bericht, a.korpus)
    _konsole(bericht, ziel, a.mit_vorschlaegen)
    return 0


if __name__ == "__main__":
    sys.exit(main())
