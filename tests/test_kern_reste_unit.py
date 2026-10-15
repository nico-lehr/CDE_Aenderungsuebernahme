"""Unit-Tests K-REST: kleine Restarbeiten in Excel-Modell, Matcher, Planer und PDF-Leser (nur erfundene Daten, Regel 16).

Die Red-Team-Tests (tests/redteam/test_korpus_analyse.py) pruefen die ganzen Faelle; hier stehen die Bausteine und
die Gegenproben ("was darf NICHT blockieren").
"""
from __future__ import annotations

import pytest
from openpyxl import Workbook

from src import config
from src.excel_model import ExcelModell
from src.matcher import Matcher
from src.models import MITTEL, HOCH
from src.textutil import norm


def _tokens(text: str) -> list[tuple[int, str]]:
    return [(i, norm(w)) for i, w in enumerate(text.split()) if norm(w)]


def _excel(tmp_path, zeilen, name="a.xlsx", blatt="CDE", nachher=None):
    wb = Workbook()
    ws = wb.active
    ws.title = blatt
    for z in zeilen:
        ws.append(z)
    if nachher:
        nachher(ws)
    p = tmp_path / name
    wb.save(str(p))
    return p


_AUFBAU = [["Nr.", "Titel", "Inhalt"],
           ["1", "Allgemeines", "Der Betrieb erfolgt jederzeit durch geschultes Personal vor Ort."],
           ["2", "Stoermeldungen", "Die Reaktionszeit darf 150 ms nicht ueberschreiten."],
           ["3", "Normen", "Die Norm gilt fuer alle Anlagen im Betrieb."]]


# ===========================================================================================================
# K-XL 1: Nummern- und Titelspalte sind kein Ziel fuer Fliesstext (RT-120/121)
# ===========================================================================================================
def test_kandidaten_in_nummern_und_titelspalte_sind_markiert(tmp_path):
    xl = ExcelModell(_excel(tmp_path, _AUFBAU))
    lay = xl.layouts["CDE"]
    assert (lay.section_spalte, lay.titel_spalte, lay.inhalt_spalte) == (1, 2, 3)
    for k in xl.kandidaten:
        assert k.kopf_spalte == (k.spalte in (1, 2)), (k.art, k.spalte)
    assert any(k.kopf_spalte for k in xl.kandidaten) and any(not k.kopf_spalte for k in xl.kandidaten)


def test_listennummer_trifft_nummernzelle_nur_als_vorschlag(tmp_path):
    xl = ExcelModell(_excel(tmp_path, _AUFBAU))
    m = Matcher(xl, config.lade("zuordnung"))
    t = m.finde(_tokens("2"), "2", ist_kopf=False)
    assert t is not None and t.kandidat.spalte == 1 and t.kandidat.kopf_spalte
    assert t.stufe == MITTEL
    assert any("Nummern- oder Titelspalte" in g for g in t.gruende)


def test_ueberschrift_darf_weiter_automatisch_zugeordnet_werden(tmp_path):
    xl = ExcelModell(_excel(tmp_path, _AUFBAU))
    m = Matcher(xl, config.lade("zuordnung"))
    t = m.finde(_tokens("2 Stoermeldungen"), "2", ist_kopf=True)
    assert t is not None and t.stufe == HOCH
    assert not any("Nummern- oder Titelspalte" in g for g in t.gruende)


def test_fliesstext_im_inhalt_bleibt_automatisch(tmp_path):
    xl = ExcelModell(_excel(tmp_path, _AUFBAU))
    m = Matcher(xl, config.lade("zuordnung"))
    t = m.finde(_tokens("Die Reaktionszeit darf 150 ms nicht ueberschreiten."), "2", ist_kopf=False)
    assert t is not None and t.stufe == HOCH and not t.kandidat.kopf_spalte
