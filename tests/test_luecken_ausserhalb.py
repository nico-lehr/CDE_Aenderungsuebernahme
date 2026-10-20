"""Welle 6: Nicht gelesene Stellen (Tabelle, mehrere Spalten ...) auf Seiten OHNE CDE-Inhalt (Deckblatt, Pruef- und Revisionstabelle,
Anhang) sind nur Info. Sie machen die Bilanz nicht "unvollstaendig" und die Ampel nicht rot, bleiben aber im Plan und im Protokoll
sichtbar (nichts geht still verloren). Eine Luecke auf einer Seite mit CDE-Inhalt bleibt eine echte Luecke.

Nur erfundene Daten (Regel 16); Mappe und Word-PDF bauen die Helfer aus tests/test_register.py.
"""
from __future__ import annotations

import pytest

from src import vertrauen
from src.excel_model import ExcelModell
from src.models import UEBERNOMMEN, luecke
from src.pdf_reader import analysiere_pdf
from src.planner import plane
from src.protokoll import bilanz
from tests.test_register import GRUND, excel_absaetze, register_mappe, wort_pdf

pytestmark = pytest.mark.legende_neu

VORSPANN = ["Pumpenhaus Leitfaden {-Ausgabe A-}{+Ausgabe B+} fuer den Betrieb",
            "Revision {-A.2.0 vom Maerz-}"]


def _analyse(tmp_path):
    pdf = list(GRUND)
    pdf[2] = ("sub", "A.", "Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.")
    pa = analysiere_pdf(wort_pdf(tmp_path / "a.pdf", pdf, VORSPANN))
    xl = ExcelModell(register_mappe(tmp_path / "alt.xlsx", excel_absaetze(pdf)))
    return pa, xl


def _cde_seite(plan) -> int:
    (innen,) = [a for a in plan.aenderungen if a.bereich != "ausserhalb"]
    assert innen.status == UEBERNOMMEN
    return innen.seite


def test_luecke_im_vorspann_ist_nur_info_und_eine_im_text_bleibt_echt(tmp_path):
    pa, xl = _analyse(tmp_path)
    seite = _cde_seite(plane(pa, xl))
    assert seite > 1
    pa.luecken += [luecke("tabelle", 1, "Pruefkreis-Tabelle im Vorspann", 40),          # Deckblatt-Seite
                   luecke("mehrspaltig", seite, "Zwei Spalten im Text", 12)]             # Seite mit CDE-Inhalt
    plan = plane(pa, xl)
    assert [(l["art"], l["seite"]) for l in plan.luecken] == [("mehrspaltig", seite)]
    assert [(l["art"], l["seite"]) for l in plan.luecken_ausserhalb] == [("tabelle", 1)]
    assert len(plan.alle_luecken) == 2
    b = bilanz(plan.aenderungen, plan.markierte_woerter, plan.ignoriert, plan.luecken, plan.luecken_ausserhalb)
    assert b["luecken"] == 1 and b["vollstaendig"] is False and b["luecken_ausserhalb"] == 1 and b["stimmt"]


def test_nur_vorspann_luecken_lassen_die_bilanz_vollstaendig_und_die_ampel_nicht_rot(tmp_path):
    pa, xl = _analyse(tmp_path)
    pa.luecken += [luecke("tabelle", 1, "Pruefkreis-Tabelle im Vorspann", 40)]
    plan = plane(pa, xl)
    assert plan.luecken == [] and len(plan.luecken_ausserhalb) == 1
    b = bilanz(plan.aenderungen, plan.markierte_woerter, plan.ignoriert, plan.luecken, plan.luecken_ausserhalb)
    assert b["vollstaendig"] is True and b["stimmt"] and b["luecken_ausserhalb"] == 1
    daten = {"aenderungen": [a.zu_dict() for a in plan.aenderungen], "bilanz": b, "luecken": plan.luecken, "spec": plan.spec,
             "stichprobe": []}
    assert vertrauen.ampel(daten)["stufe"] != "rot"


def test_luecke_auf_der_ersten_cde_seite_bleibt_echt(tmp_path):
    """Streng: nur Seiten VOR der ersten bzw. NACH der letzten CDE-Seite sind "ausserhalb"; auf derselben Seite kann Inhalt stehen."""
    pa, xl = _analyse(tmp_path)
    erste = _cde_seite(plane(pa, xl))
    pa.luecken += [luecke("tabelle", erste, "Tabelle vor dem Text auf der ersten CDE-Seite", 30)]
    plan = plane(pa, xl)
    assert [l["seite"] for l in plan.luecken] == [erste] and plan.luecken_ausserhalb == []


def test_ohne_sicheren_bereich_gilt_jede_luecke_als_echt(tmp_path):
    """Word-Dateien (Seite 0) und Dokumente ohne sicheren CDE-Bereich: nichts wird als "ausserhalb" weggeordnet."""
    from tests._docx_bau import register_docx
    from src.docx_reader import analysiere_docx
    pdf = list(GRUND)
    pdf[2] = ("sub", "A.", "Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.")
    xl = ExcelModell(register_mappe(tmp_path / "alt.xlsx", excel_absaetze(pdf)))
    pa = analysiere_docx(register_docx(tmp_path / "w.docx", pdf))
    pa.luecken += [luecke("tabelle", 0, "Tabelle in der Word-Datei", 20), luecke("tabelle", 1, "Tabelle (Seite unbekannt)", 20)]
    plan = plane(pa, xl)
    assert len(plan.luecken) == 2 and plan.luecken_ausserhalb == []


def test_schon_eingetragen_ist_keine_handarbeit_in_der_ampel():
    def aen(i, status, bereich, codes):
        return {"id": f"A{i:04d}", "status": status, "bereich": bereich, "pruefcodes": codes, "typ": "text_ersetzt"}
    plan = {"aenderungen": [aen(1, "uebernommen", "", []), aen(2, "nicht_uebernommen", "", ["schon_eingetragen"]),
                            aen(3, "nicht_uebernommen", "", ["nicht_gefunden"]), aen(4, "nicht_uebernommen", "ausserhalb", ["ausserhalb_cde"])],
            "bilanz": {"stimmt": True, "vollstaendig": True}, "luecken": [], "stichprobe": []}
    r = vertrauen.ampel(plan)
    assert r["offen"] == 1                                    # nur "nicht gefunden" macht Arbeit
    assert r["stufe"] == "gelb" and any("1 Änderung wird nicht eingetragen" in g for g in r["gruende"])
    plan["aenderungen"][2]["pruefcodes"] = ["schon_eingetragen"]
    assert vertrauen.ampel(plan)["offen"] == 0


def test_luecke_im_bereich_macht_die_ampel_rot():
    plan = {"aenderungen": [{"id": "A0001", "status": "uebernommen", "bereich": "", "pruefcodes": [], "typ": "text_ersetzt"}],
            "bilanz": {"stimmt": True, "vollstaendig": False, "luecken": 1}, "luecken": [luecke("tabelle", 9, "Tabelle", 30)],
            "stichprobe": []}
    assert vertrauen.ampel(plan)["stufe"] == "rot"
