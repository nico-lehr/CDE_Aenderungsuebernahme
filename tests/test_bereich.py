"""Bereich der CDE (ADR 0006 Punkt 6): Aenderungen im PDF-Vorspann (Deckblatt, Pruef- und Revisionstabelle, Hinweiskaesten)
und im Anhang gehoeren nicht in die CDE. Sie stehen als EINE Zeile je Absatz im Protokoll (bereich = "ausserhalb"), zaehlen in
der Bilanz, sind aber kein offener Punkt. Nebenblaetter ohne Anker bekommen keine neuen Zeilen.

Nur erfundene Daten (Regel 16); die Mappe und das Word-PDF baut tests/test_register.py.
"""
from __future__ import annotations

import pytest

from src import config
from src.models import NICHT, PRUEFEN, UEBERNOMMEN
from src.protokoll import bilanz
from tests.test_register import GRUND, _plan

pytestmark = pytest.mark.legende_neu

VORSPANN = ["Pumpenhaus Leitfaden {-Ausgabe A-}{+Ausgabe B+} fuer den Betrieb",
            "Revision {-A.2.0 vom Maerz-}",
            "{+Hinweis fuer Bearbeiter: Platzhalter vor der Abgabe ersetzen.+}",
            "Pruefung durch {-Planung-}{+Fachplanung Ost+}"]


def _register_aenderungen(plan):
    return [a for a in plan.aenderungen if a.bereich != "ausserhalb"]


def test_vorspann_gehoert_nicht_zur_cde(tmp_path):
    pdf = list(GRUND)
    pdf[2] = ("sub", "A.", "Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.")
    plan, xl, _ = _plan(tmp_path, pdf, vorspann=VORSPANN)
    aussen = [a for a in plan.aenderungen if a.bereich == "ausserhalb"]
    assert len(aussen) == 4                                               # eine Zeile je Vorspann-Absatz
    for a in aussen:
        assert a.status == NICHT and not a.ops and a.pruefcodes == ["ausserhalb_cde"]
        assert a.seite == 1 and "nicht zur CDE" in a.grund_einfach and "Deckblatt/Vorspann" in a.grund
    assert [a.typ for a in aussen] == ["text_ersetzt", "text_geloescht", "absatz_neu", "text_ersetzt"]
    # keine neue Zeile in einem Nebenblatt, kein "nicht gefunden"
    assert not any(o.art == "zeile_einfuegen" for a in plan.aenderungen for o in a.ops)
    assert not any("nicht_gefunden" in a.pruefcodes for a in plan.aenderungen)
    (innen,) = _register_aenderungen(plan)
    assert innen.status == UEBERNOMMEN and innen.ziel == "Abgleich!G10"
    # Bilanz: die Woerter des Vorspanns zaehlen mit
    b = bilanz(plan.aenderungen, plan.markierte_woerter, plan.ignoriert)
    assert b["stimmt"] and b["woerter_nicht_uebernommen"] == sum(a.woerter for a in aussen) > 0


def test_gestrichene_part_zeile_im_vorspann_wird_nicht_auf_das_register_bezogen(tmp_path):
    grund = GRUND + [("part", "2", "PRODUKTE"), ("section", "2.1", "PUMPEN"), ("sub", "A.", "Pumpen sind frostsicher.")]
    plan, xl, _ = _plan(tmp_path, grund, vorspann=["Inhalt: Teil 1 Allgemeines", "{-PART 2 PRODUKTE-}"])
    (a,) = plan.aenderungen
    assert a.bereich == "ausserhalb" and a.status == NICHT and not a.ops


def test_anhang_hinter_dem_letzten_anker(tmp_path):
    from src.excel_model import ExcelModell
    from src.pdf_reader import analysiere_pdf
    from src.planner import plane
    from tests._word_pdf import WordDok
    from tests.test_register import excel_absaetze, register_mappe, stuecke
    xlsx = register_mappe(tmp_path / "a.xlsx", excel_absaetze(GRUND))
    # PDF mit Anhang auf einer eigenen Seite hinter der Spezifikation
    d = WordDok()
    for art, label, text in GRUND:
        if art == "part":
            d.kapitel(stuecke(f"PART {label} ") + stuecke(text))
        elif art == "section":
            d.ueberschrift(stuecke(label), stuecke(text))
        elif art in ("sub", "sub_titel"):
            d.listenpunkt(1, stuecke(label), [stuecke(text)], abstand=0)
        else:
            d.listenpunkt(2 if label[0].isdigit() else 3, stuecke(label), [stuecke(text)], abstand=0)
    d.neue_seite()
    d.absatz([stuecke("Anhang Pruefprotokoll {-Vorlage 1-}{+Vorlage 2+}")], ebene=1)
    plan = plane(analysiere_pdf(d.speichern(tmp_path / "b.pdf")), ExcelModell(xlsx))
    (a,) = plan.aenderungen
    assert a.bereich == "ausserhalb" and "Anhang" in a.grund and a.status == NICHT


def test_inhaltsverzeichnis_ohne_seitenzahlen_verschiebt_den_anfang_nicht(tmp_path):
    """Welle 3: Ein Inhaltsverzeichnis im Vorspann wiederholt die Ueberschriften genau (ohne Seitenzahlen). Bei gleich langer
    Kette gewinnt die echte Gliederung mit ihren Folgeabsaetzen; die Vorspann-Aenderung hinter dem Verzeichnis bleibt
    ausserhalb, die echte Ueberschrift ist sicher zugeordnet."""
    from src.excel_model import ExcelModell
    from src.pdf_reader import analysiere_pdf
    from src.planner import plane
    from tests._word_pdf import WordDok
    from tests.test_register import excel_absaetze, register_mappe, stuecke
    pdf = list(GRUND)
    pdf[1] = ("section", "1.1", "PUMPEN{+HAUS+}")
    xlsx = register_mappe(tmp_path / "a.xlsx", excel_absaetze(pdf))
    d = WordDok()
    d.absatz([stuecke("Inhalt")], ebene=1, abstand=6)
    d.kapitel(stuecke("PART 1 ALLGEMEINES"))
    d.ueberschrift(stuecke("1.1"), stuecke("PUMPEN"))
    d.ueberschrift(stuecke("1.2"), stuecke("LAGERUNG"))
    d.absatz([stuecke("Revision {-A.2.0-}{+A.3.0+} vom Fruehjahr")], ebene=1, abstand=6)
    d.neue_seite()
    for art, label, text in pdf:
        if art == "part":
            d.kapitel(stuecke(f"PART {label} ") + stuecke(text))
        elif art == "section":
            d.ueberschrift(stuecke(label), stuecke(text))
        elif art in ("sub", "sub_titel"):
            d.listenpunkt(1, stuecke(label), [stuecke(text)], abstand=0)
        else:
            d.listenpunkt(2 if label[0].isdigit() else 3, stuecke(label), [stuecke(text)], abstand=0)
    plan = plane(analysiere_pdf(d.speichern(tmp_path / "b.pdf")), ExcelModell(xlsx))
    rev = next(a for a in plan.aenderungen if "A.2.0" in a.alt)
    assert rev.bereich == "ausserhalb" and not rev.ops
    titel = next(a for a in plan.aenderungen if a.neu == "HAUS")
    assert titel.status == UEBERNOMMEN and titel.ziel == "Abgleich!G9" and titel.seite == 2


def test_neuer_absatz_am_ende_der_spezifikation_bleibt_drin(tmp_path):
    pdf = GRUND + [("liste", "{+3.+}", "{+Die Kranich-Plakette ist sichtbar.+}")]
    plan, xl, _ = _plan(tmp_path, pdf)
    (a,) = plan.aenderungen
    assert a.bereich == "" and a.typ == "absatz_neu" and a.ops and a.ops[0].art == "zeile_einfuegen"


def test_zu_wenige_anker_dann_gilt_die_regel_nicht(tmp_path, monkeypatch):
    monkeypatch.setitem(config.lade("zuordnung")["planer"], "umfang_min_anker", 100)
    plan, _, _ = _plan(tmp_path, list(GRUND), vorspann=VORSPANN)
    assert not any(a.bereich == "ausserhalb" for a in plan.aenderungen)


def test_umfang_aus_schaltet_die_regel_ab(tmp_path, monkeypatch):
    monkeypatch.setitem(config.lade("zuordnung")["planer"], "umfang", "aus")
    plan, _, _ = _plan(tmp_path, list(GRUND), vorspann=VORSPANN)
    assert not any(a.bereich == "ausserhalb" for a in plan.aenderungen)


def test_kein_neuer_absatz_in_einem_nebenblatt_ohne_anker(tmp_path):
    """Ein Absatz mitten im Register, dessen alter Text nur im Hinweis-Blatt steht (Fundstelle ausserhalb der Reihenfolge):
    der neue Absatz danach bekommt dort keine neue Zeile."""
    pdf = GRUND[:5] + [("absatz", "", "Je Absatz gibt es eine {-Zeile-}{+Zeile im Register+}."),
                       ("absatz", "", "{+Der Kranich wird gezaehlt.+}")] + GRUND[5:]
    from tests.test_register import excel_absaetze
    plan, xl, _ = _plan(tmp_path, pdf, excel_absaetze(GRUND))
    neu = next(a for a in plan.aenderungen if "Kranich" in a.neu)
    assert neu.status == NICHT and "blatt_ohne_anker" in neu.pruefcodes and not neu.ops
    assert neu.grund_einfach and neu.bereich == ""
    assert not any(o.art == "zeile_einfuegen" and o.blatt != "Abgleich" for a in plan.aenderungen for o in a.ops)
    assert all(a.status in (NICHT, PRUEFEN, UEBERNOMMEN) for a in plan.aenderungen)
