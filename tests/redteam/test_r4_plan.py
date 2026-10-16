"""Red-Team R4, Teil PLAN: Zuordnung und Absatz-Sperre (erfundene Daten, Regel 16).

Ende-zu-Ende (analysiere_pdf -> ExcelModell -> plane -> uebernehmen(..., absaetze=plan.absaetze)).
Pruefregel: "richtig ODER sichtbar unsicher, nie still falsch".
"""
import pytest

from src.models import UEBERNOMMEN
from tests.redteam._bau_r4 import (aufbau, automatisch, ende_zu_ende, excel_sections, richtig_oder_unsicher, scheitern,
                                   zelle)

Z2 = "Zweiter Absatz mit genug Text darin für den Test."


def _excel(tmp_path, absaetze):
    return excel_sections(tmp_path, [("1", "Eins", absaetze), ("2", "Zwei", [Z2])])


def _elemente(markup):
    return [("h", "1 Eins"), ("p", markup), ("h", "2 Zwei"), ("p", Z2)]


# ---------------------------------------------------------------------------------------------
# RT-203: gruenes Wort (unbekannte Farbe) im selben Absatz wie eine blaue Aenderung -> Absatz-Sperre
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-203 gruenes Wort im Absatz sperrt die blaue Aenderung nicht (hoch/automatisch); Besitzer PLAN")
def test_rt203_gruenes_wort_sperrt_den_absatz(tmp_path):
    vorher = "Die Steuerung überwacht die Anlage täglich und alle Werte."
    alt = _excel(tmp_path, [vorher])
    l = ende_zu_ende(tmp_path, _elemente("Die Steuerung überwacht die Anlage {@gn:täglich@} und {+insgesamt+} alle Werte."), alt)
    gruen = [t for b in l.pdf.blocks for t in b.tokens if "täglich" in (t.old, t.new)]
    aufbau(bool(gruen) and all(t.status == "keep" for t in gruen), f"gruenes Wort muss als 'unveraendert' gelesen sein: {gruen}")
    blau = [a for a in l.aen if "insgesamt" in a.neu]
    aufbau(bool(blau), f"blaue Aenderung fehlt: {l.kurz()}")
    if any(automatisch(a) for a in blau):
        scheitern(f"RT-203: blaue Aenderung trotz unerklaerter Farbe im selben Absatz automatisch eingetragen "
                  f"({l.kurz()}); Zelle jetzt {zelle(l.out, 2, 3)!r}")
    if any(a.status != UEBERNOMMEN for a in blau) and zelle(l.out, 2, 3) != vorher:
        scheitern("RT-203: Zelle veraendert, obwohl die Aenderung nicht uebernommen ist")


# ---------------------------------------------------------------------------------------------
# RT-204: Excel weicht ein Wort AUSSERHALB der Aenderung vom PDF ab ("Druckwert" statt "Druckwerte") -> A1
# ---------------------------------------------------------------------------------------------
PDF_LANG = ("Die Steuerung meldet die gemessenen Druckwerte der Pumpen und der Lüftung in festen Abständen an die {-alte-} "
            "{+neue+} Leitstelle, die alle Meldungen sammelt, auswertet und im Störungsfall den Bereitschaftsdienst sofort "
            "informiert.")
EXCEL_LANG = ("Die Steuerung meldet die gemessenen Druckwert der Pumpen und der Lüftung in festen Abständen an die alte "
              "Leitstelle, die alle Meldungen sammelt, auswertet und im Störungsfall den Bereitschaftsdienst sofort informiert.")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-204 Excel weicht ausserhalb der Aenderung ab, trotzdem hoch/automatisch (A1 fehlt); Besitzer PLAN")
def test_rt204_excel_weicht_ausserhalb_ab_nicht_automatisch(tmp_path):
    alt = _excel(tmp_path, [EXCEL_LANG])
    l = ende_zu_ende(tmp_path, _elemente(PDF_LANG), alt)
    betroffen = [a for a in l.aen if "neue" in a.neu or "alte" in a.alt]
    aufbau(bool(betroffen), f"Aenderung fehlt: {l.kurz()}")
    if any(automatisch(a) for a in betroffen):
        scheitern(f"RT-204: alter PDF-Absatz steht so nicht in der Excel ('Druckwert'), trotzdem automatisch: {l.kurz()}")
    if zelle(l.out, 2, 3) != EXCEL_LANG:
        scheitern("RT-204: Zelle veraendert, obwohl die Aenderung nicht automatisch ist")


# ---------------------------------------------------------------------------------------------
# RT-205: geschuetztes Leerzeichen (NBSP) in der Excel direkt an der Einfuegestelle (Schutz-Test, besteht schon)
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
@pytest.mark.parametrize("pdf, erwartet", [
    pytest.param("Die Steuerung meldet 50 kW {+sofort+} an die Leitstelle.",
                 "Die Steuerung meldet 50 kW sofort an die Leitstelle.", id="pdf_leerzeichen_danach"),
    pytest.param("Die Steuerung meldet {+sofort+} 50 kW an die Leitstelle.",
                 "Die Steuerung meldet sofort 50 kW an die Leitstelle.", id="pdf_nbsp_davor"),
    pytest.param("Die Steuerung meldet 50 kW {+sofort+} an die Leitstelle.",
                 "Die Steuerung meldet 50 kW sofort an die Leitstelle.", id="pdf_nbsp_danach"),
])
def test_rt205_nbsp_an_der_einfuegestelle_bleibt(tmp_path, pdf, erwartet):
    vorher = "Die Steuerung meldet 50 kW an die Leitstelle."
    alt = _excel(tmp_path, [vorher])
    l = ende_zu_ende(tmp_path, _elemente(pdf), alt)
    betroffen = [a for a in l.aen if "sofort" in a.neu]
    richtig_oder_unsicher(l, betroffen, 2, 3, erwartet, vorher, "RT-205")


# ---------------------------------------------------------------------------------------------
# RT-213: derselbe Satz zweimal in derselben Zelle, Aenderung am zweiten Vorkommen; zweiter Lauf aendert nichts
# (Schutz-Test, besteht schon)
# ---------------------------------------------------------------------------------------------
SATZ = "Die Werte werden täglich geprüft."
MITTE = "Die Anlage läuft im Normalbetrieb ohne Unterbrechung."


@pytest.mark.redteam
def test_rt213_dublette_in_derselben_zelle_und_zweiter_lauf(tmp_path):
    vorher = f"{SATZ} {MITTE} {SATZ}"
    alt = _excel(tmp_path, [vorher])
    el = _elemente(f"{SATZ} {MITTE} Die Werte werden {{-täglich-}} {{+stündlich+}} geprüft.")
    l = ende_zu_ende(tmp_path, el, alt)
    betroffen = [a for a in l.aen if "stündlich" in a.neu]
    erwartet = f"{SATZ} {MITTE} Die Werte werden stündlich geprüft."
    richtig_oder_unsicher(l, betroffen, 2, 3, erwartet, vorher, "RT-213 erster Lauf")
    l2 = ende_zu_ende(tmp_path / "zwei", el, l.out)
    if zelle(l2.out, 2, 3) != zelle(l.out, 2, 3):
        scheitern(f"RT-213: zweiter Lauf auf dem Ergebnis aendert die Zelle erneut: {zelle(l2.out, 2, 3)!r} ({l2.kurz()})")
    if any(automatisch(a) for a in l2.aen):
        scheitern(f"RT-213: zweiter Lauf meldet wieder automatische Aenderungen: {l2.kurz()}")
