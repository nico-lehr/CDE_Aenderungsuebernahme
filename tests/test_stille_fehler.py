"""Regressionstests gegen stille Fehler (Auftrag 003, Nacharbeit): je Ursache ein kleines erfundenes Beispiel (Regel 16), kein Paar
aus einem Korpus (Regel 11). Die Beispiele sind in der Form nachgebaut, in der der PDF-Leser sie liefert."""
from __future__ import annotations

from types import SimpleNamespace

import pytest

from src.gegenprobe import pruefe_absatz_markiert
from src.messen import Diff, _Auswertung
from src.markierung import NORMAL, Stueck, markiere_ops, projiziere
from src.models import Absatz, Lauf

K, A, D = (lambda t: Lauf("k", t)), (lambda t: Lauf("a", t, "A1")), (lambda t: Lauf("d", t, "A1"))
SP = Lauf("s", " ")


def _kurz(stuecke) -> str:
    return "".join({"normal": s.text, "alt": "{-" + s.text + "-}", "neu": "{+" + s.text + "+}"}[s.format] for s in stuecke)


def _projektion(excel: str, laeufe: list[Lauf]) -> str:
    proj = projiziere([excel], laeufe, {"A1"})
    assert proj is not None, "die Zuordnung zur Excel muss gelingen (sonst faellt der Plan auf die Operationen zurueck)"
    m = markiere_ops([Stueck(NORMAL, excel)], [], bereiche=[(0, len(excel), proj[0])])
    assert "".join(z[0] for z in m.zeichen if z[2] >= 0) == excel                 # jedes Excel-Zeichen steht unveraendert da
    ab = Absatz("P1", "CDE", [(2, 3, 0, len(excel))], laeufe)
    assert pruefe_absatz_markiert(ab, {(2, 3): (excel, m.zeichen)}, {"A1"}) is None
    return _kurz(m.stuecke)


# ================================================================ Ursache 1: Polster zwischen neuer Listennummer und neuem ersten Wort
POLSTER = ["\t", "\xa0" * 14 + " ", "\xa0" * 6]


@pytest.mark.parametrize("polster", POLSTER)
def test_neue_listennummer_polster_bleibt_normal_und_neues_wort_hat_sein_markiertes_leerzeichen(polster):
    """Word-PDF: alte Nummer gestrichen, neue Nummer unterstrichen und angeklebt ("c.a."), dann das unmarkierte Polster der Excel
    (im PDF eine Luecke), dann das NEUE erste Wort samt seinem unterstrichenen Leerzeichen. Das Polster ist Teil des alten Textes
    und bleibt normal; das Leerzeichen hinter dem neuen Wort ist rot (wie in Word). Frueher gehoerte das Polster als Trenner
    zwischen zwei neuen Stuecken zur Markierung, dem alten Text fehlte der Leerraum, die Zuordnung scheiterte und der Plan fiel
    auf die Operationen zurueck, die das markierte Leerzeichen hinter dem Wort normal liessen."""
    excel = "d." + polster + "Ausführung des Ventils."
    laeufe = [D("d."), A("f."), SP, A("Besondere"), A(" "), K("Ausführung"), SP, K("des"), SP, K("Ventils.")]
    assert _projektion(excel, laeufe) == "{-d.-}{+f.+}" + polster + "{+Besondere +}Ausführung des Ventils."


def test_neue_listennummer_neues_wort_ohne_markiertes_leerzeichen():
    excel = "3.\tAllgemeine Regeln."
    laeufe = [D("3."), A("4."), SP, A("Weitere"), SP, K("Allgemeine"), SP, K("Regeln.")]
    # unmarkierter Trenner hinter dem neuen Wort: bleibt normal; das Polster (Tab) ebenfalls
    assert _projektion(excel, laeufe) == "{-3.-}{+4.+}\t{+Weitere+} Allgemeine Regeln."


def test_leerraum_zwischen_zwei_neuen_woertern_gehoert_weiter_zur_markierung():
    """Der Grundsatz bleibt: ein Trenner zwischen zwei Zeichen derselben markierten Klasse ist Teil der Markierung, solange der alte
    Text an dieser Stelle schon Leerraum hat (hier davor und dahinter)."""
    excel = "nach NORM 4711."
    laeufe = [K("nach"), SP, A("ZUSATZ"), SP, A("EN"), SP, K("NORM"), SP, K("4711.")]
    assert _projektion(excel, laeufe) == "nach {+ZUSATZ EN+} NORM 4711."


def test_mehrere_luecken_im_neuen_block_nur_die_erste_bleibt_alter_leerraum():
    """Nummer, Polster, zwei neue Woerter, markiertes Leerzeichen: das Polster ist der Leerraum des alten Textes, das Leerzeichen
    zwischen den neuen Woertern gehoert zur Markierung."""
    excel = "c.\tText."
    laeufe = [D("c."), A("a."), SP, A("Neues"), SP, A("Wort"), A(" "), K("Text.")]
    assert _projektion(excel, laeufe) == "{-c.-}{+a.+}\t{+Neues Wort +}Text."


def test_neues_wort_hinter_polster_mit_unmarkiertem_trenner_dahinter():
    excel = "c.\tText."
    laeufe = [D("c."), A("a."), SP, A("Neues"), SP, A("Wort"), SP, K("Text.")]
    assert _projektion(excel, laeufe) == "{-c.-}{+a.+}\t{+Neues Wort+} Text."


# ================================================================ Messwerkzeug: zwei Artefakte (neue Fehlerklassen)
# Auf einem Paar mit sehr langem Blatt (Zeile > 500) meldete die Messung "Kollateral" in der Textspalte und "keine Wirkung" an einer
# neuen Zeile (Section-Spalte), obwohl die Ergebnis-Excel Zelle fuer Zelle dem Soll gleicht. Ursache war das Messwerkzeug: die
# Zeilen-Ausrichtung paart eine neue Listenzeile mit der aehnlichen alten Zeile direkt dahinter.
NB = "\xa0" * 17 + " "


def test_messung_polster_hinter_geaenderter_listennummer_bleibt_streng():
    """Ein geaenderter Listenpunkt steht als "alt"+"neu" ("a)b)") vor dem Polster. Wird er als NEUE Zeile eingeordnet, darf die
    Messung das Polster nicht zu Leerzeichen glaetten: das Polster steht schon so in vorher (`_polster_muster`)."""
    fake = SimpleNamespace(pdf_zeichen=set("ab). Sensoren und Dichtungen."), _POLSTER=_Auswertung._POLSTER)
    polster = _Auswertung._polster_muster([["", "a)" + NB + "Sensoren und Dichtungen."]])
    assert NB in polster
    soll = "a)b)" + NB + "Sensoren dauerhaft Dichtungen."
    assert _Auswertung._blind(fake, soll, "", polster) == soll                  # mit einer Nummer ging es schon, jetzt auch mit zwei
    assert _Auswertung._blind(fake, "b)" + NB + "Sensoren.", "", polster) == "b)" + NB + "Sensoren."
    # ein Polster, das so in keiner Zeile von vorher steht, ist weiter "PDF-blind" (RT-135)
    anders = "a)b)" + "\xa0" * 5 + " Sensoren."
    assert _Auswertung._blind(fake, anders, "", polster) == "a)b)" + " " * 6 + "Sensoren."


def _fake_messung(i_zellen: dict, s_zellen: dict, zuordnung: dict, vorher: dict):
    s, i = Diff(), Diff()
    s.zellen, i.zellen, s.zuordnung, i.zuordnung = dict(s_zellen), dict(i_zellen), dict(zuordnung), dict(zuordnung)
    return SimpleNamespace(S=[s], I=[i], _vorher=lambda bi, z, c: vorher.get((z, c), ""))


def test_messung_neue_zeile_als_gepaarte_zeile_mit_gleichen_konstanten():
    """Die neue Zeile wurde als geaenderte alte Zeile 700 gepaart. Ihre Section- und Sub-Section-Zelle sind dieselben wie in der alten
    Zeile (nicht in den Unterschieden), nur der Text ist neu: das ist keine fehlende Wirkung -- wenn Soll UND Ergebnis dort gleich
    stehen."""
    vorher = {(700, 5): "3.2", (700, 6): "E", (700, 7): "a) Sensoren."}
    neu = {(700, 7): "a) Filter und Getriebe."}
    gesucht = {5: "3.2", 6: "E", 7: "a) Filter und Getriebe."}
    fake = _fake_messung(neu, neu, {700: 709}, vorher)
    assert _Auswertung._neu_als_gepaarte_zeile(fake, 0, gesucht)
    # Text im Ergebnis anders als im Soll: weiter ein Fehler
    fake = _fake_messung({(700, 7): "a) Filter."}, neu, {700: 709}, vorher)
    assert not _Auswertung._neu_als_gepaarte_zeile(fake, 0, gesucht)
    # eine Konstante der neuen Zeile, die die alte Zeile NICHT hat: kein Treffer
    fake = _fake_messung(neu, neu, {700: 709}, {**vorher, (700, 5): "3.3"})
    assert not _Auswertung._neu_als_gepaarte_zeile(fake, 0, gesucht)
    # nichts hat sich geaendert (alles schon in der alten Zeile): die neue Zeile fehlt, also keine Wirkung
    fake = _fake_messung({}, {}, {700: 709}, {**vorher, (700, 7): "a) Filter und Getriebe."})
    assert not _Auswertung._neu_als_gepaarte_zeile(fake, 0, gesucht)
