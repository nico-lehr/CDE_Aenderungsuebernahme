"""Leerraum-Abweichungen zwischen Excel und altem PDF-Text (Auftrag "ZUORDNEN", Code a1_leerraum). Nur erfundene Daten (Regel 16).

Erkenntnis aus der Messung: `a1_leerraum` hiess bisher "an der ersten Abweichung steht auf einer Seite ein Leerzeichen". Das ist auch
dann der Fall, wenn dort ein Buchstabe oder ein Satzzeichen VOR dem Leerzeichen fehlt ("Temperatu darf" statt "Temperatur darf"):
Der Grund-Satz sagte dann "Leerzeichen", obwohl ein Buchstabe fehlt. Jetzt heisst die Abweichung nur dann Leerraum, wenn die Zeichen
ausser dem Leerraum in beiden Texten genau gleich sind.

Eine Abweichung im Leerraum macht die Aenderung NIE automatisch (Regel 27): Die Excel steht dort nicht so, wie das PDF den alten
Text zeigt; ein Mensch gibt den fertigen Vorschlag frei. Der Vorschlag laesst den Leerraum der Excel ausserhalb der markierten Stelle
unveraendert.

Die PDFs baut tests/_word_pdf.py in der Form der Word-Aenderungsverfolgung.
"""
from __future__ import annotations

import copy
from types import SimpleNamespace

import pytest

from src import codes, messen
from src.abgleich import a1_art
from src.applier import uebernehmen
from src.excel_model import ExcelModell
from src.models import PRUEFEN, UEBERNOMMEN
from src.pdf_reader import analysiere_pdf
from src.planner import plane
from tests._word_pdf import WordDok
from tests.redteam._bau import schreibe_excel

pytestmark = [pytest.mark.legende_neu, pytest.mark.markierung_an]


# ------------------------------------------------------------------------------------------------ Art der Abweichung
@pytest.mark.parametrize("pdf,excel,erwartet", [
    ("Die Anlage gilt fuer alle Pumpen.", "Die Anlage gilt fuer alle Pumpen.", ""),
    ("Die Anlage gilt\tfuer  alle Pumpen.", "Die Anlage gilt fuer alle Pumpen.", ""),       # Tab, doppelt, geschuetzt: gleich
    # nur Leerzeichen fehlen oder sind zu viel (gleiche Zeichen ausser dem Leerraum)
    ("Pumpen. Sie sind sicher.", "Pumpen.Sie sind sicher.", "leerraum"),
    ("Die Anlage gilt fuer alle Pumpen", "Die Anlagegilt fuer alle Pumpen", "leerraum"),
    ("Die Anlage gilt fuer alle Pumpen", "Die Anlage gilt fuer alle Pum pen", "leerraum"),
    ("Die Anlage gilt fuer alle Pumpen im Keller", "Die Anlage gilt fuer alle Pumpen imKeller", "leerraum"),
    # fehlt ein Buchstabe, ein Satzzeichen oder eine Ziffer VOR einem Leerzeichen, ist es kein Leerraum-Unterschied
    ("Die Temperatur darf nicht steigen", "Die Temperatu darf nicht steigen", "buchstabe"),
    ("Die Temperatur darf nicht steigen", "Die Temperaturen darf nicht steigen", "buchstabe"),
    ("Das darf nicht ueberschritten werden", "Das darf nicht, ueberschritten werden", "satzzeichen"),
    ("Der Druck betraegt 25 bar", "Der Druck betraegt 2 bar", "ziffer"),
    ("Der Druck betraegt 25 bar", "Der Druck betraegt 250 bar", "ziffer"),
    ("Die Anlage gilt fuer alle Pumpen", "Die anlage gilt fuer alle Pumpen", "gross_klein"),
    ("Die Anlage gilt fuer alle Pumpen", "Die Anlage gilt fuer alle Armaturen", "buchstabe"),
    ("Die Anlage gilt fuer alle Pumpen und Armaturen", "Die Anlage gilt fuer alle Pumpen", "pdf_laenger"),
    ("Die Anlage gilt fuer alle Pumpen", "Die Anlage gilt fuer alle Pumpen und Armaturen", "excel_laenger"),
])
def test_a1_art_nennt_leerraum_nur_bei_gleichen_zeichen(pdf, excel, erwartet):
    assert a1_art(pdf, excel) == erwartet


def test_a1_art_stimmt_mit_dem_alten_wert_ueberein_ausser_bei_leerraum():
    """Alle Arten ausser "leerraum" bleiben wie `vergleiche(...).art` (nur "leerraum" wird genauer)."""
    from src.gegenprobe import vergleiche
    paare = [("Die Anlage gilt fuer alle", "Die anlage gilt fuer alle"), ("Wert 25 mm", "Wert 26 mm"),
             ("Wert 25 mm.", "Wert 25 mm"), ("Wert 25 mm", "Wert 25 mm und mehr"), ("Ein Satz.", "Ein Satz!")]
    for pdf, excel in paare:
        alt = vergleiche(pdf, excel).art
        assert alt != "leerraum"
        assert a1_art(pdf, excel) == alt


# ------------------------------------------------------------------------------------------------ Grund-Satz
def _ae(codes_):
    return SimpleNamespace(status=PRUEFEN, pruefcodes=list(codes_), grund="")


@pytest.mark.parametrize("art,wort", [("leerraum", "nur Leerzeichen"), ("buchstabe", "ein Buchstabe"), ("satzzeichen", "ein Satzzeichen"),
                                      ("ziffer", "eine Ziffer"), ("gross_klein", "Groß-/Kleinschreibung"),
                                      ("pdf_laenger", "PDF hat mehr Text"), ("excel_laenger", "Excel hat mehr Text")])
def test_karten_satz_nennt_die_abweichung_in_einfachem_deutsch(art, wort):
    satz = codes.grund_einfach(_ae(["excel_weicht_ab", "a1_" + art]))
    assert satz == f"Die Excel weicht hier vom alten PDF-Text ab ({wort}). Vorschlag bitte prüfen."
    assert len(satz.split()) <= codes.KURZ_MAX_WOERTER and "_" not in satz


def test_karten_satz_ohne_a1_code_bleibt_wie_bisher():
    assert codes.grund_einfach(_ae(["excel_weicht_ab"])) == codes.kurz("excel_weicht_ab")
    assert codes.grund_einfach(_ae(["excel_weicht_ab", "a1_abweichung"])) == codes.kurz("excel_weicht_ab")


def test_hinweis_satz_zum_leerraum_ist_einfaches_deutsch():
    assert codes.satz("a1_leerraum") == "Der alte Text steht in der Excel etwas anders (Leerzeichen fehlen oder zu viele)."
    assert codes.satz("a1_buchstabe") == "Der alte Text steht in der Excel etwas anders (ein Buchstabe)."


# ------------------------------------------------------------------------------------------------ Planer: Ende zu Ende
def _lauf(tmp_path, zeilen_pdf, excel_text, name="a"):
    """Word-PDF (Ueberschrift 1.1 + ein Absatz) und Excel (eine Zeile mit `excel_text`) -> (Plan, Zellstuecke nach Freigabe aller
    Vorschlaege als [(format, text)], Zellstuecke der Automatik)."""
    d = WordDok()
    d.ueberschrift([("1.1", "k")], [("ANLAGE", "k")])
    d.absatz(zeilen_pdf)
    pdf = d.speichern(tmp_path / f"{name}.pdf")
    alt = schreibe_excel(tmp_path / f"{name}.xlsx", [["Nr.", "Titel", "Inhalt"], ["1.1", "ANLAGE", None], ["", "", excel_text]])
    plan = plane(analysiere_pdf(pdf), ExcelModell(alt))
    frei = tmp_path / f"{name}_frei.xlsx"
    uebernehmen(copy.deepcopy(plan.aenderungen), alt, frei, {a.id for a in plan.aenderungen if a.status == PRUEFEN},
                absaetze=plan.absaetze)
    auto = tmp_path / f"{name}_auto.xlsx"
    uebernehmen(copy.deepcopy(plan.aenderungen), alt, auto, None, absaetze=plan.absaetze)
    return (plan, messen._stuecke(messen._lese_excel(frei, True)[0][2][2]), messen._stuecke(messen._lese_excel(auto, True)[0][2][2]))


ERSETZT = [[("Die Anlage gilt fuer alle Pumpen im ", "k"), ("Keller", "d"), ("Hof", "n"), (" und im Werk.", "k")]]


def _genau_ein_vorschlag(plan):
    (a,) = plan.aenderungen
    assert a.status == PRUEFEN and a.konfidenz == "mittel" and a.ops, (a.status, a.pruefcodes)
    assert "excel_weicht_ab" in a.pruefcodes
    return a


def test_fehlender_buchstabe_vor_dem_leerzeichen_heisst_buchstabe_nicht_leerzeichen(tmp_path):
    """Die Excel hat "Pump" statt "Pumpen" (ein Buchstabe fehlt vor einem Leerzeichen). Frueher: a1_leerraum / "Leerzeichen"."""
    plan, frei, auto = _lauf(tmp_path, ERSETZT, "Die Anlage gilt fuer alle Pump im Keller und im Werk.")
    a = _genau_ein_vorschlag(plan)
    assert "a1_buchstabe" in a.pruefcodes and "a1_leerraum" not in a.pruefcodes
    assert a.grund_einfach == "Die Excel weicht hier vom alten PDF-Text ab (ein Buchstabe). Vorschlag bitte prüfen."
    assert any("ein Buchstabe" in h for h in a.hinweise) and not any("Leerzeichen" in h for h in a.hinweise)
    assert frei == [("normal", "Die Anlage gilt fuer alle Pump im "), ("alt", "Keller"), ("neu", "Hof"), ("normal", " und im Werk.")]
    assert auto == [("normal", "Die Anlage gilt fuer alle Pump im Keller und im Werk.")]          # Automatik schreibt nichts


def test_fehlendes_satzzeichen_vor_dem_leerzeichen_heisst_satzzeichen(tmp_path):
    plan, frei, auto = _lauf(tmp_path, ERSETZT, "Die Anlage gilt fuer alle Pumpen, im Keller und im Werk.")
    a = _genau_ein_vorschlag(plan)
    assert "a1_satzzeichen" in a.pruefcodes and "a1_leerraum" not in a.pruefcodes
    assert "(ein Satzzeichen)" in a.grund_einfach
    assert frei == [("normal", "Die Anlage gilt fuer alle Pumpen, im "), ("alt", "Keller"), ("neu", "Hof"), ("normal", " und im Werk.")]


@pytest.mark.parametrize("excel,vorn,hinten", [
    # Leerzeichen fehlen weit weg von der Aenderung: bleiben so stehen
    ("Die Anlagegilt fuer alle Pumpen im Keller und im Werk.", "Die Anlagegilt fuer alle Pumpen im ", " und im Werk."),
    ("Die Anlage gilt fuer alle Pum pen im Keller und im Werk.", "Die Anlage gilt fuer alle Pum pen im ", " und im Werk."),
    # Leerzeichen fehlt direkt vor bzw. hinter der Aenderung: die Excel behaelt ihr Muster, die Naht wird nicht "repariert"
    ("Die Anlage gilt fuer alle Pumpen imKeller und im Werk.", "Die Anlage gilt fuer alle Pumpen im", " und im Werk."),
    ("Die Anlage gilt fuer alle Pumpen im Kellerund im Werk.", "Die Anlage gilt fuer alle Pumpen im ", "und im Werk."),
])
def test_nur_leerzeichen_abweichung_wird_vorschlag_und_laesst_den_leerraum_der_excel_stehen(tmp_path, excel, vorn, hinten):
    plan, frei, auto = _lauf(tmp_path, ERSETZT, excel)
    a = _genau_ein_vorschlag(plan)
    assert "a1_leerraum" in a.pruefcodes
    assert a.grund_einfach == "Die Excel weicht hier vom alten PDF-Text ab (nur Leerzeichen). Vorschlag bitte prüfen."
    assert frei == [("normal", vorn), ("alt", "Keller"), ("neu", "Hof"), ("normal", hinten)]
    assert auto == [("normal", excel)]                              # nie automatisch (Regel 27), Zelle bleibt wie sie ist


def test_leerzeichen_im_gestrichenen_text_wird_mit_gestrichen(tmp_path):
    """Die Excel hat "Kel ler" (Leerzeichen im Wort); das PDF streicht "Keller". Der ganze Excel-Text der Stelle wird gestrichen."""
    pdf = [[("Die Anlage gilt fuer alle Pumpen im ", "k"), ("Kel", "d"), ("ler", "d"), ("Hof", "n"), (" und im Werk.", "k")]]
    plan, frei, auto = _lauf(tmp_path, pdf, "Die Anlage gilt fuer alle Pumpen im Kel ler und im Werk.")
    a = _genau_ein_vorschlag(plan)
    assert "a1_leerraum" in a.pruefcodes
    assert frei == [("normal", "Die Anlage gilt fuer alle Pumpen im "), ("alt", "Kel ler"), ("neu", "Hof"), ("normal", " und im Werk.")]
    assert auto == [("normal", "Die Anlage gilt fuer alle Pumpen im Kel ler und im Werk.")]


def test_einheit_und_abkuerzung_ohne_leerzeichen_bleiben_unveraendert(tmp_path):
    """Einheit ("25mm") und Abkuerzung ("z.B.") ohne Leerzeichen weit weg von der Aenderung: haeufige Abweichungen in echten
    Tabellen. Nur Leerzeichen sind anders -> a1_leerraum, Vorschlag, die Excel behaelt ihre Schreibweise."""
    excel = "Die Nennweite von 25mm gilt z.B. fuer alle Pumpen im Keller und im Werk."
    pdf = [[("Die Nennweite von 25 mm gilt z. B. fuer alle Pumpen im ", "k"), ("Keller", "d"), ("Hof", "n"), (" und im Werk.", "k")]]
    plan, frei, auto = _lauf(tmp_path, pdf, excel)
    a = _genau_ein_vorschlag(plan)
    assert "a1_leerraum" in a.pruefcodes
    assert a.grund_einfach == "Die Excel weicht hier vom alten PDF-Text ab (nur Leerzeichen). Vorschlag bitte prüfen."
    assert frei == [("normal", "Die Nennweite von 25mm gilt z.B. fuer alle Pumpen im "), ("alt", "Keller"), ("neu", "Hof"),
                    ("normal", " und im Werk.")]
    assert auto == [("normal", excel)]


def test_leerzeichen_im_gestrichenen_wert_wird_mit_gestrichen(tmp_path):
    """Das PDF streicht "25 mm" und schreibt "30 mm"; die Excel hat "25mm". Gestrichen wird die Stelle der Excel ("25mm")."""
    pdf = [[("Die Nennweite von ", "k"), ("25 mm", "d"), ("30 mm", "n"), (" gilt fuer alle Pumpen im Keller und im Werk.", "k")]]
    plan, frei, auto = _lauf(tmp_path, pdf, "Die Nennweite von 25mm gilt fuer alle Pumpen im Keller und im Werk.")
    a = _genau_ein_vorschlag(plan)
    assert "a1_leerraum" in a.pruefcodes
    assert frei == [("normal", "Die Nennweite von "), ("alt", "25mm"), ("neu", "30 mm"),
                    ("normal", " gilt fuer alle Pumpen im Keller und im Werk.")]
    assert auto == [("normal", "Die Nennweite von 25mm gilt fuer alle Pumpen im Keller und im Werk.")]


@pytest.mark.parametrize("excel", [
    "Die Anlage gilt fuer alle Pumpen im Kelle und im Werk.",      # Buchstabe fehlt weit vor der Stelle (a1_buchstabe)
    "Die Anlage gilt fuer alle Pumpen imKeller und im Werk.",      # Leerzeichen fehlt weit vor der Stelle (a1_leerraum)
])
def test_gestrichene_wortgruppe_vor_dem_satzpunkt_nimmt_ihr_leerzeichen_mit(tmp_path, excel):
    """Word streicht das Leerzeichen vor der gestrichenen Wortgruppe mit ("Keller{- und im Werk-}."). Der Vorschlag streicht es
    auch; frueher blieb es unmarkiert stehen ("Keller ." in der Neu-Ansicht), weil der PDF-Absatz fuer die Markierung wegen des
    Leerzeichens vor dem Punkt verworfen wurde."""
    pdf = [[("Die Anlage gilt fuer alle Pumpen im Keller", "k"), (" und im Werk", "d"), (".", "k")]]
    plan, frei, auto = _lauf(tmp_path, pdf, excel)
    a = _genau_ein_vorschlag(plan)
    assert len(plan.absaetze) == 1 and a.id in plan.absaetze[0].aenderungen
    vorn = excel[:excel.index(" und im Werk")]
    assert frei == [("normal", vorn), ("alt", " und im Werk"), ("normal", ".")]
    assert auto == [("normal", excel)]


def test_genau_gleicher_text_mit_anderem_polster_bleibt_automatisch(tmp_path):
    """Gegenprobe: Tab, geschuetztes und doppeltes Leerzeichen sind in der Normalform GLEICH -- das ist keine Abweichung, die
    Aenderung wird wie bisher automatisch eingetragen."""
    plan, frei, auto = _lauf(tmp_path, ERSETZT, "Die Anlage gilt fuer\talle Pumpen  im Keller und im Werk.")
    (a,) = plan.aenderungen
    assert a.status == UEBERNOMMEN and not a.pruefcodes
    assert auto == [("normal", "Die Anlage gilt fuer\talle Pumpen  im "), ("alt", "Keller"), ("neu", "Hof"), ("normal", " und im Werk.")]


# ------------------------------------------------------------------------------------------------ Listennummer + Polster
def _lauf_liste(tmp_path, label, zeilen_pdf, excel_text, name="l"):
    """Wie `_lauf`, aber der Absatz ist ein Listenpunkt: Listennummer (`label`) und Text im haengenden Einzug."""
    d = WordDok()
    d.ueberschrift([("1.1", "k")], [("ANLAGE", "k")])
    d.listenpunkt(1, label, zeilen_pdf)
    pdf = d.speichern(tmp_path / f"{name}.pdf")
    alt = schreibe_excel(tmp_path / f"{name}.xlsx", [["Nr.", "Titel", "Inhalt"], ["1.1", "ANLAGE", None], ["", "", excel_text]])
    plan = plane(analysiere_pdf(pdf), ExcelModell(alt))
    frei = tmp_path / f"{name}_frei.xlsx"
    uebernehmen(copy.deepcopy(plan.aenderungen), alt, frei, {a.id for a in plan.aenderungen if a.status == PRUEFEN},
                absaetze=plan.absaetze)
    return plan, messen._stuecke(messen._lese_excel(frei, True)[0][2][2])


TEXT_LISTE = "Die Anlage gilt fuer alle Pumpen im Keller und im Werk."
NEUES_WORT_AM_ANFANG = [[("Alle", "n"), (" " + TEXT_LISTE, "k")]]


@pytest.mark.parametrize("polster", ["    ", "  ", "       "])
@pytest.mark.parametrize("abweichung", [
    ("Pumpen", "Pump", "a1_buchstabe"),                 # Buchstabe fehlt vor einem Leerzeichen (frueher: "Leerzeichen")
    ("Anlage gilt", "Anlagegilt", "a1_leerraum"),       # Leerzeichen fehlt weit weg von der Aenderung
    ("Die Anlage", "DieAnlage", "a1_leerraum"),         # Leerzeichen fehlt direkt hinter dem ersten Wort
], ids=["buchstabe", "leerraum_fern", "leerraum_nah"])
def test_neues_erstes_wort_im_listenpunkt_steht_hinter_dem_ganzen_polster(tmp_path, polster, abweichung):
    """Das neue Wort steht in der Liste hinter dem Polster der Listennummer ("a." + geschuetzte/doppelte Leerzeichen), nicht
    mitten darin. Wie bei einer Excel OHNE Abweichung bleibt das Polster der Excel ganz erhalten (Nico traegt das neue Wort
    dahinter ein). Frueher schrieb der Vorschlag das Wort zwischen die Leerzeichen des Polsters und verkuerzte es."""
    von, nach, code = abweichung
    plan, frei = _lauf_liste(tmp_path, [("a.", "k")], NEUES_WORT_AM_ANFANG, "a." + polster + TEXT_LISTE.replace(von, nach))
    a = _genau_ein_vorschlag(plan)
    assert code in a.pruefcodes
    assert len(plan.absaetze) == 1 and a.id in plan.absaetze[0].aenderungen       # PDF-Absatz fuer die Markierung ist angelegt
    assert frei == [("normal", "a." + polster), ("neu", "Alle"), ("normal", " " + TEXT_LISTE.replace(von, nach))]


def test_neues_erstes_wort_im_listenpunkt_mit_tab_als_polster(tmp_path):
    plan, frei = _lauf_liste(tmp_path, [("a.", "k")], NEUES_WORT_AM_ANFANG, "a.\t" + TEXT_LISTE.replace("Pumpen", "Pump"))
    a = _genau_ein_vorschlag(plan)
    assert "a1_buchstabe" in a.pruefcodes
    assert frei == [("normal", "a.\t"), ("neu", "Alle"), ("normal", " " + TEXT_LISTE.replace("Pumpen", "Pump"))]


def test_listenpunkt_ohne_abweichung_ergibt_dasselbe_bild_automatisch(tmp_path):
    """Gegenstueck: mit genau passender Excel ist die Aenderung automatisch und sieht genauso aus."""
    polster = "    "
    plan, frei = _lauf_liste(tmp_path, [("a.", "k")], NEUES_WORT_AM_ANFANG, "a." + polster + TEXT_LISTE)
    (a,) = plan.aenderungen
    assert a.status == UEBERNOMMEN
    assert frei == [("normal", "a." + polster), ("neu", "Alle"), ("normal", " " + TEXT_LISTE)]
