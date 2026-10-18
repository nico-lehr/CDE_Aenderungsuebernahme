"""Excel-Markierung (src/markierung.py, ADR 0005): Zustand mit Markierung, Lesen, Schreiben, Rundreise. Nur erfundene Woerter."""
import random
import shutil
import subprocess

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from src.gegenprobe import normalform, zustand
from src.markierung import (ALT, NEU, NORMAL, Stueck, aktiv, ansichten, ansicht, faerbe_zelle, format_von_schrift,
                            schreibe, zustand_markiert)
from src.models import Lauf
from tests.test_gegenprobe import L


def stuecke(spec: str, angewendet=("D", "N")):
    return [(s.format, s.text) for s in zustand_markiert(L(spec), set(angewendet))]


def test_standard_ist_farbe_in_der_konfiguration_aber_tests_laufen_ohne():
    assert not aktiv()          # conftest schaltet fuer die frueheren Tests ab


@pytest.mark.markierung_an
def test_markierung_an_marker_nutzt_den_standard():
    assert aktiv()


def test_ersetzung_alt_direkt_vor_neu():
    assert stuecke("darf {-12-} {+7+} ms") == [(NORMAL, "darf "), (ALT, "12"), (NORMAL, " "), (NEU, "7"), (NORMAL, " ms")]


def test_verschmolzene_ersetzung_ohne_trenner_wie_im_pdf():
    # Nico, 18.10.: Alt und Neu stehen genau so hintereinander wie im PDF (kein zusaetzliches Leerzeichen)
    assert stuecke("x {-alt-}|{+neu+} y") == [(NORMAL, "x "), (ALT, "alt"), (NEU, "neu"), (NORMAL, " y")]


def test_konfigurierter_trenner_wird_eingefuegt():
    cfg = {"markierung": {"aenderungen": "farbe", "trenner": " "}}
    st = [(s.format, s.text) for s in zustand_markiert(L("x {-alt-}|{+neu+} y"), {"D", "N"}, cfg)]
    assert st == [(NORMAL, "x "), (ALT, "alt"), (NORMAL, " "), (NEU, "neu"), (NORMAL, " y")]


def test_einfuegung_rot_loeschung_blau_durchgestrichen_stehen_lassen():
    assert stuecke("Die Anlage {+insgesamt+}|.", {"N"}) == [(NORMAL, "Die Anlage "), (NEU, "insgesamt"), (NORMAL, ".")]
    assert stuecke("Die Anlage {-insgesamt-}|.", {"D"}) == [(NORMAL, "Die Anlage "), (ALT, "insgesamt"), (NORMAL, ".")]


def test_nicht_angewendet_alt_bleibt_normal_neu_fehlt():
    assert stuecke("a {-b-} {+c+} d", set()) == [(NORMAL, "a b d")]


def test_ansichten_alt_ist_das_original_neu_ist_der_neue_absatz():
    laeufe = L("Die {-alte-} {+neue+} Pumpe {-und-} {+oder+} Leitung.")
    st = zustand_markiert(laeufe, {"D", "N"})
    assert normalform(ansicht(st, ALT))[0] == normalform(zustand(laeufe, set()))[0]
    assert normalform(ansicht(st, NEU))[0] == normalform(zustand(laeufe, {"D", "N"}))[0]


def test_zufall_beide_ansichten_stimmen_fuer_jede_teilmenge():
    rnd = random.Random(11)
    woerter = ["Ventil", "Pumpe", "Druck", "Wert", "Anlage", "Prüfung", "Leitung"]
    for _ in range(300):
        laeufe: list[Lauf] = []
        for i in range(rnd.randint(1, 9)):
            if laeufe:
                laeufe.append(Lauf("s", rnd.choice([" ", " ", " "])))
            art = rnd.choice("kkkda")
            laeufe.append(Lauf(art, rnd.choice(woerter), f"A{rnd.randint(1, 3)}" if art in "da" else ""))
        ids = sorted({l.aenderung for l in laeufe if l.aenderung})
        for k in range(len(ids) + 1):
            S = set(rnd.sample(ids, k))
            st = zustand_markiert(laeufe, S)
            # neue Ansicht == Absatz laut PDF fuer S; alte Ansicht == ganzer alter Absatz (nichts geht verloren)
            assert normalform(ansicht(st, NEU))[0] == normalform(zustand(laeufe, S))[0]
            assert normalform(ansicht(st, ALT))[0] == normalform(zustand(laeufe, set()))[0]


def test_format_von_schrift_strich_ist_alt_jede_farbe_rot_ohne_strich_ist_neu():
    assert format_von_schrift(Font(strike=True, color="FF000000")) == ALT
    assert format_von_schrift(Font(strike=True, color="FFFF0000")) == ALT
    assert format_von_schrift(Font(color="FFFF0000")) == NEU
    assert format_von_schrift(Font(color="FFC00000")) == NEU
    assert format_von_schrift(Font(color="FF0000FF")) == NORMAL      # blau ohne Strich ist in der Excel nichts Besonderes
    assert format_von_schrift(Font(color="FF000000")) == NORMAL
    assert format_von_schrift(None) == NORMAL


def _zelle(ws="A1", **kw):
    wb = Workbook()
    z = wb.active[ws]
    z.font = Font(name="Arial", sz=12, b=kw.get("fett", False))
    z.fill = PatternFill("solid", fgColor="EEEEEE")
    z.alignment = Alignment(wrap_text=True, vertical="top")
    return wb, z


def test_schreiben_rundreise_behaelt_schrift_fuellung_und_formate(tmp_path):
    wb, z = _zelle(fett=True)
    st = [Stueck(NORMAL, "Die Anlage "), Stueck(ALT, "alt"), Stueck(NORMAL, " "), Stueck(NEU, "neu"), Stueck(NORMAL, " läuft.")]
    schreibe(z, st)
    p = tmp_path / "m.xlsx"
    wb.save(p)
    z2 = load_workbook(p, rich_text=True).active["A1"]
    alt, neu, zs = ansichten(z2)
    assert alt == "Die Anlage alt  läuft." and neu == "Die Anlage  neu läuft."
    assert [s.format for s in zs] == [NORMAL, ALT, NORMAL, NEU, NORMAL]
    assert z2.fill.fgColor.rgb.endswith("EEEEEE")
    blocke = [b for b in z2.value]
    assert all(getattr(b.font, "rFont", "Arial") == "Arial" and getattr(b.font, "sz", 12) == 12 for b in blocke if hasattr(b, "font"))
    assert blocke[1].font.strike and blocke[1].font.color.rgb.endswith("00B0F0")
    assert blocke[3].font.color.rgb.endswith("FF0000") and not blocke[3].font.strike


def test_ganze_zelle_ein_format_wird_zellschrift_kein_rich_text(tmp_path):
    wb, z = _zelle()
    schreibe(z, [Stueck(NEU, "Ganz neue Zeile")])
    assert z.value == "Ganz neue Zeile" and z.font.color.rgb.endswith("FF0000") and z.font.name == "Arial"
    wb2, z2 = _zelle("B2")
    schreibe(z2, [Stueck(ALT, "Ganz weg")])
    assert z2.font.strike and z2.font.color.rgb.endswith("00B0F0")
    assert ansichten(z2)[:2] == ("Ganz weg", "")


def test_faerbe_zelle_zahl_bleibt_zahl():
    wb, z = _zelle()
    z.value = 2.1
    z.number_format = "0.00"
    faerbe_zelle(z, ALT)
    assert z.value == 2.1 and z.number_format == "0.00" and z.font.strike


def test_leere_und_normale_zelle():
    wb, z = _zelle()
    schreibe(z, [])
    assert z.value is None
    schreibe(z, [Stueck(NORMAL, "nur "), Stueck(NORMAL, "Text")])
    assert z.value == "nur Text" and not z.font.strike


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice nicht installiert")
def test_libreoffice_zeigt_rot_blau_und_strich(tmp_path):
    """Beweis, dass Excel-Programme die Markierung wirklich so darstellen: xlsx -> PDF -> Farbe/Strich je Zeichen."""
    import pdfplumber
    wb, z = _zelle()
    schreibe(z, [Stueck(NORMAL, "Die "), Stueck(ALT, "alt"), Stueck(NORMAL, " "), Stueck(NEU, "neu"), Stueck(NORMAL, " Pumpe.")])
    wb.save(tmp_path / "m.xlsx")
    profil = tmp_path / "lo"
    subprocess.run(["soffice", "--headless", f"-env:UserInstallation=file://{profil}", "--convert-to", "pdf",
                    "--outdir", str(tmp_path), str(tmp_path / "m.xlsx")], check=True, capture_output=True, timeout=180)
    with pdfplumber.open(tmp_path / "m.pdf") as pdf:
        seite = pdf.pages[0]
        farben = {c["text"]: c["non_stroking_color"] for c in seite.chars if c["text"] in "altneu"}
        assert tuple(round(v, 2) for v in farben["a"]) == (0.0, 0.69, 0.94)   # alt = Hellblau 00B0F0
        assert tuple(farben["n"]) == (1.0, 0.0, 0.0)      # neu = rot
        assert len(seite.lines) + len(seite.rects) >= 1   # Durchstreichung
