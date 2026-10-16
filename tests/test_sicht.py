"""SICHT (src/sicht.py): zweiter Leser (C1) und Farb-Inventur (C3). Nur erfundene Daten (Regel 16).

Je Regel ein Fall, der ohne SICHT still verloren ginge, und je ein Fehlalarm-Fall (Link, blaue Design-Ueberschriften,
Bild/Logo, Aenderungsbalken, Tabellenkopf, dunkles Farbband), bei dem SICHT NICHTS melden darf.
Der Aufbau jedes Falls wird mit pdfium gerendert geprueft (`wirkt`): die Stelle muss wirklich so aussehen.
"""
from __future__ import annotations

import copy

import pytest

from src import config
from src.models import ADD, KEEP
from src.pdf_reader import _Z, analysiere_pdf, hinweis_text
from src.sicht import entscheide
from tests._pdf_roh import SichtDok
from tests.fixtures.pdf_writer import breite
from tests.redteam._bau_r4 import GROESSE, LINKS, RECHTS, funktion_ps, wirkt

VORN = "Die Steuerung überwacht die Anlage "
WORT = "täglich"
HINTEN = " und meldet alle Werte."
SICHT_CODES = ("farbe_uneinig", "tinte_unerklaert")


def _cfg(aktiv: bool = True) -> dict:
    cfg = copy.deepcopy(config.lade("farben"))
    cfg["sicht"]["aktiv"] = aktiv
    return cfg


def _dok(stueck=(WORT, "k"), ressourcen: str = "", objekte=None, zweiter=(("ebenfalls ", "0 0 1 rg"),)):
    """Zwei Abschnitte; im ersten steht `stueck` mitten im Satz. Gibt (Dokument, Kasten des Stuecks) zurueck."""
    d = SichtDok(ressourcen=ressourcen, objekte=list(objekte or []))
    d.ueberschrift("1 Eins")
    top = d.top
    d.absatz([[(VORN, "k"), stueck, (HINTEN, "k")]])
    x0 = LINKS + breite(VORN, GROESSE)
    kasten = (x0 + 1, top + 1, x0 + breite(stueck[0], GROESSE) - 1, top + GROESSE - 1)
    d.ueberschrift("2 Zwei")
    d.absatz([[("Zweiter Absatz mit genug Text darin ", "k"), *zweiter, ("für den Test.", "k")]])
    d.absatz([[("Dritter Absatz ohne jede Markierung, nur damit die Seite Text hat.", "k")]])
    return d, kasten


def _lies(tmp_path, dok, aktiv: bool = True, name: str = "a.pdf"):
    p = tmp_path / name
    p.write_bytes(dok.bytes())
    return analysiere_pdf(p, _cfg(aktiv))


def _token(analyse, text: str):
    return next(t for b in analyse.blocks for t in b.tokens if text in (t.old, t.new))


def _codes(analyse) -> set[str]:
    return {c for b in analyse.blocks for t in b.tokens for c in t.hinweis.split(";") if c}


def _unerklaert(analyse) -> list[dict]:
    return [l for l in analyse.luecken if l["art"] == "unerklaerte_markierung"]


def _still(analyse) -> None:
    """Fehlalarm-Pruefung: SICHT hat nichts gemeldet."""
    assert not (_codes(analyse) & set(SICHT_CODES)), _codes(analyse)
    assert not _unerklaert(analyse), _unerklaert(analyse)
    assert not [w for w in analyse.warnungen if "zweite" in w or "Inventur" in w], analyse.warnungen


# ------------------------------------------------------------------------------------------- Konfiguration
def test_sicht_ist_eingeschaltet_und_hat_texte():
    cfg = config.lade("farben")
    assert cfg["sicht"]["aktiv"] is True
    for code in SICHT_CODES:
        assert hinweis_text(code)


# ------------------------------------------------------------------------------------------- C1 zweiter Leser
def test_c1_devicen_text_mit_zwei_kanaelen_wird_blau(tmp_path):
    dok, kasten = _dok((WORT, "/Bl cs 1 0.7 scn"), "/ColorSpace << /Bl [/DeviceN [/Cyan /Magenta] /DeviceCMYK {OBJ0} 0 R] >>",
                       [funktion_ps("{ 0 0 }", 2, 4)])
    assert wirkt(dok, *kasten) == "blau"
    assert _token(_lies(tmp_path, dok, aktiv=False), WORT).status == KEEP      # ohne SICHT: still unveraendert
    a = _lies(tmp_path, dok)
    assert _token(a, WORT).status == ADD
    assert _token(a, "ebenfalls").status == ADD
    _still(a)


def test_c1_nur_kontur_text_wird_blau(tmp_path):
    dok, kasten = _dok((WORT, "0 g 0 0 1 RG 0.3 w 1 Tr"))
    assert wirkt(dok, *kasten) == "blau"
    assert _token(_lies(tmp_path, dok, aktiv=False), WORT).status == KEEP
    a = _lies(tmp_path, dok)
    assert _token(a, WORT).status == ADD
    _still(a)


def test_c1_fuellung_und_kontur_in_verschiedenen_farben_ist_uneinig(tmp_path):
    # Modus 2: schwarze Fuellung, kraeftige rote Kontur -> wirkt rot, pdfplumber sieht nur Schwarz
    dok, kasten = _dok((WORT, "0 g 1 0 0 RG 0.8 w 2 Tr"))
    assert wirkt(dok, *kasten) == "rot"
    a = _lies(tmp_path, dok)
    t = _token(a, WORT)
    assert "farbe_uneinig" in t.hinweis
    blk = next(b for b in a.blocks if t in b.tokens)
    assert hinweis_text("farbe_uneinig") in blk.hinweise


def _z(hex_: str = "#000000", klasse_: str | None = None, unsichtbar: bool = False) -> _Z:
    z = _Z(text="x", x0=0, x1=5, top=0, bottom=10, size=10, font="F")
    z.hex, z.farbklasse, z.unsichtbar = hex_, klasse_, unsichtbar
    z.farbe_bed = {"blau": "neu", "rot": "entfaellt"}.get(klasse_ or "")
    return z


BLAU, ROT, SCHWARZ, GRAU_BLAU = (0.0, 0.3, 0.9), (0.9, 0.1, 0.1), (0.0, 0.0, 0.0), (0.45, 0.47, 0.5)


@pytest.mark.parametrize("fall,z,fuell,kontur,modus,erg,klasse_nachher,hinweis", [
    ("keine_farbe_pdfplumber", _z(""), BLAU, SCHWARZ, 0, "uebernommen", "blau", False),
    ("sonderfarbe_pdfplumber", _z("sonderfarbe"), ROT, SCHWARZ, 0, "uebernommen", "rot", False),
    ("nur_kontur", _z(), SCHWARZ, BLAU, 1, "uebernommen", "blau", False),
    ("nur_kontur_clip", _z(), SCHWARZ, BLAU, 5, "uebernommen", "blau", False),
    ("pdfplumber_blau_pdfium_schwarz", _z("#0050e6", "blau"), SCHWARZ, SCHWARZ, 0, "uneinig", "blau", True),
    ("pdfplumber_schwarz_pdfium_rot", _z(), ROT, SCHWARZ, 0, "uneinig", None, True),
    ("pdfium_blass_ist_nicht_klar_unmarkiert", _z("#0050e6", "blau"), GRAU_BLAU, SCHWARZ, 0, "", "blau", False),
    ("beide_einig_blau", _z("#0050e6", "blau"), BLAU, SCHWARZ, 0, "", "blau", False),
    ("beide_einig_schwarz", _z(), SCHWARZ, SCHWARZ, 0, "", None, False),
    ("pdfium_unsichtbar", _z(""), BLAU, BLAU, 3, "", None, False),
    ("pdfplumber_unsichtbar_pdfium_blau", _z(unsichtbar=True), BLAU, SCHWARZ, 0, "uneinig", None, True),
    ("fuellung_und_kontur_verschieden", _z(), SCHWARZ, ROT, 2, "uneinig", None, True),
])
def test_c1_regeln(fall, z, fuell, kontur, modus, erg, klasse_nachher, hinweis):
    cfg = config.lade("farben")
    assert entscheide(z, fuell, kontur, modus, cfg) == erg, fall
    assert z.farbklasse == klasse_nachher, fall
    assert ("farbe_uneinig" in z.hinweis) == hinweis, fall
    if erg == "uebernommen":
        assert z.farbe_bed in ("neu", "entfaellt")


# ------------------------------------------------------------------------------------------- C3 Farb-Inventur
def test_c3_kurvenflaeche_in_absatz_ohne_aenderung_wird_luecke(tmp_path):
    dok, kasten = _dok((WORT, ("flaeche", "0.72 0.86 1 rg", "kurve")))
    assert wirkt(dok, *kasten) == "blau"
    assert not _unerklaert(_lies(tmp_path, dok, aktiv=False))                 # ohne SICHT: nichts gemeldet
    a = _lies(tmp_path, dok)
    assert "tinte_unerklaert" in _token(a, WORT).hinweis
    lue = _unerklaert(a)
    assert len(lue) == 1 and lue[0]["seite"] == 1 and lue[0]["woerter"] >= 1
    assert "Seite 1" in lue[0]["beschreibung"]
    assert "tinte_unerklaert" not in _token(a, "ebenfalls").hinweis          # andere Absaetze bleiben unberuehrt


def test_c3_kurvenflaeche_in_absatz_mit_aenderung_gibt_hinweis(tmp_path):
    d = SichtDok()
    d.ueberschrift("1 Eins")
    d.absatz([[(VORN, "k"), (WORT, ("flaeche", "1 0.75 0.8 rg", "kurve")), (" und meldet ", "k"), ("neue ", "0 0 1 rg"),
               ("Werte.", "k")]])
    d.absatz([[("Zweiter Absatz ohne jede Markierung, nur damit die Seite Text hat.", "k")]])
    a = _lies(tmp_path, d)
    assert "tinte_unerklaert" in _token(a, WORT).hinweis
    blk = next(b for b in a.blocks if _token(a, WORT) in b.tokens)
    assert blk.hat_aenderung and hinweis_text("tinte_unerklaert") in blk.hinweise
    assert not _unerklaert(a)                                                  # der Absatz wird ueber den Hinweis gesperrt


def test_c3_farbe_ohne_text_wird_luecke_ohne_woerter(tmp_path):
    dok, _ = _dok()
    dok.schattierung(300, 400, 340, 440, (0.9, 0.2, 0.1))                    # fuer pdfplumber unsichtbar
    assert wirkt(dok, 300, 400, 340, 440) == "rot"
    a = _lies(tmp_path, dok)
    lue = _unerklaert(a)
    assert len(lue) == 1 and lue[0]["woerter"] == 0
    assert not (_codes(a) & set(SICHT_CODES))


def test_c3_schattierung_hinter_text_gibt_hinweis(tmp_path):
    dok, kasten = _dok()
    dok.schattierung(kasten[0] - 1, kasten[1] - 2, kasten[2] + 1, kasten[3] + 2, (0.6, 0.8, 1.0))
    assert wirkt(dok, *kasten) == "blau"
    a = _lies(tmp_path, dok)
    assert "tinte_unerklaert" in _token(a, WORT).hinweis
    assert "tinte_unerklaert" not in _token(a, "Dritter").hinweis


def test_aus_schalter(tmp_path):
    dok, _ = _dok((WORT, ("flaeche", "0.72 0.86 1 rg", "kurve")))
    a = _lies(tmp_path, dok, aktiv=False)
    assert not (_codes(a) & set(SICHT_CODES)) and not _unerklaert(a)
    cfg = _cfg(True)
    cfg["sicht"]["inventur"]["aktiv"] = False
    p = tmp_path / "b.pdf"
    p.write_bytes(dok.bytes())
    assert not _unerklaert(analysiere_pdf(p, cfg))


# ------------------------------------------------------------------------------------------- Fehlalarme
def test_fehlalarm_hyperlink(tmp_path):
    d = SichtDok()
    d.ueberschrift("1 Eins")
    top = d.top
    d.absatz([[("Mehr dazu steht auf ", "k"), ("der Projektseite", "0 0 0.8 rg"), (" im Netz.", "k")]])
    x0 = LINKS + breite("Mehr dazu steht auf ", GROESSE)
    x1 = x0 + breite("der Projektseite", GROESSE)
    d.linie(x0, x1, top + GROESSE * 0.95, "0 0 0.8 RG", 0.6)                   # Unterstreichung des Links
    d.link(x0, top, x1, top + GROESSE)
    d.absatz([[("Zweiter Absatz mit ", "k"), ("neuem ", "0 0 1 rg"), ("Text.", "k")]])
    assert wirkt(d, x0, top, x1, top + GROESSE) == "blau"
    a = _lies(tmp_path, d)
    assert "link" in _token(a, "Projektseite").hinweis
    _still(a)


def test_fehlalarm_blaue_design_ueberschriften(tmp_path):
    d = SichtDok()
    for i in range(1, 5):
        d.top += 6
        d._text(LINKS, d.top, f"{i} Abschnitt Nummer {i}", 14.0, True, "0.1 0.3 0.8 rg")
        d.top += 26
        d.absatz([[(f"Gewöhnlicher Text im Abschnitt {i}, ganz ohne Markierung.", "k")]])
    a = _lies(tmp_path, d)
    assert "design_blau" in _codes(a)
    _still(a)


def test_fehlalarm_bild_und_logo(tmp_path):
    dok, _ = _dok()
    dok.bild(420, 90, 520, 150, (30, 90, 220))                               # farbiges Bild neben dem Text
    dok.bild(470, 20, 530, 45, (220, 40, 40), name="Im2")                    # Logo im Kopf
    assert wirkt(dok, 420, 90, 520, 150) == "blau"
    _still(_lies(tmp_path, dok))


def test_fehlalarm_aenderungsbalken(tmp_path):
    dok, kasten = _dok()
    # Balken als Linie (pdfplumber sieht sie) und als Schattierung (pdfplumber sieht sie nicht)
    dok.inhalt.append(f"q 0 0 1 RG 1.5 w 40 {842 - kasten[1] + 2:.2f} m 40 {842 - kasten[3] - 2:.2f} l S Q\n")
    dok.schattierung(30, kasten[1] - 2, 32, kasten[3] + 14, (0.0, 0.2, 0.9))
    assert wirkt(dok, 29, kasten[1], 33, kasten[3]) == "blau"
    _still(_lies(tmp_path, dok))


def test_fehlalarm_tabellenkopf_und_dunkles_farbband(tmp_path):
    d = SichtDok()
    d.ueberschrift("1 Eins")
    top = d.top
    d.flaeche(LINKS - 4, top - 6, RECHTS, top + GROESSE + 14, "0.55 0.75 1 rg")       # heller Kopf, viel Leerraum
    d.absatz([[("Spalte A", "k")]], abstand_danach=16)
    top2 = d.top
    d.flaeche(LINKS - 4, top2 - 4, RECHTS, top2 + GROESSE + 4, "0.05 0.1 0.35 rg")    # dunkles Band, weisse Schrift
    d.absatz([[("Hinweis im dunklen Band", "1 g")]], abstand_danach=16)
    d.absatz([[("Normaler Absatz mit ", "k"), ("neuem ", "0 0 1 rg"), ("Wort.", "k")]])
    assert wirkt(d, RECHTS - 40, top, RECHTS - 10, top + GROESSE) == "blau"
    _still(_lies(tmp_path, d))


def test_fehlalarm_farbiger_text_in_legendenfremder_farbe(tmp_path):
    dok, _ = _dok((WORT, "0 0.6 0 rg"))                                       # gruen: meldet der Leser selbst
    a = _lies(tmp_path, dok)
    assert any(l["art"] == "unbekannte_farbe" for l in a.luecken)
    _still(a)
