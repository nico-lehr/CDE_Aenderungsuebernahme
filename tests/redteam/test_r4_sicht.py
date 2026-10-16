"""Red-Team R4, Teil SICHT: Markierungen, die pdfplumber nicht als Farbe sieht (erfundene Daten, Regel 16).

Separation-Flaeche mit Ton 1,0, DeviceN-Text mit zwei Kanaelen, Markierung als Kurven-Pfad, Nur-Kontur-Text.
Jedes Test-PDF wird im Aufbau mit pdfium gerendert: die Stelle muss wirklich blau bzw. rot AUSSEHEN.
Zu jedem Fall gibt es eine Kontrolle (gleiche Markierung in DeviceRGB bzw. als Rechteck / gefuellter Text): sie
besteht heute und zeigt, dass der Testaufbau stimmt. Im Dokument steht zusaetzlich eine normale blaue Aenderung
("ebenfalls"), damit der Lauf nicht nur wegen "keine Aenderungen erkannt" auffaellt.
"""
import pytest

from tests.redteam._bau_r4 import (GROESSE, LINKS, RohDok, aufbau, breite, ende_zu_ende, erkannt_oder_gemeldet,
                                   excel_sections, funktion_ps, separation, wirkt)

VORN = "Die Steuerung überwacht die Anlage "
WORT = "täglich"
HINTEN = " und meldet alle Werte."
MIT = VORN + WORT + HINTEN
OHNE = "Die Steuerung überwacht die Anlage und meldet alle Werte."
Z2 = "Zweiter Absatz mit genug Text darin für den Test."
Z2_VORN, Z2_HINTEN = "Zweiter Absatz mit genug Text darin ", "für den Test."


def _dok(stueck, ressourcen: str = "", objekte=None) -> tuple[RohDok, tuple[float, float, float, float]]:
    d = RohDok(ressourcen=ressourcen, objekte=list(objekte or []))
    d.ueberschrift("1 Eins")
    top = d.top
    d.absatz([[(VORN, "k"), stueck, (HINTEN, "k")]])
    x0 = LINKS + breite(VORN, GROESSE)
    kasten = (x0 + 1, top + 1, x0 + breite(WORT, GROESSE) - 1, top + GROESSE - 1)
    d.ueberschrift("2 Zwei")
    d.absatz([[(Z2_VORN, "k"), ("ebenfalls ", "0 0 1 rg"), (Z2_HINTEN, "k")]])
    return d, kasten


def _lauf(tmp_path, dok, excel_text):
    alt = excel_sections(tmp_path, [("1", "Eins", [excel_text]), ("2", "Zwei", [Z2])])
    l = ende_zu_ende(tmp_path, dok, alt)
    aufbau(any("ebenfalls" in a.neu for a in l.aen), f"die normale blaue Aenderung muss erkannt sein: {l.kurz()}")
    return l


def _fremd(a):
    return "ebenfalls" in a.neu


def _passt(a):
    return WORT in a.alt or WORT in a.neu


XF = pytest.mark.xfail


# ---------------------------------------------------------------------------------------------
# RT-207: Hinterlegung als Flaeche in Separation-Farbe "Rot" mit Ton 1,0 (wirkt rot -> entfaellt)
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
@pytest.mark.parametrize("art", [
    pytest.param("kontrolle_rgb", id="kontrolle_rgb"),
    pytest.param("separation", id="separation", marks=XF(strict=True, raises=AssertionError,
                 reason="RT-207 rote Flaeche in Separation-Farbe (Ton 1,0) wird nicht erkannt und nicht gemeldet; Besitzer SICHT")),
])
def test_rt207_hinterlegung_in_separation_farbe(tmp_path, art):
    if art == "separation":
        dok, kasten = _dok((WORT, ("flaeche", "/Rot cs 1 scn", "re")),
                           "/ColorSpace << " + separation("Rot", (0, 0.3, 0.2, 0)) + " >>")
    else:
        dok, kasten = _dok((WORT, ("flaeche", "1 0.75 0.8 rg", "re")))
    aufbau(wirkt(dok, *kasten) == "rot", f"Hinterlegung muss rot aussehen, ist {wirkt(dok, *kasten)}")
    l = _lauf(tmp_path, dok, MIT)
    erkannt_oder_gemeldet(l, 2, 3, OHNE, MIT, "RT-207", passt=_passt, fremd=_fremd)


# ---------------------------------------------------------------------------------------------
# RT-208: Text in DeviceN-Farbe mit zwei Farbkanaelen (Cyan + Magenta -> wirkt blau -> neu)
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
@pytest.mark.parametrize("art", [
    pytest.param("kontrolle_cmyk", id="kontrolle_cmyk"),
    pytest.param("devicen", id="devicen", marks=XF(strict=True, raises=AssertionError,
                 reason="RT-208 blauer DeviceN-Text (2 Kanaele) gilt still als unveraendert; Besitzer SICHT")),
])
def test_rt208_text_in_devicen_farbe(tmp_path, art):
    if art == "devicen":
        dok, kasten = _dok((WORT, "/Bl cs 1 0.7 scn"),
                           "/ColorSpace << /Bl [/DeviceN [/Cyan /Magenta] /DeviceCMYK {OBJ0} 0 R] >>",
                           [funktion_ps("{ 0 0 }", 2, 4)])        # (c, m) -> (c, m, 0, 0)
    else:
        dok, kasten = _dok((WORT, "1 0.7 0 0 k"))
    aufbau(wirkt(dok, *kasten) == "blau", f"Text muss blau aussehen, ist {wirkt(dok, *kasten)}")
    l = _lauf(tmp_path, dok, OHNE)
    erkannt_oder_gemeldet(l, 2, 3, MIT, OHNE, "RT-208", passt=_passt, fremd=_fremd)


# ---------------------------------------------------------------------------------------------
# RT-209: Markierung als Kurven-/Pfad-Flaeche (abgerundetes Feld aus Bezier-Kurven, kein Rechteck) hinter Text
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
@pytest.mark.parametrize("form", [
    pytest.param("re", id="kontrolle_rechteck"),
    pytest.param("kurve", id="kurve", marks=XF(strict=True, raises=AssertionError,
                 reason="RT-209 blaue Hinterlegung als Kurven-Pfad gilt still als unveraendert; Besitzer SICHT")),
])
def test_rt209_markierung_als_kurvenflaeche(tmp_path, form):
    dok, kasten = _dok((WORT, ("flaeche", "0.72 0.86 1 rg", form)))
    aufbau(wirkt(dok, *kasten) == "blau", f"Hinterlegung muss blau aussehen, ist {wirkt(dok, *kasten)}")
    l = _lauf(tmp_path, dok, OHNE)
    erkannt_oder_gemeldet(l, 2, 3, MIT, OHNE, "RT-209", passt=_passt, fremd=_fremd)


# ---------------------------------------------------------------------------------------------
# RT-210: Nur-Kontur-Text (Render-Modus 1) mit blauer Kontur, Fuellfarbe bleibt schwarz
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
@pytest.mark.parametrize("modus", [
    pytest.param("kontrolle_fuellung", id="kontrolle_fuellung"),
    pytest.param("kontur", id="kontur", marks=XF(strict=True, raises=AssertionError,
                 reason="RT-210 Text nur mit blauer Kontur (Tr 1) gilt still als unveraendert; Besitzer SICHT")),
])
def test_rt210_nur_kontur_text(tmp_path, modus):
    befehl = "0 g 0 0 1 RG 0.3 w 1 Tr" if modus == "kontur" else "0 0 1 rg"
    dok, kasten = _dok((WORT, befehl))
    aufbau(wirkt(dok, *kasten) == "blau", f"Text muss blau aussehen, ist {wirkt(dok, *kasten)}")
    l = _lauf(tmp_path, dok, OHNE)
    erkannt_oder_gemeldet(l, 2, 3, MIT, OHNE, "RT-210", passt=_passt, fremd=_fremd)
