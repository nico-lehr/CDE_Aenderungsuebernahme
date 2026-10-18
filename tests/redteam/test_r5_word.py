"""Red-Team R5: Grenzfaelle beim Lesen von Word-PDFs mit Aenderungsverfolgung (ADR 0006). Erfundene Daten (Regel 16).

Pruefregel auf Ebene des PDF-Lesers: jedes Wort ist RICHTIG gelesen ODER der Absatz traegt einen sperrenden Hinweis
(config zuordnung.yaml planer.unsichere_hinweise) bzw. das PDF eine Luecke -- nie still falsch.
Strikte xfails nur fuer bewusst ungeloeste Faelle (Begruendung im Test).
"""
from __future__ import annotations

import pytest

from src import config
from src.models import ADD, DEL, KEEP
from src.pdf_reader import analysiere_pdf
from tests._word_pdf import EBENEN, WORT_BLAU, WORT_GRUEN, WORT_ROT, ZEILE, WordDok, breite, fuellzeilen, voll
from tests.redteam._bau_r4 import aufbau, scheitern

pytestmark = [pytest.mark.legende_neu, pytest.mark.redteam]

SPERR = [str(c) for c in config.lade("zuordnung")["planer"]["unsichere_hinweise"]]


def _lies(tmp_path, d: WordDok):
    return analysiere_pdf(d.speichern(tmp_path / "rt.pdf"))


def _toks(a):
    return [t for b in a.blocks for t in b.tokens]


def _tok(a, text: str):
    t = next((t for t in _toks(a) if text in (t.old, t.new)), None)
    aufbau(t is not None, f"Wort {text!r} nicht gefunden")
    return t


def _gesperrt(a, text: str) -> bool:
    """Traegt der Absatz mit `text` einen sperrenden Hinweis?"""
    b = next(b for b in a.blocks if any(text in (t.old, t.new) for t in b.tokens))
    codes = {c for t in b.tokens for c in t.hinweis.split(";") if c}
    return any(c == s or c.startswith(s + "_") for c in codes for s in SPERR)


def richtig_oder_gesperrt(a, erwartet: dict[str, str], was: str) -> str:
    """erwartet: Wort -> Status. 'richtig', 'gesperrt' oder AssertionError (still falsch)."""
    falsch = {w: (_tok(a, w).status, s) for w, s in erwartet.items() if _tok(a, w).status != s}
    if not falsch:
        return "richtig"
    if all(_gesperrt(a, w) for w in falsch):
        return "gesperrt"
    scheitern(f"{was}: still falsch gelesen {falsch}")
    return ""


# ---------------------------------------------------------------------------------------------
# RT-300: gestrichelte Verbindungslinie, die VOR der Leiste endet, quer durch Text auf der Grundlinie
# ---------------------------------------------------------------------------------------------
def test_rt300_verbindungslinie_endet_vor_der_leiste(tmp_path):
    d = WordDok()
    top = d.top
    d.absatz([voll([("Pumpe ", "k"), ("Ventil", "d"), (" Becken Zebrafink Speicher", "k")])] + fuellzeilen(2))
    d.verbindungslinie(EBENEN[1] + 3, top, bis_x=405.0)
    a = _lies(tmp_path, d)
    richtig_oder_gesperrt(a, {"Pumpe": KEEP, "Ventil": DEL, "Becken": KEEP, "Zebrafink": KEEP}, "RT-300")
    assert not {c for t in _toks(a) for c in t.hinweis.split(";")} & {"unterstrichen", "linie_unklar"}


# ---------------------------------------------------------------------------------------------
# RT-301: Ankerpunkt einer Verbindungslinie mitten auf einem gestrichenen Wort
# ---------------------------------------------------------------------------------------------
def test_rt301_ankerpunkt_auf_gestrichenem_wort(tmp_path):
    d = WordDok()
    tops = d.absatz([voll([("Pumpe ", "k"), ("Ventilgruppe", "d"), (" Becken", "k")])] + fuellzeilen(1))
    d.sprechblase(tops[0] + 20, "hat formatiert:", "Nicht Hervorheben",
                  anker=(EBENEN[1] + breite("Pumpe Ventil"), tops[0]))
    a = _lies(tmp_path, d)
    assert richtig_oder_gesperrt(a, {"Ventilgruppe": DEL, "Pumpe": KEEP}, "RT-301") == "richtig"
    assert a.luecken == []


# ---------------------------------------------------------------------------------------------
# RT-302: Text in der Leiste exakt auf der Hoehe einer Fliesstextzeile (gleiche Grundlinie, gleiche Schrift)
# ---------------------------------------------------------------------------------------------
def test_rt302_randtext_auf_gleicher_zeilenhoehe_wird_nicht_angeklebt(tmp_path):
    d = WordDok()
    tops = d.absatz([voll([("Die ", "k"), ("Pumpe", "n"), (" ", "k")])] + fuellzeilen(3, 1))
    d.randtext(tops[1], "Formatiert: Schriftart Fett")
    a = _lies(tmp_path, d)
    assert len(a.blocks) == 1
    assert [r["art"] for r in a.randnotizen] == ["formatiert"]
    assert all(t.old not in ("Schriftart", "Fett") for t in _toks(a))


# ---------------------------------------------------------------------------------------------
# RT-303: zwei Autoren am selben Wort: blau gestrichen UND rot unterstrichen -> Widerspruch (sperrt)
# ---------------------------------------------------------------------------------------------
def test_rt303_zwei_autoren_strich_und_unterstreichung_am_selben_wort(tmp_path):
    d = WordDok()
    x, top = EBENEN[1], d.top
    d.zeile(x, top, voll([("Pumpe ", "k"), ("Ventil", "d", WORT_BLAU), (" Becken", "k")]))
    w0 = breite("Pumpe ")
    d.unterstreichung(x + w0, x + w0 + breite("Ventil"), top, WORT_ROT)
    d.top += ZEILE
    a = _lies(tmp_path, d)
    assert _gesperrt(a, "Ventil")
    assert "gemischt_widerspruch" in _tok(a, "Ventil").hinweis


# ---------------------------------------------------------------------------------------------
# RT-304: verschobener Text (Word: gruen doppelt gestrichen bzw. doppelt unterstrichen)
# ---------------------------------------------------------------------------------------------
def test_rt304_verschobener_text_doppelte_linien(tmp_path):
    d = WordDok()
    x, top = EBENEN[1], d.top
    d.zeile(x, top, voll([("Pumpe ", "k"), ("Ventil", "d", WORT_GRUEN), (" Becken ", "k"), ("Speicher", "n", WORT_GRUEN)]))
    w_v0, w_v = breite("Pumpe "), breite("Ventil")
    d.rechteck(x + w_v0, top + 0.42 * 8.4 - 0.2, x + w_v0 + w_v, top + 0.42 * 8.4 + 0.2, WORT_GRUEN)   # zweiter Strich
    w_s0, w_s = breite("Pumpe Ventil Becken "), breite("Speicher")
    d.rechteck(x + w_s0, top + 1.02 * 8.4 - 0.18, x + w_s0 + w_s, top + 1.02 * 8.4 + 0.18, WORT_GRUEN)  # zweite Linie
    d.top += ZEILE
    a = _lies(tmp_path, d)
    richtig_oder_gesperrt(a, {"Ventil": DEL, "Speicher": ADD, "Becken": KEEP}, "RT-304")


# ---------------------------------------------------------------------------------------------
# RT-305 (bewusst ungeloest): Hyperlink-Formatvorlage (blau unterstrichen) OHNE Link-Kommentar im Word-PDF
# ---------------------------------------------------------------------------------------------
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-305: Ohne Link-Kommentar ist ein blau unterstrichener Hyperlink von neuem Text eines zweiten "
                          "Autors nicht zu unterscheiden (Linie entscheidet). Die echten Beispiele haben keine Links.")
def test_rt305_hyperlink_formatvorlage_ohne_link_kommentar(tmp_path):
    d = WordDok()
    d.absatz([voll([("Siehe ", "k"), ("beispiel.invalid", "n", (0.02, 0.39, 0.76)), (" und ", "k"), ("Pumpe", "n")])])
    a = _lies(tmp_path, d)
    richtig_oder_gesperrt(a, {"beispiel.invalid": KEEP, "Pumpe": ADD}, "RT-305")


# ---------------------------------------------------------------------------------------------
# RT-306 (bewusst ungeloest): einseitiges Word-PDF -- die Kopfzeile kehrt nicht wieder und steht nicht im 7-%-Rand
# ---------------------------------------------------------------------------------------------
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-306: Kopf-/Fusszeilen werden ueber die Wiederkehr auf >= 50 % der Seiten erkannt; eine "
                          "einzelne Seite liefert dafuer keinen Beleg. Die Kopfzeile wird als Text gelesen (der Planer "
                          "findet sie nicht in der CDE: sichtbar 'nicht uebernommen').")
def test_rt306_einseitiges_dokument_kopfzeile(tmp_path):
    d = WordDok(mit_kopf_fuss=True)
    d.absatz([voll([("Seite ", "k"), ("Pumpe", "n")])] + fuellzeilen(2))
    a = _lies(tmp_path, d)
    if any(t.old == "Projekt" for t in _toks(a)):
        scheitern("RT-306: Kopfzeile als Absatztext gelesen")


# ---------------------------------------------------------------------------------------------
# RT-307: Markup-Leiste nur auf einer von drei Seiten -> Legende gilt fuer das ganze Dokument
# ---------------------------------------------------------------------------------------------
def test_rt307_leiste_nur_auf_einer_seite(tmp_path):
    d = WordDok(mit_leiste=False)
    d.absatz([voll([("Erste ", "k"), ("Pumpe", "n")])] + fuellzeilen(2))
    d.neue_seite(leiste=True)
    d.absatz([voll([("Zweite ", "k"), ("Ventil", "d")])] + fuellzeilen(2, 1))
    d.neue_seite(leiste=False)
    d.absatz([voll([("Dritte ", "k"), ("Becken", "f"), (" ", "k"), ("Speicher", "n")])] + fuellzeilen(2, 2))
    a = _lies(tmp_path, d)
    assert a.legende_modus == "linie"
    richtig_oder_gesperrt(a, {"Pumpe": ADD, "Ventil": DEL, "Becken": KEEP, "Speicher": ADD}, "RT-307")


# ---------------------------------------------------------------------------------------------
# RT-308: doppelte Buchstaben-Nummer geaendert (Word zaehlt nach "Z." mit "AA.", "BB." ...: "BB." gestrichen, "AA."
#         neu) und Nummer mit Klammer ("3)" -> "2)")
# ---------------------------------------------------------------------------------------------
@pytest.mark.parametrize("alt,neu,ebene", [("BB.", "AA.", 1), ("3)", "2)", 4)])
def test_rt308_geaenderte_nummern_anderer_art(tmp_path, alt, neu, ebene):
    d = WordDok()
    d.listenpunkt(ebene, [("Z.", "k")] if ebene == 1 else [("1)", "k")], fuellzeilen(1, 0, ebene + 1))
    top = d.top
    d.zeile(EBENEN[ebene], top, [(alt, "d", WORT_GRUEN), (neu, "n", WORT_GRUEN)])
    ende = EBENEN[ebene] + breite(alt + neu)
    tab = next(x for x in EBENEN + (245.0, 272.0) if x >= ende + 4)          # wie Word: Text zum naechsten Tabstopp
    d.zeile(tab, top, [("Der Kessel heizt das Wasser.", "k")])
    d.top += ZEILE
    a = _lies(tmp_path, d)
    assert len(a.blocks) == 2
    lab = a.blocks[1].tokens[0]
    assert (lab.old, lab.new) == (alt, neu)


# ---------------------------------------------------------------------------------------------
# RT-309: Section-Nummer und Titel geaendert ("1.7 VENTILE" -> "1.8 VENTILE UND BECKEN")
# ---------------------------------------------------------------------------------------------
def test_rt309_section_nummer_und_titel_geaendert(tmp_path):
    d = WordDok()
    d.ueberschrift([("1.7", "d"), ("1.8", "n")], [("VENTILE", "k"), (" UND BECKEN", "n")])
    d.listenpunkt(1, [("A.", "k")], fuellzeilen(2, 0, 2))
    a = _lies(tmp_path, d)
    kopf = a.blocks[0]
    assert kopf.kind == "ueberschrift" and (kopf.section_alt, kopf.section_neu) == ("1.7", "1.8")
    assert [t.status for t in kopf.tokens if t.old in ("UND", "BECKEN") or t.new in ("UND", "BECKEN")] == [ADD, ADD]


# ---------------------------------------------------------------------------------------------
# RT-310: farbiger Text einer frueheren Runde (dunkelblau, ohne Linie) direkt neben neuem Text
# ---------------------------------------------------------------------------------------------
def test_rt310_farbe_frueherer_runde_ohne_linie(tmp_path):
    d = WordDok()
    d.absatz([voll([("Pumpe ", "k"), ("Ventil Becken", "f", (0.12, 0.22, 0.39)), (" ", "k"), ("Speicher", "n")])])
    a = _lies(tmp_path, d)
    assert richtig_oder_gesperrt(a, {"Ventil": KEEP, "Becken": KEEP, "Speicher": ADD}, "RT-310") == "richtig"
    assert not _gesperrt(a, "Speicher")                     # die alte Farbe sperrt den Absatz nicht


# ---------------------------------------------------------------------------------------------
# RT-311: gelbe Hinterlegung hinter neuem, unterstrichenem Text
# ---------------------------------------------------------------------------------------------
def test_rt311_gelb_hinter_neuem_text(tmp_path):
    d = WordDok()
    x, top = EBENEN[1], d.top
    d.zeile(x, top, voll([("Wert ", "k"), ("[42]", "n"), (" bar", "k")]))
    d.gelb(x + breite("Wert "), x + breite("Wert [42]"), top)
    d.top += ZEILE
    a = _lies(tmp_path, d)
    assert richtig_oder_gesperrt(a, {"[42]": ADD, "Wert": KEEP}, "RT-311") == "richtig"
    assert a.luecken == []


# ---------------------------------------------------------------------------------------------
# RT-312: Kommentar-Hinterlegung (blasses Rot, flaechig) hinter Text ohne Linie
# ---------------------------------------------------------------------------------------------
def test_rt312_blasse_kommentar_hinterlegung(tmp_path):
    d = WordDok()
    x, top = EBENEN[1], d.top
    d.zeile(x, top, voll([("Pumpe ", "k"), ("Ventil Becken", "k"), (" ", "k"), ("Speicher", "n")]))
    d.rechteck(x + breite("Pumpe "), top - 0.5, x + breite("Pumpe Ventil Becken"), top + 9.0, (0.98, 0.86, 0.86), hinten=True)
    d.top += ZEILE
    a = _lies(tmp_path, d)
    assert richtig_oder_gesperrt(a, {"Ventil": KEEP, "Speicher": ADD}, "RT-312") == "richtig"
    assert a.luecken == []


# ---------------------------------------------------------------------------------------------
# RT-313 (bewusst ungeloest, ADR 0006): Word-Markup mit SCHWARZER Autorfarbe -- schwarze Unterstreichung = "normal"
# ---------------------------------------------------------------------------------------------
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-313: Laut ADR 0006 ist eine schwarze Unterstreichung 'normal' (nur Hinweis). Stellt jemand "
                          "die Autorfarbe in Word auf Schwarz, gehen neue Woerter so verloren. Die echten Beispiele sind farbig.")
def test_rt313_schwarze_autorfarbe(tmp_path):
    d = WordDok()
    d.absatz([voll([("Pumpe ", "k"), ("Ventil", "d", (0.0, 0.0, 0.0)), (" ", "k"), ("Becken", "u")])])
    a = _lies(tmp_path, d)
    richtig_oder_gesperrt(a, {"Ventil": DEL, "Becken": ADD}, "RT-313")
