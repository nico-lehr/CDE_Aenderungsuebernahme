"""PDF-Leser fuer Word-Aenderungsverfolgung (ADR 0006, Auftrag 003): Legende "Linie entscheidet", Markup-Leiste mit
Sprechblasen, Kopf-/Fusszeilen mit eigenen Aenderungen, Listenpunkte und Section-Ueberschriften.

Alle Tests laufen mit der Standard-Legende aus config/farben.yaml (`legende_neu`, modus auto). Die Test-PDFs baut
tests/_word_pdf.py in genau den gemessenen Formen; nur erfundene Daten (Regel 16).
"""
from __future__ import annotations

import pytest

from src import config
from src.models import ADD, DEL, KEEP, MOD
from src.pdf_reader import _mit_modus, analysiere_pdf
from tests._word_pdf import EBENEN, WORT_BLAU, WORT_GRUEN, WORT_ROT, ZEILE, WordDok, breite, fuellzeilen, voll

pytestmark = pytest.mark.legende_neu

SPERR = [str(c) for c in config.lade("zuordnung")["planer"]["unsichere_hinweise"]]


def _lies(tmp_path, d: WordDok, cfg: dict | None = None, name: str = "word.pdf"):
    return analysiere_pdf(d.speichern(tmp_path / name), cfg)


def _toks(a):
    return [t for b in a.blocks for t in b.tokens]


def _tok(a, text: str):
    return next(t for t in _toks(a) if text in (t.old, t.new))


def _block(a, text: str):
    return next(b for b in a.blocks if any(text in (t.old, t.new) for t in b.tokens))


def _codes(t) -> set[str]:
    return set(filter(None, t.hinweis.split(";")))


def _sperrend(codes: set[str]) -> set[str]:
    return {c for c in codes if any(c == s or c.startswith(s + "_") for s in SPERR)}


def _alle_codes(a) -> set[str]:
    return set().union(*(_codes(t) for t in _toks(a)))


def _cfg(modus: str) -> dict:
    return _mit_modus(config.lade("farben"), modus)


# =================================================================================================
# Legende "linie": Status je Wort und Zeichen
# =================================================================================================
def test_unterstrichen_ist_neu_durchgestrichen_entfaellt_in_derselben_farbe(tmp_path):
    d = WordDok()
    d.absatz([voll([("Die ", "k"), ("Pumpe", "n"), (" ", "k"), ("Ventil", "d"), (" foerdert", "k")])])
    a = _lies(tmp_path, d)
    assert a.legende_modus == "linie" and a.markup_leiste
    assert _tok(a, "Pumpe").status == ADD
    assert _tok(a, "Ventil").status == DEL
    assert _tok(a, "foerdert").status == KEEP
    assert not _sperrend(_alle_codes(a))
    assert a.luecken == []


def test_klammern_einzeln_gestrichen_bleiben_zeichengenau(tmp_path):
    d = WordDok()
    d.absatz([voll([("Druck ", "k"), ("[", "d"), ("40", "k"), ("]", "d"), (" bar", "k")])])
    a = _lies(tmp_path, d)
    t = _tok(a, "[40]")
    assert (t.old, t.new, t.status) == ("[40]", "40", MOD)
    assert t.teile == [(DEL, "["), (KEEP, "40"), (DEL, "]")]
    assert not _sperrend(_codes(t))


def test_gestrichener_platzhalter_direkt_vor_neuem_wert_ist_ein_wort(tmp_path):
    d = WordDok()
    d.absatz([voll([("Hoehe ", "k"), ("[40]", "d"), ("42.9", "n"), (" Meter", "k")])])
    a = _lies(tmp_path, d)
    t = _tok(a, "[40]")
    assert (t.old, t.new, t.status) == ("[40]", "42.9", MOD)
    assert t.teile == [(DEL, "[40]"), (ADD, "42.9")]
    assert not _sperrend(_codes(t))


def test_einzelne_satzzeichen_werden_nicht_geglaettet(tmp_path):
    d = WordDok()
    d.absatz([voll([("Pumpe", "k"), (".", "n"), (" Ventil", "d"), (",", "k"), (" Zebrafink", "k"), (",", "d")])])
    a = _lies(tmp_path, d)
    p = _tok(a, "Pumpe")
    assert (p.old, p.new) == ("Pumpe", "Pumpe.")
    v = _tok(a, "Ventil,")
    assert (v.old, v.new, v.teile) == ("Ventil,", ",", [(DEL, "Ventil"), (KEEP, ",")])
    z = _tok(a, "Zebrafink,")
    assert (z.old, z.new) == ("Zebrafink,", "Zebrafink")
    assert not _sperrend(_alle_codes(a))


def test_zweiter_autor_in_blau_mit_linien(tmp_path):
    d = WordDok()
    d.absatz([voll([("Pumpe", "n", WORT_BLAU), (" ", "k"), ("Ventil", "d", WORT_BLAU), (" und ", "k"),
                    ("Zebrafink", "n", WORT_ROT)])])
    a = _lies(tmp_path, d)
    assert _tok(a, "Pumpe").status == ADD
    assert _tok(a, "Ventil").status == DEL
    assert _tok(a, "Zebrafink").status == ADD


def test_farbe_ohne_linie_bleibt_unveraendert_mit_hinweis_ohne_sperre(tmp_path):
    d = WordDok()
    d.absatz([voll([("Pumpe ", "k"), ("Ventil", "f"), (" Zebrafink", "k"), (" ", "k"), ("Becken", "n")])])
    a = _lies(tmp_path, d)
    v = _tok(a, "Ventil")
    assert v.status == KEEP and "farbe_ohne_linie" in _codes(v)
    assert not _sperrend(_codes(v))
    assert _tok(a, "Becken").status == ADD
    assert any("farbig, aber weder" in w for w in a.warnungen)
    assert not [l for l in a.luecken if l["art"] == "unbekannte_farbe"]


def test_schwarze_unterstreichung_ist_nur_hinweis(tmp_path):
    d = WordDok()
    d.absatz([voll([("Pumpe ", "k"), ("Ventil", "u"), (" ", "k"), ("Becken", "n")])])
    a = _lies(tmp_path, d)
    v = _tok(a, "Ventil")
    assert v.status == KEEP and "unterstrichen" in _codes(v)


def test_strich_und_farbige_unterstreichung_am_selben_wort_sperrt(tmp_path):
    d = WordDok()
    d.absatz([voll([("Pumpe ", "k"), ("Ventil", "nd"), (" ", "k"), ("Becken", "n")])])
    a = _lies(tmp_path, d)
    v = _tok(a, "Ventil")
    assert v.status == DEL
    assert "gemischt_widerspruch" in _sperrend(_codes(v))


@pytest.mark.parametrize("kuerzer", [1.2, 2.2], ids=["endet_im_zeichen", "verfehlt_das_zeichen"])
def test_linie_zu_kurz_ist_rand_unklar(tmp_path, kuerzer):
    """Die Unterstreichung endet mitten im letzten Buchstaben bzw. verfehlt ihn fast ganz: nie still lesen."""
    d = WordDok()
    x = EBENEN[1]
    d.zeile(x, d.top, voll([("Pumpe ", "k"), ("Ventil", "f"), (" Becken", "n")]))
    w0 = breite("Pumpe ")
    d.unterstreichung(x + w0, x + w0 + breite("Ventil") - kuerzer, d.top, WORT_ROT)     # "l" ist 2,3 pt breit
    d.top += ZEILE
    a = _lies(tmp_path, d)
    v = _tok(a, "Ventil")
    assert "rand_unklar" in _sperrend(_codes(v))


# =================================================================================================
# Leerzeichen als Aenderung (nie still verlieren)
# =================================================================================================
def _ansichten(block) -> tuple[str, str]:
    """(alter, neuer Absatz) aus den Laeufen, wie der Abgleich sie bildet (Leerraum zusammengefasst)."""
    from src.abgleich import laeufe_aus_tokens
    laeufe = laeufe_aus_tokens(block.tokens)
    alt = "".join(l.text for l in laeufe if l.art in ("k", "d", "s"))
    neu = "".join(l.text for l in laeufe if l.art in ("k", "a", "s"))
    return " ".join(alt.split()), " ".join(neu.split())


@pytest.mark.parametrize("stuecke,alt,neu", [
    ([("Pumpe", "k"), (" ", "d"), ("Ventil", "k")], "Pumpe Ventil", "PumpeVentil"),
    ([("Pumpe", "k"), (" ", "n"), ("Ventil", "k")], "PumpeVentil", "Pumpe Ventil"),
    ([("Pumpe", "k"), (" ", "d"), ("-", "n"), ("Ventil", "k")], "Pumpe Ventil", "Pumpe-Ventil"),
    ([("Pumpe ", "k"), ("Ventil ", "d"), ("Becken", "k")], "Pumpe Ventil Becken", "Pumpe Becken"),
    ([("Pumpe ", "k"), ("neues ", "n"), ("Becken", "k")], "Pumpe Becken", "Pumpe neues Becken"),
], ids=["leer_weg", "leer_neu", "leer_weg_strich_neu", "wort_mit_leer_weg", "wort_mit_leer_neu"])
def test_markiertes_leerzeichen_aendert_die_ansichten_richtig(tmp_path, stuecke, alt, neu):
    d = WordDok()
    d.absatz([voll([("Die ", "k")] + stuecke + [(" ", "k")])])
    a = _lies(tmp_path, d)
    va, vn = _ansichten(a.blocks[0])
    assert va.startswith("Die " + alt + " ") and vn.startswith("Die " + neu + " ")
    assert not _sperrend(_alle_codes(a))


def test_linie_nur_ueber_einer_wortluecke_ohne_leerzeichen_sperrt(tmp_path):
    d = WordDok()
    x, top = EBENEN[1], d.top
    w = breite("Pumpe")
    luecke = breite(" ")
    d.zeile(x, top, [("Pumpe", "k")])
    d.zeile(x + w + luecke, top, voll([("Ventil", "k")], 1))
    d.durchstreichung(x + w + 0.2, x + w + luecke - 0.2, top, WORT_ROT)        # gestrichenes Leerzeichen ohne Zeichen (1,7 pt)
    d.top += ZEILE
    a = _lies(tmp_path, d)
    assert "linie_unklar" in _sperrend(_codes(_tok(a, "Pumpe")))
    assert "linie_unklar" in _codes(_tok(a, "Ventil"))


def test_unterstreichung_je_wort_mit_luecken_segment_ist_kein_befund(tmp_path):
    """Manche Programme zeichnen die Unterstreichung je Wort plus ein Stueck ueber der Luecke: das ist eine Linie."""
    d = WordDok()
    x, top = EBENEN[1], d.top
    w1, wl, w2 = breite("Pumpe"), breite(" "), breite("Ventil")
    d.text(x, top, "Pumpe", farbe=WORT_ROT)
    d.text(x + w1 + wl, top, "Ventil", farbe=WORT_ROT)
    d.zeile(x + w1 + wl + w2, top, voll([(" Becken Zebrafink", "k")], 1))
    d.unterstreichung(x, x + w1, top, WORT_ROT)
    d.unterstreichung(x + w1, x + w1 + wl, top, WORT_ROT)
    d.unterstreichung(x + w1 + wl, x + w1 + wl + w2, top, WORT_ROT)
    d.top += ZEILE
    a = _lies(tmp_path, d)
    assert _tok(a, "Pumpe").status == ADD and _tok(a, "Ventil").status == ADD
    assert not _sperrend(_alle_codes(a))


# =================================================================================================
# Markup-Leiste: Sprechblasen, Verbindungslinien, Ankerpunkte
# =================================================================================================
def _absatz_mit_blasen(d: WordDok, n: int = 4) -> list[float]:
    zeilen = [voll([("Die ", "k"), ("Pumpe", "n"), (" ", "k")])] + fuellzeilen(n - 1, 1)
    tops = d.absatz(zeilen)
    d.sprechblase(tops[0] - 2, "hat formatiert:", "Nicht Hervorheben", anker=(EBENEN[1] + 30, tops[0]))
    d.sprechblase(tops[1] + 2, "Kommentiert [ZF1]:", "Bitte die Pumpe pruefen", anker=(EBENEN[1] + 120, tops[1]))
    d.sprechblase(tops[2] + 4, "Formatiert:", "Schriftart: Fett", anker=(EBENEN[1] + 200, tops[2]))
    return tops


def test_sprechblasen_stehen_nie_im_text_aber_in_randnotizen(tmp_path):
    d = WordDok()
    _absatz_mit_blasen(d)
    a = _lies(tmp_path, d)
    woerter = {x for t in _toks(a) for x in (t.old, t.new)}
    for fremd in ("formatiert:", "Hervorheben", "Kommentiert", "[ZF1]:", "Fett"):
        assert fremd not in woerter
    arten = [(r["art"], r["text"]) for r in a.randnotizen]
    assert arten == [("formatiert", "hat formatiert: Nicht Hervorheben"),
                     ("kommentar", "Kommentiert [ZF1]: Bitte die Pumpe pruefen"),
                     ("formatiert", "Formatiert: Schriftart: Fett")]
    assert all(r["seite"] == 1 and isinstance(r["y"], float) for r in a.randnotizen)
    assert any("Sprechblasen" in w and "1 Kommentare" in w for w in a.warnungen)


def test_freier_text_in_der_leiste_ist_randnotiz(tmp_path):
    d = WordDok()
    tops = d.absatz([voll([("Die ", "k"), ("Pumpe", "n")])] + fuellzeilen(2))
    d.randtext(tops[1], "Seitenumbruch oberhalb")
    a = _lies(tmp_path, d)
    assert [(r["art"], r["text"]) for r in a.randnotizen] == [("sonstig", "Seitenumbruch oberhalb")]
    assert all("oberhalb" not in (t.old, t.new) for t in _toks(a))
    assert a.luecken == []


def test_geloeschter_text_nur_in_der_sprechblase_ist_luecke(tmp_path):
    """Word-Option "Loeschungen in Sprechblasen": die Loeschung steht nur in der Leiste -> nie still verloren."""
    d = WordDok()
    tops = d.absatz([voll([("Die ", "k"), ("Pumpe", "n")])] + fuellzeilen(2))
    d.sprechblase(tops[1], "Gelöscht:", "Ventil und Becken", anker=(EBENEN[1] + 60, tops[1]))
    a = _lies(tmp_path, d)
    l = [x for x in a.luecken if x["art"] == "sonstiges"]
    assert len(l) == 1 and l[0]["seite"] == 1 and l[0]["woerter"] == 3
    assert "Ventil" not in l[0]["beschreibung"]          # keine Dokumentinhalte in der Luecke (Regel 16)


def test_absatz_neben_sprechblasen_bleibt_ein_block(tmp_path):
    d = WordDok()
    _absatz_mit_blasen(d, 5)
    a = _lies(tmp_path, d)
    assert len(a.blocks) == 1
    assert _tok(a, "Pumpe").status == ADD


def test_kurze_zeile_bezieht_sich_auf_die_inhaltsbreite_links_der_leiste(tmp_path):
    """Grosser Zeilenabstand (Absatzende erkennt man dort an der kurzen Zeile): volle Zeilen enden am Textrand (~395 pt),
    weit vor dem Seitenrand -- gemessen an der Inhaltsbreite links der Leiste sind sie nicht kurz."""
    d = WordDok()
    zeilen = [voll([("Die ", "k"), ("Pumpe", "n"), (" ", "k")])] + fuellzeilen(3, 1)
    for i, z in enumerate(zeilen):
        d.zeile(EBENEN[1], d.top + 15.0 * i, z)                   # 15 pt Abstand bei 8,4 pt Schrift
    d.top += 15.0 * len(zeilen) + 12
    d.absatz([voll([("Zweiter ", "k"), ("Absatz", "n")])])
    a = _lies(tmp_path, d)
    assert len(a.blocks) == 2
    assert _tok(a, "Pumpe").status == ADD


def test_verbindungslinien_und_ankerpunkte_sind_keine_markierung(tmp_path):
    d = WordDok()
    tops = d.absatz([voll([("Pumpe ", "k"), ("Ventil", "d"), (" Becken Zebrafink", "k")])] + fuellzeilen(2))
    # Verbindungslinie laeuft auf der Grundlinie quer durch schwarzen Text
    d.sprechblase(tops[0] + 15, "hat formatiert:", "Nicht Hervorheben", anker=(EBENEN[1] + 2, tops[0]))
    d.sprechblase(tops[1] + 25, "hat formatiert:", "Hervorheben", anker=(EBENEN[1] + 2, tops[1]))
    a = _lies(tmp_path, d)
    assert [t.old for t in _toks(a) if t.status != KEEP] == ["Ventil"]
    codes = _alle_codes(a)
    assert not codes & {"unterstrichen", "linie_unklar", "tinte_unerklaert"}
    assert not _sperrend(codes)
    assert a.luecken == []


def test_flacher_ankerpunkt_am_anfang_der_verbindungslinie(tmp_path):
    """pdfplumber meldet einen Ankerpunkt aus zwei Boegen als flachen Kasten (Hoehe 0) -- er sitzt aber am Anfang einer
    Verbindungslinie und ist damit keine Linie im Text (kein linie_unklar mitten im Wort)."""
    d = WordDok()
    tops = d.absatz([voll([("Pumpe ", "k"), ("Kranichgruppe", "k"), (" ", "k"), ("Becken", "n")])] + fuellzeilen(1))
    ax = EBENEN[1] + breite("Pumpe Kranich")
    d.verbindungslinie(ax, tops[0], tops[0] + 30)
    d.ankerpunkt(ax, tops[0], flach=True)
    a = _lies(tmp_path, d)
    assert "linie_unklar" not in _alle_codes(a)
    assert _tok(a, "Kranichgruppe").status == KEEP and _tok(a, "Becken").status == ADD


def test_gelbe_hinterlegung_und_aenderungsbalken_ohne_sperre(tmp_path):
    d = WordDok()
    x = EBENEN[1]
    top = d.top
    stuecke = voll([("Wert ", "k"), ("[Platzhalter]", "k"), (" ", "k"), ("Pumpe", "n")])
    d.zeile(x, top, stuecke)
    w0 = breite("Wert ")
    d.gelb(x + w0, x + w0 + breite("[Platzhalter]"), top)
    d.aenderungsbalken(top, top + 2 * ZEILE)
    d.top += ZEILE
    d.absatz(fuellzeilen(1))
    a = _lies(tmp_path, d)
    p = _tok(a, "[Platzhalter]")
    assert p.status == KEEP and "hinterlegt" in _codes(p)
    assert not _sperrend(_alle_codes(a))
    assert a.luecken == []


# =================================================================================================
# Listen und Ueberschriften
# =================================================================================================
def test_mehrere_listenpunkte_werden_mehrere_bloecke(tmp_path):
    d = WordDok()
    d.listenpunkt(1, [("A.", "k")], fuellzeilen(2, 0, 2))
    d.listenpunkt(2, [("1.", "k")], fuellzeilen(2, 1, 3), abstand=0)
    d.listenpunkt(3, [("a.", "k")], fuellzeilen(1, 2, 4), abstand=0)
    d.listenpunkt(3, [("b.", "k")], [voll([("Die ", "k"), ("Pumpe", "n")], 4)], abstand=0)
    d.listenpunkt(4, [("1)", "k")], fuellzeilen(1, 3, 5), abstand=0)
    d.listenpunkt(5, [("a)", "k")], fuellzeilen(1, 0, 6), abstand=0)
    d.listenpunkt(1, [("B.", "k")], fuellzeilen(1, 1, 2), abstand=0)
    d.listenpunkt(1, [("AA.", "k")], fuellzeilen(1, 2, 2), abstand=0)
    a = _lies(tmp_path, d)
    assert [b.tokens[0].old for b in a.blocks] == ["A.", "1.", "a.", "b.", "1)", "a)", "B.", "AA."]
    assert all(b.kind == "absatz" for b in a.blocks)
    assert a.luecken == []                     # Nummer + Tabulator-Luecke ist kein zweispaltiger Text


def test_gruene_listennummer_ist_ein_label_wort(tmp_path):
    d = WordDok()
    d.listenpunkt(1, [("b.", "k")], fuellzeilen(1, 0, 2))
    d.listenpunkt(1, [("c.", "d", WORT_GRUEN)], [[("Das Ventil schliesst bei Ueberdruck selbst.", "d")]], abstand=0)
    d.listenpunkt(1, [("d.", "d", WORT_GRUEN), ("c.", "n", WORT_GRUEN)], fuellzeilen(1, 1, 2), abstand=0)
    a = _lies(tmp_path, d)
    assert len(a.blocks) == 3
    lab = a.blocks[2].tokens[0]
    assert (lab.old, lab.new, lab.status) == ("d.", "c.", MOD)
    assert lab.teile == [(DEL, "d."), (ADD, "c.")]
    assert a.blocks[1].ganz_geloescht
    assert not _sperrend(_alle_codes(a))


def test_gruene_nummer_mit_kleiner_luecke_wird_zusammengefuegt(tmp_path):
    d = WordDok()
    top = d.top
    x = EBENEN[1]
    d.zeile(x, top, [("c.", "d", WORT_GRUEN)])
    d.zeile(x + breite("c.") + 2.0, top, [("a.", "n", WORT_GRUEN)])        # 2 pt Luecke (mehr als ein Wortabstand)
    d.zeile(EBENEN[2], top, voll([("Das Ventil schliesst", "k")], 2))
    d.top += ZEILE
    a = _lies(tmp_path, d)
    lab = a.blocks[0].tokens[0]
    assert (lab.old, lab.new) == ("c.", "a.")
    assert a.blocks[0].tokens[1].old == "Das"


def test_section_und_part_ueberschriften(tmp_path):
    d = WordDok()
    d.kapitel([("PART 1 GENERAL", "k")])
    d.ueberschrift([("1.6", "k")], [("PUMPENHAUS UND VENTILE", "k")])
    d.listenpunkt(1, [("A.", "k")], fuellzeilen(2, 0, 2))
    d.ueberschrift([("1.7", "d"), ("1.8", "n")], [("ZEBRAFINK", "k")])
    d.listenpunkt(1, [("A.", "k")], fuellzeilen(1, 1, 2))
    a = _lies(tmp_path, d)
    kinds = [(b.kind, b.section_alt, b.section_neu) for b in a.blocks]
    assert kinds == [("ueberschrift", "Part 1", "Part 1"), ("ueberschrift", "1.6", "1.6"), ("absatz", "1.6", "1.6"),
                     ("ueberschrift", "1.7", "1.8"), ("absatz", "1.7", "1.8")]
    assert [t.old for t in a.blocks[1].tokens] == ["1.6", "PUMPENHAUS", "UND", "VENTILE"]


def test_zahl_am_zeilenanfang_im_fliesstext_ist_keine_ueberschrift(tmp_path):
    d = WordDok()
    d.absatz([voll([("Die Pumpe hebt das Wasser um", "k")]), voll([("1.5 Meter an und das Ventil", "k")])])
    a = _lies(tmp_path, d)
    assert len(a.blocks) == 1 and a.blocks[0].kind == "absatz"


# =================================================================================================
# Kopf-/Fusszeilen und Seitenumbruch
# =================================================================================================
def test_kopf_und_fuss_mit_eigenen_aenderungen_werden_ueber_drei_seiten_ignoriert(tmp_path):
    d = WordDok(mit_kopf_fuss=True)
    for i in range(3):
        if i:
            d.neue_seite()
        d.absatz([voll([("Seite ", "k"), ("Pumpe", "n")], 1, i)] + fuellzeilen(2, i))
    a = _lies(tmp_path, d)
    woerter = {x for t in _toks(a) for x in (t.old, t.new)}
    assert not woerter & {"Projekt", "Ostflügel", "Revision", "[Datum]Mai"}
    kf = [i for i in a.ignoriert if i["grund"] == "kopf_fuss"]
    assert sorted({i["seite"] for i in kf}) == [1, 2, 3] and len(kf) == 9
    assert all(i["woerter"] == 1 for i in kf)
    assert [t.status for t in _toks(a) if t.status != KEEP] == [ADD, ADD, ADD]
    # dieselben Baender gelten in der Farb-Inventur (keine unerklaerte Farbe in Kopf/Fuss)
    assert a.luecken == [] and "tinte_unerklaert" not in _alle_codes(a)
    assert all(any(t <= 125 <= b for t, b in a.kopf_fuss_baender[s]) for s in (1, 2, 3))


def test_gliederung_oben_auf_jeder_seite_ist_keine_kopfzeile(tmp_path):
    """Jede Seite beginnt an derselben Hoehe mit "PART 1 GENERAL", "1.1 PUMPENHAUS" und demselben Listenpunkt: das ist
    Gliederung (wiederkehrend, aber nie Kopf-/Fusszeile) -- nichts davon darf verschluckt werden."""
    d = WordDok(mit_kopf_fuss=True)
    for i in range(3):
        if i:
            d.neue_seite()
        d.kapitel([("PART 1 GENERAL", "k")])
        d.ueberschrift([("1.1", "k")], [("PUMPENHAUS", "k")])
        d.listenpunkt(1, [("A.", "k")], [voll([("Die ", "k"), ("Pumpe", "n"), (" ", "k")], 2)] + fuellzeilen(1, i, 2))
    a = _lies(tmp_path, d)
    assert [b.kind for b in a.blocks] == ["ueberschrift", "ueberschrift", "absatz"] * 3
    assert [t.status for t in _toks(a) if t.status != KEEP] == [ADD, ADD, ADD]
    assert len(a.ignoriert) == 9                     # nur die echten Kopf-/Fusszeilen


def test_randzeile_nur_auf_einer_von_drei_seiten_bleibt_text(tmp_path):
    d = WordDok()
    for i in range(3):
        if i:
            d.neue_seite()
        if i == 1:
            d.zeile(EBENEN[0], 124.0, [("Einmalige Notiz zum ", "k"), ("Kranich", "n")])
        d.absatz(fuellzeilen(2, i))
    a = _lies(tmp_path, d)
    assert _tok(a, "Kranich").status == ADD
    assert not a.ignoriert


def test_absatz_ueber_den_seitenumbruch(tmp_path):
    d = WordDok(mit_kopf_fuss=True)
    d.top = 585.0
    d.absatz([voll([("Die ", "k"), ("Pumpe", "n"), (" ", "k")])] + fuellzeilen(3, 1))
    d.neue_seite()
    d.absatz([[("weiter im Text und endet hier.", "k")]])
    a = _lies(tmp_path, d)
    assert len(a.blocks) == 2
    assert a.blocks[0].fortsetzung in ("sicher", "moeglich")


# =================================================================================================
# Legende waehlen (auto) und fest einstellen
# =================================================================================================
def test_auto_ohne_leiste_mit_unter_und_durchstreichung_derselben_farbe_ist_linie(tmp_path):
    d = WordDok(mit_leiste=False)
    d.absatz([voll([("Pumpe ", "k"), ("Ventil", "d"), (" ", "k"), ("Becken", "n")])])
    a = _lies(tmp_path, d)
    assert a.legende_modus == "linie" and not a.markup_leiste
    assert "derselben Farbe" in a.legende_grund
    assert _tok(a, "Ventil").status == DEL and _tok(a, "Becken").status == ADD


def test_auto_ohne_leiste_mit_blauem_text_ohne_linie_ist_farbe(tmp_path):
    d = WordDok(mit_leiste=False)
    d.absatz([voll([("Pumpe ", "k"), ("Ventil", "f", WORT_BLAU), (" ", "k"), ("Becken", "d", WORT_ROT)])])
    a = _lies(tmp_path, d)
    assert a.legende_modus == "farbe"
    assert _tok(a, "Ventil").status == ADD
    assert _tok(a, "Becken").status == DEL
    assert not any(w.startswith("Legende:") for w in a.warnungen)


def test_auto_viel_farbe_ohne_linie_bleibt_farbe(tmp_path):
    """Von Hand markiert: neuer Text rot ohne Linie, EIN Wort zusaetzlich rot unterstrichen, eines rot gestrichen ->
    die Mehrheit des farbigen Textes hat keine Linie, also bleibt die Farbe die Bedeutung."""
    d = WordDok(mit_leiste=False)
    d.absatz([voll([("Pumpe ", "k"), ("Ventil Becken Zebrafink Speicher", "f"), (" ", "k"), ("alt", "d"), (" ", "k"),
                    ("neu", "n")])])
    a = _lies(tmp_path, d)
    assert a.legende_modus == "farbe"
    assert _tok(a, "Zebrafink").status == ADD


@pytest.mark.parametrize("modus,erwartet", [("linie", KEEP), ("farbe", ADD)])
def test_legende_fest_eingestellt(tmp_path, modus, erwartet):
    d = WordDok(mit_leiste=False)
    d.absatz([voll([("Pumpe ", "k"), ("Ventil", "f"), (" ", "k"), ("Becken", "d"), (" ", "k"),
                    ("Kranich", "f", WORT_GRUEN)])])
    a = _lies(tmp_path, d, _cfg(modus))
    assert a.legende_modus == modus and a.legende_grund == ""
    assert _tok(a, "Ventil").status == erwartet
    assert _tok(a, "Becken").status == DEL
    gruen = _tok(a, "Kranich")
    assert gruen.status == KEEP
    if modus == "linie":      # Farbe bedeutet nichts: Hinweis ohne Sperre, keine "unbekannte Farbe"
        assert "farbe_ohne_linie" in _codes(gruen) and "farbe_ohne_linie" in _codes(_tok(a, "Ventil"))
        assert not [l for l in a.luecken if l["art"] == "unbekannte_farbe"]
    else:                     # ADR 0005: andere Farben bleiben unbekannt und sperren
        assert "unbekannte_farbe" in _sperrend(_codes(gruen))


def test_farbuebersicht_nennt_die_bedeutung_der_legende(tmp_path):
    d = WordDok()
    d.absatz([voll([("Pumpe ", "k"), ("Ventil", "d")])])
    a = _lies(tmp_path, d)
    rot = next(f for f in a.farben if f["farbe"] == "#d13438")
    assert "Linie" in rot["bedeutung"]


def test_ohne_leiste_bleibt_der_zeilenbau_wie_bisher(tmp_path):
    """Ohne Markup-Leiste: keine Randnotizen, nichts abgeschnitten (auch Text weit rechts bleibt Text)."""
    d = WordDok(mit_leiste=False)
    d.zeile(EBENEN[1], d.top, [("Pumpe", "k")])
    d.zeile(470.0, d.top, [("Ventil", "n")])
    d.top += ZEILE
    a = _lies(tmp_path, d)
    assert a.randnotizen == [] and not a.markup_leiste
    assert _tok(a, "Ventil").status == ADD


# =================================================================================================
# Alte Legende (Korpora von Auftrag 001/002) bleibt unveraendert
# =================================================================================================
def test_alte_legende_liest_rot_weiter_als_entfaellt(tmp_path):
    d = WordDok()
    d.absatz([voll([("Pumpe ", "k"), ("Ventil", "n")])])
    a = _lies(tmp_path, d, _cfg("alt"))
    assert a.legende_modus == "alt"
    assert _tok(a, "Ventil").status == DEL
