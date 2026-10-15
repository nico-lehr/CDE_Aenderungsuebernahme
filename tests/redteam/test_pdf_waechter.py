"""Waechter fuer den PDF-Leser: Eigenschaften, die HEUTE gelten und weiter gelten muessen (kein xfail).

Bricht einer dieser Tests, wurde beim Reparieren der RT-60 bis RT-109 etwas kaputt gemacht. Viele sind das
Gegenstueck zu einem Fehlertest (gleicher Aufbau, aber ohne den Stoerfall). Nur erfundene Daten (Regel 16);
Zufall nur mit `random.Random(seed)`.
"""
import random

import pytest

from src.models import ADD, DEL, KEEP, MOD
from tests.fixtures.pdf_writer import PdfDokument, markierte_woerter
from tests.redteam._bau import aufbau
from tests.redteam._bau_pdf import (RohSeite, VOKAB, eindeutiger_text, SEITE_H, alle_tokens, annot, annot_gedreht90, bloecke_mit, dok, geaendert, kopf_fuss_zone,
                                    lese, roh_pdf, token, tokens_mit, woerter, zeile, zeile_gedreht90)


def _status(a, *ws):
    return [token(a, w).status if token(a, w) else None for w in ws]


# ---------------------------------------------------------------- Kopf-/Fusszeilen
def test_waechter_seite_n_von_m_und_wiederholte_kopfzeile_werden_ignoriert(tmp_path):
    d = PdfDokument(kopf_farbe=(0.0, 0.0, 1.0))                  # blaue Kopf-/Fusszeile auf jeder Seite
    d.block("Erste Seite mit Text.")
    d.seitenumbruch()
    d.block("Zweite Seite mit Text.")
    aufbau(d.seitenzahl == 2, "zwei Seiten erwartet")
    a = lese(tmp_path, d)
    assert not kopf_fuss_zone(a), "Kopf-/Fusszeilen sind als Bloecke stehen geblieben"
    assert not woerter(a, ADD), "blaue Kopf-/Fusszeile zaehlt als neu"
    assert a.ignoriert and all(i["grund"] == "kopf_fuss" for i in a.ignoriert), "Ignorierte Woerter muessen gezaehlt sein"
    assert sum(i["woerter"] for i in a.ignoriert) == 2 * (3 + 4), "3 Kopf- und 4 Fusswoerter je Seite"


# ---------------------------------------------------------------- einfache Markierungen
@pytest.mark.parametrize("variante", ["farbe", "strich", "hintergrund", "annot"])
def test_waechter_einfache_markierung_alle_varianten(tmp_path, variante):
    d = dok(variante=variante)
    d.block("Die Kapazitaet betraegt {-100 kWh-} {+120 kWh+} je Rack.")
    a = lese(tmp_path, d)
    assert _status(a, "100", "120") == [DEL, ADD], variante
    b = bloecke_mit(a, "Rack.")[0]
    assert b.alt_text == "Die Kapazitaet betraegt 100 kWh je Rack."
    assert b.neu_text == "Die Kapazitaet betraegt 120 kWh je Rack."


def test_waechter_blau_rot_hex_farben(tmp_path):
    d = dok(farbe_neu="1F4E9F", farbe_weg="C00000")
    d.block("Das ist {+neu+} und {-alt-} im Satz.")
    a = lese(tmp_path, d)
    assert _status(a, "neu", "alt") == [ADD, DEL]


@pytest.mark.parametrize("farbraum,grund", [("cmyk", "gray"), ("cmyk", "cmyk"), ("rgb", "gray")])
def test_waechter_cmyk_und_grau_als_farbraum(tmp_path, farbraum, grund):
    d = dok(farbraum=farbraum, grundfarbraum=grund)
    d.block("Das ist {+neu+} und {-alt-} im Satz.")
    a = lese(tmp_path, d)
    assert _status(a, "neu", "alt") == [ADD, DEL]


def test_waechter_verschmolzene_ersetzung_gleicher_laenge(tmp_path):
    d = dok()
    d.block("Die Kapazitaet betraegt {-100-}{+120+} kWh je Rack.")
    a = lese(tmp_path, d)
    t = next(t for t in alle_tokens(a) if t.status == MOD)
    assert (t.old, t.new) == ("100", "120")


def test_waechter_blau_unterstrichen_und_rot_unterstrichen_zaehlen_nach_farbe(tmp_path):
    d = dok()
    d.block("Das ist {@bu:neu@} und {@ru:alt@} im Satz.")
    a = lese(tmp_path, d)
    assert _status(a, "neu", "alt") == [ADD, DEL]


def test_waechter_strikeout_trifft_die_richtigen_woerter(tmp_path):
    """Gegenstueck zu RT-91/RT-92: gleicher Seitenaufbau, hier mit StrikeOut -> 'wichtiger Satz' ist DEL."""
    inhalt = zeile(56, 64, "Das ist ein wichtiger Satz mit Markierung hier.")
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt, annots=[annot("StrikeOut", 110, 64, 179, 77, "1 0 0")])]))
    assert _status(a, "wichtiger", "Satz", "ein", "mit") == [DEL, DEL, KEEP, KEEP]


def test_waechter_aufrechte_gedrehte_seite_wird_gelesen(tmp_path):
    """Gegenstueck zu RT-93: aufrechter Text auf einer /Rotate-90-Seite wird als Satz gelesen."""
    inhalt = zeile_gedreht90(56, 64, "Erste Zeile bleibt gleich hier.") + zeile_gedreht90(56, 79, "Zweite Zeile ist ")
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt, attribute="/Rotate 90")]))
    assert bloecke_mit(a, "Erste", "Zweite"), "Seite mit /Rotate 90 und aufrechtem Text wird nicht als Text gelesen"


@pytest.mark.parametrize("dy,mediabox", [(0, "0 0 595 842"), (100, "0 100 595 942"), (0, "50 0 645 842")])
def test_waechter_highlight_mit_versetzter_mediabox(tmp_path, dy, mediabox):
    """Annotation bei MediaBox-Ursprung ungleich 0/0 trifft dieselben Woerter (CropBox/MediaBox-Versatz kein Fehler)."""
    inhalt = ""
    for i, t in enumerate(["Erste Zeile bleibt gleich hier.", "Zweite Zeile ist neuer Text", "Dritte Zeile bleibt auch gleich."]):
        basis = SEITE_H - (64 + 15 * i + 11 * 0.8) + dy
        inhalt += f"BT /F1 11 Tf 0 0 0 rg 56 {basis:.2f} Td ({t}) Tj ET\n"
    y1, y0 = SEITE_H - 79 + dy, SEITE_H - 90 + dy
    an = (f"/Subtype /Highlight /Rect [56 {y0} 300 {y1}] /QuadPoints [56 {y1} 300 {y1} 56 {y0} 300 {y0}] /C [0 0 1] /F 4")
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt, annots=[an], mediabox=mediabox)]))
    assert _status(a, "Zweite", "neuer", "Text", "Erste", "Dritte") == [ADD, ADD, ADD, KEEP, KEEP]


# ---------------------------------------------------------------- Warnungen, die heute schon da sind
def test_waechter_scan_seite_erzeugt_mindestens_die_warnung(tmp_path):
    from tests.fixtures.pdf_writer import scan_bild
    d = dok()
    d.block("Normale Seite mit {+einer Aenderung+} im Text.")
    d.scan_seite(scan_bild([("p", "Nur Pixel auf dieser Seite.")]))
    a = lese(tmp_path, d)
    assert any("Seite 2" in w and "keinen auslesbaren Text" in w for w in a.warnungen), a.warnungen
    assert _status(a, "Aenderung") == [ADD] or [t.status for t in tokens_mit(a, "Aenderung")] == [ADD]


def test_waechter_unbekannte_farbe_erzeugt_mindestens_die_warnung(tmp_path):
    d = dok(farbe_neu="00FFFF")
    d.block("Das ist {+neuer Text hier+} im Satz und {-alt-}.")
    a = lese(tmp_path, d)
    assert any("#00ffff" in w and "Legende" in w for w in a.warnungen), a.warnungen


def test_waechter_dokument_ohne_aenderung_warnt(tmp_path):
    d = dok()
    d.block("Ein Satz ohne jede Markierung.")
    a = lese(tmp_path, d)
    assert any("keine Änderungen erkannt" in w for w in a.warnungen)


# ---------------------------------------------------------------- Layout
def test_waechter_absatz_ohne_fussnote_bleibt_ein_block(tmp_path):
    """Gegenstueck zu RT-62 (gleicher Text, nur ohne Fussnoten)."""
    d = dok()
    d.block(eindeutiger_text(62, 38, "Erstwort", "Endwort", "{+neuwort+}"))
    d.block(eindeutiger_text(63, 30, "Zweitwort", "Schlusswort"))
    a = lese(tmp_path, d)
    assert bloecke_mit(a, "Erstwort", "Endwort", "neuwort")
    assert bloecke_mit(a, "Zweitwort", "Schlusswort")
    assert len([b for b in a.blocks if b.kind == "absatz"]) == 2


def test_waechter_lockere_liste_mit_absatzabstand_ist_getrennt(tmp_path):
    d = dok()
    d.block("Einleitung vor der Liste ohne Neuerung.")
    for i, wort in enumerate(("Alpha", "Beta", "Gamma"), 1):
        d.block(f"Das {wort} Listenelement ist hier eingetragen und lang genug.", marker=f"{i}.")
    a = lese(tmp_path, d)
    assert not bloecke_mit(a, "Alpha", "Beta") and not bloecke_mit(a, "Beta", "Gamma")
    assert len(a.blocks) == 4


def test_waechter_gedankenstrich_mitten_in_der_zeile_teilt_nicht(tmp_path):
    d = dok()
    d.block("Der Betrieb laeuft – stabil und wird dokumentiert.")
    a = lese(tmp_path, d)
    assert len(a.blocks) == 1 and bloecke_mit(a, "laeuft", "stabil")


def test_waechter_geschuetztes_leerzeichen_bleibt_nicht_im_wort(tmp_path):
    d = dok()
    d.block("Die Last betraegt 5 kW im Betrieb.")
    a = lese(tmp_path, d)
    assert _status(a, "5", "kW") == [KEEP, KEEP]
    assert not any(" " in t.old + t.new for t in alle_tokens(a))


def test_waechter_schwarze_silbentrennung_mit_strich_wird_zusammengefuegt(tmp_path):
    texte = ["Der Betreiber stellt die Speicherkapazitaet bereit und dokumentiert Verantwortlichkeiten "
             "Uebergabeprotokolle sowie Wartungsintervalle der gesamten Anlage regelmaessig und vollstaendig."] * 3
    d = dok(trennung="ascii")
    for t in texte:
        d.block(t)
    getrennt = [w for w, stile, z in d.getrennt if stile == ("k",) and z == "-"]
    aufbau(getrennt, "im Test-PDF wurde kein schwarzes Wort getrennt")
    a = lese(tmp_path, d)
    for w in set(getrennt):
        assert token(a, w) is not None, f"getrenntes schwarzes Wort {w!r} wurde nicht zusammengefuegt"


def test_waechter_tabellenrahmen_ohne_farbe_sind_keine_durchstreichung(tmp_path):
    d = dok()
    d.tabelle([["Name", "Wert"], ["Strom", "5 kW"], ["Spannung", "230 V"]], breiten=[100, 100], kopf_farbe=None, band_farbe=None)
    a = lese(tmp_path, d)
    assert not any(geaendert(t) for t in alle_tokens(a))


def test_waechter_zwei_seiten_absatz_verliert_kein_wort(tmp_path):
    d = dok()
    text = ("Dies ist ein Absatz der ueber einen Seitenumbruch geht und mehrere Zeilen hat damit das Ganze testbar ist "
            "und noch weiter geht fuer eine dritte Zeile")
    d.block(text, umbruch_nach_zeile=1)
    a = lese(tmp_path, d)
    gelesen = [t.old for t in alle_tokens(a)]
    assert gelesen == text.split()


# ---------------------------------------------------------------- Zufall (fester Seed) und Wiederholbarkeit
@pytest.mark.parametrize("seed", [101, 202, 303, 404, 505])
def test_waechter_zufallsdokument_markierte_woerter_stimmen(tmp_path, seed):
    rng = random.Random(seed)
    d = dok()
    soll_neu, soll_weg, soll_summe = [], [], 0
    for _ in range(4):
        teile = []
        for _ in range(rng.randint(8, 16)):
            w = rng.choice(VOKAB) + str(rng.randint(0, 99))        # Zahl dran: jedes Wort ist (fast) einmalig
            art = rng.random()
            if art < 0.15:
                teile.append("{+" + w + "+}")
                soll_neu.append(w)
            elif art < 0.30:
                teile.append("{-" + w + "-}")
                soll_weg.append(w)
            else:
                teile.append(w)
        markup = " ".join(teile)
        soll_summe += markierte_woerter(markup)
        d.block(markup)
    a = lese(tmp_path, d)
    assert sorted(t.new for t in woerter(a, ADD)) == sorted(soll_neu)
    assert sorted(t.old for t in woerter(a, DEL)) == sorted(soll_weg)
    assert len(woerter(a, ADD)) + len(woerter(a, DEL)) == soll_summe == len(soll_neu) + len(soll_weg)


def test_waechter_lesen_ist_wiederholbar(tmp_path):
    d = dok()
    d.ueberschrift("1.1 Geltungsbereich")
    d.block("Die Kapazitaet betraegt {-100 kWh-} {+120 kWh+} je Rack und gilt fuer alle Module im Betrieb.")
    a1 = lese(tmp_path / "eins", d)
    a2 = lese(tmp_path / "zwei", d)
    assert a1.blocks == a2.blocks and a1.luecken == a2.luecken and a1.warnungen == a2.warnungen
