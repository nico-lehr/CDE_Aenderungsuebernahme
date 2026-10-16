"""Leser-Treue R1-R8 (Auftrag 002, Welle 1, Paket PDF): Schema-3-Felder trenner/teile/trennstelle/fortsetzung und die
behobenen Lesefehler. Nur erfundene Daten (Regel 16); die PDFs entstehen im Test.
"""
from __future__ import annotations

import copy

from src import config
from src.models import ADD, DEL, KEEP, MOD
from src.pdf_reader import analysiere_pdf
from tests.fixtures.pdf_writer import SEITE_H, breite
from tests.redteam._bau_pdf import RohSeite, alle_tokens, dok, lese, roh_pdf, token, zeile

RECHTS = 539.0
WOERTER = ("Kuehlung Schaltschrank Messwert Sicherung Lueftung Zuleitung Steuerung Abnahme Wartung Anschluss Spannung "
           "Frequenz Schutzleiter Erdung Batterie Wechselrichter Transformator Schalter Verteiler Kabeltrasse Brandmelder "
           "Sensor Pruefung Dokumentation Betriebszustand Fernwartung Stoerung Protokoll Lieferant Betreiber").split()


def _volle_zeile(start: int, anfang: str = "", ende: str = "") -> str:
    """Eine Zeile, die (rechtsbuendig gesetzt) den rechten Rand fuellt: der naechste Wortschatz-Eintrag passt nicht mehr."""
    teile = [anfang] if anfang else []
    i = start
    while True:
        w = WOERTER[i % len(WOERTER)]
        probe = " ".join(teile + [w] + ([ende] if ende else []))
        if breite(probe, 11) > 470:
            break
        teile.append(w)
        i += 1
    return " ".join(teile + ([ende] if ende else []))


def _rechtsbuendig(top: float, text: str, farbe: str = "0 0 0 rg") -> str:
    return zeile(RECHTS - breite(text, 11), top, text, 11.0, farbe)


def _mitte_y(top: float, groesse: float = 11.0) -> float:
    return SEITE_H - (top + groesse * 0.55)


# =================================================================================================
# R1 Trenner
# =================================================================================================
NBSP_SCHRIFT = ("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Type /Encoding "
                "/BaseEncoding /WinAnsiEncoding /Differences [160 /nbspace] >> >>")


def test_trenner_geschuetztes_leerzeichen_und_leerzeichen(tmp_path):
    # /nbspace hat in dieser Schrift keine Breite: der Abstand kommt (wie oft in echten PDFs) aus der TJ-Verschiebung
    inhalt = "BT /F1 11 Tf 0 g 56 700 Td [(Wert 5\240) -278 (kW hier)] TJ ET\n"
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt, schrift=NBSP_SCHRIFT)]))
    assert [(t.old, t.trenner) for t in alle_tokens(a)] == [("Wert", " "), ("5", "\xa0"), ("kW", " "), ("hier", " ")]


def test_leerzeichen_zeichen_trennt_auch_bei_kleiner_luecke(tmp_path):
    """Enger Blocksatz (negative Wortabstaende): die Luecke ist kleiner als die Schwelle, das Leerzeichen-Zeichen trennt."""
    inhalt = "BT /F1 11 Tf 0 g -2.2 Tw 56 700 Td (eng gesetzt hier) Tj ET\n"
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert [t.old for t in alle_tokens(a)] == ["eng", "gesetzt", "hier"]


def test_phantom_leerzeichen_ohne_vorschub_trennen_kein_wort(tmp_path):
    """Wie LibreOffice an weichen Trennstellen: ein Leerzeichen, das das Zeichen davor ueberdeckt (kein Vorschub), und
    Leerzeichen, die erst NACH dem Trennstrich gezeichnet werden, trennen das Wort nicht."""
    inhalt = ("BT /F1 11 Tf 0 g 56 700 Td [(Die Batte) 278 ( ) (rie hier)] TJ ET\n"
              "BT /F1 11 Tf 0 g 56 685 Td [(ist zu pro) -80 (-) 333 ( ) 278 ( )] TJ ET\n"
              "BT /F1 11 Tf 0 g 56 670 Td (tokollieren und fertig.) Tj ET\n")
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    woerter = [t.old for t in alle_tokens(a)]
    assert "Batterie" in woerter and "protokollieren" in woerter, woerter
    assert token(a, "protokollieren").trennstelle == 3


def test_luecke_ohne_leerzeichen_trennt_weiter(tmp_path):
    """Woerter, die einzeln gesetzt sind (ohne Leerzeichen-Zeichen), werden weiter an der Luecke getrennt."""
    d = dok()
    d.block("Die Anlage {+laeuft+} gut.")
    a = lese(tmp_path, d)
    assert [(t.new, t.trenner) for t in alle_tokens(a)] == [("Die", " "), ("Anlage", " "), ("laeuft", " "), ("gut.", " ")]


# =================================================================================================
# R2 Zeichen-Laeufe
# =================================================================================================
def test_farbiges_wort_mit_schwarzem_punkt_hat_zeichen_laeufe(tmp_path):
    d = dok()
    d.block("Die Anlage {+insgesamt+}. Und {-weg-}. {-Netzleistung-}{+Leistung+} hier.")
    a = lese(tmp_path, d)
    neu = token(a, "insgesamt.")
    assert (neu.old, neu.new, neu.status) == ("", "insgesamt.", ADD)            # wie bisher (Rueckwaertsvertraeglichkeit)
    assert neu.teile == [(ADD, "insgesamt"), (KEEP, ".")] and not neu.hinweis
    weg = token(a, "weg.")
    assert weg.status == DEL and weg.teile == [(DEL, "weg"), (KEEP, ".")]
    ers = token(a, "Leistung")
    assert ers.status == MOD and ers.teile == [(DEL, "Netzleistung"), (ADD, "Leistung")]
    assert token(a, "hier.").teile == [] and token(a, "Anlage").teile == []


def test_geometrie_glaettung_meldet_rand_unklar(tmp_path):
    d = dok(variante="hintergrund")
    d.block("Die Anlage {+insgesamt+}. Und mehr.")
    a = lese(tmp_path, d)
    t = token(a, "insgesamt.")
    assert t.status == ADD and t.teile == [] and "rand_unklar" in t.hinweis
    assert any("Rand" in h for h in a.blocks[0].hinweise)


def test_ganz_markiertes_wort_ohne_rand_hat_keinen_hinweis(tmp_path):
    d = dok(variante="hintergrund")
    d.block("Die Anlage {+laeuft+} gut.")
    a = lese(tmp_path, d)
    assert not [t for t in alle_tokens(a) if t.hinweis]


# =================================================================================================
# R3 Hochzahlen
# =================================================================================================
def _hochzahl(farbe: str, hoch: float = 4.0) -> bytes:
    w = breite("Die Zentrale", 11)
    basis = 64 + 11 * 0.8 - hoch                      # Grundlinie der Hochzahl (von oben)
    inhalt = (zeile(56, 64, "Die Zentrale") + zeile(56 + w, basis - 7 * 0.8, "1", 7.0, farbe)
              + zeile(56 + w + breite("1", 7), 64, " meldet weiter."))
    return roh_pdf([RohSeite(inhalt=inhalt)])


def test_hochzahl_gehoert_nicht_zum_wort(tmp_path):
    a = lese(tmp_path, _hochzahl("0 0 0 rg"))
    assert [t.old for t in alle_tokens(a)] == ["Die", "Zentrale", "meldet", "weiter."]
    assert not [t for t in alle_tokens(a) if t.hinweis]


def test_markierte_hochzahl_wird_gemeldet(tmp_path):
    a = lese(tmp_path, _hochzahl("0 0 1 rg"))
    t = token(a, "Zentrale")
    assert t.status == KEEP and "hochzahl_markiert" in t.hinweis


def test_kleines_zeichen_auf_der_grundlinie_bleibt_im_wort(tmp_path):
    a = lese(tmp_path, _hochzahl("0 0 0 rg", hoch=0.0))
    assert token(a, "Zentrale1") is not None


# =================================================================================================
# R4 Absatzgrenzen und Seiten-Fortsetzung
# =================================================================================================
def _seite_mit_absatz(top: float, anzahl: int, letzte: str, start: int = 0) -> tuple[str, float]:
    inhalt = ""
    for i in range(anzahl):
        inhalt += _rechtsbuendig(top + 15 * i, _volle_zeile(start + 3 * i))
    return inhalt + zeile(56, top + 15 * anzahl, letzte), top + 15 * (anzahl + 1)


def test_harter_umbruch_vor_einem_wort_das_noch_gepasst_haette_ist_neuer_absatz(tmp_path):
    inhalt, top = _seite_mit_absatz(64, 3, "Kurze letzte Zeile mit Satzende.")
    inhalt += _rechtsbuendig(top, _volle_zeile(5, "Neu")) + _rechtsbuendig(top + 15, _volle_zeile(9))
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert len(a.blocks) == 2 and a.blocks[1].tokens[0].old == "Neu"


def test_volle_zeile_vor_der_naechsten_bleibt_ein_absatz(tmp_path):
    inhalt, top = _seite_mit_absatz(64, 3, "")
    inhalt += _rechtsbuendig(top - 15, _volle_zeile(11)) + _rechtsbuendig(top, _volle_zeile(14, "weiter"))
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert len(a.blocks) == 1 and not [t for t in alle_tokens(a) if t.hinweis]


def test_knappe_absatzgrenze_wird_gemeldet(tmp_path):
    inhalt, top = _seite_mit_absatz(64, 3, "")
    platz = breite("Neu", 11) + 0.28 * 11 + 2.0                 # passt, aber nur knapp (Toleranz 0,5 em)
    kurz = "Etwas Text davor"
    inhalt += zeile(RECHTS - platz - breite(kurz, 11), top - 15, kurz)
    inhalt += _rechtsbuendig(top, _volle_zeile(5, "Neu"))
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert len(a.blocks) == 1
    assert "absatzgrenze_unklar" in token(a, "Neu").hinweis and a.blocks[0].hinweise


def _zwei_seiten(ende_seite1: str, anfang_seite2: str, voll: bool = True) -> bytes:
    s1 = "".join(_rechtsbuendig(64 + 15 * i, _volle_zeile(3 * i)) for i in range(3))
    letzte = _volle_zeile(20, ende=ende_seite1) if voll else ende_seite1
    s1 += _rechtsbuendig(109, letzte) if voll else zeile(56, 109, letzte)
    s2 = _rechtsbuendig(64, _volle_zeile(7, anfang_seite2)) + _rechtsbuendig(79, _volle_zeile(12))
    return roh_pdf([RohSeite(inhalt=s1), RohSeite(inhalt=s2)])


def test_fortsetzung_sicher_ohne_satzende_und_klein_weiter(tmp_path):
    a = lese(tmp_path, _zwei_seiten("und", "weiter"))
    assert [b.page for b in a.blocks] == [1, 2]
    assert a.blocks[0].fortsetzung == "sicher" and a.blocks[1].fortsetzung == ""
    assert any("nächsten Seite" in h for h in a.blocks[0].hinweise)          # der bisherige Hinweis bleibt


def test_fortsetzung_moeglich_bei_grossbuchstabe_oder_satzende_in_voller_zeile(tmp_path):
    assert lese(tmp_path / "gross", _zwei_seiten("die", "Anlage")).blocks[0].fortsetzung == "moeglich"
    assert lese(tmp_path / "punkt", _zwei_seiten("Ende.", "Neuer")).blocks[0].fortsetzung == "moeglich"


def test_keine_fortsetzung_nach_kurzer_letzter_zeile_mit_satzende(tmp_path):
    a = lese(tmp_path, _zwei_seiten("Kurz und fertig.", "Neuer", voll=False))
    assert a.blocks[0].fortsetzung == "" and not a.blocks[0].hinweise


def test_fortsetzung_sicher_bei_silbentrennung_ueber_die_seite(tmp_path):
    a = lese(tmp_path, _zwei_seiten("Kapazi-", "taet"))
    t = a.blocks[0].tokens[-1]
    assert t.old == "Kapazitaet" and t.trennstelle == 6 and t.page == 1
    assert a.blocks[0].fortsetzung == "sicher" and a.blocks[1].tokens[0].old != "taet"


# =================================================================================================
# R5 Durchstreichung ohne Breitenkappe
# =================================================================================================
def test_strich_ueber_eine_ganze_zeile_wird_erkannt(tmp_path):
    text = _volle_zeile(0)
    x0 = RECHTS - breite(text, 11)
    y = _mitte_y(64)
    inhalt = _rechtsbuendig(64, text) + "0 0 0 RG 0.8 w %.2f %.2f m %.2f %.2f l S\n" % (x0 - 0.5, y, RECHTS + 0.5, y)
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert RECHTS - x0 > 0.6 * 595
    assert all(t.status == DEL for t in alle_tokens(a)) and not [t for t in alle_tokens(a) if t.hinweis]


def test_flache_kurve_ist_eine_durchstreichung(tmp_path):
    y = _mitte_y(64)
    w = breite("Wartung", 11)
    inhalt = zeile(56, 64, "Wartung hier") + "0 0 0 RG 0.8 w 56 %.2f m %.2f %.2f l %.2f %.2f l S\n" % (
        y, 56 + w / 2, y, 56 + w, y)
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert [t.status for t in alle_tokens(a)] == [DEL, KEEP]


def test_linie_die_die_meisten_zeichen_nicht_im_band_kreuzt_ist_unklar(tmp_path):
    """Eine Linie auf Mittelhoehe grosser Zeichen laeuft durch den oberen Rand kleiner Zeichen: keine Durchstreichung."""
    w = breite("AB", 20)
    inhalt = zeile(56, 64, "AB", 20.0) + zeile(56 + w + 2, 64 + 20 * 0.8 - 8 * 0.8, "kleinwortkleinwort", 8.0)
    basis = SEITE_H - (64 + 16)
    y = basis + 5.0
    inhalt += "0 0 0 RG 0.8 w 56 %.2f m %.2f %.2f l S\n" % (y, 56 + w + 2 + breite("kleinwortkleinwort", 8), y)
    pfad = tmp_path / "t.pdf"
    pfad.write_bytes(roh_pdf([RohSeite(inhalt=inhalt)]))
    cfg = copy.deepcopy(config.lade("farben"))
    cfg["durchstreichung"]["max_ueberstand_em"] = 99.0          # nur die Kreuz-Regel (R5) darf hier entscheiden
    a = analysiere_pdf(pfad, cfg)
    assert all(t.status == KEEP for t in alle_tokens(a))
    assert "linie_unklar" in token(a, "AB").hinweis
    cfg["durchstreichung"]["min_kreuz_anteil"] = 0.0             # Gegenprobe: ohne Kreuz-Regel wuerde gestrichen
    assert token(analysiere_pdf(pfad, cfg), "AB").status == DEL


# =================================================================================================
# R6 Farbraum fuer Flaechen und Linien
# =================================================================================================
TYP2_BLAU = "<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0 0] /C1 [1 0.8 0 0] /N 1 >>"
TYP3 = "<< /FunctionType 3 /Domain [0 1] /Functions [" + TYP2_BLAU + "] /Bounds [] /Encode [0 1] >>"


def _sonder(funktion: str, zeichnen: str) -> bytes:
    raeume = "/ColorSpace << /Cs1 [/Separation /FirmenFarbe /DeviceCMYK " + funktion + "] >>"
    inhalt = zeichnen + "0 g 0 G\n" + zeile(56, 64, "Wort hier und weiter")
    return roh_pdf([RohSeite(inhalt=inhalt, ressourcen=raeume)])


def _flaeche_ueber_wort() -> str:
    w = breite("Wort", 11)
    return "/Cs1 cs 1 scn %.2f %.2f %.2f %.2f re f\n" % (56, SEITE_H - 64 - 12, w, 13)


def test_sonderfarb_flaeche_ton_1_ist_markierung(tmp_path):
    a = lese(tmp_path, _sonder(TYP2_BLAU, _flaeche_ueber_wort()))
    assert [t.status for t in alle_tokens(a)][:2] == [ADD, KEEP]


def test_nicht_berechenbare_sonderfarb_flaeche_wird_gemeldet(tmp_path):
    a = lese(tmp_path, _sonder(TYP3, _flaeche_ueber_wort()))
    assert token(a, "Wort").status == KEEP
    assert "unbekannte_farbe" in token(a, "Wort").hinweis and "unbekannte_farbe" in [l["art"] for l in a.luecken]


def test_linie_mit_nicht_berechenbarer_farbe_wird_gemeldet(tmp_path):
    y = _mitte_y(64)
    linie = "/Cs1 CS 1 SCN 0.8 w 56 %.2f m %.2f %.2f l S\n" % (y, 56 + breite("Wort", 11), y)
    a = lese(tmp_path, _sonder(TYP3, linie))
    t = token(a, "Wort")
    assert t.status == DEL and "linienfarbe_unklar" in t.hinweis


def test_hellgraue_sonderfarb_linie_ist_zierlinie(tmp_path):
    grau = "<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0 0] /C1 [0 0 0 0.25] /N 1 >>"
    y = _mitte_y(64)
    linie = "/Cs1 CS 1 SCN 0.8 w 56 %.2f m %.2f %.2f l S\n" % (y, 56 + breite("Wort", 11), y)
    a = lese(tmp_path, _sonder(grau, linie))
    t = token(a, "Wort")
    assert t.status == KEEP and "linie_unklar" in t.hinweis


# =================================================================================================
# R8 Silbentrennung
# =================================================================================================
def _getrennt(z1: str, z2: str, farbe: str = "0 0 0 rg", zusatz: str = "") -> bytes:
    inhalt = zeile(56, 64, z1, farbe=farbe) + zeile(56, 79, z2, farbe=farbe)
    if zusatz:
        inhalt += zeile(56, 200, zusatz)
    return roh_pdf([RohSeite(inhalt=inhalt)])


def test_grossbuchstabe_nach_dem_strich_behaelt_den_bindestrich(tmp_path):
    a = lese(tmp_path, _getrennt("Meldung per E-", "Mail an alle.", farbe="0 0 1 rg"))
    t = token(a, "E-Mail")
    assert t is not None and t.status == ADD and t.trennstelle == 1 and t.new[t.trennstelle] == "-"
    assert not t.hinweis


def test_abkuerzung_und_ziffer_vor_dem_strich_behalten_ihn(tmp_path):
    a = lese(tmp_path / "xml", _getrennt("Das Format ist XML-", "basiert und gut."))
    assert token(a, "XML-basiert").trennstelle == 3
    a = lese(tmp_path / "zahl", _getrennt("Eine 2-", "stufige Anlage."))
    assert token(a, "2-stufige") is not None


def test_kleinbuchstabe_danach_fuegt_ohne_strich_zusammen(tmp_path):
    a = lese(tmp_path, _getrennt("Die Kapazi-", "taet ist gross."))
    t = token(a, "Kapazitaet")
    assert t.trennstelle == 6 and t.trenner == " " and token(a, "taet") is None


def test_neues_getrenntes_wort_aus_dem_wortschatz_ist_sicher(tmp_path):
    a = lese(tmp_path, _getrennt("Die Sicherheitsarchi-", "tektur hier.", farbe="0 0 1 rg",
                                 zusatz="Die Sicherheitsarchitektur gilt."))
    t = next(t for t in alle_tokens(a) if t.status == ADD and t.new == "Sicherheitsarchitektur")
    assert t.trennstelle == 16 and "trennung_unklar" not in t.hinweis


def test_neues_getrenntes_wort_ausserhalb_des_wortschatzes_wird_gemeldet(tmp_path):
    a = lese(tmp_path, _getrennt("Die Sicherheitsarchi-", "tektur hier.", farbe="0 0 1 rg"))
    t = token(a, "Sicherheitsarchitektur")
    assert t.status == ADD and "trennung_unklar" in t.hinweis and a.blocks[0].hinweise


def test_neues_wort_mit_bindestrich_im_wortschatz_behaelt_ihn(tmp_path):
    a = lese(tmp_path, _getrennt("Die Lade-", "station hier.", farbe="0 0 1 rg", zusatz="Die Lade-station gilt."))
    t = next(t for t in alle_tokens(a) if t.status == ADD and "station" in t.new)
    assert t.new == "Lade-station" and t.trennstelle == 4 and "trennung_unklar" not in t.hinweis


def test_zusammengefuegte_zeichen_laeufe_bleiben_erhalten(tmp_path):
    d = dok(trennung="ascii")
    d.block("Wir pruefen heute ausfuehrlich die Anforderung ob die {+Kapazitaetserweiterung Sicherheitsarchitektur "
            "Netzanbindung Verfuegbarkeitsgarantie+}. Und sonst nichts weiter hier.")
    a = lese(tmp_path, d)
    for wort, _, _ in d.getrennt:
        t = next(t for t in alle_tokens(a) if t.new.startswith(wort))
        assert t.status == ADD and 0 < t.trennstelle < len(t.new)
        assert not t.teile or "".join(x for _, x in t.teile) == t.new


# =================================================================================================
# Nacharbeit: gedrehte Seiten (Schriftgroesse) und Fussnoten am Seitenende
# =================================================================================================
def test_quer_stehender_text_hat_die_richtige_schriftgroesse_und_bleibt_ein_absatz(tmp_path):
    """Seite mit /Rotate 90 und quer gesetztem Text: pdfminer meldet als Groesse die Zeichenbreite (~6 statt 11).
    Damit galt der Zeilenabstand als 'weit' und jede kuerzere Zeile beendete den Absatz."""
    text = " ".join(WOERTER[:24]) + " und " + " ".join(WOERTER[:10]) + "."
    d = dok()
    d.block(text)
    d.block("Zweiter Absatz mit etwas Text hier.")
    d.seite_rotieren(90)
    a = lese(tmp_path, d)
    assert [len(b.tokens) for b in a.blocks] == [len(text.split()), 6]


def test_fussnote_am_seitenende_verhindert_die_fortsetzung_nicht(tmp_path):
    s1 = "".join(_rechtsbuendig(64 + 15 * i, _volle_zeile(3 * i)) for i in range(3))
    s1 += _rechtsbuendig(109, _volle_zeile(20, ende="und"))
    s1 += zeile(56, 780, "1 Eine erfundene Fussnote zur Erlaeuterung.", 8.0)
    s2 = _rechtsbuendig(64, _volle_zeile(7, "weiter")) + _rechtsbuendig(79, _volle_zeile(12) + ".")
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=s1), RohSeite(inhalt=s2)]))
    assert [b.page for b in a.blocks] == [1, 2, 1] and [b.idx for b in a.blocks] == [0, 1, 2]
    assert a.blocks[0].fortsetzung == "sicher" and a.blocks[1].tokens[0].old == "weiter"
    assert "Fussnote" in a.blocks[2].alt_text and a.blocks[2].fortsetzung == ""
