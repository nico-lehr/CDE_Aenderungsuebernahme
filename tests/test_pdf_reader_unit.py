"""Unit-Tests fuer src/pdf_reader.py (K-PDF): eine Gruppe je Reparatur, dazu Gegenproben ("darf NICHT melden").

Nur erfundene Daten (Regel 16). Die PDFs entstehen im Test (tests/fixtures/pdf_writer.py bzw. die Roh-PDF-Helfer aus
tests/redteam/_bau_pdf.py, hier nur benutzt, nicht veraendert).
"""
from __future__ import annotations

import pytest

from src import config
from src.models import ADD, DEL, KEEP, MOD
from src.pdf_reader import analysiere_pdf
from tests.fixtures.pdf_writer import PdfDokument, muster_bild
from tests.redteam._bau_pdf import RohSeite, alle_tokens, dok, lese, roh_pdf, token, zeile

BILD = "q 595 0 0 842 0 0 cm /Im1 Do Q\n"
TEXT = ("Dieser Satz steht auf der Seite und ist lang genug damit die Seite als Textseite gilt "
        "und niemand sie fuer einen Scan haelt weil genug Woerter darauf stehen und weiter und weiter")


def _arten(a) -> list[str]:
    return [l["art"] for l in a.luecken]


def _luecke(a, art: str) -> dict:
    return next(l for l in a.luecken if l["art"] == art)


def _text_seite(zeilen: list[str], farbe: str = "0 0 0 rg", top: float = 64.0) -> str:
    return "".join(zeile(56, top + 15 * i, t, 11.0, farbe) for i, t in enumerate(zeilen))


# =================================================================================================
# Scan-Seiten (scan_seite)
# =================================================================================================
def test_seite_ohne_text_nur_bild_ist_scan_seite(tmp_path):
    pdf = roh_pdf([RohSeite(inhalt=_text_seite([TEXT])), RohSeite(inhalt=BILD, bild=True)])
    a = lese(tmp_path, pdf)
    assert [(l["art"], l["seite"]) for l in a.luecken] == [("scan_seite", 2)]
    assert any("Seite 2" in w and "keinen auslesbaren Text" in w for w in a.warnungen)    # alte Warnung bleibt


def test_bildseite_mit_aufgedruckter_seitenzahl_ist_scan_seite(tmp_path):
    fuss = zeile(250, 800, "Seite 2 von 2", 9.0, "0.4 0.4 0.4 rg")
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=_text_seite([TEXT])), RohSeite(inhalt=BILD + fuss, bild=True)]))
    assert [(l["art"], l["seite"]) for l in a.luecken] == [("scan_seite", 2)]
    assert any("Seite 2" in w for w in a.warnungen)             # nicht mehr stillschweigend


@pytest.mark.parametrize("farbe,modus", [("0 0 0 rg", 3), ("1 1 1 rg", None)], ids=["unsichtbar", "weiss"])
def test_textebene_ueber_bild_ist_scan_seite(tmp_path, farbe, modus):
    ocr = zeile(56, 64, "Erkannter Text einer gescannten Seite hier unten.", farbe=farbe, modus=modus)
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=_text_seite([TEXT])), RohSeite(inhalt=BILD + ocr, bild=True)]))
    assert ("scan_seite", 2) in [(l["art"], l["seite"]) for l in a.luecken]
    assert token(a, "Erkannter") is not None                    # der Text bleibt lesbar, die Luecke warnt


def test_unsichtbarer_text_zaehlt_nicht_als_farbige_markierung(tmp_path):
    ocr = zeile(56, 300, "Verborgener Text in Blau.", farbe="0 0 1 rg", modus=3)
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=_text_seite([TEXT]) + ocr)]))
    assert token(a, "Verborgener").status == KEEP


def test_textseite_mit_hintergrundbild_ist_keine_scan_seite(tmp_path):
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=BILD + _text_seite([TEXT, TEXT]), bild=True)]))
    assert a.luecken == []


def test_leere_seite_ohne_bild_ist_keine_luecke(tmp_path):
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=_text_seite([TEXT])), RohSeite(inhalt="")]))
    assert a.luecken == []
    assert any("Seite 2" in w for w in a.warnungen)


# =================================================================================================
# Bilder (bild)
# =================================================================================================
def test_grosses_bild_im_text_ist_luecke_bild(tmp_path):
    d = dok()
    d.block("Text vor dem Bild.")
    d.bild(muster_bild(240, 140, seed=3), 200.0)
    d.block("Text nach dem Bild.")
    a = lese(tmp_path, d)
    assert _arten(a) == ["bild"] and a.luecken[0]["seite"] == 1 and a.luecken[0]["woerter"] == 0


def test_kleines_logo_und_bild_im_seitenkopf_sind_keine_luecke(tmp_path):
    logo = "q 60 0 0 30 56 600 cm /Im1 Do Q\n"          # 1800 pt^2 von ~500000: unter 2 %
    kopf = "q 200 0 0 20 56 800 cm /Im1 Do Q\n"          # Bild im oberen Seitenrand
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=logo + kopf + _text_seite([TEXT]), bild=True)]))
    assert a.luecken == []


# =================================================================================================
# Tabellen und Spalten
# =================================================================================================
def test_tabelle_mit_rahmen_ist_luecke_und_zaehlt_geaenderte_woerter(tmp_path):
    d = dok()
    d.block("Vorspann ohne Aenderung.")
    d.tabelle([["Name", "Wert"], ["Strom {+geaendert+}", "5 kW"], ["Spannung", "230 V"]], kopf_farbe=None, band_farbe=None)
    a = lese(tmp_path, d)
    assert _arten(a) == ["tabelle"]
    assert _luecke(a, "tabelle")["woerter"] == 1


def test_einzelne_linien_und_ein_kasten_sind_keine_tabelle(tmp_path):
    inhalt = (_text_seite([TEXT, TEXT])
              + "0.5 w 56 600 m 156 600 l S\n"                    # Fussnotenstrich
              + "0.5 w 56 400 m 539 400 l S\n"                     # Trennlinie
              + "56 200 200 100 re S\n")                           # ein Kasten = eine Zelle
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert a.luecken == []


LINKS = ["Linkswort " + " ".join(["Netzanschluss", "Kuehlung", "Messwert", "Sicherung", "Lueftung", "Zuleitung",
                                  "Steuerung", "Abnahme", "Wartung", "Anschluss", "Spannung", "Frequenz"]) + " Linksende"]
RECHTS = ["Rechtswort " + " ".join(["Batterie", "Wechselrichter", "Transformator", "Schalter", "Verteiler",
                                    "Kabeltrasse", "Brandmelder", "Sensor", "Pruefung", "Dokumentation"]) + " Rechtsende"]


def test_zwei_spalten_sind_luecke_mehrspaltig(tmp_path):
    d = dok()
    d.zwei_spalten(LINKS, RECHTS)
    a = lese(tmp_path, d)
    assert _arten(a) == ["mehrspaltig"]


def test_zwei_spalten_mit_ungleich_langen_absaetzen(tmp_path):
    """Die Zeilen der Spalten stehen nicht auf gleicher Hoehe (eine Spalte hat zwei Absaetze mehr)."""
    d = dok()
    d.zwei_spalten(LINKS * 2 + ["Noch ein Absatz links mit {+neuem+} Wort."], RECHTS)
    a = lese(tmp_path, d)
    assert "mehrspaltig" in _arten(a)


def test_gewoehnlicher_text_ueberschriften_und_listen_sind_keine_spalten(tmp_path):
    d = dok()
    d.ueberschrift("1 Allgemeines")
    d.block(TEXT)
    d.block(TEXT, marker="1.")
    d.block(TEXT, marker="2.")
    d.block(TEXT, marker="3.")
    d.ueberschrift("2 Weiteres")
    d.block(TEXT)
    a = lese(tmp_path, d)
    assert a.luecken == []


# =================================================================================================
# Farben, die keiner Legende zugeordnet sind (unbekannte_farbe)
# =================================================================================================
@pytest.mark.parametrize("hexfarbe", ["00FFFF", "7F00FF", "009900", "00003C", "808080", "F5B5B5"],
                         ids=["cyan", "violett", "gruen", "dunkelblau", "grau", "pastellrot"])
def test_unbekannte_textfarbe_ist_luecke_mit_woerterzahl(tmp_path, hexfarbe):
    d = dok(farbe_neu=hexfarbe)
    d.block("Das ist {+neuer Text hier+} im Satz.")
    a = lese(tmp_path, d)
    assert [t.status for t in alle_tokens(a)] == [KEEP] * 7
    l = _luecke(a, "unbekannte_farbe")
    assert l["woerter"] == 3 and l["seite"] == 1
    assert "#" in l["beschreibung"] and "neuer" not in l["beschreibung"]       # Farbe ja, Dokumentinhalt nie


def test_grundfarbe_des_dokuments_ist_keine_markierung(tmp_path):
    inhalt = _text_seite([TEXT, TEXT], farbe="0 0.4 0 rg")             # der ganze Text ist dunkelgruen
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert a.luecken == []


def test_graue_kopfzeile_und_schwarz_sind_keine_luecke(tmp_path):
    d = PdfDokument()                                                   # graue Kopf-/Fusszeile (einseitig)
    d.block("Ein Satz ohne Markierung.")
    assert lese(tmp_path, d).luecken == []


def test_gelbe_flaeche_hinter_text_ist_unbekannte_farbe(tmp_path):
    inhalt = "1 1 0 rg 92 %.2f 76 13 re f\n0 g\n" % (842 - 77) + zeile(56, 64, "Das ist neuer Text hier im Satz.")
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert _luecke(a, "unbekannte_farbe")["woerter"] == 3


def test_luecke_nennt_die_seite(tmp_path):
    d = dok(farbe_neu="009900")
    d.block("Seite eins ohne Farbe.")
    d.seitenumbruch()
    d.block("Seite zwei mit {+gruenem Text+}.")
    a = lese(tmp_path, d)
    assert [(l["art"], l["seite"], l["woerter"]) for l in a.luecken] == [("unbekannte_farbe", 2, 2)]


# =================================================================================================
# Unlesbare Zeichen (nicht_lesbarer_text), Ligaturen
# =================================================================================================
_SCHRIFT = ("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Type /Encoding "
            "/BaseEncoding /WinAnsiEncoding /Differences [128 /fi /gl1234 /uniF0B7] >> >>")


def _zeichen(satz: str):
    return roh_pdf([RohSeite(inhalt=zeile(56, 64, satz, kodierung="latin-1"), schrift=_SCHRIFT)])


def test_cid_zeichen_sind_luecke_und_wort_hat_hinweis(tmp_path):
    a = lese(tmp_path, _zeichen("Die Daten \x81 sind hier."))
    l = _luecke(a, "nicht_lesbarer_text")
    assert l["woerter"] == 1 and l["seite"] == 1
    t = token(a, "(cid:129)")
    assert "nicht_lesbar" in t.hinweis
    assert any("nicht gelesen" in h for b in a.blocks for h in b.hinweise)


def test_private_use_zeichen_mitten_im_text_ist_luecke(tmp_path):
    a = lese(tmp_path, _zeichen("Die Daten \x82 sind hier."))
    assert _luecke(a, "nicht_lesbarer_text")["woerter"] == 1


def test_private_use_aufzaehlungspunkt_am_zeilenanfang_ist_keine_luecke(tmp_path):
    a = lese(tmp_path, _zeichen("\x82 Erster Listenpunkt steht hier."))
    assert a.luecken == []


def test_ligatur_bekommt_hinweis_ohne_luecke_und_bleibt_unveraendert(tmp_path):
    a = lese(tmp_path, _zeichen("Die Anlage ist \x80xiert und stabil."))
    t = next(t for t in alle_tokens(a) if "xiert" in t.old)
    assert t.old == "ﬁxiert" and "ligatur" in t.hinweis
    assert a.luecken == []


def test_normaler_text_hat_keine_hinweise(tmp_path):
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=_text_seite([TEXT]))]))
    assert a.luecken == [] and all(not t.hinweis for t in alle_tokens(a)) and all(not b.hinweise for b in a.blocks)


# =================================================================================================
# Kopf-/Fusszeilen ohne Wiederholungsmuster
# =================================================================================================
def test_einseitige_kopf_und_fusszeile_werden_ignoriert_und_gezaehlt(tmp_path):
    d = PdfDokument(kopf_text="Anlage 3 - Technische Spezifikation", kopf_farbe=(0.0, 0.0, 1.0),
                    fuss_vorlage="Dokument CDE-0815 Blatt {nr}")
    d.block("Normaler Text der Seite.")
    a = lese(tmp_path, d)
    assert [b.page for b in a.blocks] == [1] and a.blocks[0].neu_text == "Normaler Text der Seite."
    assert sum(i["woerter"] for i in a.ignoriert) == 5 + 4      # blaue Kopfwoerter + blaue Fusswoerter, gezaehlt
    assert not any(t.status == ADD for t in alle_tokens(a))


def test_text_im_oberen_rand_ohne_abstand_ist_keine_kopfzeile(tmp_path):
    inhalt = zeile(56, 30, "Erste Zeile ganz oben auf der Seite.") + zeile(56, 45, "Zweite Zeile folgt direkt darunter.")
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert a.ignoriert == [] and a.blocks[0].alt_text.startswith("Erste Zeile")


# =================================================================================================
# Falsche "neu/entfaellt"-Treffer: Zierlinie, Tabellenkopf, Hyperlink, blaue Ueberschriften
# =================================================================================================
def test_hellgraue_zierlinie_durch_ueberschrift_ist_keine_durchstreichung(tmp_path):
    d = dok()
    d.ueberschrift("1.1 Geltungsbereich Anforderungen")
    d.block("Normaler Text fuer die Anforderung hier.")
    d._s.inhalt.append((d._linie_op(56, 300, 70 + 9, (0.7, 0.7, 0.7), 0.5) + "\n").encode("latin-1"))
    a = lese(tmp_path, d)
    kopf = next(b for b in a.blocks if b.kind == "ueberschrift")
    assert all(t.status == KEEP for t in kopf.tokens)
    assert all("linie_unklar" in t.hinweis for t in kopf.tokens) and kopf.hinweise


def test_dunkle_linie_die_weit_ueber_den_text_hinausragt_ist_keine_durchstreichung(tmp_path):
    inhalt = zeile(56, 64, "Kurz hier") + "0 0 0 RG 0.8 w 56 %.2f m 400 %.2f l S\n" % (842 - 69.5, 842 - 69.5)
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert [t.status for t in alle_tokens(a)] == [KEEP, KEEP]
    assert all("linie_unklar" in t.hinweis for t in alle_tokens(a))


def test_echte_durchstreichung_bleibt_entfaellt(tmp_path):
    for variante in ("strich", "farbe"):
        d = dok(variante=variante)
        d.block("Das ist {-alter Text hier-} im Satz.")
        a = lese(tmp_path / variante, d)
        assert [token(a, w).status for w in ("alter", "Text", "hier", "Das")] == [DEL, DEL, DEL, KEEP], variante
        assert all(not t.hinweis for t in alle_tokens(a)), variante


def test_blauer_tabellenkopf_und_baenderung_sind_gestaltung(tmp_path):
    d = dok()
    d.tabelle([["Name", "Wert", "Einheit"], ["Strom", "5", "kW"], ["Spannung", "230", "V"], ["Frequenz", "50", "Hz"],
               ["Druck", "2", "bar"]], kopf_farbe="4472C4", band_farbe="B4C6E7")
    a = lese(tmp_path, d)
    assert not [t for t in alle_tokens(a) if t.status != KEEP]
    assert "tabelle" in _arten(a)
    assert any("design_flaeche" in t.hinweis for t in alle_tokens(a))


def test_blauer_kasten_ausserhalb_einer_tabelle_ist_gestaltung_mit_luecke(tmp_path):
    inhalt = ("0.72 0.86 1 rg 50 640 495 70 re f\n0 g\n"
              + zeile(56, 140, "Hinweis in einem blauen Kasten mit Text") + zeile(56, 155, "und einer zweiten Zeile darin")
              + zeile(56, 170, "und einer dritten Zeile darin hier."))
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    assert not [t for t in alle_tokens(a) if t.status != KEEP]
    l = _luecke(a, "sonstiges")
    assert l["woerter"] == 18


def test_markierung_direkt_am_wort_bleibt_neu(tmp_path):
    d = dok(variante="hintergrund")
    d.block("Die Kapazitaet betraegt {-100 kWh-} {+120 kWh+} je Rack.")
    a = lese(tmp_path, d)
    assert [token(a, w).status for w in ("100", "120")] == [DEL, ADD]
    assert not [t for t in alle_tokens(a) if t.hinweis]


def test_hyperlink_in_linkblau_ist_nicht_neu_aber_blaues_wort_ohne_link_schon(tmp_path):
    d = dok()
    d.block("Siehe {@l:www.beispiel.example@} und {+neues Wort+} dazu.")
    a = lese(tmp_path, d)
    link = token(a, "www.beispiel.example")
    assert link.status == KEEP and "link" in link.hinweis
    assert token(a, "neues").status == ADD and token(a, "Wort").status == ADD


def _blaue_ueberschriften(n: int, blau: str = "0.12 0.3 0.62 rg", text_neu: bool = False) -> bytes:
    inhalt, top = "", 64.0
    for nr in range(1, n + 1):
        inhalt += zeile(56, top, f"{nr} Titel{nr}", 14.0, blau, fett=True)
        inhalt += zeile(56, top + 24, f"Text unter der Ueberschrift {nr} ohne jede Aenderung.",
                        farbe="0 0 1 rg" if text_neu else "0 0 0 rg")
        top += 60
    return roh_pdf([RohSeite(inhalt=inhalt)])


def test_durchgehend_blaue_ueberschriften_sind_gestaltung(tmp_path):
    a = lese(tmp_path, _blaue_ueberschriften(4))
    assert not [t for t in alle_tokens(a) if t.status != KEEP]
    kopf = [b for b in a.blocks if b.kind == "ueberschrift"]
    assert len(kopf) == 4 and all(b.section_alt == b.section_neu and b.section_alt for b in kopf)
    assert all("design_blau" in t.hinweis for b in kopf for t in b.tokens)
    assert [(l["art"], l["woerter"]) for l in a.luecken] == [("sonstiges", 8)]


@pytest.mark.parametrize("fall", ["zu_wenige", "text_auch_blau"])
def test_blaue_ueberschriften_bleiben_neu_wenn_es_keine_gestaltung_sein_kann(tmp_path, fall):
    pdf = _blaue_ueberschriften(2) if fall == "zu_wenige" else _blaue_ueberschriften(4, text_neu=True)
    a = lese(tmp_path, pdf)
    kopf = [b for b in a.blocks if b.kind == "ueberschrift"]
    assert kopf and all(t.status == ADD for b in kopf for t in b.tokens)
    assert a.luecken == []


# =================================================================================================
# Verschmolzene Ersetzungen: kein Wortteil geht verloren
# =================================================================================================
def _einzel(tmp_path, markup: str, **kw):
    d = dok(**kw)
    d.block(markup)
    return lese(tmp_path, d)


def test_ersetzung_ohne_leerzeichen_behaelt_alten_und_neuen_text(tmp_path):
    a = _einzel(tmp_path, "Die {-Netzleistung-}{+Leistung+} betraegt 5 kW.")
    t = next(t for t in alle_tokens(a) if "eistung" in t.old + t.new)
    assert (t.old, t.new, t.status) == ("Netzleistung", "Leistung", MOD)
    b = a.blocks[0]
    assert b.alt_text == "Die Netzleistung betraegt 5 kW." and b.neu_text == "Die Leistung betraegt 5 kW."


@pytest.mark.parametrize("variante", ["farbe", "strich", "hintergrund", "annot"])
def test_ersetzung_ohne_leerzeichen_in_allen_stilen(tmp_path, variante):
    a = _einzel(tmp_path, "Die {-Netzleistung-}{+Leistung+} betraegt 5 kW.", variante=variante)
    b = a.blocks[0]
    assert b.alt_text == "Die Netzleistung betraegt 5 kW." and b.neu_text == "Die Leistung betraegt 5 kW."


def test_unveraenderter_wortanfang_bleibt_im_alten_stand(tmp_path):
    a = _einzel(tmp_path, "Die Anlage ist ab{+gesichert+} und sicher.")
    t = token(a, "abgesichert")
    assert (t.old, t.new, t.status) == ("ab", "abgesichert", MOD)
    assert a.blocks[0].alt_text == "Die Anlage ist ab und sicher."


def test_teilmarkierung_am_wortende_bleibt_ein_aendern(tmp_path):
    a = _einzel(tmp_path, "Die Rack{+s+} werden geprueft und {-Modul-}{+Module+} ebenso.")
    t = token(a, "Racks")
    assert (t.old, t.new, t.status) == ("Rack", "Racks", MOD)
    t2 = token(a, "Module")
    assert (t2.old, t2.new) == ("Modul", "Module")


def test_gleich_lange_ersetzung_bleibt_ein_aendern(tmp_path):
    a = _einzel(tmp_path, "Es sind {-100-}{+120+} kWh.")
    t = token(a, "120")
    assert (t.old, t.new, t.status) == ("100", "120", MOD)


def test_satzzeichen_am_markierten_wort_bleibt_ein_wort(tmp_path):
    a = _einzel(tmp_path, "Es sind {+120 kWh+}. Danach folgt Text.")
    t = token(a, "kWh.")
    assert t.status == ADD and t.new == "kWh." and t.old == ""


def test_linie_die_knapp_den_letzten_buchstaben_verfehlt_streicht_das_ganze_wort(tmp_path):
    from tests.fixtures.pdf_writer import breite
    y = 842 - (64 + 11 * 0.55)

    def pdf(abdeckung_bis: str) -> bytes:
        inhalt = zeile(56, 64, "Wartung") + "0 G 0.8 w 56 %.2f m %.2f %.2f l S\n" % (y, 56 + breite(abdeckung_bis, 11), y)
        return roh_pdf([RohSeite(inhalt=inhalt)])
    a = lese(tmp_path / "eins", pdf("Wartun"))                  # ein Buchstabe fehlt: Linie knapp daneben
    assert (token(a, "Wartung").old, token(a, "Wartung").status) == ("Wartung", DEL)
    a = lese(tmp_path / "zwei", pdf("Wart"))                    # drei Buchstaben fehlen: bleibt Teilmarkierung
    t = alle_tokens(a)[0]
    assert t.status == MOD and t.old == "Wartung" and t.new == "ung" and "teilweise_markiert" in t.hinweis
