"""Red-Team R2 (PDF-Leser), Teil 2: Farben, Markierungsformen, Annotationen (Faelle f, g, h).

Alles ERFUNDEN (Regel 16), nichts in src/ geaendert. Konvention siehe test_pdf_leser.py:
`@pytest.mark.redteam` + `@pytest.mark.xfail(strict=True, raises=AssertionError, reason="RT-nn ...; Besitzer K-PDF")`.

Massstab fuer Farben (g): Eine Markierung wird entweder ERKANNT oder als Luecke `unbekannte_farbe` mit Woerterzahl
gemeldet. Die blosse Warnung ("Farbe ... nicht in der Legende") reicht nicht: sie macht die Bilanz nicht unvollstaendig.
"""
import pytest

from src.models import ADD, DEL
from tests.fixtures.pdf_writer import PdfDokument
from tests.redteam._bau import aufbau, scheitern
from tests.redteam._bau_pdf import (RohSeite, SEITE_H, alle_tokens, annot, annot_gedreht90, dok, geaendert, hinweis_bei, kurz,
                                    lese, lueckenwoerter, meldungen, roh_pdf, token, tokens_mit, woerter, zeile,
                                    zeile_gedreht90)

ZIEL = ("neuer", "Text", "hier")
SATZ_NEU = "Das ist {+neuer Text hier+} im Satz."


def _ziel_tokens(a):
    ts = [token(a, w) for w in ZIEL]
    aufbau(all(ts), f"Die Woerter {ZIEL} wurden nicht gelesen")
    return ts


def _erkannt_oder_luecke(a, was: str) -> None:
    """Erfolg: alle drei Zielwoerter als Aenderung erkannt ODER Luecke `unbekannte_farbe` mit Woerterzahl."""
    ts = _ziel_tokens(a)
    if all(geaendert(t) for t in ts):
        return
    if "unbekannte_farbe" in {l["art"] for l in a.luecken} and lueckenwoerter(a, "unbekannte_farbe") > 0:
        return
    scheitern(f"{was}: nicht erkannt (Status {[t.status for t in ts]}) und keine Luecke 'unbekannte_farbe' "
              f"(Luecken: {[l['art'] for l in a.luecken]}; nur Warnungen: {[w[:40] for w in a.warnungen]})")


# ==================================================================================================
# f) Blau, das KEINE Aenderung ist (Design-Farben)
# ==================================================================================================
@pytest.mark.redteam
def test_rt76_blauer_tabellenkopf_ist_keine_aenderung(tmp_path):
    d = dok()
    d.block("Vorspann ohne Aenderung.")
    d.tabelle([["Name", "Wert", "Einheit"], ["Strom", "5", "kW"], ["Spannung", "230", "V"], ["Frequenz", "50", "Hz"]],
              kopf_farbe="4472C4", band_farbe=None)
    a = lese(tmp_path, d)
    aufbau(token(a, "Spannung") is not None, "Tabelle nicht gelesen")
    neu = [t for t in woerter(a, ADD) if not t.hinweis]
    if neu and not meldungen(a):
        scheitern(f"{len(neu)} Woerter des Tabellenkopfs zaehlen still als 'neu': " + " ".join(t.new for t in neu))


@pytest.mark.redteam
def test_rt77_blaue_baenderung_ist_keine_aenderung(tmp_path):
    d = dok()
    d.block("Vorspann ohne Aenderung.")
    d.tabelle([["Name", "Wert", "Einheit"], ["Strom", "5", "kW"], ["Spannung", "230", "V"], ["Frequenz", "50", "Hz"],
               ["Druck", "2", "bar"]], kopf_farbe=None, band_farbe="B4C6E7")
    a = lese(tmp_path, d)
    aufbau(token(a, "Frequenz") is not None, "Tabelle nicht gelesen")
    neu = [t for t in woerter(a, ADD) if not t.hinweis]
    if neu and not meldungen(a):
        scheitern(f"{len(neu)} Woerter der gebaenderten Zeilen zaehlen still als 'neu': " + " ".join(t.new for t in neu))


@pytest.mark.redteam
def test_rt78_hyperlink_blau_ist_keine_aenderung(tmp_path):
    d = dok()
    d.block("Siehe {@l:www.beispiel.example@} fuer mehr Informationen zum Thema.")
    a = lese(tmp_path, d)
    t = token(a, "www.beispiel.example")
    aufbau(t is not None, "Link-Wort nicht gelesen")
    if geaendert(t) and not hinweis_bei(a, [t]) and not a.luecken:
        scheitern(f"Der Hyperlink zaehlt als Aenderung: {kurz(t)}")


@pytest.mark.redteam
def test_rt79_durchgehend_blaue_ueberschriften(tmp_path):
    blau = "0.1216 0.3059 0.6235 rg"          # 1F4E9F
    inhalt = ""
    top = 64.0
    for nr, titel in enumerate(("Allgemeines", "Anforderungen", "Betrieb", "Abnahme"), 1):
        inhalt += zeile(56, top, f"{nr} {titel}", 14.0, blau, fett=True)
        inhalt += zeile(56, top + 24, f"Text unter der Ueberschrift {nr} ohne jede Aenderung.")
        inhalt += zeile(56, top + 39, "Noch eine Zeile Text zu diesem Abschnitt.")
        top += 80
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    aufbau(token(a, "Abnahme") is not None, "Ueberschriften nicht gelesen")
    neu = [t for t in woerter(a, ADD) if not t.hinweis]
    if neu and not meldungen(a):
        scheitern(f"{len(neu)} Woerter durchgehend blauer Ueberschriften zaehlen still als 'neu': " + " ".join(t.new for t in neu))


# ==================================================================================================
# g) Farben, die nicht erkannt werden: je Farbfamilie ein Test
# ==================================================================================================
def _neu_in_farbe(tmp_path, hexfarbe: str):
    d = dok(farbe_neu=hexfarbe)
    d.block(SATZ_NEU)
    return lese(tmp_path, d)


@pytest.mark.redteam
def test_rt80_cyan(tmp_path):
    _erkannt_oder_luecke(_neu_in_farbe(tmp_path, "00FFFF"), "Cyan")


@pytest.mark.redteam
def test_rt81_violett(tmp_path):
    _erkannt_oder_luecke(_neu_in_farbe(tmp_path, "7F00FF"), "Violett")


@pytest.mark.redteam
def test_rt82_orange(tmp_path):
    d = dok(farbe_weg="FF8000")
    d.block("Das ist {@r:alter Text hier@} im Satz.")
    a = lese(tmp_path, d)
    ts = [token(a, w) for w in ("alter", "Text", "hier")]
    aufbau(all(ts), "Woerter nicht gelesen")
    if not all(geaendert(t) for t in ts) and lueckenwoerter(a, "unbekannte_farbe") == 0:
        scheitern(f"Orange Text nicht erkannt (Status {[t.status for t in ts]}), keine Luecke 'unbekannte_farbe'")


@pytest.mark.redteam
def test_rt83_pink(tmp_path):
    d = dok(farbe_weg="E6007E")
    d.block("Das ist {@r:alter Text hier@} im Satz.")
    a = lese(tmp_path, d)
    ts = [token(a, w) for w in ("alter", "Text", "hier")]
    aufbau(all(ts), "Woerter nicht gelesen")
    if not all(geaendert(t) for t in ts) and lueckenwoerter(a, "unbekannte_farbe") == 0:
        scheitern(f"Pinker Text nicht erkannt (Status {[t.status for t in ts]}), keine Luecke 'unbekannte_farbe'")


@pytest.mark.redteam
def test_rt84_dunkelblau_helligkeit_unter_025(tmp_path):
    a = _neu_in_farbe(tmp_path, "00003C")
    _erkannt_oder_luecke(a, "Dunkelblau (Helligkeit 0,24)")


@pytest.mark.redteam
def test_rt85_pastellrot_saettigung_unter_035(tmp_path):
    d = dok(farbe_weg="F5B5B5")
    d.block("Das ist {@r:alter Text hier@} im Satz.")
    a = lese(tmp_path, d)
    ts = [token(a, w) for w in ("alter", "Text", "hier")]
    aufbau(all(ts), "Woerter nicht gelesen")
    if not all(geaendert(t) for t in ts) and lueckenwoerter(a, "unbekannte_farbe") == 0:
        scheitern(f"Pastellroter Text nicht erkannt (Status {[t.status for t in ts]}), keine Luecke, Warnungen: {a.warnungen}")


def _gelb_pdf(variante: str) -> bytes:
    text = zeile(56, 64, "Das ist neuer Text hier im Satz.")
    if variante == "annotation":
        return roh_pdf([RohSeite(inhalt=text, annots=[annot("Highlight", 92, 64, 168, 77, "1 1 0")])])
    rechteck = f"1 1 0 rg 92 {SEITE_H - 77:.2f} 76 13 re f\n"
    return roh_pdf([RohSeite(inhalt=rechteck + "0 g\n" + text)])


@pytest.mark.redteam
@pytest.mark.parametrize("variante", ["annotation", "flaeche"])
def test_rt86_gelbes_highlight(tmp_path, variante):
    _erkannt_oder_luecke(lese(tmp_path, _gelb_pdf(variante)), f"Gelbes Highlight ({variante})")


@pytest.mark.redteam
def test_rt87_gruener_text(tmp_path):
    d = dok()
    d.block("Das ist {@gn:neuer Text hier@} im Satz.")
    a = lese(tmp_path, d)
    aufbau(any("009900" in w for w in a.warnungen), "Die Warnung zur gruenen Farbe fehlt (Aufbau)")
    _erkannt_oder_luecke(a, "Gruener Text")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-88 Reader repariert (Separation: Tint 1,0 = volle Farbe, geprueft in test_pdf_reader_unit.py), "
                          "der TEST selbst ist fehlerhaft: zeile(...) setzt im Textobjekt '0 0 0 rg' und ueberschreibt die "
                          "Sonderfarbe, der Text ist im PDF wirklich schwarz. Marker bleibt, bis der Test mit "
                          "zeile(..., farbe='') gebaut wird; Besitzer K-PDF/R2")
def test_rt88_spotfarbe_separation_tint_1(tmp_path):
    spot = ("/ColorSpace << /Cs1 [/Separation /FirmenBlau /DeviceCMYK "
            "<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0 0] /C1 [1 0.8 0 0] /N 1 >>] >>")
    inhalt = (zeile(56, 64, "Das ist ") + "/Cs1 cs 1 scn\n" + zeile(92, 64, "neuer Text hier") + "0 g\n"
              + zeile(168, 64, " im Satz."))
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt, ressourcen=spot)]))
    _erkannt_oder_luecke(a, "Spotfarbe (Separation, Tint 1,0)")


@pytest.mark.redteam
def test_rt89_graustufen_markierung(tmp_path):
    a = _neu_in_farbe(tmp_path, "808080")
    _erkannt_oder_luecke(a, "Mittelgrauer Text")


# ==================================================================================================
# h) Mischformen und Annotationen
# ==================================================================================================
@pytest.mark.redteam
def test_rt90_blau_plus_strich_ohne_hinweis(tmp_path):
    d = dok()
    d.block("Das ist {@bs:blau durchgestrichen@} im Satz.")
    a = lese(tmp_path, d)
    ts = [token(a, w) for w in ("blau", "durchgestrichen")]
    aufbau(all(ts), "Woerter nicht gelesen")
    if not all(t.hinweis for t in ts) and not hinweis_bei(a, ts) and not a.luecken:
        scheitern("Blau+Strich ohne Hinweis gelesen: " + "; ".join(kurz(t) for t in ts))


@pytest.mark.redteam
def test_rt91_unterstreichung_wird_gemeldet(tmp_path):
    inhalt = (zeile(56, 64, "Das ist ein wichtiger Satz mit Unterstreichung hier.")
              + "0 0 0 RG 0.8 w 110.4 %.2f m 179 %.2f l S\n" % (SEITE_H - 74.2, SEITE_H - 74.2)
              + zeile(56, 94, "Danach kommt ein ", 11.0) + zeile(152, 94, "neues", 11.0, "0 0 1 rg"))
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt)]))
    ts = [token(a, w) for w in ("wichtiger", "Satz")]
    aufbau(all(ts) and token(a, "neues") is not None and token(a, "neues").status == ADD, "Aufbau nicht gelesen")
    if not any(geaendert(t) for t in ts) and not meldungen(a, mit_warnungen=True):
        scheitern("Unterstrichene Woerter 'wichtiger Satz' bleiben unveraendert, ohne jeden Hinweis")


@pytest.mark.redteam
@pytest.mark.parametrize("art", ["Underline", "Squiggly", "Caret", "Ink", "FreeText"])
def test_rt92_annotationen_anderer_typen_werden_gemeldet(tmp_path, art):
    extra = " /DA (/Helv 0 Tf 0 g) /Contents (bitte pruefen)" if art == "FreeText" else ""
    inhalt = (zeile(56, 64, "Das ist ein wichtiger Satz mit Markierung hier.")
              + zeile(56, 94, "Danach kommt ein ") + zeile(152, 94, "neues", 11.0, "0 0 1 rg"))
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt, annots=[annot(art, 110, 64, 179, 77, "1 0 0") + extra])]))
    ts = [token(a, w) for w in ("wichtiger", "Satz")]
    aufbau(all(ts) and token(a, "neues").status == ADD, "Aufbau nicht gelesen")
    if not any(geaendert(t) for t in ts) and not meldungen(a, mit_warnungen=True):
        scheitern(f"Kommentar vom Typ {art} liegt auf 'wichtiger Satz' und wird ohne jede Meldung uebergangen")


@pytest.mark.redteam
@pytest.mark.parametrize("art,farbe", [("Highlight", "0 0 1"), ("StrikeOut", "1 0 0")])
def test_rt93_annotation_auf_gedrehter_seite(tmp_path, art, farbe):
    inhalt = (zeile_gedreht90(56, 64, "Erste Zeile bleibt gleich hier.") + zeile_gedreht90(56, 79, "Zweite Zeile ist ")
              + zeile_gedreht90(146, 79, "neuer Text") + zeile_gedreht90(56, 94, "Dritte Zeile mit ")
              + zeile_gedreht90(150, 94, "blauem", 11.0, "0 0 1 rg") + zeile_gedreht90(200, 94, "Wort."))
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt, attribute="/Rotate 90",
                                         annots=[annot_gedreht90(art, 146, 79, 220, 90, farbe)])]))
    ts = [token(a, "neuer"), token(a, "Text")]
    aufbau(all(ts) and token(a, "blauem").status == ADD, "Aufbau: aufrechter Text auf gedrehter Seite nicht gelesen")
    if not any(geaendert(t) for t in ts) and not meldungen(a, mit_warnungen=True):
        scheitern(f"{art} auf gedrehter Seite trifft die Woerter nicht ({[t.status for t in ts]}), ohne Meldung")


@pytest.mark.redteam
def test_rt94_gedrehte_seite_mit_querstehendem_text(tmp_path):
    d = dok()
    d.block("Das ist ein Satz mit mehreren Woertern und {+neuen+} Worten drin.")
    d.seite_rotieren(90)
    a = lese(tmp_path, d)
    ganz = token(a, "Woertern") is not None and token(a, "neuen") is not None
    if not ganz and not meldungen(a, mit_warnungen=True):
        scheitern(f"Gedrehte Seite wird in {len(a.blocks)} Bloecke zu je {max(len(b.tokens) for b in a.blocks)} "
                  f"Zeichen zerlegt (Wort 'Woertern' nicht lesbar), ohne Warnung oder Luecke")
