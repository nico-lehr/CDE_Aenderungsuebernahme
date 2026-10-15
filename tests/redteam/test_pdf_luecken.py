"""Red-Team R2 (PDF-Leser), Teil 3: Luecken, die gemeldet werden muessen (Faelle d, k, m und Bilder).

Alles ERFUNDEN (Regel 16), nichts in src/ geaendert. Konvention siehe test_pdf_leser.py.

Massstab (ADR 0002): Was das Tool nicht (sicher) lesen kann, wird als LUECKE gemeldet (`PdfAnalyse.luecken`,
Art aus `models.LUECKEN_ARTEN`); dann zeigt die Bilanz `vollstaendig = False`. Eine blosse Warnung genuegt nicht.
"""
import pytest

from tests.fixtures.pdf_writer import muster_bild, scan_bild
from tests.redteam._bau import aufbau, scheitern
from tests.redteam._bau_pdf import (RohSeite, alle_tokens, bloecke_mit, dok, hinweis_bei, lese, luecken_arten, roh_pdf,
                                    token, tokens_mit, vollstaendig, zeile)

LINKS = "Linkswort " + " ".join(["Netzanschluss", "Kuehlung", "Messwert", "Sicherung", "Lueftung", "Zuleitung",
                                 "Steuerung", "Abnahme", "Wartung", "Anschluss", "Spannung", "Frequenz"]) + " {+linksneu+} Linksende"
RECHTS = "Rechtswort " + " ".join(["Batterie", "Wechselrichter", "Transformator", "Schalter", "Verteiler", "Kabeltrasse",
                                   "Brandmelder", "Sensor", "Pruefung", "Dokumentation", "Protokoll", "Lieferant"]) + " Rechtsende"


def _bilanz_meldet_luecke(a, art: str, was: str) -> None:
    if art not in luecken_arten(a):
        scheitern(f"{was}: keine Luecke '{art}' gemeldet (Luecken: {sorted(luecken_arten(a))}; "
                  f"Warnungen: {[w[:45] for w in a.warnungen]}) -- die Bilanz meldet 'vollstaendig'={vollstaendig(a)}")
    if vollstaendig(a):
        scheitern(f"{was}: Luecke vorhanden, aber Bilanz meldet 'vollstaendig'")


# ==================================================================================================
# d) Zwei Spalten / Tabelle
# ==================================================================================================
@pytest.mark.redteam
def test_rt95_zwei_spalten_werden_als_luecke_gemeldet(tmp_path):
    d = dok()
    d.zwei_spalten([LINKS], [RECHTS])
    a = lese(tmp_path, d)
    aufbau(token(a, "Linkswort") is not None and token(a, "Rechtswort") is not None, "Spalten nicht gelesen")
    if not bloecke_mit(a, "Linkswort", "Rechtswort"):      # Spalten sauber getrennt gelesen waere auch richtig
        return
    _bilanz_meldet_luecke(a, "mehrspaltig", "Zwei Spalten wurden vermischt")


@pytest.mark.redteam
def test_rt96_tabelle_wird_als_luecke_gemeldet(tmp_path):
    d = dok()
    d.block("Vorspann ohne Aenderung.")
    d.tabelle([["Name", "Wert"], ["Strom {+geaendert+}", "5 kW"], ["Spannung", "230 V"], ["Frequenz", "50 Hz"]],
              kopf_farbe=None, band_farbe=None)
    a = lese(tmp_path, d)
    aufbau(token(a, "Spannung") is not None and token(a, "geaendert") is not None, "Tabelle nicht gelesen")
    _bilanz_meldet_luecke(a, "tabelle", "Tabelle")


@pytest.mark.redteam
def test_rt97_bild_im_fliesstext_wird_als_luecke_gemeldet(tmp_path):
    d = dok()
    d.block("Text vor dem Bild mit {+einer Aenderung+} darin.")
    d.bild(muster_bild(240, 140, seed=3), 200.0)
    d.block("Text nach dem Bild ohne Aenderung.")
    a = lese(tmp_path, d)
    aufbau(token(a, "Aenderung") is not None and token(a, "nach") is not None, "Texte nicht gelesen")
    _bilanz_meldet_luecke(a, "bild", "Bild")


# ==================================================================================================
# k) Scan-Seiten und unsichtbare Textebenen
# ==================================================================================================
def _scan_dok(tmp_path):
    d = dok()
    d.block("Normale Seite mit {+einer Aenderung+} im Text.")
    d.scan_seite(scan_bild([("h", "2 Scan"), ("p", "Auf der Scan-Seite steht {+neuer Text+} und {-alter Text-} nur als Pixel.")]))
    aufbau(d.seitenzahl == 2, "Scan-Seite fehlt")
    return lese(tmp_path, d)


@pytest.mark.redteam
def test_rt98_scan_seite_wird_als_luecke_gemeldet(tmp_path):
    a = _scan_dok(tmp_path)
    aufbau(any("Seite 2" in w and "kein" in w for w in a.warnungen), "Die vorhandene Warnung fehlt (Aufbau)")
    _bilanz_meldet_luecke(a, "scan_seite", "Scan-Seite")


def _seite_text(fuss: str = "") -> str:
    inh = zeile(56, 64, "Normale Seite mit ") + zeile(150, 64, "neuem", 11.0, "0 0 1 rg") + zeile(190, 64, "Text.")
    return inh + (zeile(250, 800, fuss, 9.0, "0.4 0.4 0.4 rg") if fuss else "")


BILD = "q 595 0 0 842 0 0 cm /Im1 Do Q\n"


@pytest.mark.redteam
def test_rt99_scan_seite_mit_seitenzahl_wird_als_luecke_gemeldet(tmp_path):
    pdf = roh_pdf([RohSeite(inhalt=_seite_text("Seite 1 von 2")),
                   RohSeite(inhalt=BILD + zeile(250, 800, "Seite 2 von 2", 9.0, "0.4 0.4 0.4 rg"), bild=True)])
    a = lese(tmp_path, pdf)
    aufbau(token(a, "neuem") is not None, "Seite 1 nicht gelesen")
    _bilanz_meldet_luecke(a, "scan_seite", "Bildseite mit Seitenzahl")


@pytest.mark.redteam
def test_rt100_unsichtbare_textebene_ueber_bild(tmp_path):
    pdf = roh_pdf([RohSeite(inhalt=_seite_text("Seite 1 von 2")),
                   RohSeite(inhalt=BILD + zeile(56, 64, "Unsichtbare Textebene einer gescannten Seite hier.", modus=3)
                            + zeile(250, 800, "Seite 2 von 2", 9.0, "0.4 0.4 0.4 rg"), bild=True)])
    a = lese(tmp_path, pdf)
    aufbau(token(a, "Unsichtbare") is not None, "Die unsichtbare Textebene wird nicht gelesen (Aufbau)")
    _bilanz_meldet_luecke(a, "scan_seite", "Bildseite mit unsichtbarem Text")


@pytest.mark.redteam
def test_rt101_weisse_textebene_ueber_bild(tmp_path):
    pdf = roh_pdf([RohSeite(inhalt=_seite_text("Seite 1 von 2")),
                   RohSeite(inhalt=BILD + zeile(56, 64, "Weisse Textebene einer gescannten Seite hier.", farbe="1 1 1 rg")
                            + zeile(250, 800, "Seite 2 von 2", 9.0, "0.4 0.4 0.4 rg"), bild=True)])
    a = lese(tmp_path, pdf)
    aufbau(token(a, "Weisse") is not None, "Die weisse Textebene wird nicht gelesen (Aufbau)")
    _bilanz_meldet_luecke(a, "scan_seite", "Bildseite mit weissem Text")


# ==================================================================================================
# m) Zeichenqualitaet
# ==================================================================================================
# Schrift mit eigener Zuordnung: Code 128 = Ligatur 'fi', 129 = unbekannter Glyphenname, 130 = Private-Use-Zeichen U+F0B7
_ZEICHEN_SCHRIFT = ("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Type /Encoding "
                    "/BaseEncoding /WinAnsiEncoding /Differences [128 /fi /gl1234 /uniF0B7] >> >>")


def _zeichen_pdf(satz: str) -> bytes:
    inhalt = zeile(56, 64, satz, kodierung="latin-1") + zeile(56, 94, "Danach kommt ein ") + zeile(152, 94, "neues", 11.0, "0 0 1 rg")
    return roh_pdf([RohSeite(inhalt=inhalt, schrift=_ZEICHEN_SCHRIFT)])


def _roh_zeichen(t: str) -> bool:
    return "(cid:" in t or any(0xFB00 <= ord(c) <= 0xFB06 or 0xE000 <= ord(c) <= 0xF8FF for c in t)


def _zeichen_ok(a, anfang: str) -> bool:
    ts = [t for t in tokens_mit(a, anfang) if _roh_zeichen(t.old + t.new)]
    return (not ts) or "nicht_lesbarer_text" in luecken_arten(a) or hinweis_bei(a, ts)


@pytest.mark.redteam
def test_rt102_cid_zeichen_werden_gemeldet(tmp_path):
    a = lese(tmp_path, _zeichen_pdf("Die Daten \x81 sind hier."))
    aufbau(token(a, "neues") is not None and token(a, "Daten") is not None, "Satz nicht gelesen")
    roh = [t.old for t in alle_tokens(a) if "(cid:" in t.old]
    if not _zeichen_ok(a, "(cid:"):
        scheitern(f"Roh uebernommenes Token {roh}, keine Luecke 'nicht_lesbarer_text', kein Hinweis")


@pytest.mark.redteam
def test_rt103_ligatur_wird_normalisiert_oder_gemeldet(tmp_path):
    a = lese(tmp_path, _zeichen_pdf("Die Anlage ist \x80xiert und stabil."))
    aufbau(token(a, "neues") is not None and token(a, "Anlage") is not None, "Satz nicht gelesen")
    roh = [t.old for t in alle_tokens(a) if _roh_zeichen(t.old)]
    if not _zeichen_ok(a, "xiert"):
        scheitern(f"Ligatur roh uebernommen: {[ascii(x) for x in roh]} (statt 'fixiert'), kein Hinweis, keine Luecke")


@pytest.mark.redteam
def test_rt104_private_use_zeichen_werden_gemeldet(tmp_path):
    a = lese(tmp_path, _zeichen_pdf("Die Daten \x82 sind hier."))
    aufbau(token(a, "neues") is not None and token(a, "Daten") is not None, "Satz nicht gelesen")
    roh = [t.old for t in alle_tokens(a) if _roh_zeichen(t.old)]
    if not _zeichen_ok(a, ""):
        scheitern(f"Private-Use-Zeichen {[ascii(x) for x in roh]} als Wort uebernommen, kein Hinweis, keine Luecke")
