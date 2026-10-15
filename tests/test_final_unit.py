"""Unit-Tests K-FINAL: Funde der unabhaengigen Pruefer (nur erfundene Daten, Regel 16)."""
from __future__ import annotations

from openpyxl import Workbook

from src import config
from src.excel_model import ExcelModell
from src.matcher import Matcher
from src.models import HOCH, MITTEL, PRUEFEN, UEBERNOMMEN
from src.textutil import norm
from tests.redteam._bau import analysiere, baue_alt, bearbeite, lauf, uebernimm, blatt


def _tokens(text: str) -> list[tuple[int, str]]:
    return [(i, norm(w)) for i, w in enumerate(text.split()) if norm(w)]


def _excel(tmp_path, zeilen, name="a.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "CDE"
    for z in zeilen:
        ws.append(z)
    p = tmp_path / name
    wb.save(str(p))
    return p


_ZWILLING = "Die Reaktionszeit der Anlage darf im Normalbetrieb hundertfuenfzig Millisekunden nicht ueberschreiten."
_AUFBAU = [["Nr.", "Titel", "Inhalt"],
           ["1", "Allgemeines", "Der Betrieb erfolgt jederzeit durch geschultes Personal vor Ort."],
           ["2", "Stoermeldungen", _ZWILLING],
           ["3", "Normen", "Die Norm gilt fuer alle Anlagen im Betrieb."],
           ["4", "Reaktion", "Im Normalbetrieb darf die Anlage hoechstens hundertfuenfzig Millisekunden brauchen."]]


# ===========================================================================================================
# 1. Matcher: eine falsche Section ist nie automatisch
# ===========================================================================================================
def test_zwilling_in_anderer_section_ist_nur_mittel(tmp_path):
    xl = ExcelModell(_excel(tmp_path, _AUFBAU))
    m = Matcher(xl, config.lade("zuordnung"))
    t = m.finde(_tokens(_ZWILLING), "4", ist_kopf=False)         # PDF-Section 4, der Text steht aber in Section 2
    assert t is not None and t.kandidat.section == "2" and t.section_passt is False and t.eindeutig
    assert t.stufe == MITTEL
    assert any("Section passt nicht" in g for g in t.gruende)


def test_passende_section_bleibt_hoch(tmp_path):
    xl = ExcelModell(_excel(tmp_path, _AUFBAU))
    t = Matcher(xl, config.lade("zuordnung")).finde(_tokens(_ZWILLING), "2", ist_kopf=False)
    assert t is not None and t.section_passt is True and t.stufe == HOCH


def test_ohne_section_im_pdf_bleibt_hoch(tmp_path):
    xl = ExcelModell(_excel(tmp_path, _AUFBAU))
    t = Matcher(xl, config.lade("zuordnung")).finde(_tokens(_ZWILLING), None, ist_kopf=False)
    assert t is not None and t.section_passt is None and t.stufe == HOCH


def test_schalter_section_unpassend_stufe_wirkt(tmp_path):
    xl = ExcelModell(_excel(tmp_path, _AUFBAU))
    cfg = config.lade("zuordnung")
    cfg = {**cfg, "matcher": {**cfg["matcher"], "section_unpassend_stufe": "hoch"}}
    t = Matcher(xl, cfg).finde(_tokens(_ZWILLING), "4", ist_kopf=False)
    assert t is not None and t.stufe == HOCH


_SATZ = "Die Reaktionszeit der Anlage {-darf-} {+soll+} im Normalbetrieb hundertfuenfzig Millisekunden nicht ueberschreiten."
_ELEM_ZWILLING = [
    ("h", "1 Einleitung"),
    ("p", "Der Betrieb erfolgt jederzeit durch geschultes Personal vor Ort."),
    ("h", "2 Stoermeldungen"),
    ("p", _ZWILLING),                       # Section 2 bleibt im PDF unveraendert
    ("h", "3 Normen"),
    ("p", "Die Norm gilt fuer alle Anlagen im Betrieb."),
    ("h", "4 Reaktion"),
    ("p", _SATZ),
]


def test_zwilling_in_section_2_wird_nicht_fuer_section_4_geschrieben(tmp_path):
    """Nachbau des Pruefer-Falls: Die Aenderung gehoert zu Section 4, deren Excel-Text abweicht (umformuliert);
    Section 2 hat denselben alten Satz. Vorher wurde die Aenderung automatisch in Section 2 geschrieben."""
    def umformulieren(pfad):
        def fn(wb):
            ws = wb["CDE"]
            for row in ws.iter_rows(min_row=2):
                if str(row[0].value).strip() == "4":
                    row[2].value = "Im Normalbetrieb darf die Anlage hoechstens hundertfuenfzig Millisekunden brauchen."
        bearbeite(pfad, fn)

    alt = baue_alt(tmp_path, _ELEM_ZWILLING, "L1", umformulieren)
    l = uebernimm(analysiere(tmp_path, _ELEM_ZWILLING, alt))            # ohne Freigabe
    # Der Satz in Section 2 steht in der Ausgabe unveraendert da (nur ein Mensch darf ihn aendern)
    zeilen = {str(r[0].value).strip(): r[2].value for r in blatt(l.out).iter_rows(min_row=2)}
    assert "darf im Normalbetrieb" in zeilen["2"] and "soll" not in zeilen["2"], zeilen["2"]
    zu_4 = [a for a in l.aen if a.section == "4" and a.typ != "section_neu" and "soll" in a.neu]
    assert zu_4 and all(a.status == PRUEFEN for a in zu_4), [(a.status, a.konfidenz, a.ziel) for a in l.aen]


def test_umnummerierung_bleibt_automatisch(tmp_path):
    """Die Section wird umnummeriert (alt 2.3 -> neu 2.2): der Vergleich nutzt die ALTE Nummer, es bleibt automatisch."""
    elem = [
        ("h", "1 Einleitung"),
        ("p", "Dieses Dokument beschreibt die Anforderungen an das Speichersystem fuer alle Standorte des Kunden."),
        ("h", "2.1 Leistung"),
        ("p", "Die Nennleistung betraegt 50 kW und darf kurzzeitig ueberschritten werden."),
        ("h", "{-2.3-} {+2.2+} Sicherheit"),
        ("p", "Das System muss gegen {-Ueberhitzung-} {+Ueberspannung+} geschuetzt sein und dies im Betrieb nachweisen."),
        ("h", "3 Lieferung"),
        ("p", "Die Lieferung erfolgt frei Haus und wird vom Kunden abgenommen."),
    ]
    l = lauf(tmp_path, elem)
    a = [x for x in l.aen if "Ueberspannung" in x.neu]
    assert a and a[0].status == UEBERNOMMEN and a[0].konfidenz == HOCH, [(x.typ, x.status, x.grund) for x in l.aen]


# ===========================================================================================================
# 2. Applier: Gruppen-Mitfreigabe nur durch einen Menschen; keine Ueberschreibung; kein Ersetzen von plan.json
# ===========================================================================================================
_ELEM_GRUPPE = [
    ("h", "1 Einleitung"),
    ("p", "Dieses Dokument beschreibt die Anforderungen an das Speichersystem fuer alle Standorte."),
    ("h", "{+2 Dokumentation+}"),
    ("p", "{+Alle Unterlagen werden in deutscher Sprache geliefert.+}"),
    ("p", "{+Pruefberichte liegen dem Angebot bei.+}"),
    ("h", "3 Lieferung"),
    ("p", "Die Lieferung erfolgt frei Haus und wird vom Kunden abgenommen."),
]


def _gruppe(tmp_path):
    alt = baue_alt(tmp_path, _ELEM_GRUPPE, "L1")
    l = analysiere(tmp_path, _ELEM_GRUPPE, alt)
    kopf = next(a for a in l.aen if a.typ == "section_neu")
    folgen = [a for a in l.aen if a.typ == "absatz_neu"]
    assert len(folgen) == 2
    return l, kopf, folgen


def _texte(pfad):
    return [str(c.value) for row in blatt(pfad).iter_rows() for c in row if c.value is not None]


def test_gemischte_gruppe_ohne_freigabe_wird_nicht_mitgenommen(tmp_path):
    """Handgebauter Plan: neue Section 'uebernommen', ein Folgeabsatz 'pruefen', keine Freigabe -> nichts wird geschrieben."""
    l, kopf, folgen = _gruppe(tmp_path)
    for a in [kopf] + folgen:
        a.status = UEBERNOMMEN
    folgen[1].status = PRUEFEN
    uebernimm(l)                                              # keine Freigaben
    assert not any("Pruefberichte" in t or "Dokumentation" in t for t in _texte(l.out))
    assert [a.status for a in [kopf] + folgen] == [PRUEFEN, PRUEFEN, PRUEFEN]
    assert all(a.hinweise for a in [kopf] + folgen if a is not folgen[1]), "Hinweis fehlt"


def test_gemischte_gruppe_mit_menschlicher_freigabe_kommt_zusammen(tmp_path):
    l, kopf, folgen = _gruppe(tmp_path)
    for a in [kopf] + folgen:
        a.status = UEBERNOMMEN
    folgen[1].status = PRUEFEN
    uebernimm(l, freigaben={folgen[1].id})
    texte = " ".join(_texte(l.out))
    assert "Pruefberichte" in texte and "Dokumentation" in texte and "Unterlagen" in texte
    assert all(a.status == UEBERNOMMEN for a in [kopf] + folgen)


def test_uebernehmen_ueberschreibt_nie_eine_vorhandene_zieldatei(tmp_path):
    import pytest
    from src.applier import uebernehmen
    l, kopf, folgen = _gruppe(tmp_path)
    ziel = tmp_path / "ergebnis.xlsx"
    ziel.write_bytes(b"alter Inhalt")
    with pytest.raises(FileExistsError, match="nicht überschrieben"):
        uebernehmen(l.aen, l.alt, ziel)
    assert ziel.read_bytes() == b"alter Inhalt"
