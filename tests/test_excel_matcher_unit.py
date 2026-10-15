"""Einzeltests fuer excel_model.py und matcher.py (K-XL). Nur erfundene Daten (Regel 16).

Sie halten fest: das Verhalten mit den Standardwerten aus config/zuordnung.yaml, jede Reparatur der Red-Team-Faelle
RT-27..RT-38 im Kleinen und die neue Schnittstelle ExcelModell.einfuege_risiken().
"""
import copy
import time

import pytest
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

from src import config
from src.excel_model import ExcelModell, zahl_als_text
from src.matcher import Matcher
from src.models import HOCH, MITTEL
from src.textutil import norm

CFG = config.lade("zuordnung")


# ------------------------------------------------------------------ Hilfen
def excel(tmp_path, zeilen, blatt="CDE", name="t.xlsx", vorher=None):
    wb = Workbook()
    ws = wb.active
    ws.title = blatt
    for z in zeilen:
        ws.append(z)
    if vorher:
        vorher(wb)
    p = tmp_path / name
    wb.save(str(p))
    return p


def sections(texte, kopf=("Nr.", "Titel", "Inhalt")):
    """Standardaufbau: eine Zeile je Section, Inhalt in Spalte C."""
    return [list(kopf)] + [[str(i), f"Titel {i}", t] for i, t in enumerate(texte, 1)]


def block(text):
    return [(i, norm(w)) for i, w in enumerate(text.split()) if norm(w)]


def finde(xl, text, section=None, cfg=CFG, ist_kopf=False):
    return Matcher(xl, cfg).finde(block(text), section, ist_kopf)


def cfg_mit(**aenderungen):
    c = copy.deepcopy(CFG)
    for pfad, wert in aenderungen.items():
        abschnitt, schluessel = pfad.split("__")
        c[abschnitt][schluessel] = wert
    return c


S1 = "Alpha beta gamma delta epsilon zeta eta theta"          # 8 Woerter


# ------------------------------------------------------------------ Standardwerte (Verhalten unveraendert)
def test_standardwerte_in_der_config_sind_die_bisherigen_festen_werte():
    assert CFG["matcher"]["gewicht_abdeckung"] == 0.9 and CFG["matcher"]["gewicht_dice"] == 0.1
    assert CFG["matcher"]["section_bonus"] == 0.03 and CFG["matcher"]["untergrenze"] == 0.4
    assert CFG["matcher"]["max_kandidaten_section"] == 400 and CFG["kandidaten_pro_block"] == 10
    assert CFG["excel"]["modell_schwelle"] == 0.6 and CFG["excel"]["titel_max_zeichen"] == 160


def test_score_mit_standardwerten(tmp_path):
    xl = ExcelModell(excel(tmp_path, sections([S1, "Ganz andere Worte hier unten stehen noch"])))
    t = finde(xl, "alpha beta XX delta epsilon zeta")        # 6 Woerter, 5 passen, Kandidat hat 8
    cov, dice = 5 / 6, 2 * 5 / (6 + 8)
    assert t.score == pytest.approx(0.9 * cov + 0.1 * dice)
    assert t.abdeckung == pytest.approx(cov) and t.excel_abdeckung == pytest.approx(5 / 8)


def test_gewichte_kommen_aus_der_config(tmp_path):
    xl = ExcelModell(excel(tmp_path, sections([S1, "Ganz andere Worte hier unten stehen noch"])))
    t = finde(xl, "alpha beta XX delta epsilon zeta", cfg=cfg_mit(matcher__gewicht_abdeckung=1.0, matcher__gewicht_dice=0.0))
    assert t.score == pytest.approx(5 / 6)


def test_untergrenze_kommt_aus_der_config(tmp_path):
    xl = ExcelModell(excel(tmp_path, sections([S1, "Ganz andere Worte hier unten stehen noch"])))
    text = "alpha beta xx yy"                                   # 2 von 4 -> Score ca. 0.48
    assert finde(xl, text) is not None
    assert finde(xl, text, cfg=cfg_mit(matcher__untergrenze=0.5)) is None


def test_section_bonus_kommt_aus_der_config(tmp_path):
    # zwei Sections mit fast gleichem Text; der Bonus entscheidet nur, wenn er gross genug ist
    xl = ExcelModell(excel(tmp_path, sections(["Das System speichert alle Daten sicher ab hier",
                                               "Das System speichert alle Daten sicher ab dort"])))
    text = "Das System speichert alle Daten sicher ab dort"
    t = finde(xl, text, section="2")
    assert t.kandidat.section == "2" and t.section_passt is True and t.eindeutig
    t0 = finde(xl, text, section="1", cfg=cfg_mit(matcher__section_bonus=0.0))
    assert t0.kandidat.section == "2"                              # ohne Bonus gewinnt der bessere Text


def test_modell_schwelle_und_titel_max_aus_der_config(tmp_path, monkeypatch):
    zeilen = [["Nr.", "Titel", "Inhalt"], ["1", "Eins", "Erster Abschnitt mit genug Text hier drin."],
              ["2", "Zwei", None], ["3", "Drei", "Dritter Abschnitt mit genug Text hier drin."], ["4", "Vier", None]]
    p = excel(tmp_path, zeilen)
    assert ExcelModell(p).layouts["CDE"].modell == "zeile_pro_absatz"           # 0.5 < 0.6
    monkeypatch.setattr(config, "lade", lambda name: cfg_mit(excel__modell_schwelle=0.4))
    assert ExcelModell(p).layouts["CDE"].modell == "zeile_pro_section"
    monkeypatch.setattr(config, "lade", lambda name: cfg_mit(excel__titel_max_zeichen=3))
    assert ExcelModell(p).layouts["CDE"].titel_spalte is None


# ------------------------------------------------------------------ RT-27: doppelte Absaetze
def test_gleicher_absatz_zweimal_in_einer_zelle_ist_nicht_eindeutig(tmp_path):
    satz = "Die Anlage wird vor dem Versand einzeln geprüft."
    xl = ExcelModell(excel(tmp_path, sections([f"{satz}\n{satz}", "Zweiter Absatz mit genug Text darin."])))
    t = finde(xl, satz)
    assert t.score >= 0.9 and t.eindeutig is False and t.stufe == MITTEL
    assert any("mehrmals" in g for g in t.gruende)


def test_verschiedene_absaetze_in_einer_zelle_bleiben_eindeutig(tmp_path):
    xl = ExcelModell(excel(tmp_path, sections(["Erster Absatz handelt von Akkus.\nZweiter Absatz handelt von Kabeln und Steckern.",
                                            "Dritter Abschnitt mit genug Text hier drin."])))
    t = finde(xl, "Zweiter Absatz handelt von Kabeln und Steckern.")
    assert t.eindeutig and t.stufe == HOCH


# ------------------------------------------------------------------ RT-28: Anfangs-Treffer
LANG = ("Die Anlage wird vor dem Versand einzeln geprüft und danach dokumentiert; anschließend wird ein "
        "Prüfbericht erstellt und dem Kunden übergeben.")


def test_anfangs_treffer_taugt_nicht_als_anker_fuer_neuen_absatz(tmp_path):
    xl = ExcelModell(excel(tmp_path, sections([LANG, "Zweiter Absatz mit genug Text darin."])))
    t = finde(xl, "Die Anlage wird vor dem Versand einzeln geprüft und danach")
    assert t.stufe == HOCH                       # ein Wort darin zu aendern ist sicher
    assert t.anker_stufe == MITTEL and t.anker_gruende
    assert t.excel_abdeckung < CFG["matcher"]["excel_abdeckung_min"]


def test_ein_wort_geaendert_bleibt_als_anker_hoch(tmp_path):
    xl = ExcelModell(excel(tmp_path, sections([LANG, "Zweiter Absatz mit genug Text darin."])))
    t = finde(xl, LANG)
    assert t.stufe == HOCH and t.anker_stufe == HOCH and t.excel_abdeckung == pytest.approx(1.0)


# ------------------------------------------------------------------ RT-29 / RT-30: Section-Zahlen
@pytest.mark.parametrize("wert,format,erwartet", [
    (3, "General", "3"), (3.0, "General", "3"), (2.1, "General", "2.1"), (2.1, "0.0", "2.1"),
    (2.1, "0.00", "2.10"), (2.15, "0.0", "2.2"), (0.1 + 0.2, "General", "0.3"), (1234567.5, "General", "1234567.5")])
def test_zahl_als_text(wert, format, erwartet):
    assert zahl_als_text(wert, format) == erwartet


def _zahlen(tmp_path, f3="0.0", f4="0.00"):
    def vorher(wb):
        ws = wb.active
        ws["A3"].number_format, ws["A4"].number_format = f3, f4
    texte = ["Erster Abschnitt mit genug Text hier drin.", "Abschnitt zwei Punkt eins mit viel Text hier.",
             "Abschnitt zwei Punkt zehn mit anderem Text hier.", "Abschnitt zwei Punkt zwei mit noch mehr Text.",
             "Dritter Abschnitt mit genug Text hier drin."]
    zeilen = [["Nr.", "Titel", "Inhalt"]] + [[nr, f"Titel {i}", tx] for i, (nr, tx) in
                                              enumerate(zip([1, 2.1, 2.1, 2.2, 3], texte), 1)]
    return excel(tmp_path, zeilen, vorher=vorher)


def test_section_zahlen_2_1_und_2_10_bleiben_verschieden(tmp_path):
    xl = ExcelModell(_zahlen(tmp_path))
    assert xl.zeilen_section[("CDE", 3)] == "2.1" and xl.zeilen_section[("CDE", 4)] == "2.10"
    assert xl.zeilen_section[("CDE", 5)] == "2.2"
    assert xl.ist_zahl("CDE", 4, 1) and not xl.ist_zahl("CDE", 4, 2)
    assert any("als Zahlen" in w for w in xl.warnungen)


def test_kandidat_weiss_dass_die_zelle_eine_zahl_war(tmp_path):
    xl = ExcelModell(_zahlen(tmp_path))
    kopf = [k for k in xl.kandidaten if k.art == "kopf" and k.zeile == 5][0]
    inhalt = [k for k in xl.kandidaten if k.art == "absatz" and k.zeile == 5 and k.spalte == 3][0]
    assert kopf.zelle_ist_zahl is True and inhalt.zelle_ist_zahl is False


def test_gleich_aussehende_section_zahlen_sind_unsicher_und_nie_hoch(tmp_path):
    xl = ExcelModell(_zahlen(tmp_path, f3="General", f4="General"))         # beide zeigen "2.1"
    lay = xl.layouts["CDE"]
    assert lay.unsichere_sections == {"2.1"} and any("mehrfach" in w for w in xl.warnungen)
    t = finde(xl, "Abschnitt zwei Punkt eins mit viel Text hier.", section="2.1")
    assert t.kandidat.section_unsicher and t.section_passt is None
    assert t.stufe == MITTEL and any("Zahl" in g for g in t.gruende)


# ------------------------------------------------------------------ RT-31..33: Aufbau mit Sicherheit
KOMMENTAR = "Das ist ein langer interner Kommentar des Bearbeiters zu diesem Punkt mit vielen Details und Hinweisen darin."
TEXTE3 = ["Erster Abschnitt mit genug Text hier drin.", "Zweiter Abschnitt mit genug Text hier drin.",
          "Dritter Abschnitt mit genug Text hier drin."]


def _mit_kommentar(tmp_path, kopf4):
    zeilen = [["Nr.", "Titel", "Inhalt", kopf4]] + [[str(i), f"Titel {i}", t, KOMMENTAR] for i, t in enumerate(TEXTE3, 1)]
    return ExcelModell(excel(tmp_path, zeilen))


def test_sauberer_standardaufbau_ist_hoch_ohne_warnung(tmp_path):
    xl = ExcelModell(excel(tmp_path, sections(TEXTE3)))
    lay = xl.layouts["CDE"]
    assert lay.sicherheit == "hoch" and not lay.sicherheit_gruende and not xl.warnungen
    b = xl.beschreibung()[0]
    assert b["sicherheit"] == "hoch" and b["sicherheit_gruende"] == [] and b["blatt"] == "CDE"
    assert finde(xl, TEXTE3[1]).stufe == HOCH


def test_kommentarspalte_mit_kopfwort_ist_nicht_der_inhalt(tmp_path):
    xl = _mit_kommentar(tmp_path, "Kommentar")
    lay = xl.layouts["CDE"]
    assert lay.inhalt_spalte == 3 and lay.sicherheit == "hoch"


def test_kommentarspalte_ohne_kopfwort_macht_den_aufbau_unsicher(tmp_path):
    xl = _mit_kommentar(tmp_path, "Weiteres")
    lay = xl.layouts["CDE"]
    assert lay.sicherheit == "mittel"
    assert any("Aufbau" in w and "unsicher" in w and "Spalte" in w for w in xl.warnungen)
    t = finde(xl, TEXTE3[1])                                  # exakter Treffer, aber unsicherer Aufbau
    assert t.score >= 0.9 and t.stufe == MITTEL and any("Aufbau" in g for g in t.gruende)


def test_lfd_spalte_ist_nicht_die_section_spalte(tmp_path):
    zeilen = [["Lfd", "Nr.", "Titel", "Text"]]
    lfd = 0
    for nr, titel in (("1", "Einleitung"), ("1.1", "Ziel"), ("2", "Technik"), ("2.1", "Leistung")):
        lfd += 1
        zeilen.append([lfd, nr, titel, None])
        for k in (1, 2):
            lfd += 1
            zeilen.append([lfd, None, None, f"Absatz {k} zu {titel} mit genug Text damit er als Inhalt zaehlt."])
    xl = ExcelModell(excel(tmp_path, zeilen))
    lay = xl.layouts["CDE"]
    assert (lay.section_spalte, lay.titel_spalte, lay.inhalt_spalte) == (2, 3, 4) and lay.sicherheit == "hoch"


def test_nur_eine_lfd_spalte_ist_niedrig_mit_warnung(tmp_path):
    zeilen = [["Lfd", "Text"]] + [[i, f"Absatz {i} mit genug Text damit er als Inhalt zaehlt."] for i in range(1, 6)]
    xl = ExcelModell(excel(tmp_path, zeilen))
    assert xl.layouts["CDE"].sicherheit == "niedrig"
    assert any("laufende Nummer" in w for w in xl.warnungen)


def test_aufzaehlung_im_text_ist_keine_section(tmp_path):
    zeilen = [["Nr", "Text", "Prio"],
              ["1", "Erster Abschnitt mit genug Text hier drin.", "A"],
              [None, "1. Das System muss dies tun und noch mehr.", "B"],
              [None, "2. Das System muss das tun und noch mehr.", "B"],
              [None, "3. Das System muss jenes tun und noch mehr.", "B"],
              ["2", "Zweiter Abschnitt mit genug Text hier drin.", "A"],
              [None, "1. Das System soll etwas anderes tun und mehr.", "B"]]
    xl = ExcelModell(excel(tmp_path, zeilen))
    lay = xl.layouts["CDE"]
    assert (lay.section_spalte, lay.titel_spalte, lay.inhalt_spalte) == (1, None, 2)
    assert lay.sicherheit == "mittel"                         # Text steht teils in der Section-Zeile, teils darunter


def test_kurze_ueberschrift_mit_nummernpunkt_bleibt_eine_section(tmp_path):
    xl = ExcelModell(excel(tmp_path, [["Kapitel", "Text"], ["1. Einleitung", "Erster Abschnitt mit genug Text hier drin."],
                                      ["2. Technik", "Zweiter Abschnitt mit genug Text hier drin."]]))
    lay = xl.layouts["CDE"]
    assert lay.section_spalte == 1 and lay.kombiniert and lay.sicherheit == "hoch"


def test_blatt_ohne_gliederung_ist_niedrig_und_wird_nur_bei_viel_text_gemeldet(tmp_path):
    klein = ExcelModell(excel(tmp_path, [["Beispielprojekt"], ["Dieses Blatt gehört nicht zur Gliederung."]], name="k.xlsx"))
    assert klein.layouts["CDE"].sicherheit == "niedrig" and not klein.warnungen
    gross = ExcelModell(excel(tmp_path, [[f"Zeile {i}", f"Text zu Zeile {i} mit genug Zeichen."] for i in range(1, 6)],
                              name="g.xlsx"))
    assert gross.layouts["CDE"].sicherheit == "niedrig" and any("Gliederungsnummern" in w for w in gross.warnungen)
    assert finde(gross, "Text zu Zeile 3 mit genug Zeichen.").stufe == MITTEL        # nie automatisch


# ------------------------------------------------------------------ RT-34..36: Blaetter, versteckt
SATZ = "Das System muss alle Daten speichern und sichern."


def _zwei_blaetter(tmp_path, a, b, versteckt=()):
    def vorher(wb):
        ws = wb.create_sheet("Blatt B")
        for z in sections(b):
            ws.append(z)
        for n in versteckt:
            wb[n].sheet_state = "hidden"
    return excel(tmp_path, sections(a), blatt="Blatt A", vorher=vorher)


def test_nach_section_ist_je_blatt_getrennt(tmp_path):
    xl = ExcelModell(_zwei_blaetter(tmp_path, ["Alpha eins steht hier im Dokument.", "Anderer Inhalt zwei nur hier."],
                                    ["Beta eins steht dort im Dokument.", "Anderer Inhalt zwei nur dort."]))
    assert {b for (b, s) in xl.nach_blatt_section if s == "2"} == {"Blatt A", "Blatt B"}
    assert len(xl.nach_section["2"]) == sum(len(v) for (b, s), v in xl.nach_blatt_section.items() if s == "2")
    assert all(xl.kandidaten[i].blatt == "Blatt A" for i in xl.nach_blatt_section[("Blatt A", "2")])


def test_gleicher_text_in_zwei_blaettern_ist_nicht_eindeutig(tmp_path):
    a = ["Alpha text eins steht hier im Dokument.", "Anderer Inhalt zwei nur hier.", SATZ]
    b = ["Beta andere eins steht dort im Dokument.", SATZ, "Delta text drei steht dort im Dokument."]
    xl = ExcelModell(_zwei_blaetter(tmp_path, a, b))
    t = finde(xl, SATZ, section="2")                          # Section 2 gibt es nur in Blatt B -> trotzdem nicht sicher
    assert t.score >= 0.9 and not t.eindeutig and t.stufe == MITTEL
    assert any("Blatt" in g for g in t.gruende)


def test_section_bonus_gilt_im_eigenen_blatt(tmp_path):
    a = ["Alpha text eins steht hier im Dokument.", "Anderer Inhalt zwei nur hier.", "Gamma drei steht hier."]
    b = ["Beta andere eins steht dort im Dokument.", SATZ, "Delta text drei steht dort im Dokument."]
    xl = ExcelModell(_zwei_blaetter(tmp_path, a, b))
    t = finde(xl, SATZ, section="2")
    assert t.kandidat.blatt == "Blatt B" and t.eindeutig and t.stufe == HOCH


def test_verstecktes_blatt_nie_automatisch(tmp_path):
    a = ["Sichtbar eins steht hier irgendwo.", "Sichtbar zwei steht hier irgendwo anders.", "Sichtbar drei steht hier noch."]
    b = ["Alpha text eins steht hier im Dokument.", SATZ, "Gamma text drei steht hier im Dokument."]
    xl = ExcelModell(_zwei_blaetter(tmp_path, a, b, versteckt=("Blatt B",)))
    assert xl.layouts["Blatt B"].versteckt and any("ausgeblendet" in w for w in xl.warnungen)
    t = finde(xl, SATZ, section="2")
    assert t.kandidat.versteckt == "blatt" and t.stufe == MITTEL and any("ausgeblendet" in g for g in t.gruende)


def test_versteckte_zeile_und_spalte_nie_automatisch(tmp_path):
    def vorher(wb):
        wb.active.row_dimensions[3].hidden = True
    xl = ExcelModell(excel(tmp_path, sections(["Alpha text eins steht hier im Dokument.", SATZ, "Gamma drei steht hier."]),
                           vorher=vorher))
    assert 3 in xl.versteckte_zeilen["CDE"] and any("ausgeblendete Zeilen" in w for w in xl.warnungen)
    t = finde(xl, SATZ)
    assert t.kandidat.versteckt == "zeile" and t.stufe == MITTEL
    assert finde(xl, "Alpha text eins steht hier im Dokument.").stufe == HOCH

    def vorher2(wb):
        wb.active.column_dimensions["C"].hidden = True
    xl2 = ExcelModell(excel(tmp_path, sections([SATZ, "Zweiter Absatz mit genug Text hier.", "Dritter genug Text hier."]),
                            name="s.xlsx", vorher=vorher2))
    assert 3 in xl2.versteckte_spalten["CDE"]
    assert finde(xl2, SATZ).kandidat.versteckt == "spalte" and finde(xl2, SATZ).stufe == MITTEL


# ------------------------------------------------------------------ RT-37 / RT-38: Groesse
def test_eine_formatierte_zelle_weit_unten_kostet_fast_nichts(tmp_path):
    def vorher(wb):
        wb.active.cell(60000, 40).fill = PatternFill("solid", fgColor="FFFF00")
    p = excel(tmp_path, sections(TEXTE3), vorher=vorher)
    t0 = time.time()
    xl = ExcelModell(p)
    assert time.time() - t0 < 2
    lay = xl.layouts["CDE"]
    assert lay.max_zeile == 4 and lay.max_spalte == 3
    assert len(xl.zeilen_section) == 4 and len(xl.zellen) == 12


def test_gleiche_datei_wird_nur_einmal_gelesen_aber_nie_veraltet(tmp_path):
    p = excel(tmp_path, sections(TEXTE3))
    a = ExcelModell(p)
    assert ExcelModell(p) is a
    excel(tmp_path, sections(TEXTE3[:2]))                    # gleiche Datei, anderer Inhalt
    b = ExcelModell(p)
    assert b is not a and len(b.layouts["CDE"].kopf_zeilen) == 2


def test_kandidaten_tokens_sind_wie_vorher(tmp_path):
    xl = ExcelModell(excel(tmp_path, sections(["Erster  Satz hier.\nZweiter Satz da.", "Zweiter Abschnitt mit genug Text hier drin."])))
    k = [k for k in xl.kandidaten if k.art == "zelle"][0]
    assert [t.text for t in k.tokens] == ["Erster", "Satz", "hier.", "Zweiter", "Satz", "da."]
    assert [t.norm for t in k.tokens] == k.norms and k.tokens[0].start == 0 and k.tokens[0].end == 6
    kopf = [k for k in xl.kandidaten if k.art == "kopf"][0]
    assert [(t.spalte, t.text) for t in kopf.tokens] == [(1, "1"), (2, "Titel"), (2, "1")]


# ------------------------------------------------------------------ einfuege_risiken
def _eigenschaften(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sauber"
    for z in sections(TEXTE3):
        ws.append(z)

    ws = wb.create_sheet("Merge")
    ws.append(["a", "b"])
    ws.merge_cells("A3:B5")

    ws = wb.create_sheet("Formel")
    ws["B2"] = "=SUM(A1:A9)"
    ws["A20"] = 1
    ws = wb.create_sheet("Anders")
    ws["A1"] = "='Formel'!A5+1"

    ws = wb.create_sheet("Bedingt")
    ws["A1"] = 1
    ws.conditional_formatting.add("A1:A20", CellIsRule(operator="greaterThan", formula=["5"], fill=PatternFill("solid", bgColor="FF0000")))

    ws = wb.create_sheet("Pruefung")
    dv = DataValidation(type="list", formula1='"a,b"')
    ws.add_data_validation(dv)
    dv.add("B1:B20")

    ws = wb.create_sheet("Filter")
    ws.append(["x", "y"])
    ws.append([1, 2])
    ws.auto_filter.ref = "A1:B20"

    ws = wb.create_sheet("Tab")
    ws.append(["Name", "Wert"])
    for i in range(8):
        ws.append([f"n{i}", i])
    ws.add_table(Table(displayName="Werte", ref="A1:B9"))

    ws = wb.create_sheet("Namen")
    ws["A1"] = 1
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names["Bereich"] = DefinedName("Bereich", attr_text="Namen!$A$1:$A$20")

    ws = wb.create_sheet("Links")
    ws["A10"] = "Link"
    ws["A10"].hyperlink = "https://example.org/"

    ws = wb.create_sheet("Gruppe")
    for i in range(1, 11):
        ws.cell(i, 1, i)
    ws.row_dimensions[8].hidden = True
    ws.row_dimensions[9].outlineLevel = 1

    ws = wb.create_sheet("Bild")
    ws["A1"], ws["A2"] = "x", 1
    ch = BarChart()
    ch.add_data(Reference(ws, min_col=1, min_row=1, max_row=2), titles_from_data=True)
    ws.add_chart(ch, "E10")
    p = tmp_path / "eig.xlsx"
    wb.save(str(p))
    return ExcelModell(p)


def _hat(risiken, wort):
    return any(wort in r for r in risiken)


def test_einfuege_risiken_sauberes_blatt_ist_unbedenklich(tmp_path):
    xl = _eigenschaften(tmp_path)
    assert xl.einfuege_risiken("Sauber", 2) == [] and xl.einfuege_risiken("Sauber", 99) == []


def test_einfuege_risiken_unbekanntes_blatt(tmp_path):
    xl = _eigenschaften(tmp_path)
    r = xl.einfuege_risiken("Gibt es nicht", 1)
    assert len(r) == 1 and "gibt es in der Excel nicht" in r[0]


def test_einfuege_risiken_verbundene_zellen(tmp_path):
    xl = _eigenschaften(tmp_path)
    assert _hat(xl.einfuege_risiken("Merge", 4), "durch die Zeile")          # mitten durch A3:B5
    assert _hat(xl.einfuege_risiken("Merge", 1), "Unterhalb sind verbundene")
    assert not _hat(xl.einfuege_risiken("Merge", 6), "verbundene")           # Bereich liegt darueber


def test_einfuege_risiken_formeln(tmp_path):
    xl = _eigenschaften(tmp_path)
    assert _hat(xl.einfuege_risiken("Formel", 1), "Formeln")                  # Formel in Zeile 2 liegt darunter
    assert _hat(xl.einfuege_risiken("Formel", 5), "Formeln")                  # SUM(A1:A9) reicht bis Zeile 9
    assert not _hat(xl.einfuege_risiken("Formel", 30), "Formeln")             # alles darueber, kein Bezug auf Zeile >= 30
    assert _hat(xl.einfuege_risiken("Formel", 4), "Formeln")                  # Bezug aus anderem Blatt (Anders!A1 -> Formel!A5)


def test_einfuege_risiken_bezug_aus_anderem_blatt(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Daten"
    ws["A1"] = 1
    ws2 = wb.create_sheet("Rechnung")
    ws2["A1"] = "=Daten!A30"
    p = tmp_path / "bezug.xlsx"
    wb.save(str(p))
    xl2 = ExcelModell(p)
    assert _hat(xl2.einfuege_risiken("Daten", 10), "Formeln")                 # Rechnung zeigt auf Daten!A30
    assert not xl2.einfuege_risiken("Daten", 31)                              # Zeile 31 liegt unter dem Bezug


def test_einfuege_risiken_bedingte_formatierung_pruefung_filter_tabelle(tmp_path):
    xl = _eigenschaften(tmp_path)
    assert _hat(xl.einfuege_risiken("Bedingt", 5), "bedingte Formatierung") and not xl.einfuege_risiken("Bedingt", 21)
    assert _hat(xl.einfuege_risiken("Pruefung", 5), "Datenüberprüfung") and not xl.einfuege_risiken("Pruefung", 21)
    assert _hat(xl.einfuege_risiken("Filter", 5), "Autofilter") and not xl.einfuege_risiken("Filter", 21)
    assert _hat(xl.einfuege_risiken("Tab", 5), "Excel-Tabelle 'Werte'") and not xl.einfuege_risiken("Tab", 10)


def test_einfuege_risiken_namen_links_gruppen_bilder(tmp_path):
    xl = _eigenschaften(tmp_path)
    assert _hat(xl.einfuege_risiken("Namen", 5), "benannte Bereich 'Bereich'") and not xl.einfuege_risiken("Namen", 21)
    assert _hat(xl.einfuege_risiken("Links", 5), "Hyperlinks") and not xl.einfuege_risiken("Links", 11)
    assert _hat(xl.einfuege_risiken("Gruppe", 5), "gruppierte Zeilen") and not xl.einfuege_risiken("Gruppe", 11)
    assert _hat(xl.einfuege_risiken("Bild", 5), "Bilder oder Diagramme") and not xl.einfuege_risiken("Bild", 11)


def test_einfuege_risiken_aendert_nichts(tmp_path):
    xl = _eigenschaften(tmp_path)
    vorher = (list(xl.warnungen), xl.beschreibung(), len(xl.kandidaten))
    for blatt in list(xl.layouts):
        for z in (1, 5, 50):
            xl.einfuege_risiken(blatt, z)
    assert (list(xl.warnungen), xl.beschreibung(), len(xl.kandidaten)) == vorher
