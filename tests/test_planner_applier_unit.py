"""Unit-Tests K-PL (Planen / Anwenden / Protokoll): je eine Reparatur, nur erfundene Daten (Regel 16).

Die Red-Team-Tests (tests/redteam/test_pl_planer_applier.py) pruefen die ganzen Faelle. Hier stehen die kleinen
Bausteine und die Randfaelle, die dort nicht vorkommen.
"""
from __future__ import annotations

import copy

import pytest
from openpyxl import Workbook, load_workbook

from src import config
from src.applier import _wende_text_an, _wortgrenze, uebernehmen
from src.models import (NICHT, PRUEFEN, UEBERNOMMEN, WIRKUNG_ABWEICHUNG, WIRKUNG_KEINE, WIRKUNG_OK, Aenderung, Operation,
                        luecke)
from src.planner import neutext_probleme
from src.protokoll import bilanz, bilanz_satz, schreibe_xlsx
from tests.redteam._bau import (alle_texte, analysiere, baue_alt, blatt, lauf, schreibe_excel, uebernimm, zelle,
                                zeile_mit)


def _basis_excel(tmp_path, text="Das System muss laufen und stabil bleiben."):
    return schreibe_excel(tmp_path / "a.xlsx", [["Nr.", "Titel", "Inhalt"], ["1", "Eins", text],
                                                ["2", "Zwei", "Zweiter Absatz mit genug Text darin."]])


def _ae(aid, ops, status=UEBERNOMMEN, typ="text_ersetzt", alt="", neu="x", ziel="CDE!C2", **kw):
    return Aenderung(id=aid, typ=typ, seite=1, bbox=(0, 0, 1, 1), section="1", alt=alt, neu=neu, woerter=1,
                     status=status, konfidenz="hoch", ops=ops, ziel=ziel, **kw)


# ---------------------------------------------------------------------------------------------
# 1. Wirkung: nie "uebernommen" ohne nachweisbare Wirkung
# ---------------------------------------------------------------------------------------------
def test_wirkung_ok_wird_gesetzt_und_nachkontrolle_ok(tmp_path):
    p = _basis_excel(tmp_path)
    rep = Operation("ersetzen", "CDE", 2, 3, 4, 10, "Gerät", 1, aenderung="A1", erwartet="System")
    a = _ae("A1", [rep], alt="System", neu="Gerät")
    uebernehmen([a], p, tmp_path / "out.xlsx")
    assert zelle(tmp_path / "out.xlsx", 2, 3) == "Das Gerät muss laufen und stabil bleiben."
    assert (a.status, a.wirkung, a.nachkontrolle) == (UEBERNOMMEN, WIRKUNG_OK, "ok")


def test_operation_nur_mit_leerzeichen_hat_keine_wirkung_und_ist_nie_uebernommen(tmp_path):
    p = _basis_excel(tmp_path)
    ins = Operation("einfuegen", "CDE", 2, 3, 3, 3, " ", 1, aenderung="A1")
    a = _ae("A1", [ins], typ="text_eingefuegt", neu="")
    uebernehmen([a], p, tmp_path / "out.xlsx")
    assert a.wirkung == WIRKUNG_KEINE
    assert a.status == NICHT and a.grund and a.hinweise
    assert a.nachkontrolle == "abweichung"


def test_aenderung_ohne_operationen_ist_nie_uebernommen(tmp_path):
    p = _basis_excel(tmp_path)
    a = _ae("A1", [], typ="absatz_neu")
    uebernehmen([a], p, tmp_path / "out.xlsx")
    assert a.status == NICHT and a.wirkung == WIRKUNG_KEINE


def test_einfuegung_mitten_im_wort_ist_abweichung_und_pruefen(tmp_path):
    p = _basis_excel(tmp_path, "Das Gehäuse steht hier.")
    ins = Operation("einfuegen", "CDE", 2, 3, 8, 8, " s", 1, aenderung="A1")      # mitten in "Gehäuse"
    a = _ae("A1", [ins], typ="text_eingefuegt", neu="s")
    res = uebernehmen([a], p, tmp_path / "out.xlsx")
    assert a.wirkung == WIRKUNG_ABWEICHUNG and a.status == PRUEFEN and a.konfidenz == "niedrig"
    assert a.hinweise and res.warnungen


def test_wortgrenze():
    assert _wortgrenze("ab cd", 0) and _wortgrenze("ab cd", 2) and _wortgrenze("ab cd", 3) and _wortgrenze("ab cd", 5)
    assert not _wortgrenze("abcd", 2)


def test_zwei_aenderungen_in_einer_zelle_beide_mit_wirkung(tmp_path):
    p = _basis_excel(tmp_path)
    r1 = Operation("ersetzen", "CDE", 2, 3, 4, 10, "Gerät", 1, aenderung="A1", erwartet="System")
    r2 = Operation("ersetzen", "CDE", 2, 3, 11, 15, "darf", 2, aenderung="A2", erwartet="muss")
    a1, a2 = _ae("A1", [r1]), _ae("A2", [r2])
    uebernehmen([a1, a2], p, tmp_path / "out.xlsx")
    assert zelle(tmp_path / "out.xlsx", 2, 3) == "Das Gerät darf laufen und stabil bleiben."
    assert a1.wirkung == a2.wirkung == WIRKUNG_OK


def test_abgedeckte_aenderung_in_geleerter_zelle_bleibt_uebernommen(tmp_path):
    p = _basis_excel(tmp_path)
    leer = Operation("zelle_leeren", "CDE", 2, 3, aenderung="A1",
                     erwartet="Das System muss laufen und stabil bleiben.")
    rep = Operation("ersetzen", "CDE", 2, 3, 4, 10, "Gerät", 2, aenderung="A2", erwartet="System")
    a1 = _ae("A1", [leer], typ="absatz_geloescht", alt="Das System muss laufen und stabil bleiben.", neu="")
    a2 = _ae("A2", [rep])
    uebernehmen([a1, a2], p, tmp_path / "out.xlsx")
    assert zelle(tmp_path / "out.xlsx", 2, 3) == ""
    assert a1.wirkung == WIRKUNG_OK and a2.wirkung == WIRKUNG_OK


# ---------------------------------------------------------------------------------------------
# 2. Folgeabsaetze neuer Sections: Status und Freigabe gekoppelt
# ---------------------------------------------------------------------------------------------
EL_SEC = [("h", "1 Eins"), ("p", "Erster Absatz mit genug Text darin."), ("h", "2 Zwei"),
          ("p", "Bekannter Absatz mit genug Text darin."),
          ("p", "Unbekannter Absatz der in der Excel gar nicht vorkommt."),
          ("h", "{+3 Neue Section+}"), ("p", "{+Folgeabsatz der neuen Section steht hier.+}")]
EX_SEC = [e for e in EL_SEC if "Unbekannter" not in e[1]]
FOLGE = "Folgeabsatz der neuen Section steht hier."


def _sec(tmp_path, freigaben):
    alt = baue_alt(tmp_path, EX_SEC, "L1")
    l = analysiere(tmp_path, EL_SEC, alt)
    kopf = next(a for a in l.aen if a.typ == "section_neu")
    folge = next(a for a in l.aen if a.typ == "absatz_neu")
    return l, kopf, folge, {{"kopf": kopf.id, "folge": folge.id}[f] for f in freigaben}


def test_folgeabsatz_hat_markierung_statt_leerer_operationen(tmp_path):
    l, kopf, folge, _ = _sec(tmp_path, [])
    assert kopf.status == folge.status == PRUEFEN
    assert len(folge.ops) == 1 and folge.ops[0].abgedeckt and folge.ops[0].art == "zeile_einfuegen"
    assert folge.ops[0].seq == kopf.ops[0].seq and not folge.ops[0].zellen


@pytest.mark.parametrize("freigeben", [["kopf"], ["folge"], ["kopf", "folge"]])
def test_gruppe_ist_gemeinsam_wirksam(tmp_path, freigeben):
    l, kopf, folge, ids = _sec(tmp_path, freigeben)
    uebernimm(l, ids)
    assert kopf.status == folge.status == UEBERNOMMEN
    assert kopf.wirkung == folge.wirkung == WIRKUNG_OK
    assert sum(FOLGE in t for t in alle_texte(l.out)) == 1


def test_ohne_freigabe_bleibt_die_gruppe_unberuehrt(tmp_path):
    l, kopf, folge, ids = _sec(tmp_path, [])
    uebernimm(l, ids)
    assert kopf.status == folge.status == PRUEFEN
    assert not any(FOLGE in t for t in alle_texte(l.out))


def test_folgeabsatz_allein_uebernommen_ohne_section_wird_zurueckgestuft(tmp_path):
    """Alter/fremder Plan: Folgeabsatz 'uebernommen', Section nur 'pruefen' -> nichts wird geschrieben, kein falsches 'ok'."""
    l, kopf, folge, _ = _sec(tmp_path, [])
    folge.status = UEBERNOMMEN
    uebernimm(l, set())
    assert folge.status == PRUEFEN and folge.hinweise
    assert not any(FOLGE in t for t in alle_texte(l.out))


def test_schema1_folgeabsatz_ohne_markierung_wird_ueber_die_zielzeile_gekoppelt(tmp_path):
    l, kopf, folge, _ = _sec(tmp_path, ["kopf"])
    folge.ops = []                               # so sah ein Plan vor der Markierung aus
    uebernimm(l, {kopf.id})
    assert folge.status == UEBERNOMMEN and folge.wirkung == WIRKUNG_OK


def test_section_mit_steuerzeichen_nimmt_den_folgeabsatz_mit(tmp_path):
    l, kopf, folge, _ = _sec(tmp_path, ["kopf"])
    kopf.ops[0].zellen[2] = "Neue\x07 Section"
    uebernimm(l, {kopf.id})
    assert kopf.status == NICHT and folge.status == NICHT
    assert not any(FOLGE in t for t in alle_texte(l.out))


# ---------------------------------------------------------------------------------------------
# 3. Idempotenz
# ---------------------------------------------------------------------------------------------
def _zweimal(tmp_path, el, layout):
    l1 = lauf(tmp_path / "eins", el, layout)
    l2 = uebernimm(analysiere(tmp_path / "zwei", el, l1.out), name="out2.xlsx")
    return l1, l2


def test_zweiter_lauf_neue_section_wird_nicht_doppelt_angelegt(tmp_path):
    el = [("h", "1 Eins"), ("p", "Erster Absatz mit genug Text darin."), ("h", "{+2 Neue Section+}"),
          ("p", "{+Folgeabsatz der neuen Section steht hier.+}"), ("h", "3 Drei"), ("p", "Dritter Absatz mit genug Text darin.")]
    for layout in ("L1", "L3"):
        l1, l2 = _zweimal(tmp_path / layout, el, layout)
        assert all(a.status == UEBERNOMMEN for a in l1.aen), layout
        assert all(a.status == NICHT and "steht schon" in a.grund for a in l2.aen), (layout, [(a.typ, a.status, a.grund) for a in l2.aen])
        assert sum("Neue Section" in t for t in alle_texte(l2.out)) == 1


def test_zweiter_lauf_ersetzen_sagt_dass_der_neue_text_schon_da_ist(tmp_path):
    el = [("h", "1 Eins"), ("p", "Das System {-muss-} {+soll+} die Daten speichern."), ("h", "2 Zwei"),
          ("p", "Zweiter Absatz mit genug Text darin.")]
    l1, l2 = _zweimal(tmp_path, el, "L1")
    assert "steht schon" in l2.aen[0].grund and l2.aen[0].status == NICHT


def test_kurzer_doppelter_absatz_ist_nur_ein_vorschlag(tmp_path):
    """Unter `planer.doppelt_min_woerter` Woerter gilt ein gleicher Absatz nicht sicher als Doppelung -> pruefen mit Hinweis."""
    el = [("h", "1 Eins"), ("p", "Erster Absatz mit genug Text darin."), ("p", "{+Ja gut.+}"), ("h", "2 Zwei"),
          ("p", "Zweiter Absatz mit genug Text darin.")]
    alt = baue_alt(tmp_path, [("h", "1 Eins"), ("p", "Erster Absatz mit genug Text darin."), ("p", "Ja gut."), ("h", "2 Zwei"),
                              ("p", "Zweiter Absatz mit genug Text darin.")], "L1")
    l = analysiere(tmp_path, el, alt)
    neu = next(a for a in l.aen if a.typ == "absatz_neu")
    assert neu.status == PRUEFEN and any("kurzer Absatz" in h for h in neu.hinweise)


def test_neuer_absatz_der_nur_woanders_vorkommt_wird_normal_geplant(tmp_path):
    """Gleicher Text in einer ANDEREN Section ist kein Duplikat."""
    el = [("h", "1 Eins"), ("p", "Erster Absatz mit genug Text darin."), ("h", "2 Zwei"),
          ("p", "Zweiter Absatz mit genug Text darin."), ("p", "{+Erster Absatz mit genug Text darin.+}")]
    alt = baue_alt(tmp_path, [e for e in el if "{+" not in e[1]], "L3")
    l = analysiere(tmp_path, el, alt)
    neu = next(a for a in l.aen if a.typ == "absatz_neu")
    assert neu.status != NICHT


# ---------------------------------------------------------------------------------------------
# 4. Woerter ohne Normform ("•", ".")
# ---------------------------------------------------------------------------------------------
def _el_punkt(text):
    return [("h", "1 Eins"), ("p", text), ("h", "2 Zwei"), ("p", "Zweiter Absatz mit genug Text darin.")]


def test_gestrichener_punkt_am_ende_wird_entfernt(tmp_path):
    l = lauf(tmp_path, _el_punkt("Das System läuft stabil im Betrieb und liefert Daten {-•-}"), "L1",
             alt=baue_alt(tmp_path, _el_punkt("Das System läuft stabil im Betrieb und liefert Daten •"), "L1"))
    assert zelle(l.out, 2, 3) == "Das System läuft stabil im Betrieb und liefert Daten"
    assert l.aen[0].status == UEBERNOMMEN and l.aen[0].wirkung == WIRKUNG_OK


def test_gestrichenes_zeichen_das_in_der_excel_fehlt_wird_nicht_geschrieben(tmp_path):
    el = _el_punkt("Das System läuft stabil im Betrieb {-•-} und liefert Daten.")
    alt = baue_alt(tmp_path, _el_punkt("Das System läuft stabil im Betrieb und liefert Daten."), "L1")
    l = uebernimm(analysiere(tmp_path, el, alt))
    ae = l.aen[0]
    assert ae.status == NICHT and ae.grund and ae.wirkung == ""
    assert zelle(l.out, 2, 3) == "Das System läuft stabil im Betrieb und liefert Daten."


def test_gestrichener_punkt_samt_woertern_am_rand(tmp_path):
    """'{-• Das System-}': das Aufzaehlungszeichen vor dem ersten Wort wird mit entfernt."""
    el = _el_punkt("{-• Das System-} Die Anlage läuft stabil im Betrieb und liefert Daten.")
    alt = baue_alt(tmp_path, _el_punkt("• Das System Die Anlage läuft stabil im Betrieb und liefert Daten."), "L1")
    l = uebernimm(analysiere(tmp_path, el, alt))
    assert zelle(l.out, 2, 3) == "Die Anlage läuft stabil im Betrieb und liefert Daten."
    assert l.aen[0].wirkung == WIRKUNG_OK


def test_gestrichenes_zeichen_am_rand_fehlt_in_der_excel_gibt_nicht_uebernommen(tmp_path):
    el = _el_punkt("{-• Das System-} Die Anlage läuft stabil im Betrieb und liefert Daten.")
    alt = baue_alt(tmp_path, _el_punkt("Das System Die Anlage läuft stabil im Betrieb und liefert Daten."), "L1")
    l = analysiere(tmp_path, el, alt)
    assert l.aen[0].status == NICHT and "Zeichen" in l.aen[0].grund


def test_ersetztes_zeichen_am_absatzanfang_wird_ersetzt(tmp_path):
    el = _el_punkt("{-•-} {+-+} Das System läuft stabil im Betrieb und liefert Daten.")
    alt = baue_alt(tmp_path, _el_punkt("• Das System läuft stabil im Betrieb und liefert Daten."), "L1")
    l = uebernimm(analysiere(tmp_path, el, alt))
    assert zelle(l.out, 2, 3) == "- Das System läuft stabil im Betrieb und liefert Daten."


# ---------------------------------------------------------------------------------------------
# 5. Gefaehrlicher Neutext
# ---------------------------------------------------------------------------------------------
def test_neutext_probleme_erkennt_die_faelle():
    assert neutext_probleme(["Alles gut hier."])["verdaechtig"] == []
    assert neutext_probleme(["a\x07b"])["steuerzeichen"]
    assert neutext_probleme(["Teil (cid:12) da"])["verdaechtig"]
    assert neutext_probleme(["Die ﬁnale Prüfung"])["verdaechtig"]
    assert neutext_probleme(["Müll � hier"])["verdaechtig"]
    assert neutext_probleme(["=SUMME(A1)"])["formel"] and not neutext_probleme(["a =b"])["formel"]
    assert neutext_probleme(["\tTab und\nZeilenumbruch sind erlaubt"])["steuerzeichen"] is False


def test_formel_text_bleibt_als_text_mit_quote_prefix(tmp_path):
    p = _basis_excel(tmp_path)
    rep = Operation("ersetzen", "CDE", 2, 3, 0, 3, "=Das", 1, aenderung="A1", erwartet="Das")
    a = _ae("A1", [rep], alt="Das", neu="=Das")
    uebernehmen([a], p, tmp_path / "out.xlsx")
    z = blatt(tmp_path / "out.xlsx").cell(2, 3)
    assert z.data_type == "s" and z.value.startswith("=Das System") and z.quotePrefix
    assert a.status == UEBERNOMMEN and a.wirkung == WIRKUNG_OK and a.hinweise


def test_formel_in_der_zelle_wird_nie_veraendert(tmp_path):
    p = schreibe_excel(tmp_path / "f.xlsx", [["Nr.", "Titel", "Inhalt"], ["1", "Eins", "=A1&B1"]])
    rep = Operation("ersetzen", "CDE", 2, 3, 0, 3, "x", 1, aenderung="A1", erwartet="=A1")
    a = _ae("A1", [rep])
    uebernehmen([a], p, tmp_path / "out.xlsx")
    assert a.status == NICHT and "Formel" in a.grund


def test_steuerzeichen_im_neutext_stuerzt_nicht_ab_applier(tmp_path):
    p = _basis_excel(tmp_path)
    ins = Operation("einfuegen", "CDE", 2, 3, 3, 3, " Neu\x07er Text", 1, aenderung="A1")
    a = _ae("A1", [ins], typ="text_eingefuegt")
    uebernehmen([a], p, tmp_path / "out.xlsx")
    assert a.status == NICHT and "Steuerzeichen" in a.grund
    assert zelle(tmp_path / "out.xlsx", 2, 3) == "Das System muss laufen und stabil bleiben."


def test_steuerzeichen_im_neutext_planer_nicht_uebernommen():
    """Der Planer schickt Steuerzeichen nie durch (hier ueber die Pruefung, die `_setze` benutzt)."""
    assert neutext_probleme(["x", "y\x00"], config.lade("zuordnung"))["steuerzeichen"]


@pytest.mark.parametrize("text", ["Teil (cid:12) hier", "Die ﬁnale Prüfung", "Kaputt � hier"])
def test_verdaechtiger_neutext_wird_im_applier_nicht_automatisch_geschrieben(tmp_path, text):
    p = _basis_excel(tmp_path)
    ins = Operation("einfuegen", "CDE", 2, 3, 3, 3, " " + text, 1, aenderung="A1")
    a = _ae("A1", [ins], typ="text_eingefuegt", neu=text)
    uebernehmen([a], p, tmp_path / "out.xlsx")
    assert a.status == PRUEFEN and a.hinweise
    assert zelle(tmp_path / "out.xlsx", 2, 3) == "Das System muss laufen und stabil bleiben."


def test_verdaechtiger_neutext_wird_nach_menschlicher_freigabe_geschrieben(tmp_path):
    p = _basis_excel(tmp_path)
    ins = Operation("einfuegen", "CDE", 2, 3, 3, 3, " Die ﬁnale", 1, aenderung="A1")
    a = _ae("A1", [ins], status=PRUEFEN, typ="text_eingefuegt")
    uebernehmen([a], p, tmp_path / "out.xlsx", {"A1"})
    assert a.status == UEBERNOMMEN and a.bestaetigt and a.wirkung == WIRKUNG_OK
    assert "ﬁnale" in zelle(tmp_path / "out.xlsx", 2, 3)


# ---------------------------------------------------------------------------------------------
# 6. Zelle leeren nur, wenn ALLE Woerter gestrichen sind
# ---------------------------------------------------------------------------------------------
def test_zelle_leeren_mit_ungestrichenem_wort_wird_nicht_geleert(tmp_path):
    p = _basis_excel(tmp_path)
    leer = Operation("zelle_leeren", "CDE", 2, 3, aenderung="A1", erwartet="Das System muss laufen und stabil bleiben.")
    a = _ae("A1", [leer], typ="absatz_geloescht", alt="Das System muss laufen bleiben.", neu="")
    uebernehmen([a], p, tmp_path / "out.xlsx")
    assert a.status == PRUEFEN and a.hinweise
    assert zelle(tmp_path / "out.xlsx", 2, 3) == "Das System muss laufen und stabil bleiben."


def test_zelle_leeren_bei_geaenderter_zelle_ist_konflikt(tmp_path):
    p = _basis_excel(tmp_path)
    leer = Operation("zelle_leeren", "CDE", 2, 3, aenderung="A1", erwartet="Etwas ganz anderes")
    a = _ae("A1", [leer], typ="absatz_geloescht", alt="Etwas ganz anderes", neu="")
    uebernehmen([a], p, tmp_path / "out.xlsx")
    assert a.status == NICHT and "geändert" in a.grund


def test_planer_leert_zelle_nur_bei_vollstaendig_gestrichenem_absatz(tmp_path):
    el = [("h", "1 Eins"), ("p", "{-Alle Racks werden vor dem Versand geprüft.-}"), ("h", "2 Zwei"),
          ("p", "Zweiter Absatz mit genug Text darin.")]
    alt = baue_alt(tmp_path, [("h", "1 Eins"), ("p", "Alle Racks werden vor dem Versand geprüft."), ("h", "2 Zwei"),
                              ("p", "Zweiter Absatz mit genug Text darin.")], "L3")
    l = uebernimm(analysiere(tmp_path, el, alt))
    ae = l.aen[0]
    assert ae.status == UEBERNOMMEN and ae.ops[0].art == "zelle_leeren" and ae.ops[0].erwartet
    assert zeile_mit(l.out, "Alle Racks") is None


def test_teil_loeschung_mit_zusatzwoertern_ist_nur_vorschlag_und_entfernt_gezielt(tmp_path):
    el = [("h", "1 Eins"), ("p", "{-Alle Racks werden jährlich vor dem Versand geprüft.-}"), ("h", "2 Zwei"),
          ("p", "Zweiter Absatz mit genug Text darin.")]
    alt = baue_alt(tmp_path, [("h", "1 Eins"), ("p", "Alle Racks werden zusätzlich jährlich noch vor dem Versand geprüft."),
                              ("h", "2 Zwei"), ("p", "Zweiter Absatz mit genug Text darin.")], "L3")
    l = analysiere(tmp_path, el, alt)
    ae = l.aen[0]
    assert ae.status == PRUEFEN and ae.hinweise and all(o.art == "ersetzen" for o in ae.ops)
    uebernimm(l, {ae.id})
    assert zelle(l.out, 3, 3) == "zusätzlich noch"      # nur die gestrichenen Woerter sind weg
    assert ae.wirkung == WIRKUNG_OK


# ---------------------------------------------------------------------------------------------
# 7. Reihenfolge bei gleicher Startstelle
# ---------------------------------------------------------------------------------------------
@pytest.mark.parametrize("seq_ins, seq_rep", [(1, 2), (2, 1), (5, 5)])
def test_gleicher_start_einfuegen_steht_vor_dem_ersetzten(seq_ins, seq_rep):
    ins = Operation("einfuegen", "CDE", 2, 3, 4, 4, "neue ", seq_ins, aenderung="A1")
    rep = Operation("ersetzen", "CDE", 2, 3, 4, 10, "Gerät", seq_rep, aenderung="A2", erwartet="System")
    assert _wende_text_an("Das System muss laufen.", [ins, rep]) == "Das neue Gerät muss laufen."
    assert _wende_text_an("Das System muss laufen.", [rep, ins]) == "Das neue Gerät muss laufen."


def test_zwei_einfuegungen_gleicher_start_nach_seq():
    a = Operation("einfuegen", "CDE", 2, 3, 4, 4, "eins ", 1, aenderung="A1")
    b = Operation("einfuegen", "CDE", 2, 3, 4, 4, "zwei ", 2, aenderung="A2")
    assert _wende_text_an("Das System.", [b, a]) == "Das eins zwei System."


# ---------------------------------------------------------------------------------------------
# 8. Format neuer Zeilen
# ---------------------------------------------------------------------------------------------
def test_stilzeile_einer_neuen_absatzzeile_ist_nie_die_ueberschrift(tmp_path):
    el = [("h", "1 Eins"), ("p", "{+Neu direkt nach der Überschrift eins.+}"), ("p", "Absatz eins mit genug Text darin."),
          ("h", "2 Zwei"), ("p", "Absatz zwei mit genug Text darin.")]
    alt = baue_alt(tmp_path, [e for e in el if "{+" not in e[1]], "L3")
    l = analysiere(tmp_path, el, alt)
    op = l.aen[0].ops[0]
    kopf_zeilen = l.xl.layouts["CDE"].kopf_zeilen
    assert op.stil_von not in kopf_zeilen and l.xl.text("CDE", op.stil_von, 3).startswith("Absatz eins")


# ---------------------------------------------------------------------------------------------
# 9. Protokoll
# ---------------------------------------------------------------------------------------------
def _ae_einfach(status=UEBERNOMMEN, **kw):
    return Aenderung(id="A0001", typ="text_ersetzt", seite=1, bbox=(0, 0, 1, 1), section="1", alt="a", neu="b", woerter=2,
                     status=status, **kw)


def test_bilanz_satz_nur_bei_vollstaendig_und_stimmig():
    ok = bilanz([_ae_einfach()], 2, [])
    assert "nichts ging still verloren" in bilanz_satz(ok)[1]
    mit = bilanz([_ae_einfach()], 2, [], [luecke("scan_seite", 2, "Seite 2 enthält keinen lesbaren Text.")])
    assert mit["stimmt"] and not mit["vollstaendig"]
    assert "nichts ging still verloren" not in bilanz_satz(mit)[1] and "UNVOLLSTÄNDIG" in bilanz_satz(mit)[1]
    falsch = bilanz([_ae_einfach()], 5, [])
    assert "nichts ging still verloren" not in bilanz_satz(falsch)[1] and "NEIN" in bilanz_satz(falsch)[1]


def _zellen_text(pfad):
    wb = load_workbook(str(pfad))
    return {ws.title: [[c.value for c in row] for row in ws.iter_rows()] for ws in wb.worksheets}


def test_protokoll_xlsx_unvollstaendig_mit_luecken_hinweisen_und_wirkung(tmp_path):
    a = _ae_einfach(hinweise=["Bitte mit dem PDF vergleichen."], wirkung=WIRKUNG_OK, nachkontrolle="ok")
    lk = [luecke("scan_seite", 3, "Seite 3 enthält keinen lesbaren Text.", 0)]
    bil = bilanz([a], 2, [], lk)
    pfad = tmp_path / "p.xlsx"
    schreibe_xlsx(pfad, {"pdf_name": "a.pdf", "excel_name": "b.xlsx"}, [a], bil, [], [], None, lk)
    inhalt = _zellen_text(pfad)
    bilanz_zellen = [str(c) for row in inhalt["Bilanz"] for c in row if c is not None]
    assert not any("nichts ging still verloren" in c for c in bilanz_zellen)
    assert any("UNVOLLSTÄNDIG" in c for c in bilanz_zellen)
    assert any("Seite 3 enthält keinen lesbaren Text." in c for c in bilanz_zellen)
    kopf = inhalt["Änderungen"][0]
    zeile = dict(zip(kopf, inhalt["Änderungen"][1]))
    assert zeile["Wirkung"] == "ok" and "Bitte mit dem PDF vergleichen." in zeile["Warnungen"]


def test_protokoll_xlsx_vollstaendig_sagt_nichts_ging_verloren(tmp_path):
    a = _ae_einfach()
    bil = bilanz([a], 2, [])
    pfad = tmp_path / "p.xlsx"
    schreibe_xlsx(pfad, {}, [a], bil, [], [], None)
    zellen = [str(c) for row in _zellen_text(pfad)["Bilanz"] for c in row if c is not None]
    assert any("nichts ging still verloren" in c for c in zellen)


def test_protokoll_xlsx_ueberlebt_steuerzeichen_und_formeltext(tmp_path):
    a = _ae_einfach()
    a.neu, a.alt, a.grund = "Text\x07 mit Steuerzeichen", "=SUMME(A1)", "=1+1"
    bil = bilanz([a], 2, [])
    pfad = tmp_path / "p.xlsx"
    schreibe_xlsx(pfad, {"pdf_name": "=evil"}, [a], bil, [], ["=Warnung\x01"], None)
    wb = load_workbook(str(pfad))
    formeln = [(ws.title, c.coordinate) for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.data_type == "f"]
    assert not formeln


def test_bilanz_zaehlt_wirkung_als_abweichung():
    a = _ae_einfach(status=NICHT, wirkung=WIRKUNG_KEINE)
    assert bilanz([a], 2, [])["nachkontrolle_abweichungen"] == 1


def test_markierung_gelb_liest_farbe_aus_der_konfiguration(tmp_path):
    p = _basis_excel(tmp_path)
    rep = Operation("ersetzen", "CDE", 2, 3, 4, 10, "Gerät", 1, aenderung="A1", erwartet="System")
    cfg = copy.deepcopy(config.lade("zuordnung"))
    cfg["markierung"]["geaenderte_zellen"] = "gelb"
    cfg["anwenden"]["markierung_farbe"] = "AABBCC"
    uebernehmen([_ae("A1", [rep])], p, tmp_path / "out.xlsx", cfg=cfg)
    assert blatt(tmp_path / "out.xlsx").cell(2, 3).fill.fgColor.rgb.endswith("AABBCC")


def test_kommentar_markierung_liest_autor_und_ueberschreibt_keinen_kommentar(tmp_path):
    from openpyxl.comments import Comment
    p = _basis_excel(tmp_path)
    wb = load_workbook(str(p))
    wb["CDE"]["C2"].comment = Comment("Alter Kommentar von Hand", "Nico")
    wb.save(str(p))
    rep = Operation("ersetzen", "CDE", 2, 3, 4, 10, "Gerät", 1, aenderung="A1", erwartet="System")
    cfg = copy.deepcopy(config.lade("zuordnung"))
    cfg["markierung"]["geaenderte_zellen"] = "kommentar"
    cfg["anwenden"]["kommentar_autor"] = "Test-Tool"
    uebernehmen([_ae("A1", [rep], alt="System", neu="Gerät")], p, tmp_path / "out.xlsx", cfg=cfg)
    k = blatt(tmp_path / "out.xlsx").cell(2, 3).comment
    assert "Alter Kommentar von Hand" in k.text and "PDF S.1" in k.text
    # Zelle ohne vorhandenen Kommentar: Autor kommt aus der Konfiguration
    rep2 = Operation("ersetzen", "CDE", 3, 3, 0, 7, "Dritter", 1, aenderung="A2", erwartet="Zweiter")
    uebernehmen([_ae("A2", [rep2], alt="Zweiter", neu="Dritter")], p, tmp_path / "out2.xlsx", cfg=cfg)
    assert blatt(tmp_path / "out2.xlsx").cell(3, 3).comment.author == "Test-Tool"


def test_einstellung_neue_section_gibt_es_nicht_mehr():
    assert "neue_section" not in config.lade("zuordnung")["einfuegen"]
    assert config.lade("zuordnung")["planer"]["doppelt_min_woerter"] >= 1


def test_kommentar_markierung_ueberlebt_steuerzeichen_im_alten_text(tmp_path):
    p = _basis_excel(tmp_path)
    rep = Operation("ersetzen", "CDE", 2, 3, 4, 10, "Gerät", 1, aenderung="A1", erwartet="System")
    cfg = copy.deepcopy(config.lade("zuordnung"))
    cfg["markierung"]["geaenderte_zellen"] = "kommentar"
    a = _ae("A1", [rep], alt="Sys\x07tem", neu="Gerät")
    uebernehmen([a], p, tmp_path / "out.xlsx", cfg=cfg)
    assert a.status == UEBERNOMMEN and "\x07" not in blatt(tmp_path / "out.xlsx").cell(2, 3).comment.text
