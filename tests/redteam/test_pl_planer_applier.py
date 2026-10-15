"""Red-Team K-PL (planner / applier / protokoll): Fehlerfaelle mit erfundenen Daten.

Jeder Test hier faellt heute mit einem AssertionError (xfail strict). Wird der Fehler repariert,
wird der Test XPASS = rot -> der Besitzer loescht dann den xfail-Marker.
"""
import json

import pytest
from openpyxl.styles import Font, PatternFill

from src import runs
from src.applier import uebernehmen
from src.models import Aenderung, Operation
from tests.fixtures.pdf_writer import baue_pdf
from tests.redteam._bau import (alle_texte, analysiere, analysiere_mit_pdf, aufbau, baue_alt, baue_pdf_ligatur, bearbeite,
                                blatt, lauf, scheitern, schreibe_excel, uebernimm, zelle, zeile_mit)

# ---------------------------------------------------------------------------------------------
# RT-01 / RT-02: neue Section im Modell "eine Zeile pro Section" -- Folgeabsaetze haben keine Operation
# ---------------------------------------------------------------------------------------------
EL_SEC = [("h", "1 Eins"), ("p", "Erster Absatz mit genug Text darin."), ("h", "2 Zwei"),
          ("p", "Bekannter Absatz mit genug Text darin."),
          ("p", "Unbekannter Absatz der in der Excel gar nicht vorkommt."),   # nur im PDF -> Anker "nicht direkt" -> pruefen
          ("h", "{+3 Neue Section+}"), ("p", "{+Folgeabsatz der neuen Section steht hier.+}")]
EX_SEC = [e for e in EL_SEC if "Unbekannter" not in e[1]]
FOLGE = "Folgeabsatz der neuen Section steht hier."


def _section_lauf(tmp_path, freigeben: str):
    alt = baue_alt(tmp_path, EX_SEC, "L1")
    l = analysiere(tmp_path, EL_SEC, alt)
    kopf = next(a for a in l.aen if a.typ == "section_neu")
    folge = next(a for a in l.aen if a.typ == "absatz_neu")
    aufbau(kopf.status == "pruefen" and folge.status == "pruefen", "neue Section muss auf 'pruefen' stehen")
    uebernimm(l, {kopf.id if freigeben == "kopf" else folge.id})
    return l, kopf, folge


def test_rt01_folgeabsatz_allein_freigegeben_ohne_wirkung(tmp_path):
    l, kopf, folge = _section_lauf(tmp_path, "folge")
    in_excel = any(FOLGE in t for t in alle_texte(l.out))
    if folge.status == "uebernommen" and not in_excel:
        scheitern(f"Folgeabsatz hat Status 'uebernommen' (Nachkontrolle: {folge.nachkontrolle!r}), steht aber nicht in der Excel")


def test_rt02_nur_ueberschrift_freigegeben_folgeabsatz_schleicht_mit(tmp_path):
    l, kopf, folge = _section_lauf(tmp_path, "kopf")
    in_excel = any(FOLGE in t for t in alle_texte(l.out))
    if in_excel and folge.status != "uebernommen":
        scheitern(f"Folgeabsatz steht in der Excel, obwohl er nicht freigegeben wurde (Status {folge.status!r})")


# ---------------------------------------------------------------------------------------------
# RT-03..05: zweiter Lauf mit demselben PDF auf dem Ergebnis -> reine Einfuegungen doppelt
# ---------------------------------------------------------------------------------------------
def _zweimal(tmp_path, elemente, layout, text):
    l1 = lauf(tmp_path / "eins", elemente, layout)
    aufbau(sum(t.count(text) for t in alle_texte(l1.out)) == 1, "erster Lauf muss den Text genau einmal eintragen")
    l2 = uebernimm(analysiere(tmp_path / "zwei", elemente, l1.out), name="out2.xlsx")
    return sum(t.count(text) for t in alle_texte(l2.out)), l2


def test_rt03_zweiter_lauf_text_eingefuegt_nicht_doppelt(tmp_path):
    el = [("h", "1 Eins"), ("p", "Dieses Dokument beschreibt die Anforderungen. {+Es gilt für alle Standorte des Kunden.+}"),
          ("h", "2 Zwei"), ("p", "Zweiter Absatz mit genug Text darin.")]
    n, l2 = _zweimal(tmp_path, el, "L1", "Es gilt für alle Standorte des Kunden.")
    assert n == 1, f"Satz steht nach dem zweiten Lauf {n}-mal in der Excel (Status {l2.aen[0].status}/{l2.aen[0].konfidenz})"


def test_rt04_zweiter_lauf_absatz_neu_in_zelle_nicht_doppelt(tmp_path):
    el = [("h", "1 Eins"), ("p", "Erster Absatz."), ("p", "{+Neuer Absatz hier unten.+}"), ("h", "2 Zwei"),
          ("p", "Zweiter Absatz mit genug Text darin.")]
    n, l2 = _zweimal(tmp_path, el, "L1", "Neuer Absatz hier unten.")
    assert n == 1, f"Neuer Absatz steht nach dem zweiten Lauf {n}-mal in der Excel"


def test_rt05_zweiter_lauf_absatz_neu_als_zeile_nicht_doppelt(tmp_path):
    el = [("h", "1 Eins"), ("p", "Erster Absatz."), ("p", "{+Neuer Absatz hier unten.+}"), ("h", "2 Zwei"),
          ("p", "Zweiter Absatz mit genug Text darin.")]
    n, l2 = _zweimal(tmp_path, el, "L3", "Neuer Absatz hier unten.")
    assert n == 1, f"Neuer Absatz steht nach dem zweiten Lauf {n}-mal in der Excel"


# ---------------------------------------------------------------------------------------------
# RT-06 / RT-07: gestrichenes Aufzaehlungszeichen allein
# ---------------------------------------------------------------------------------------------
def test_rt06_gestrichener_punkt_mitten_im_text(tmp_path):
    el = [("h", "1 Eins"), ("p", "Das System läuft stabil im Betrieb {-•-} und liefert Daten."), ("h", "2 Zwei"),
          ("p", "Zweiter Absatz mit genug Text darin.")]
    l = lauf(tmp_path, el, "L1")
    ae = l.aen[0]
    if ae.status == "uebernommen" and "•" in zelle(l.out, 2, 3):
        scheitern(f"Status 'uebernommen'/{ae.nachkontrolle}, aber Zelle ist jetzt {zelle(l.out, 2, 3)!r}")


def test_rt07_gestrichener_punkt_am_anfang(tmp_path):
    el = [("h", "1 Eins"), ("p", "{-•-} Das System läuft stabil im Betrieb und liefert Daten."), ("h", "2 Zwei"),
          ("p", "Zweiter Absatz mit genug Text darin.")]
    l = lauf(tmp_path, el, "L1")
    ae = l.aen[0]
    if ae.status == "uebernommen" and "•" in zelle(l.out, 2, 3):
        scheitern(f"Status 'uebernommen'/{ae.nachkontrolle}, aber Zelle ist jetzt {zelle(l.out, 2, 3)!r}")


@pytest.mark.parametrize("alt, neu", [
    pytest.param("•", "-"),
    pytest.param(".", "!"),
])
def test_rt53_ersetztes_zeichen_ohne_normform_bleibt_nicht_stehen(tmp_path, alt, neu):
    """Woerter, die nach der Normalisierung leer sind ('•', '.', '-'), zaehlen fuer den Planer nicht als 'alter Text'."""
    el = [("h", "1 Eins"), ("p", "Das System läuft stabil im Betrieb {-" + alt + "-} {+" + neu + "+} und liefert Daten."),
          ("h", "2 Zwei"), ("p", "Zweiter Absatz mit genug Text darin.")]
    l = lauf(tmp_path, el, "L1")
    ae = l.aen[0]
    text = zelle(l.out, 2, 3)
    aufbau(neu in text.split(), f"neues Zeichen muss in der Zelle stehen: {text!r}")
    if ae.status == "uebernommen" and alt in text.split():
        scheitern(f"Status 'uebernommen'/{ae.nachkontrolle}, aber das alte Zeichen steht noch da: {text!r}")


# ---------------------------------------------------------------------------------------------
# RT-08..12: Neutext, der in Excel etwas anderes bewirkt oder unsauber ist
# ---------------------------------------------------------------------------------------------
def _neuer_absatz(neu: str, layout: str = "L3"):
    return [("h", "1 Eins"), ("p", "Erster Absatz."), ("p", "{+" + neu + "+}"), ("h", "2 Zwei"),
            ("p", "Zweiter Absatz mit genug Text darin.")]


def test_rt08_neutext_mit_gleichheitszeichen_wird_keine_formel(tmp_path):
    l = lauf(tmp_path, _neuer_absatz("=SUMME(A1:A9)"), "L3")
    ws = blatt(l.out)
    formeln = [(c.coordinate, c.value) for row in ws.iter_rows() for c in row if c.data_type == "f"]
    assert not formeln, f"In der Excel steht jetzt eine Formel statt Text: {formeln} (Status {l.aen[0].status})"


def test_rt09_ersetzter_zellanfang_mit_gleichheitszeichen_wird_keine_formel(tmp_path):
    el = [("h", "1 Eins"), ("p", "{-Dieses-} {+=Dieses+} Dokument beschreibt die Anforderungen an das System."),
          ("h", "2 Zwei"), ("p", "Zweiter Absatz mit genug Text darin.")]
    l = lauf(tmp_path, el, "L1")
    ws = blatt(l.out)
    formeln = [(c.coordinate, c.value) for row in ws.iter_rows() for c in row if c.data_type == "f"]
    assert not formeln, f"In der Excel steht jetzt eine Formel statt Text: {formeln} (Status {l.aen[0].status})"


def test_rt10_steuerzeichen_im_neutext_stuerzt_nicht_ab(tmp_path):
    ordner = tmp_path / "fx"
    ordner.mkdir()
    el = _neuer_absatz("Normaler neuer Absatz hier unten.")
    alt = baue_alt(ordner, el, "L3")
    baue_pdf(ordner / "a.pdf", el, "farbe")
    run_id = runs.neuer_lauf(ordner / "a.pdf", alt)
    runs.analysiere(run_id)
    pj = runs.lauf_dir(run_id) / "plan.json"
    j = json.loads(pj.read_text(encoding="utf-8"))
    ae = next(a for a in j["aenderungen"] if a["typ"] == "absatz_neu")
    # So koennte ein PDF-Leser ein Steuerzeichen (BEL) aus dem PDF durchreichen:
    ae["neu"] = "Normaler neuer Absatz\x07 hier unten."
    for op in ae["ops"]:
        op["zellen"] = {k: "Normaler neuer Absatz\x07 hier unten." for k in op["zellen"]}
    pj.write_text(json.dumps(j, ensure_ascii=False), encoding="utf-8")
    try:
        runs.uebernehmen_lauf(run_id, [])
    except Exception as e:                     # noqa: BLE001
        leer = [p.name for p in runs.lauf_dir(run_id).glob("ergebnis_*") if not (p / "ergebnis.json").exists()]
        scheitern(f"Absturz beim Uebernehmen: {type(e).__name__}: {e}; halbfertige Ordner: {leer}")


def test_rt11_cid_muell_wird_nicht_automatisch_uebernommen(tmp_path):
    l = lauf(tmp_path, _neuer_absatz("Teil (cid:12) hier unten."), "L3")
    ae = l.aen[0]
    steht = any("(cid:" in t for t in alle_texte(l.out))
    if ae.status == "uebernommen" and steht:
        scheitern(f"'(cid:12)' steht jetzt in der Excel, Status {ae.status}/{ae.konfidenz}, Grund {ae.grund!r}")


def test_rt12_ligatur_wird_nicht_ungeprueft_geschrieben(tmp_path):
    el = _neuer_absatz("Die €nale Prüfung folgt hier unten.")
    alt = baue_alt(tmp_path, [("h", "1 Eins"), ("p", "Erster Absatz."), ("h", "2 Zwei"),
                              ("p", "Zweiter Absatz mit genug Text darin.")], "L3")
    pdf = tmp_path / "lig.pdf"
    baue_pdf_ligatur(pdf, el)
    l = uebernimm(analysiere_mit_pdf(pdf, alt))
    ae = next(a for a in l.aen if a.typ == "absatz_neu")
    aufbau("ﬁ" in ae.neu, f"Test-PDF liefert keine Ligatur: {ae.neu!r}")
    steht = any("ﬁ" in t for t in alle_texte(l.out))
    if ae.status == "uebernommen" and steht:
        scheitern(f"Ligatur 'ﬁ' steht roh in der Excel (Status {ae.status}/{ae.konfidenz}, Grund {ae.grund!r})")


# ---------------------------------------------------------------------------------------------
# RT-13 / RT-14: Zelle leeren prueft nur erstes und letztes Wort
# ---------------------------------------------------------------------------------------------
GEL = [("h", "1 Eins"), ("p", "Erster Absatz mit Text."),
       ("p", "{-Alle Racks werden jährlich vor dem Versand geprüft.-}"), ("h", "2 Zwei"),
       ("p", "Zweiter Absatz mit genug Text darin.")]
GEL_EX = [("h", "1 Eins"), ("p", "Erster Absatz mit Text."),
          ("p", "Alle Racks werden zusätzlich jährlich noch vor dem Versand geprüft."), ("h", "2 Zwei"),
          ("p", "Zweiter Absatz mit genug Text darin.")]


@pytest.mark.parametrize("layout, nr", [
    pytest.param("L1", 13),
    pytest.param("L3", 14),
])
def test_rt13_14_zelle_leeren_verliert_nicht_gestrichene_woerter(tmp_path, layout, nr):
    alt = baue_alt(tmp_path, GEL_EX, layout)
    l = uebernimm(analysiere(tmp_path, GEL, alt))
    ae = l.aen[0]
    aufbau(ae.typ == "absatz_geloescht", "erwartet: ein gestrichener Absatz")
    text = " ".join(alle_texte(l.out))
    if ae.status == "uebernommen" and ("zusätzlich" not in text or "noch" not in text.split()):
        scheitern(f"Status 'uebernommen'/{ae.konfidenz}, aber die nicht gestrichenen Woerter 'zusätzlich'/'noch' sind weg")


# ---------------------------------------------------------------------------------------------
# RT-15: Nachkontrolle prueft nur "Neutext kommt irgendwo in der Zelle vor"
# ---------------------------------------------------------------------------------------------
@pytest.mark.parametrize("neu", [
    pytest.param("s"),
    pytest.param("2.2"),
])
def test_rt15_nachkontrolle_erkennt_falsche_stelle(tmp_path, neu):
    """Plan wurde fuer 'Das Rack ...' gemacht (Einfuegestelle hinter 'Das Rack' = Position 8). Inzwischen steht in
    der Zelle anderer Text gleicher Laenge. Die Einfuegung landet mitten in einem Wort -- Nachkontrolle sagt 'ok'."""
    p = schreibe_excel(tmp_path / "b.xlsx", [["Nr.", "Titel", "Inhalt"],
                                              ["1", "Eins", "Das Gehäuse steht hier und 2.2 s gilt."],
                                              ["2", "Zwei", "Zweiter Absatz mit genug Text darin."]])
    op = Operation("einfuegen", "CDE", 2, 3, 8, 8, " " + neu, 1, aenderung="A0001")
    ae = Aenderung(id="A0001", typ="text_eingefuegt", seite=1, bbox=(0, 0, 1, 1), section="1", alt="", neu=neu,
                   woerter=1, status="uebernommen", konfidenz="hoch", ops=[op])
    uebernehmen([ae], p, tmp_path / "out.xlsx")
    text = zelle(tmp_path / "out.xlsx", 2, 3)
    aufbau(text != "Das Gehäuse steht hier und 2.2 s gilt.", "die Einfuegung muss ausgefuehrt worden sein (sonst kein Fehlerfall)")
    if ae.status == "uebernommen" and ae.nachkontrolle == "ok":
        scheitern(f"Einfuegung mitten im Wort ({text!r}), aber Status 'uebernommen' und Nachkontrolle 'ok'")


# ---------------------------------------------------------------------------------------------
# RT-16: Ersetzen und Einfuegen an derselben Stelle (gleicher start)
# ---------------------------------------------------------------------------------------------
def test_rt16_ersetzen_und_einfuegen_gleicher_start(tmp_path):
    p = schreibe_excel(tmp_path / "a.xlsx", [["Nr.", "Titel", "Inhalt"],
                                              ["1", "Eins", "Das System muss laufen und stabil bleiben."],
                                              ["2", "Zwei", "Zweiter Absatz mit genug Text darin."]])
    ins = Operation("einfuegen", "CDE", 2, 3, 4, 4, "neue ", 2, aenderung="A0001")
    rep = Operation("ersetzen", "CDE", 2, 3, 4, 10, "Gerät", 1, aenderung="A0002", erwartet="System")
    a1 = Aenderung(id="A0001", typ="text_eingefuegt", seite=1, bbox=(0, 0, 1, 1), section="1", alt="", neu="neue",
                   woerter=1, status="uebernommen", konfidenz="hoch", ops=[ins])
    a2 = Aenderung(id="A0002", typ="text_ersetzt", seite=1, bbox=(0, 0, 1, 1), section="1", alt="System", neu="Gerät",
                   woerter=1, status="uebernommen", konfidenz="hoch", ops=[rep])
    uebernehmen([a1, a2], p, tmp_path / "out.xlsx")
    text = zelle(tmp_path / "out.xlsx", 2, 3)
    ok = ("neue" in text and text.count("Gerät") == 1 and "System" not in text and "ystem" not in text
          and "muss laufen und stabil bleiben." in text)
    assert ok, f"Beide Aenderungen muessen ankommen, ohne den Text zu zerstueckeln, ist aber: {text!r}"


# ---------------------------------------------------------------------------------------------
# RT-17: neue Zeile direkt nach einer Ueberschriftszeile uebernimmt deren Format
# ---------------------------------------------------------------------------------------------
def test_rt17_neue_zeile_nach_ueberschrift_hat_absatzformat(tmp_path):
    el = [("h", "1 Eins"), ("p", "{+Neu direkt nach der Überschrift eins.+}"), ("p", "Absatz eins mit genug Text darin."),
          ("h", "2 Zwei"), ("p", "Absatz zwei mit genug Text darin.")]

    def stil(wb):
        ws = wb["CDE"]
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value:                                   # Ueberschriftszeilen: fett, gross, farbig
                for c in range(1, 5):
                    ws.cell(r, c).font = Font(bold=True, size=14)
                    ws.cell(r, c).fill = PatternFill("solid", fgColor="CCCCFF")
    l = lauf(tmp_path, el, "L3", nachher=lambda p: bearbeite(p, stil))
    r = zeile_mit(l.out, "Neu direkt nach der Überschrift")
    aufbau(r is not None, "neuer Absatz muss in der Excel stehen")
    z = blatt(l.out).cell(r, 3)
    assert not z.font.b and z.font.sz == 11 and z.fill.fgColor.rgb in ("00000000", None), \
        f"Neue Zeile {r} hat Ueberschriftsformat: fett={z.font.b}, Groesse={z.font.sz}, Fuellung={z.fill.fgColor.rgb}"
