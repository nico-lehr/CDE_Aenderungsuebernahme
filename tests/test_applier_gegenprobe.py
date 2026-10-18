"""Applier Welle 2 (ADR 0004): Gegenprobe A3 nach dem Speichern, Invariante "ohne Absatz", neue Section-Nummer als Zahl,
Abwahl durch den Menschen. Nur erfundene Daten (Regel 16)."""
from __future__ import annotations

import copy

from openpyxl import Workbook, load_workbook

from src import config
from src.applier import uebernehmen
from src.models import NICHT, PRUEFEN, UEBERNOMMEN, WIRKUNG_ABWEICHUNG, WIRKUNG_OK, Absatz, Aenderung, Lauf, Operation
from src.gegenprobe import zustand

GEGENPROBE_HINWEIS = "Gegenprobe: Das Ergebnis in der Excel passt nicht zu dem, was das PDF verlangt. Bitte prüfen."


# ---------------------------------------------------------------- Hilfen
def L(spec: str, d: str = "A1", n: str | None = None) -> list[Lauf]:
    """Kurzschrift: Woerter mit Leerzeichen; {-x-} entfaellt (Aenderung d), {+x+} neu (Aenderung n, Standard = d);
    '|' klebt ohne Trenner; '/' steht fuer einen Zeilenumbruch-Trenner."""
    n = n or d
    out: list[Lauf] = []
    for i, teil in enumerate(spec.split(" ")):
        if i:
            out.append(Lauf("s", " "))
        for stueck in teil.split("|"):
            if stueck.startswith("{-"):
                out.append(Lauf("d", stueck[2:-2], d))
            elif stueck.startswith("{+"):
                out.append(Lauf("a", stueck[2:-2], n))
            elif stueck:
                out.append(Lauf("k", stueck))
    return out


def _excel(tmp_path, inhalt: str, weitere: list[list] | None = None, name: str = "a.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "CDE"
    for z in [["Nr.", "Titel", "Inhalt"], ["1", "Eins", inhalt], ["2", "Zwei", "Zweiter Absatz mit genug Text darin."]] \
            + (weitere or []):
        ws.append(z)
    p = tmp_path / name
    wb.save(str(p))
    return p


def _zelle(p, r, c):
    return load_workbook(str(p))["CDE"].cell(r, c).value


def _ae(aid, ops, status=UEBERNOMMEN, typ="text_ersetzt", **kw) -> Aenderung:
    return Aenderung(id=aid, typ=typ, seite=1, bbox=(0, 0, 1, 1), section="1", alt="", neu="x", woerter=1,
                     status=status, konfidenz="hoch", ops=ops, ziel="CDE!C2", **kw)


def _ers(start, ende, text, aid, orig, zeile=2, **kw) -> Operation:
    return Operation("ersetzen", "CDE", zeile, 3, start, ende, text, aenderung=aid, erwartet=orig[start:ende], **kw)


def _ein(pos, text, aid, zeile=2) -> Operation:
    return Operation("einfuegen", "CDE", zeile, 3, pos, pos, text, aenderung=aid)


def _absatz(laeufe, start, ende, zeile=2, pid="P1") -> Absatz:
    return Absatz(pid, "CDE", [(zeile, 3, start, ende)], laeufe)


def _cfg(pflicht: bool) -> dict:
    cfg = copy.deepcopy(config.lade("zuordnung"))
    cfg["anwenden"]["gegenprobe_pflicht"] = pflicht
    return cfg


def _durchgefallen(ae: Aenderung, code: str) -> None:
    assert ae.status == PRUEFEN, (ae.status, ae.grund)
    assert ae.wirkung == WIRKUNG_ABWEICHUNG and ae.nachkontrolle == "abweichung"
    assert code in ae.pruefcodes
    assert GEGENPROBE_HINWEIS in ae.hinweise


def _bestanden(ae: Aenderung) -> None:
    assert (ae.status, ae.wirkung, ae.nachkontrolle) == (UEBERNOMMEN, WIRKUNG_OK, "ok"), (ae.status, ae.grund, ae.pruefcodes)
    assert not [c for c in ae.pruefcodes if c.startswith("gegenprobe")]


# ---------------------------------------------------------------- A3 erkennt falsch angewendeten Text
def test_a3_satzzeichen_falsche_seite_wird_pruefen(tmp_path):
    t = "Die Anlage arbeitet."
    p = _excel(tmp_path, t)
    ab = _absatz(L("Die Anlage arbeitet {+leise+}|."), 0, len(t))
    ae = _ae("A1", [_ein(len(t), " leise", "A1")], typ="text_eingefuegt")       # hinter dem Punkt: "arbeitet. leise"
    res = uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Die Anlage arbeitet. leise"     # Datei wird nicht zurueckgedreht
    _durchgefallen(ae, "gegenprobe_text")
    assert any("Gegenprobe" in w and "A1" in w for w in res.warnungen)


def test_a3_satzzeichen_richtige_seite_bleibt_uebernommen(tmp_path):
    t = "Die Anlage arbeitet."
    p = _excel(tmp_path, t)
    ab = _absatz(L("Die Anlage arbeitet {+leise+}|."), 0, len(t))
    ae = _ae("A1", [_ein(len(t) - 1, " leise", "A1")], typ="text_eingefuegt")
    res = uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Die Anlage arbeitet leise."
    _bestanden(ae)
    assert not res.warnungen


def test_a3_falsches_vorkommen_wird_pruefen(tmp_path):
    t = "Der Wert und der Wert."
    p = _excel(tmp_path, t)
    ab = _absatz(L("Der Wert und der {-Wert-} {+Pegel+}|."), 0, len(t))
    assert zustand(ab.laeufe, {"A1"}) == "Der Wert und der Pegel."
    ae = _ae("A1", [_ers(4, 8, "Pegel", "A1", t)])                              # erstes statt zweites Vorkommen
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Der Pegel und der Wert."
    _durchgefallen(ae, "gegenprobe_text")


def test_a3_zeilenumbruch_im_neuen_text_wird_pruefen(tmp_path):
    t = "Erster Teil zweiter Teil."
    p = _excel(tmp_path, t)
    ab = _absatz(L("Erster Teil {+neu+} zweiter Teil."), 0, len(t))
    ae = _ae("A1", [_ein(len("Erster Teil"), "\nneu", "A1")], typ="text_eingefuegt")
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    _durchgefallen(ae, "gegenprobe_umbruch")


def test_a3_gruppe_wird_gemeinsam_zurueckgestuft(tmp_path):
    t = "Der Wert und der Wert."
    p = _excel(tmp_path, t)
    ab = _absatz(L("Der Wert und der {-Wert-} {+Pegel+}|."), 0, len(t))
    a1 = _ae("A1", [_ers(4, 8, "Pegel", "A1", t)], gruppe="G1")
    zweit = "Zweiter Absatz mit genug Text darin."
    a2 = _ae("A2", [_ers(0, 7, "Dritter", "A2", zweit, zeile=3)], gruppe="G1")
    ab2 = _absatz(L("{-Zweiter-} {+Dritter+} Absatz mit genug Text darin.", "A2"), 0, len(zweit), zeile=3, pid="P2")
    uebernehmen([a1, a2], p, tmp_path / "out.xlsx", absaetze=[ab, ab2])
    _durchgefallen(a1, "gegenprobe_text")
    assert a2.status == PRUEFEN and "gegenprobe_gruppe" in a2.pruefcodes and a2.wirkung == WIRKUNG_ABWEICHUNG


# ---------------------------------------------------------------- A3 laesst richtig angewendeten Text in Ruhe
def test_a3_mehrere_absaetze_in_einer_zelle_mit_fremder_operation_davor(tmp_path):
    t = "Erster Satz bleibt.\nZweiter Satz alt.\nDritter Satz."
    p = _excel(tmp_path, t)
    z1 = t.index("\n")
    z2 = t.index("\n", z1 + 1)
    p1 = _absatz(L("Erster Satz bleibt {+jetzt+}|.", "A1"), 0, z1, pid="P1")
    p2 = _absatz(L("Zweiter Satz {-alt-}|{+neu+}|.", "A2"), z1 + 1, z2, pid="P2")
    p3 = _absatz(L("Dritter Satz."), z2 + 1, len(t), pid="P3")                 # ohne Aenderung: wird nicht geprueft
    a1 = _ae("A1", [_ein(z1 - 1, " jetzt", "A1")], typ="text_eingefuegt")
    a = t.index("alt")
    a2 = _ae("A2", [_ers(a, a + 3, "neu", "A2", t)])
    uebernehmen([a1, a2], p, tmp_path / "out.xlsx", absaetze=[p1, p2, p3])
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Erster Satz bleibt jetzt.\nZweiter Satz neu.\nDritter Satz."
    _bestanden(a1)
    _bestanden(a2)


def test_a3_ganzer_absatz_samt_umbruch_geloescht(tmp_path):
    t = "Erster Satz.\nZweiter Satz entfällt.\nDritter Satz."
    p = _excel(tmp_path, t)
    s = t.index("Zweiter")
    e = t.index("\n", s)
    ab = _absatz(L("{-Zweiter-} {-Satz-} {-entfällt.-}"), s, e)
    ae = _ae("A1", [_ers(s, e + 1, "", "A1", t, absatz_weg=True)], typ="absatz_geloescht")
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Erster Satz.\nDritter Satz."
    _bestanden(ae)


def test_a3_ganzer_absatz_mit_umbruch_davor_geloescht(tmp_path):
    t = "Erster Satz.\nZweiter Satz entfällt."
    p = _excel(tmp_path, t)
    s = t.index("Zweiter")
    ab = _absatz(L("{-Zweiter-} {-Satz-} {-entfällt.-}"), s, len(t))
    ae = _ae("A1", [_ers(s - 1, len(t), "", "A1", t, absatz_weg=True)], typ="absatz_geloescht")
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Erster Satz."
    _bestanden(ae)


def test_a3_zwei_geloeschte_absaetze_teilen_ein_leerzeichen(tmp_path):
    t = "Eins bleibt. Zwei weg. Drei weg. Vier bleibt."
    p = _excel(tmp_path, t)
    s2, s3 = t.index("Zwei"), t.index("Drei")
    e2, e3 = s2 + len("Zwei weg."), s3 + len("Drei weg.")
    p2 = _absatz(L("{-Zwei-} {-weg.-}", "A1"), s2, e2, pid="P2")
    p3 = _absatz(L("{-Drei-} {-weg.-}", "A2"), s3, e3, pid="P3")
    a1 = _ae("A1", [_ers(s2 - 1, e2 + 1, "", "A1", t)], typ="absatz_geloescht")   # beide nehmen das Leerzeichen mit
    a2 = _ae("A2", [_ers(s3 - 1, e3, "", "A2", t)], typ="absatz_geloescht")
    uebernehmen([a1, a2], p, tmp_path / "out.xlsx", absaetze=[p2, p3])
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Eins bleibt. Vier bleibt."
    _bestanden(a1)
    _bestanden(a2)


def test_a3_mit_zeile_eingefuegt_darueber(tmp_path):
    t = "Das System muss laufen."
    p = _excel(tmp_path, t)
    ab = _absatz(L("Das {-System-} {+Gerät+} muss laufen."), 0, len(t))
    ae = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)])
    neu = Operation("zeile_einfuegen", "CDE", 2, zellen={1: "0", 2: "Vorab", 3: "Neuer Absatz ganz oben."}, stil_von=2,
                    aenderung="A9", vor=True)
    a9 = _ae("A9", [neu], status=PRUEFEN, typ="absatz_neu")
    uebernehmen([ae, a9], p, tmp_path / "out.xlsx", {"A9"}, absaetze=[ab])
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Neuer Absatz ganz oben."
    assert _zelle(tmp_path / "out.xlsx", 3, 3) == "Das Gerät muss laufen."
    _bestanden(ae)
    assert a9.status == UEBERNOMMEN


def test_a3_geleerte_zelle_bleibt_uebernommen(tmp_path):
    t = "Ganz weg."
    p = _excel(tmp_path, t)
    ab = _absatz(L("{-Ganz-} {-weg.-}"), 0, len(t))
    ae = _ae("A1", [Operation("zelle_leeren", "CDE", 2, 3, aenderung="A1", erwartet=t)], typ="absatz_geloescht")
    ae.alt = t
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    _bestanden(ae)


def test_a3_ohne_absaetze_bleibt_alte_nachkontrolle(tmp_path):
    t = "Der Wert und der Wert."
    p = _excel(tmp_path, t)
    ae = _ae("A1", [_ers(4, 8, "Pegel", "A1", t)])
    uebernehmen([ae], p, tmp_path / "out.xlsx")
    assert (ae.status, ae.wirkung) == (UEBERNOMMEN, WIRKUNG_OK)        # Schema <= 2: keine Gegenprobe moeglich


# ---------------------------------------------------------------- Invariante: keine Automatik ohne Absatz
def test_invariante_schalter_an_ohne_absatz_wird_pruefen(tmp_path):
    t = "Das System muss laufen."
    p = _excel(tmp_path, t)
    ae = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)])
    uebernehmen([ae], p, tmp_path / "out.xlsx", cfg=_cfg(True), absaetze=[])
    assert ae.status == PRUEFEN and "ohne_absatz" in ae.pruefcodes and ae.hinweise
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == t                    # nicht angewendet


def test_invariante_schalter_an_mit_absatz_wird_angewendet(tmp_path):
    t = "Das System muss laufen."
    p = _excel(tmp_path, t)
    ab = _absatz(L("Das {-System-} {+Gerät+} muss laufen."), 0, len(t))
    ae = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)])
    uebernehmen([ae], p, tmp_path / "out.xlsx", cfg=_cfg(True), absaetze=[ab])
    _bestanden(ae)


def test_invariante_schalter_aus_wie_bisher(tmp_path):
    t = "Das System muss laufen."
    p = _excel(tmp_path, t)
    ae = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)])
    uebernehmen([ae], p, tmp_path / "out.xlsx", cfg=_cfg(False), absaetze=[])
    assert (ae.status, ae.wirkung) == (UEBERNOMMEN, WIRKUNG_OK) and "ohne_absatz" not in ae.pruefcodes


def test_invariante_gilt_nicht_fuer_schema2_freigaben_und_neue_zeilen(tmp_path):
    t = "Das System muss laufen."
    p = _excel(tmp_path, t)
    a_alt = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)])
    uebernehmen([a_alt], p, tmp_path / "o1.xlsx", cfg=_cfg(True), absaetze=None)     # Schema <= 2
    assert a_alt.status == UEBERNOMMEN
    a_frei = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)], status=PRUEFEN)
    uebernehmen([a_frei], p, tmp_path / "o2.xlsx", {"A1"}, cfg=_cfg(True), absaetze=[])  # vom Menschen freigegeben
    assert a_frei.status == UEBERNOMMEN and "ohne_absatz" not in a_frei.pruefcodes
    neu = Operation("zeile_einfuegen", "CDE", 3, zellen={1: "3", 2: "Drei", 3: "Neuer Absatz unten."}, stil_von=3,
                    aenderung="A2")
    a_zeile = _ae("A2", [neu], typ="absatz_neu")
    uebernehmen([a_zeile], p, tmp_path / "o3.xlsx", cfg=_cfg(True), absaetze=[])
    assert a_zeile.status == UEBERNOMMEN and "ohne_absatz" not in a_zeile.pruefcodes


def test_invariante_nimmt_die_gruppe_mit(tmp_path):
    t = "Das System muss laufen."
    p = _excel(tmp_path, t)
    ab = _absatz(L("Das {-System-} {+Gerät+} muss laufen."), 0, len(t))
    a1 = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)], gruppe="G")
    zweit = "Zweiter Absatz mit genug Text darin."
    a2 = _ae("A2", [_ers(0, 7, "Dritter", "A2", zweit, zeile=3)], gruppe="G")       # ohne Absatz
    uebernehmen([a1, a2], p, tmp_path / "out.xlsx", cfg=_cfg(True), absaetze=[ab])
    assert a2.status == PRUEFEN and "ohne_absatz" in a2.pruefcodes
    assert a1.status == PRUEFEN                                          # gleiche Gruppe: nur gemeinsam
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == t


# ---------------------------------------------------------------- neuer Absatz in bestehender Zelle (ohne PDF-Absatz)
def _neuer_absatz(pos, text, soll) -> Aenderung:
    ae = _ae("A5", [Operation("ersetzen", "CDE", 2, 3, pos, pos, text, aenderung="A5")], typ="absatz_neu")
    ae.neu = ae.soll_absatz = soll
    return ae


def test_neuer_absatz_als_eigene_zeile_besteht_auch_mit_pflicht(tmp_path):
    t = "Erster Satz.\nZweiter Satz."
    p = _excel(tmp_path, t)
    ae = _neuer_absatz(len("Erster Satz."), "\nNeuer Satz dazwischen.", "Neuer Satz dazwischen.")
    uebernehmen([ae], p, tmp_path / "out.xlsx", cfg=_cfg(True), absaetze=[])
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Erster Satz.\nNeuer Satz dazwischen.\nZweiter Satz."
    _bestanden(ae)


def test_neuer_absatz_mit_anderer_aenderung_in_der_zelle(tmp_path):
    t = "Erster Satz alt.\nZweiter Satz."
    p = _excel(tmp_path, t)
    a = t.index("alt")
    a1 = _ae("A1", [_ers(a, a + 3, "neu", "A1", t)])
    ab = _absatz(L("Erster Satz {-alt-}|{+neu+}|."), 0, t.index("\n"))
    ae = _neuer_absatz(len(t), "\nLetzter Satz.", "Letzter Satz.")
    uebernehmen([a1, ae], p, tmp_path / "out.xlsx", cfg=_cfg(True), absaetze=[ab])
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Erster Satz neu.\nZweiter Satz.\nLetzter Satz."
    _bestanden(a1)
    _bestanden(ae)


def test_neuer_absatz_nicht_als_eigene_zeile_wird_pruefen(tmp_path):
    t = "Erster Satz.\nZweiter Satz."
    p = _excel(tmp_path, t)
    ae = _neuer_absatz(len("Erster Satz."), " Neuer Satz dazwischen.", "Neuer Satz dazwischen.")   # ohne Umbruch angehaengt
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[])
    _durchgefallen(ae, "gegenprobe_neuer_absatz")


def test_neuer_absatz_mit_falschem_text_wird_pruefen(tmp_path):
    t = "Erster Satz.\nZweiter Satz."
    p = _excel(tmp_path, t)
    ae = _neuer_absatz(len("Erster Satz."), "\nNeuer Satz dazwischen", "Neuer Satz dazwischen.")    # Punkt fehlt
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[])
    _durchgefallen(ae, "gegenprobe_neuer_absatz")


# ---------------------------------------------------------------- neue Section-Nummer als Zahl (RT-212)
def _zahl_excel(tmp_path, nummern: list[tuple[float | int, str]]):
    wb = Workbook()
    ws = wb.active
    ws.title = "CDE"
    ws.append(["Nr.", "Titel", "Inhalt"])
    for i, (n, fmt) in enumerate(nummern, start=2):
        ws.cell(i, 1).value = n
        ws.cell(i, 1).number_format = fmt
        ws.cell(i, 2).value = f"Titel {i}"
        ws.cell(i, 3).value = f"Inhalt der Zeile {i} mit genug Text darin."
    p = tmp_path / "zahl.xlsx"
    wb.save(str(p))
    return p


def _neue_section(nummer: str, nach: int) -> Aenderung:
    op = Operation("zeile_einfuegen", "CDE", nach, zellen={1: nummer, 2: "Neu", 3: "Neuer Inhalt der Section."},
                   stil_von=nach, aenderung="A1")
    return _ae("A1", [op], typ="section_neu")


def test_neue_section_nummer_wird_zahl_mit_gleichem_format(tmp_path):
    p = _zahl_excel(tmp_path, [(2.1, "0.0"), (2.2, "0.0"), (3, "0")])
    ae = _neue_section("2.3", 3)
    uebernehmen([ae], p, tmp_path / "out.xlsx")
    c = load_workbook(str(tmp_path / "out.xlsx"))["CDE"].cell(4, 1)
    assert isinstance(c.value, float) and abs(c.value - 2.3) < 1e-9 and c.number_format == "0.0"
    assert (ae.status, ae.wirkung) == (UEBERNOMMEN, WIRKUNG_OK)


def test_neue_section_nummer_bleibt_text_wenn_format_anders_aussieht(tmp_path):
    p = _zahl_excel(tmp_path, [(2, "0"), (3, "0")])
    ae = _neue_section("2.3", 2)                       # "0" zeigt keine Nachkommastellen: als Zahl stuende dort "2"
    uebernehmen([ae], p, tmp_path / "out.xlsx")
    assert _zelle(tmp_path / "out.xlsx", 3, 1) == "2.3"
    assert any("als Text" in h for h in ae.hinweise)
    assert ae.wirkung == WIRKUNG_OK


def test_neue_section_nummer_ganzzahl(tmp_path):
    p = _zahl_excel(tmp_path, [(1, "0"), (2, "0")])
    ae = _neue_section("3", 3)
    uebernehmen([ae], p, tmp_path / "out.xlsx")
    c = load_workbook(str(tmp_path / "out.xlsx"))["CDE"].cell(4, 1)
    assert c.value == 3 and not isinstance(c.value, str)
    assert ae.wirkung == WIRKUNG_OK


# ---------------------------------------------------------------- Abwahl
def test_abwahl_wird_nicht_eingetragen(tmp_path):
    t = "Das System muss laufen."
    p = _excel(tmp_path, t)
    a1 = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)])
    zweit = "Zweiter Absatz mit genug Text darin."
    a2 = _ae("A2", [_ers(0, 7, "Dritter", "A2", zweit, zeile=3)])
    uebernehmen([a1, a2], p, tmp_path / "out.xlsx", abwahl={"A1"})
    assert a1.status == NICHT and "Vom Menschen abgewählt." in a1.grund and a1.wirkung == ""
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == t
    assert a2.status == UEBERNOMMEN and _zelle(tmp_path / "out.xlsx", 3, 3).startswith("Dritter")


def test_abwahl_nimmt_die_gruppe_mit(tmp_path):
    t = "Das System muss laufen."
    p = _excel(tmp_path, t)
    a1 = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)], gruppe="G1")
    zweit = "Zweiter Absatz mit genug Text darin."
    a2 = _ae("A2", [_ers(0, 7, "Dritter", "A2", zweit, zeile=3)], gruppe="G1")
    a3 = _ae("A3", [_ers(11, 15, "darf", "A3", t)])
    uebernehmen([a1, a2, a3], p, tmp_path / "out.xlsx", abwahl={"A2"})
    assert a1.status == a2.status == NICHT
    assert "Vom Menschen abgewählt." in a1.grund and a1.hinweise                 # mitgenommen: mit Hinweis
    assert _zelle(tmp_path / "out.xlsx", 3, 3) == zweit
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Das System darf laufen."
    assert a3.status == UEBERNOMMEN


def test_abwahl_eines_vorschlags_wird_ignoriert(tmp_path):
    t = "Das System muss laufen."
    p = _excel(tmp_path, t)
    a1 = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)], status=PRUEFEN)
    uebernehmen([a1], p, tmp_path / "out.xlsx", abwahl={"A1"})
    assert a1.status == PRUEFEN


# ---------------------------------------------------------------- neue_spanne: Randfaelle (src/gegenprobe.py)
def test_neue_spanne_eigene_loeschung_ueber_den_rand():
    from src.gegenprobe import neue_spanne
    o = Operation("ersetzen", "CDE", 2, 3, 13, 36, "", aenderung="A1")           # Absatz [13, 35) samt "\n" dahinter
    assert neue_spanne(13, 35, [o], {"A1"}) == (13, 13)
    o = Operation("ersetzen", "CDE", 2, 3, 12, 35, "", aenderung="A1")           # samt "\n" davor
    assert neue_spanne(13, 35, [o], {"A1"}) == (12, 12)


# ---------------------------------------------------------------- Zufall: richtig bleibt ok, falsch wird nie still ok
def test_zufall_richtig_angewendet_besteht_falsch_angewendet_nie_still_ok(tmp_path):
    import random
    rnd = random.Random(11)
    woerter = ["Ventil", "Pumpe", "Druck", "Wert", "Anlage", "Regler", "Kessel", "Messung"]
    for n in range(30):
        saetze = [[rnd.choice(woerter) for _ in range(rnd.randint(3, 6))] for _ in range(rnd.randint(1, 3))]
        zeilen = [" ".join(s) + "." for s in saetze]
        t = "\n".join(zeilen)
        k = rnd.randrange(len(zeilen))
        start = sum(len(z) + 1 for z in zeilen[:k])
        w = rnd.randrange(len(saetze[k]))
        alt_wort = saetze[k][w]
        neu_wort = "Zusatz"
        laeufe: list[Lauf] = []
        for i, x in enumerate(saetze[k]):
            if i:
                laeufe.append(Lauf("s", " "))
            if i == w:
                laeufe += [Lauf("d", x, "A1"), Lauf("s", " "), Lauf("a", neu_wort, "A1")]
            else:
                laeufe.append(Lauf("k", x))
        laeufe.append(Lauf("k", "."))
        ab = _absatz(laeufe, start, start + len(zeilen[k]))
        pos = start + sum(len(x) + 1 for x in saetze[k][:w])
        soll = t[:pos] + neu_wort + t[pos + len(alt_wort):]
        # richtig
        p = _excel(tmp_path, t, name=f"a{n}.xlsx")
        ae = _ae("A1", [_ers(pos, pos + len(alt_wort), neu_wort, "A1", t)])
        uebernehmen([ae], p, tmp_path / f"ok{n}.xlsx", absaetze=[ab])
        assert _zelle(tmp_path / f"ok{n}.xlsx", 2, 3) == soll
        _bestanden(ae)
        # falsch: ein anderes Wort (im selben oder einem anderen Absatz der Zelle) wird ersetzt
        kandidaten = [i for i in range(len(t)) if (i == 0 or not t[i - 1].isalpha()) and t[i].isalpha() and i != pos]
        if not kandidaten:
            continue
        q = rnd.choice(kandidaten)
        ende = q
        while ende < len(t) and t[ende].isalpha():
            ende += 1
        falsch = _ae("A1", [_ers(q, ende, neu_wort, "A1", t)])
        uebernehmen([falsch], p, tmp_path / f"f{n}.xlsx", absaetze=[ab])
        ist = _zelle(tmp_path / f"f{n}.xlsx", 2, 3)
        if ist != soll:
            assert not (falsch.status == UEBERNOMMEN and falsch.wirkung == WIRKUNG_OK), (t, ist, soll)


# ================================================================ Auftrag 003: Markierung + Patch-Schreiber (ADR 0005/0007)
import pytest                                                      # noqa: E402
from openpyxl.styles import Font                                   # noqa: E402

from src.markierung import ALT, NEU, NORMAL, ansichten            # noqa: E402
from src.xlsx_patch import pruefe_treue                           # noqa: E402


def _register(tmp_path, texte: list[str], name: str = "r.xlsx", extra=None):
    """Kleines Register: Nr | Sub | Text (Arial 11 fett) | Spezifikation (gleich) | Lfd (verschieden) | Status (gleich)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "CDE"
    ws.append(["Nr.", "Sub", "Inhalt", "Spezifikation", "Lfd", "Status"])
    for i, t in enumerate(texte, start=2):
        ws.append([str(i - 1), "AB"[i % 2], t, "SPEC-7", i * 10, "offen"])
        ws.cell(i, 3).font = Font(name="Arial", sz=11, b=True)
        ws.row_dimensions[i].height = 28
    if extra:
        extra(ws)
    p = tmp_path / name
    wb.save(str(p))
    return p


def _laeufe_der_zelle(p, r, c):
    z = load_workbook(str(p), rich_text=True)["CDE"].cell(r, c)
    return z, ansichten(z)


@pytest.mark.markierung_an
def test_markiert_ersetzung_alt_vor_neu_mit_pdf_leerraum(tmp_path):
    t = "Das System muss laufen."
    p = _register(tmp_path, [t, "Zweiter Absatz mit genug Text darin."])
    ab = _absatz(L("Das {-System-} {+Gerät+} muss laufen."), 0, len(t))
    ae = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)])
    res = uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    _bestanden(ae)
    z, (alt, neu, st) = _laeufe_der_zelle(tmp_path / "out.xlsx", 2, 3)
    assert alt.replace("  ", " ") == t and neu == "Das  Gerät muss laufen."
    assert [(s.format, s.text) for s in st] == [(NORMAL, "Das "), (ALT, "System"), (NORMAL, " "), (NEU, "Gerät"),
                                                 (NORMAL, " muss laufen.")]
    assert all(b.font.b and b.font.rFont == "Arial" for b in z.value)            # Grundschrift in jedem Lauf
    assert ae.stuecke == [["normal", "Das "], ["alt", "System"], ["normal", " "], ["neu", "Gerät"], ["normal", " muss laufen."]]
    assert pruefe_treue(p, tmp_path / "out.xlsx", ["xl/worksheets/sheet1.xml", "xl/sharedStrings.xml",
                                                   "xl/_rels/workbook.xml.rels", "[Content_Types].xml"]) == []
    assert not res.warnungen


@pytest.mark.markierung_an
def test_markiert_verschmolzene_ersetzung_ohne_trenner(tmp_path):
    t = "Erster Satz alt."
    p = _register(tmp_path, [t])
    ab = _absatz(L("Erster Satz {-alt-}|{+neu+}|."), 0, len(t))
    a = t.index("alt")
    ae = _ae("A1", [_ers(a, a + 3, "neu", "A1", t)])
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    _bestanden(ae)
    _, (_, _, st) = _laeufe_der_zelle(tmp_path / "out.xlsx", 2, 3)
    assert [(s.format, s.text) for s in st] == [(NORMAL, "Erster Satz "), (ALT, "alt"), (NEU, "neu"), (NORMAL, ".")]


@pytest.mark.markierung_an
def test_markiert_falscher_plan_kommt_nicht_in_die_datei(tmp_path):
    t = "Der Wert und der Wert."
    p = _register(tmp_path, [t])
    ab = _absatz(L("Der Wert und der {-Wert-} {+Pegel+}|."), 0, len(t))
    ae = _ae("A1", [_ers(4, 8, "Pegel", "A1", t)])                               # erstes statt zweites Vorkommen
    res = uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    _durchgefallen(ae, "gegenprobe_text")
    assert any("nicht in die Ergebnis-Excel geschrieben" in h for h in ae.hinweise)
    z = load_workbook(str(tmp_path / "out.xlsx"), rich_text=True)["CDE"]["C2"]
    assert z.value == t                                                         # keine ungepruefte Automatik in der Datei
    assert any("Gegenprobe" in w for w in res.warnungen)


@pytest.mark.markierung_an
def test_markiert_gruppe_wird_gemeinsam_zurueckgehalten(tmp_path):
    t, zweit = "Der Wert und der Wert.", "Zweiter Absatz mit genug Text darin."
    p = _register(tmp_path, [t, zweit])
    ab = _absatz(L("Der Wert und der {-Wert-} {+Pegel+}|."), 0, len(t))
    ab2 = _absatz(L("{-Zweiter-} {+Dritter+} Absatz mit genug Text darin.", "A2"), 0, len(zweit), zeile=3, pid="P2")
    a1 = _ae("A1", [_ers(4, 8, "Pegel", "A1", t)], gruppe="G1")
    a2 = _ae("A2", [_ers(0, 7, "Dritter", "A2", zweit, zeile=3)], gruppe="G1")
    uebernehmen([a1, a2], p, tmp_path / "out.xlsx", absaetze=[ab, ab2])
    _durchgefallen(a1, "gegenprobe_text")
    assert a2.status == PRUEFEN and "gegenprobe_gruppe" in a2.pruefcodes and a2.wirkung == WIRKUNG_ABWEICHUNG
    assert _zelle(tmp_path / "out.xlsx", 3, 3) == zweit                          # auch die richtige bleibt draussen


@pytest.mark.markierung_an
def test_markiert_ganz_gestrichener_absatz_bleibt_durchgestrichen_stehen(tmp_path):
    t = "Ganz weg."
    p = _register(tmp_path, [t, "Bleibt stehen mit genug Text darin."])
    ab = _absatz(L("{-Ganz-} {-weg.-}"), 0, len(t))
    ae = _ae("A1", [Operation("zelle_leeren", "CDE", 2, 3, aenderung="A1", erwartet=t)], typ="absatz_geloescht")
    ae.alt = t
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    _bestanden(ae)
    z, (alt, neu, st) = _laeufe_der_zelle(tmp_path / "out.xlsx", 2, 3)
    assert alt == t and neu == "" and [s.format for s in st] == [ALT]
    assert all(b.font.b and b.font.strike for b in z.value)
    assert load_workbook(str(tmp_path / "out.xlsx"))["CDE"].max_row == 3          # die Zeile bleibt


@pytest.mark.markierung_an
def test_markiert_einfuegung_und_loeschung_mit_alten_markierungen(tmp_path):
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont

    def alt_markiert(ws):
        f = dict(rFont="Arial", sz=11, b=True)
        ws["C2"] = CellRichText(TextBlock(InlineFont(**f), "Die Pumpe "), TextBlock(InlineFont(color="FFC00000", **f),
                                "läuft "), TextBlock(InlineFont(**f), "leise im Keller."))
    t = "Die Pumpe läuft leise im Keller."
    p = _register(tmp_path, [t], extra=alt_markiert)
    ab = _absatz(L("Die Pumpe läuft {-leise-} im Keller {+heute+}|."), 0, len(t))
    s = t.index("leise")
    ae = _ae("A1", [_ers(s - 1, s + 5, "", "A1", t), _ein(len(t) - 1, " heute", "A1")])
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    _bestanden(ae)
    z = load_workbook(str(tmp_path / "out.xlsx"), rich_text=True)["CDE"]["C2"]
    laeufe = [(b.text, b.font.color.rgb if b.font.color is not None else None, bool(b.font.strike)) for b in z.value]
    # Ansichten nach Nicos Regel (ADR 0005 Nachtrag): neu = hellrot FFFF0000, alt = durchgestrichen, fruehere Farben normal
    assert "".join(x for x, f, s in laeufe if f != "FFFF0000" or s) == t
    assert "".join(x for x, f, s in laeufe if not s) == "Die Pumpe läuft im Keller heute."
    assert ("läuft", "FFC00000", False) in laeufe                               # alte Markierung bleibt, wie sie war
    assert "".join(x for x, f, s in laeufe if s and f == "FF00B0F0") == " leise" and (" heute", "FFFF0000", False) in laeufe


@pytest.mark.markierung_an
def test_markiert_neue_zeile_rot_mit_konstanten_und_leeren_spalten(tmp_path):
    p = _register(tmp_path, ["Erster Absatz mit genug Text.", "Zweiter Absatz mit genug Text.",
                             "Dritter Absatz mit genug Text.", "Vierter Absatz mit genug Text."])
    op = Operation("zeile_einfuegen", "CDE", 3, zellen={1: "2.1", 3: "Neuer Absatz dazwischen."}, stil_von=3, aenderung="A1")
    ae = _ae("A1", [op], typ="absatz_neu")
    res = uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[])
    assert (ae.status, ae.wirkung) == (UEBERNOMMEN, WIRKUNG_OK), (ae.grund, ae.hinweise)
    ws = load_workbook(str(tmp_path / "out.xlsx"), rich_text=True)["CDE"]
    z = ws["C4"]
    assert ansichten(z)[1] == "Neuer Absatz dazwischen." and [s.format for s in ansichten(z)[2]] == [NEU]
    assert all(b.font.b and b.font.rFont == "Arial" for b in z.value)            # Grundschrift der Vorlagezeile
    assert ws["D4"].value == "SPEC-7" and ws["F4"].value == "offen"               # Konstanten kopiert
    assert ws["E4"].value is None and ws["B4"].value is None and op.leer_spalten == [2, 5]
    assert any("Spalten B, E leer" in h for h in ae.hinweise)
    assert ws.row_dimensions[4].height == 28 and ws["C5"].value == "Dritter Absatz mit genug Text."
    assert not [w for w in res.warnungen if "Gegenprobe" in w]


@pytest.mark.markierung_an
def test_markiert_neue_section_nummer_als_zahl_rot(tmp_path):
    p = _zahl_excel(tmp_path, [(2.1, "0.0"), (2.2, "0.0"), (3, "0")])
    ae = _neue_section("2.3", 3)
    uebernehmen([ae], p, tmp_path / "out.xlsx")
    c = load_workbook(str(tmp_path / "out.xlsx"))["CDE"].cell(4, 1)
    assert isinstance(c.value, float) and abs(c.value - 2.3) < 1e-9 and c.number_format == "0.0"
    assert c.font.color is not None and c.font.color.rgb == "FFFF0000"
    assert (ae.status, ae.wirkung) == (UEBERNOMMEN, WIRKUNG_OK)


@pytest.mark.markierung_an
def test_markiert_zahl_ersetzt_nur_schriftfarbe(tmp_path):
    p = _zahl_excel(tmp_path, [(2, "0"), (3, "0")])
    ab = Absatz("P1", "CDE", [(2, 1, 0, 1)], L("{-2-} {+5+}"))
    ae = _ae("A1", [Operation("ersetzen", "CDE", 2, 1, 0, 1, "5", aenderung="A1", erwartet="2")])
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[ab])
    _bestanden(ae)
    c = load_workbook(str(tmp_path / "out.xlsx"))["CDE"].cell(2, 1)
    assert c.value == 5 and c.data_type == "n" and c.font.color.rgb == "FFFF0000"


@pytest.mark.markierung_an
def test_markiert_neuer_absatz_in_bestehender_zelle(tmp_path):
    t = "Erster Satz.\nZweiter Satz."
    p = _register(tmp_path, [t])
    ae = _neuer_absatz(len("Erster Satz."), "\nNeuer Satz dazwischen.", "Neuer Satz dazwischen.")
    uebernehmen([ae], p, tmp_path / "out.xlsx", cfg=_cfg(True), absaetze=[])
    _bestanden(ae)
    _, (alt, neu, st) = _laeufe_der_zelle(tmp_path / "out.xlsx", 2, 3)
    assert alt == t and neu == "Erster Satz.\nNeuer Satz dazwischen.\nZweiter Satz."
    falsch = _neuer_absatz(len("Erster Satz."), " Neuer Satz dazwischen.", "Neuer Satz dazwischen.")
    uebernehmen([falsch], p, tmp_path / "out2.xlsx", absaetze=[])
    _durchgefallen(falsch, "gegenprobe_neuer_absatz")
    assert _zelle(tmp_path / "out2.xlsx", 2, 3) == t


def test_patch_schreiber_kennt_keine_gelb_markierung(tmp_path):
    t = "Das System muss laufen."
    p = _excel(tmp_path, t)
    cfg = copy.deepcopy(config.lade("zuordnung"))
    cfg["markierung"]["geaenderte_zellen"] = "gelb"
    ae = _ae("A1", [_ers(4, 10, "Gerät", "A1", t)])
    res = uebernehmen([ae], p, tmp_path / "out.xlsx", cfg=cfg)
    assert ae.status == PRUEFEN and "nur mit dem Schreiber" in ae.grund
    assert any("geaenderte_zellen" in w for w in res.warnungen)
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == t


def test_patch_schreiber_entfernt_keine_zeilen(tmp_path):
    p = _excel(tmp_path, "Ganz weg.")
    cfg = copy.deepcopy(config.lade("zuordnung"))
    cfg["geloeschtes"]["inhalt"] = "zeile_entfernen"
    ops = [Operation("zelle_leeren", "CDE", 2, c, aenderung="A1", erwartet=v) for c, v in ((1, "1"), (2, "Eins"), (3, "Ganz weg."))]
    ae = _ae("A1", ops, typ="section_geloescht")
    ae.alt = "1 Eins Ganz weg."
    uebernehmen([ae], p, tmp_path / "out.xlsx", cfg=cfg)
    assert ae.status == PRUEFEN and "openpyxl" in ae.grund
    assert _zelle(tmp_path / "out.xlsx", 2, 3) == "Ganz weg."


@pytest.mark.schreiber_openpyxl
@pytest.mark.filterwarnings("ignore:Data Validation extension is not supported")
def test_openpyxl_schreiber_warnt_vor_verlust(tmp_path):
    from tests.test_xlsx_patch import BLATT, baue_register
    p = baue_register(tmp_path / "reg.xlsx", n=40)
    t = load_workbook(str(p))[BLATT]["C20"].value
    s = t.index("Wasser")
    ae = _ae("A1", [Operation("ersetzen", BLATT, 20, 3, s, s + 6, "Öl", aenderung="A1", erwartet="Wasser")])
    res = uebernehmen([ae], p, tmp_path / "out.xlsx")
    assert ae.status == UEBERNOMMEN
    assert any("Bilder und Formen" in w and "Kennzeichnung" in w for w in res.warnungen)


def test_patch_schreiber_neue_zeile_zieht_formeln_und_bereiche_mit(tmp_path):
    """Was beim openpyxl-Schreiber ein Hindernis war (Formeln, bedingte Formatierung, Filter, Link darunter), zieht der
    Patch-Schreiber selbst mit: automatisch und richtig."""
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill

    def eig(ws):
        for r in range(2, 6):
            ws.cell(r, 7).value = f"=LEN(C{r})"
        ws["G8"] = "=SUM(G2:G5)"
        ws.auto_filter.ref = "A1:G5"
        ws.conditional_formatting.add("C2:C5", CellIsRule(operator="equal", formula=['"x"'],
                                                          fill=PatternFill("solid", bgColor="FFFF00")))
        ws["C5"].hyperlink = "https://example.org/x"
    p = _register(tmp_path, [f"Absatz {i} mit genug Text darin." for i in range(1, 5)], extra=eig)
    op = Operation("zeile_einfuegen", "CDE", 3, zellen={1: "2.1", 3: "Neuer Absatz."}, stil_von=3, aenderung="A1")
    ae = _ae("A1", [op], typ="absatz_neu")
    uebernehmen([ae], p, tmp_path / "out.xlsx", absaetze=[])
    assert (ae.status, ae.wirkung) == (UEBERNOMMEN, WIRKUNG_OK), (ae.grund, ae.hinweise)
    ws = load_workbook(str(tmp_path / "out.xlsx"))["CDE"]
    assert [ws.cell(r, 7).value for r in range(2, 7)] == [f"=LEN(C{r})" for r in range(2, 7)]   # Vorlage-Formel uebertragen
    assert ws["G9"].value == "=SUM(G2:G6)" and ws.auto_filter.ref == "A1:G6"
    assert any("C2:C6" in str(r.sqref) for r in ws.conditional_formatting)
    assert ws["C6"].hyperlink is not None and ws["C5"].hyperlink is None
    assert op.formeln == {7: "=LEN(C3)"}


@pytest.mark.markierung_an
@pytest.mark.parametrize("layout", ["L1", "L3"])
def test_markiert_ende_zu_ende_und_zweiter_lauf_aendert_nichts(tmp_path, layout):
    """Planer + Applier mit Markierung auf einem erfundenen Szenario: alles besteht die Gegenprobe vor dem Schreiben;
    ein zweiter Lauf auf dem Ergebnis traegt nichts mehr ein (Idempotenz) -- die Datei bleibt bytegleich."""
    from src.excel_model import ExcelModell
    from src.pdf_reader import analysiere_pdf
    from src.planner import plane
    from tests.fixtures.pdf_writer import baue_pdf
    from tests.fixtures.szenario import ELEMENTE, baue_excel, zustaende

    baue_pdf(tmp_path / "a.pdf", ELEMENTE, "farbe")
    baue_excel(tmp_path / "alt.xlsx", zustaende(ELEMENTE, False), layout)
    pdf = analysiere_pdf(tmp_path / "a.pdf")
    plan = plane(pdf, ExcelModell(tmp_path / "alt.xlsx"))
    auto = [a for a in plan.aenderungen if a.status == UEBERNOMMEN]
    assert auto
    uebernehmen(plan.aenderungen, tmp_path / "alt.xlsx", tmp_path / "out.xlsx", absaetze=plan.absaetze)
    assert all(a.status == UEBERNOMMEN and a.wirkung == WIRKUNG_OK for a in auto), \
        [(a.id, a.status, a.pruefcodes) for a in auto if a.status != UEBERNOMMEN]
    plan2 = plane(pdf, ExcelModell(tmp_path / "out.xlsx"))
    uebernehmen(plan2.aenderungen, tmp_path / "out.xlsx", tmp_path / "out2.xlsx", absaetze=plan2.absaetze)
    assert not [a for a in plan2.aenderungen if a.status == UEBERNOMMEN]
    assert pruefe_treue(tmp_path / "out.xlsx", tmp_path / "out2.xlsx") == []
    assert any("schon_eingetragen" in a.pruefcodes or "schon" in a.grund for a in plan2.aenderungen)
