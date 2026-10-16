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
