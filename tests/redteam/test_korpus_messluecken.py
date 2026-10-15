"""Red-Team R3 (Analyse der stillen Fehler im Entwicklungskorpus), Teil 3: SCHWAECHEN DER MESSUNG (Sabotage-Tests).

Frage: Wo koennte `python -m src.cli messen` ein Team taeuschen, weil die Zahlen zu GUT aussehen? Methode: Ein sauberes
Basis-Paar (das Tool macht alles richtig, 0 stille Fehler) wird so SABOTIERT, wie ein fehlerhaftes Tool es tun koennte -- die
Sabotage greift nur im Messwerkzeug an (das Ergebnis des Tools wird verfaelscht), nichts in src/ wird geaendert.

  * Teil A (WAECHTER, normale Tests, muessen gruen bleiben): Sabotage, die die Messung bemerkt. Das schuetzt davor, dass
    ein spaeterer Umbau von messen.py diese Faelle verliert.
  * Teil B (Besitzer Mess-Bauer): Sabotage, die die Messung frueher NICHT bemerkte (RT-135 bis RT-142). Sie ist repariert, die
    Tests sind jetzt normale Tests (ohne xfail) und halten die neue Wirkung fest.

Alles ERFUNDEN (Regel 16). Vorschlaege zur Reparatur: tests/redteam/README_korpus.md.
"""
from __future__ import annotations

import json
from pathlib import Path
from unittest import mock

import pytest
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation

from src import config, messen
from src.models import PRUEFEN
from tests.fixtures.pdf_writer import PdfDokument, baue_pdf, scan_bild
from tests.fixtures.szenario import baue_excel, zustaende
from tests.redteam._bau import aufbau, scheitern
from tests.redteam._bau_korpus import (BASIS, KOPF, auto_lauf, baue_basis_paar, finde_zelle, messe_mit_fertigem_ergebnis,
                                       messe_mit_sabotage, paar)


def _bemerkt(res: dict) -> bool:
    """Hat die Messung IRGENDETWAS bemerkt (stiller Fehler, Kollateral, falsche Zelle, nicht erreichte Soll-Zelle)?"""
    return bool(res["stille_fehler"] or res["kollateral"] or res["auto_richtig"] != res["auto_gesamt"]
                or res["erreicht_auto"] != res["soll_zellen"])


def _basis_sauber(tmp_path) -> Path:
    ordner = baue_basis_paar(tmp_path / "basis")
    sauber = messen.messe_paar(ordner, False)
    aufbau(not _bemerkt(sauber) and sauber["soll_zellen"] >= 3 and sauber["idempotent"],
           f"Das Basis-Paar ist nicht sauber: {sauber['stille_fehler']} stille Fehler")
    return ordner


# ---- Sabotage-Funktionen (Workbook -> Workbook) ----------------------------------------------------------------------
def _cde(wb):
    return wb["CDE"]


def _ersetze_in(wb, suche: str, alt: str, neu: str) -> None:
    c = finde_zelle(_cde(wb), suche)
    aufbau(alt in str(c.value), f"'{alt}' steht nicht in der Zelle")
    c.value = str(c.value).replace(alt, neu)


SABOTAGE_ERKANNT = {
    "falsches_wort": lambda wb: _ersetze_in(wb, "monatlich", "monatlich", "woechentlich"),
    "zahl_vertauscht": lambda wb: _ersetze_in(wb, "ms nicht", "120", "12"),
    "aenderung_zurueckgenommen": lambda wb: _ersetze_in(wb, "monatlich", "monatlich", "jaehrlich"),
    "absaetze_vertauscht": lambda wb: _vertausche(wb),
    "text_in_bemerkungsspalte": lambda wb: setattr(_cde(wb).cell(3, 4), "value", "XXX"),
    "zeile_doppelt_eingefuegt": lambda wb: _doppelte_zeile(wb),
    "zeile_geloescht": lambda wb: _cde(wb).delete_rows(finde_zelle(_cde(wb), "Lieferung umfasst").row),
    "anderes_blatt_veraendert": lambda wb: wb["Deckblatt"].__setitem__("A2", "Geaendert"),
}


def _vertausche(wb) -> None:
    a, b = finde_zelle(_cde(wb), "Messprotokoll"), finde_zelle(_cde(wb), "Alle Steckverbindungen")
    a.value, b.value = b.value, a.value


def _doppelte_zeile(wb) -> None:
    ws = _cde(wb)
    r = finde_zelle(ws, "Ersatzteile").row
    ws.insert_rows(r + 1)
    for c in range(1, 5):
        ws.cell(r + 1, c).value = ws.cell(r, c).value


# ==================================================================================================
# Teil A: WAECHTER -- diese Sabotage bemerkt die Messung heute
# ==================================================================================================
def test_waechter_basis_paar_ist_sauber(tmp_path):
    res = messen.messe_paar(baue_basis_paar(tmp_path / "basis"), False)
    assert not _bemerkt(res) and res["idempotent"] and res["original_unveraendert"]


@pytest.mark.parametrize("name", sorted(SABOTAGE_ERKANNT))
def test_waechter_messung_bemerkt_sabotage(tmp_path, name):
    ordner = _basis_sauber(tmp_path)
    res = messe_mit_sabotage(ordner, SABOTAGE_ERKANNT[name])
    assert _bemerkt(res), f"Sabotage '{name}' wurde von der Messung nicht bemerkt: {res['stille_fehler']} stille Fehler"
    assert res["stille_fehler"] >= 1, f"Sabotage '{name}' ist kein stiller Fehler in der Messung"


def test_waechter_zurueckgenommene_aenderung_heisst_keine_wirkung(tmp_path):
    res = messe_mit_sabotage(_basis_sauber(tmp_path), SABOTAGE_ERKANNT["aenderung_zurueckgenommen"])
    assert "keine_wirkung" in res["stille_fehler_nach_art"]


# ==================================================================================================
# Teil B1: Die Messung vergleicht nur NORMALISIERTEN TEXT (messen._k: NFKC, unsichtbare Zeichen weg, Leerraum zusammen)
# ==================================================================================================
SABOTAGE_TEXT = {
    "absaetze_zu_einem_verschmolzen": lambda wb: _ersetze_in(wb, "Ersatzteile", "\n", " "),
    "doppelte_leerzeichen": lambda wb: _ersetze_in(wb, "Messprotokoll", " ", "  "),
    "geschuetzte_leerzeichen": lambda wb: _ersetze_in(wb, "Messprotokoll", " ", " "),
    "weicher_trennstrich_im_wort": lambda wb: _ersetze_in(wb, "Messprotokoll", "Messprotokoll", "Mess­protokoll"),
    "unsichtbares_zeichen_im_wort": lambda wb: _ersetze_in(wb, "Messprotokoll", "Messprotokoll", "Mess​protokoll"),
    "ligatur_fi": lambda wb: _ersetze_in(wb, "findet", "findet", "ﬁndet"),
}


@pytest.mark.redteam
@pytest.mark.parametrize("name", sorted(SABOTAGE_TEXT))
def test_rt135_messung_bemerkt_unsichtbare_textverfaelschung(tmp_path, name):
    res = messe_mit_sabotage(_basis_sauber(tmp_path), SABOTAGE_TEXT[name])
    if not _bemerkt(res):
        scheitern(f"Verfaelschung '{name}' bleibt unbemerkt: stille Fehler {res['stille_fehler']}, Kollateral "
                  f"{res['kollateral']}, Treffergenauigkeit {res['auto_richtig']} von {res['auto_gesamt']} = 100 %")


# ==================================================================================================
# Teil B2: Die Messung vergleicht nur WERTE -- Struktur und Format sind unsichtbar (vgl. R1: RT-18..26, RT-35/36)
# ==================================================================================================
def _verbinden(wb) -> None:
    _cde(wb).merge_cells("G2:H3")


def _pruefung(wb) -> None:
    dv = DataValidation(type="list", formula1='"a,b"')
    dv.add("D2:D6")
    _cde(wb).add_data_validation(dv)


def _fett_rot(wb) -> None:
    for zeile in _cde(wb).iter_rows(min_row=2):
        for c in zeile:
            c.font = Font(bold=True, color="FF0000")


SABOTAGE_STRUKTUR = {
    "blatt_umbenannt": lambda wb: setattr(_cde(wb), "title", "CDE_neu"),
    "blatt_ausgeblendet": lambda wb: setattr(wb["Deckblatt"], "sheet_state", "hidden"),
    "zeile_ausgeblendet": lambda wb: setattr(_cde(wb).row_dimensions[3], "hidden", True),
    "spalte_ausgeblendet": lambda wb: setattr(_cde(wb).column_dimensions["D"], "hidden", True),
    "zellen_verbunden": _verbinden,
    "datenpruefung_dazu": _pruefung,
    "schrift_fett_rot": _fett_rot,
}


STRUKTUR_ART = {"blatt_umbenannt": "blatt", "blatt_ausgeblendet": "sichtbarkeit", "zeile_ausgeblendet": "zeile_versteckt",
                "spalte_ausgeblendet": "spalte_versteckt", "zellen_verbunden": "verbunden", "datenpruefung_dazu": "validierung",
                "schrift_fett_rot": "schrift"}


@pytest.mark.redteam
@pytest.mark.parametrize("name", sorted(SABOTAGE_STRUKTUR))
def test_rt136_messung_bemerkt_strukturschaden(tmp_path, name):
    """Der Befund `struktur_geaendert` (eigene Zaehlung, ohne Textinhalte) benennt die Art des Schadens."""
    res = messe_mit_sabotage(_basis_sauber(tmp_path), SABOTAGE_STRUKTUR[name])
    art = STRUKTUR_ART[name]
    if not res.get("struktur_geaendert") or not res["struktur_nach_art"].get(art):
        scheitern(f"Strukturschaden '{name}' bleibt unbemerkt (struktur_geaendert {res.get('struktur_geaendert')}, "
                  f"{res.get('struktur_nach_art')}, stille Fehler {res['stille_fehler']}, Hinweise {res['hinweise']})")
    bericht = messen.fasse_zusammen([dict(res, name="synth_9100")], "entwicklung", "1", False)
    ampel = next(k["ampel"] for k in bericht["kriterien"] if k["name"].startswith("Struktur unverändert"))
    assert ampel == "NICHT ok"


def test_rt136_waechter_zeilen_einfuegen_verschiebt_legitim_und_ist_kein_befund(tmp_path):
    """Gegenprobe: Das Tool fuegt Zeilen in der Mitte ein (die Zeile mit eigener Hoehe darunter wandert mit). Das ist KEIN
    Strukturbefund -- sonst waere jede Einfuegung ein Phantom-Fehler. (Ausgeblendete Zeilen, verbundene Zellen und Validierung
    UNTER der Einfuegestelle setzt das Tool absichtlich auf 'pruefen'; deren Verschiebung prueft test_messen.py direkt.)"""
    elemente = [("h", "1 Eins"), ("p", "Der Nachweis erfolgt durch ein Messprotokoll der Abnahme."),
                ("h", "2 Zwei"), ("p", "Die Reaktionszeit darf 150 ms nicht ueberschreiten."),
                ("h", "{+2.5 Neu+}"), ("p", "{+Ein neuer Absatz mit Zusatz.+}"),
                ("h", "3 Drei"), ("p", "Alle Steckverbindungen sind eindeutig beschriftet."),
                ("h", "4 Vier"), ("p", "Die Lieferung umfasst alles Zubehoer.")]
    ordner = tmp_path / "paar"
    ordner.mkdir()
    baue_pdf(ordner / "aenderungen.pdf", elemente, "farbe")
    baue_excel(ordner / "nachher.xlsx", zustaende(elemente, True), "L3", True)
    baue_excel(ordner / "vorher.xlsx", zustaende(elemente, False), "L3", True)
    from openpyxl import load_workbook
    wb = load_workbook(str(ordner / "vorher.xlsx"))
    ws = wb["CDE"]
    ws.row_dimensions[7].height = 40                # unter der Einfuegestelle: wandert mit
    ws.row_dimensions[3].hidden = True              # ueber der Einfuegestelle: bleibt
    ws.merge_cells("D2:D3")
    dv = DataValidation(type="list", formula1='"A,B"')
    dv.add("D2:D3")
    ws.add_data_validation(dv)
    wb.save(str(ordner / "vorher.xlsx"))
    res = messen.messe_paar(ordner, False, tmp_path / "ergebnis.xlsx")
    aufbau(res["auto_gesamt"] >= 2 and len(load_workbook(str(tmp_path / "ergebnis.xlsx"))["CDE"]["A"]) > ws.max_row,
           "Das Tool muss Zeilen einfuegen")
    assert res["struktur_geaendert"] == 0, res["struktur_nach_art"]
    assert res["stille_fehler"] == 0


# ==================================================================================================
# Teil B3: Regel 27 -- ein Tool, das "pruefen"-Aenderungen OHNE Freigabe einarbeitet, wird von der Messung BELOHNT
# ==================================================================================================
_EL_OFFEN = [("h", "1 Eins"), ("p", "Der Nachweis erfolgt durch ein Messprotokoll der Abnahme."),
             ("h", "2 Zwei"), ("p", "Die Reaktionszeit darf 150 ms nicht ueberschreiten."),
             ("h", "3 Drei"),
             ("p", "Die Wartung erfolgt {-jaehrlich-} {+monatlich+} durch den Betreiber und {-den-} {+das+} Personal."),
             ("h", "4 Vier"), ("p", "Alle Steckverbindungen sind eindeutig beschriftet.")]


def _alle_pruefen_ohne_freigabe_anwenden():
    echt = messen.uebernehmen

    def bruch(aend, vorher, out, freigaben=None):
        if Path(out).name == "ergebnis_auto.xlsx":                  # der Automatik-Lauf bekommt heimlich alle Freigaben
            freigaben = {a.id for a in aend if a.status == PRUEFEN}
        return echt(aend, vorher, out, freigaben)

    return mock.patch.object(messen, "uebernehmen", bruch)


def _plane_mit_offener_aenderung():
    echt = messen.plane

    def offen(pdf, xl, *a, **k):
        p = echt(pdf, xl, *a, **k)
        p.aenderungen[-1].status, p.aenderungen[-1].konfidenz = PRUEFEN, "mittel"
        return p

    return mock.patch.object(messen, "plane", offen)


@pytest.mark.redteam
def test_rt137_pruefen_ohne_freigabe_angewendet_wird_gemeldet(tmp_path):
    ordner = tmp_path / "paar"
    ordner.mkdir()
    baue_pdf(ordner / "aenderungen.pdf", _EL_OFFEN, "farbe")
    baue_excel(ordner / "vorher.xlsx", zustaende(_EL_OFFEN, False), "L1", True)
    baue_excel(ordner / "nachher.xlsx", zustaende(_EL_OFFEN, True), "L1", True)
    with _plane_mit_offener_aenderung():
        ehrlich = messen.messe_paar(ordner, False)
        with _alle_pruefen_ohne_freigabe_anwenden():
            bruch = messen.messe_paar(ordner, False)
    aufbau(ehrlich["aenderungen_nach_status"].get(PRUEFEN) == 1, "Erwartet: genau eine offene Aenderung")
    aufbau(bruch["auto_gesamt"] == 1 and bruch["soll_zellen"] == 1,
           "Durch den Verstoss muss genau eine Zelle geaendert sein (sonst gilt der Fall nicht)")
    # Die Zelle ist inhaltlich richtig. Trotzdem: eigener schwerer Befund, stiller Fehler, nicht "ohne Nachfrage erreicht".
    if bruch["regel27_verstoesse"] != 1 or bruch["stille_fehler"] < 1 or bruch["erreicht_auto"] >= 1 or bruch["auto_richtig"] >= 1:
        scheitern(f"Der Verstoss gegen Regel 27 wird nicht als solcher gemeldet: regel27_verstoesse {bruch['regel27_verstoesse']}, "
                  f"stille Fehler {ehrlich['stille_fehler']} -> {bruch['stille_fehler']}, Abdeckung ohne Nachfrage "
                  f"{ehrlich['erreicht_auto']} -> {bruch['erreicht_auto']} von {bruch['soll_zellen']} Soll-Zellen")
    assert "regel27_verstoss" in bruch["stille_fehler_nach_art"]
    bericht = messen.fasse_zusammen([dict(bruch, name="synth_9101")], "entwicklung", "1", False)
    assert next(k["ampel"] for k in bericht["kriterien"] if k["name"].startswith("Regel 27")) == "NICHT ok"
    assert next(k["ampel"] for k in bericht["kriterien"] if k["name"] == "Stille Fehler") == "NICHT ok"
    assert ehrlich["regel27_verstoesse"] == 0 and ehrlich["stille_fehler"] == 0     # ohne Verstoss alles in Ordnung


@pytest.mark.redteam
@pytest.mark.parametrize("seed", [11, 58])
def test_rt138_korpus_pruefen_ohne_freigabe_angewendet(seed):
    ordner = paar(seed, "standard")
    offen = [a for a in auto_lauf(seed, "standard").plan.aenderungen if a.status == PRUEFEN]
    aufbau(bool(offen), "Das Paar hat keine offene Aenderung")
    with _alle_pruefen_ohne_freigabe_anwenden():
        bruch = messen.messe_paar(ordner, False)
    if bruch["stille_fehler"] == 0 or bruch["regel27_verstoesse"] < 1:
        scheitern(f"Seed {seed}: {len(offen)} 'pruefen'-Aenderung(en) ohne Freigabe angewendet -- die Messung zeigt "
                  f"{bruch['stille_fehler']} stille Fehler und {bruch['regel27_verstoesse']} Regel-27-Verstoesse, "
                  f"Treffergenauigkeit {bruch['auto_richtig']} von {bruch['auto_gesamt']}")


# ==================================================================================================
# Teil B4: Ein ABSTURZ im Wiederholungslauf (Idempotenz) macht die Funde des ersten Laufs ungueltig, und Kriterium 2 bleibt "ok"
# ==================================================================================================
def _starte_messung(eltern: Path, name: str) -> dict:
    """Wie `python -m src.cli messen --paar NAME --ordner ELTERN`; liest den Bericht (landet im Temp-Datenordner)."""
    code = messen.main(["--korpus", "entwicklung", "--ordner", str(eltern), "--paar", name])
    aufbau(code == 0, f"messen.main lieferte Code {code}")
    neu = sorted(config.pfad("messung_dir").iterdir())[-1]
    return json.loads((neu / "bericht.json").read_text(encoding="utf-8"))


def _ampel(bericht: dict, nr: int, name_teil: str) -> str:
    return next(k["ampel"] for k in bericht["kriterien"] if k["nr"] == nr and name_teil in k["name"])


def _stuerzt_im_wiederholungslauf(sabotage=None):
    echt = messen.uebernehmen

    def mit(aend, vorher, out, freigaben=None):
        if Path(out).name == "ergebnis_auto2.xlsx":
            raise ValueError("Absturz im Wiederholungslauf")
        r = echt(aend, vorher, out, freigaben)
        if sabotage and Path(out).name == "ergebnis_auto.xlsx":
            from openpyxl import load_workbook
            wb = load_workbook(str(out))
            sabotage(wb)
            wb.save(str(out))
        return r

    return mock.patch.object(messen, "uebernehmen", mit)


@pytest.mark.redteam
def test_rt139_absturz_im_wiederholungslauf_versteckt_stillen_fehler(tmp_path):
    ordner = baue_basis_paar(tmp_path / "eltern" / "synth_9001_basis")
    einzeln = messe_mit_sabotage(ordner, SABOTAGE_ERKANNT["falsches_wort"])
    aufbau(einzeln["stille_fehler"] >= 1, "Ohne Absturz muss die Sabotage ein stiller Fehler sein")
    with _stuerzt_im_wiederholungslauf(SABOTAGE_ERKANNT["falsches_wort"]):
        bericht = _starte_messung(tmp_path / "eltern", "synth_9001_basis")
    g = bericht["gesamt"]
    aufbau(g["paare_abgestuerzt"] == 1, "Das Paar muss als abgestuerzt gelten")
    if g["stille_fehler"] == 0:
        scheitern(f"Der stille Fehler des ersten Laufs ({einzeln['stille_fehler']}) ist weg: Bericht zeigt {g['stille_fehler']} "
                  f"stille Fehler, Paar 'nicht messbar'")
    assert g["stille_fehler"] == einzeln["stille_fehler"] and g["paare_teilweise_gemessen"] == 1
    paar = bericht["paare"][0]
    assert paar["idempotent"] is None and paar["ausgefallene_messteile"] == [{"teil": "idempotenz", "grund": "ValueError"}]
    assert "Absturz im Wiederholungslauf" not in json.dumps(bericht)


@pytest.mark.redteam
def test_rt140_abgestuerztes_paar_laesst_kriterium_stille_fehler_gruen(tmp_path):
    ordner = baue_basis_paar(tmp_path / "eltern" / "synth_9002_basis")
    with _stuerzt_im_wiederholungslauf():
        bericht = _starte_messung(tmp_path / "eltern", "synth_9002_basis")
    aufbau(bericht["gesamt"]["paare_abgestuerzt"] == 1, "Das Paar muss als abgestuerzt gelten")
    ampel2 = _ampel(bericht, 2, "Stille Fehler")
    if ampel2 == "ok":
        scheitern(f"Kriterium 2 'Stille Fehler' steht auf '{ampel2}' (Wert {bericht['gesamt']['stille_fehler']}), obwohl "
                  f"{bericht['gesamt']['paare_abgestuerzt']} von {bericht['gesamt']['paare']} Paaren nicht gemessen werden konnten")


# ==================================================================================================
# Teil B5: Ohne erwartet.json (= bei ALLEN echten Paaren) kann die Messung eine fehlende Luecken-Meldung nie bemerken
# ==================================================================================================
@pytest.mark.redteam
def test_rt141_scan_seite_ohne_erwartet_json_wird_als_plausibilitaet_gemeldet(tmp_path):
    """Ohne erwartet.json ist die Luecken-Pruefung (Kriterium 1) NICHT moeglich -- das wird ausgewiesen (nicht "ok"), und eine
    unabhaengige Plausibilitaetspruefung direkt am PDF (pdfplumber) weist die Scan-Seite als HINWEIS aus, nie als Fehler."""
    ordner = tmp_path / "paar"
    ordner.mkdir()
    d = PdfDokument(variante="farbe")
    for art, markup in BASIS:
        d.ueberschrift(markup) if art == "h" else d.block(markup)
    d.scan_seite(scan_bild([("h", "6 Scan"), ("p", "Auf der Scan-Seite steht {+neuer Text+} und {-alter Text-} nur als Pixel.")]))
    d.speichern(ordner / "aenderungen.pdf")
    baue_excel(ordner / "vorher.xlsx", zustaende(BASIS, False), "L1", True)
    baue_excel(ordner / "nachher.xlsx", zustaende(BASIS, True), "L1", True)
    echt = messen.analysiere_pdf

    def ohne_luecken(pfad):                      # so verhaelt sich der PDF-Leser heute (R2: RT-98); unabhaengig von seiner Reparatur
        a = echt(pfad)
        a.luecken = []
        return a

    with mock.patch.object(messen, "analysiere_pdf", ohne_luecken):
        res = messen.messe_paar(ordner, False)
    aufbau(not (ordner / "erwartet.json").exists() and res["luecken_gemeldet"] == 0, "Testaufbau: keine Luecke erwartet/gemeldet")
    scan = [h for h in res["plausibilitaet_hinweise"] if h["art"] == "keine_textzeichen"]
    if not scan or scan[0]["gemeldet"] or res["plausibilitaet_ungemeldet"] < 1:
        scheitern("Eine Seite ohne Textzeichen (Scan) wird von der Plausibilitaetspruefung nicht als ungemeldet ausgewiesen: "
                  f"{res['plausibilitaet_hinweise']}")
    assert res["luecken_pruefbar"] is False
    assert res["luecken_nicht_gemeldet"] == 0 and res["stille_fehler"] == 0     # nur ein Hinweis, nie ein eigener Fehler
    bericht = messen.fasse_zusammen([dict(res, name="echt_0001")], "entwicklung", "", False)
    k1 = next(k for k in bericht["kriterien"] if k["name"] == "Bilanz stimmig")
    assert k1["ampel"] == "nicht gemessen" and "erwartet.json" in k1["wert"]
    k1b = next(k for k in bericht["kriterien"] if k["name"].startswith("Plausibilit"))
    assert k1b["ampel"] == "Hinweis"


def test_rt141_mit_erwartet_json_ist_die_luecken_pruefung_moeglich(tmp_path):
    ordner = baue_basis_paar(tmp_path / "basis")
    (ordner / "erwartet.json").write_text(json.dumps({"format": 1, "profil": "standard", "luecken": []}), encoding="utf-8")
    res = messen.messe_paar(ordner, False)
    assert res["luecken_pruefbar"] is True
    bericht = messen.fasse_zusammen([dict(res, name="synth_9102")], "entwicklung", "1", False)
    assert next(k["ampel"] for k in bericht["kriterien"] if k["name"] == "Bilanz stimmig") == "ok"


# ==================================================================================================
# Teil B6: "Nicht verwaltete" Spalten: Soll-Zellen dort zaehlen weder bei der Abdeckung noch als Fehler
# ==================================================================================================
@pytest.mark.redteam
def test_rt142_unverwaltete_soll_zellen_tauchen_in_einem_kriterium_auf(tmp_path):
    zusatz = " Dieser lange Zusatz steht nur in der ersten Zeile und macht die Inhaltsspalte zur Spalte mit den meisten Aenderungen."
    text = "Der Inhalt dieser Section steht hier und ist ein ganz normaler Satz ueber die Anlage Nummer {}."
    vorher = [KOPF] + [[str(i), f"Titel {i}", text.format(i), "A"] for i in range(1, 7)]
    nachher = [KOPF] + [[str(i), f"Titel {i}", text.format(i) + (zusatz if i == 1 else ""), "B"] for i in range(1, 7)]
    ergebnis = [KOPF] + [[str(i), f"Titel {i}", text.format(i) + (zusatz if i == 1 else ""), "A"] for i in range(1, 7)]
    res = messe_mit_fertigem_ergebnis(tmp_path, vorher, nachher, ergebnis)
    aufbau(res["soll_zellen"] == 1 and res["erreicht_auto"] == 1,
           f"Testaufbau: erwartet 1 verwaltete (erreichte) Soll-Zelle, gemessen {res['soll_zellen']} / {res['erreicht_auto']}")
    # Die 6 Zellen der Bemerkungsspalte (A -> B) fehlen im Ergebnis: eigene Zaehlung, eigene Fehlerzaehlung, eigenes Kriterium
    assert res["soll_zellen_unverwaltet_sonstige"] == 6 and res["soll_zellen_unverwaltet_sonstige_verfehlt"] == 6
    bericht = messen.fasse_zusammen([dict(res, name="synth_9003")], "entwicklung", "1", False)
    k = next((k for k in bericht["kriterien"] if "verwaltet" in k["name"].lower()), None)
    if k is None or k["ampel"] == "ok":
        scheitern(f"Abdeckung {bericht['gesamt']['abdeckung_ohne_nachfrage']}, aber {res['soll_zellen_unverwaltet_sonstige']} geaenderte "
                  "Soll-Zellen (nicht verwaltete Spalte) fehlen im Ergebnis; kein Kriterium nennt sie")
    assert k["ampel"] == "NICHT ok" and "6" in k["wert"]
    assert bericht["gesamt"]["soll_zellen_alle"] == 7


def test_rt142_formelzellen_und_nummerierung_fallen_nicht_heraus(tmp_path):
    """Formelzellen und neu nummerierte Spalten werden getrennt gezaehlt (mit eigener Fehlerzaehlung); harmlose Fuellung neuer
    Zeilen bleibt 'unverwaltet' und ist aufgeschluesselt; Formeln zaehlen nicht als Soll-Woerter der Lese-Luecke."""
    zusatz = " Dieser lange Zusatz steht nur in der ersten Zeile und macht die Inhaltsspalte zur Spalte mit den meisten Aenderungen."
    text = "Der Inhalt dieser Section steht hier und ist ein ganz normaler Satz ueber die Anlage Nummer {}."
    neu_text = "Ein ganz neuer Absatz ohne Bezug zu den anderen hier."
    kopf = KOPF + ["Summe", "Lfd"]
    vorher = [kopf] + [[str(i), f"Titel {i}", text.format(i), "A", f"=LEN(C{i + 1})", str(i)] for i in range(1, 7)]
    nachher = [kopf] + [[str(i), f"Titel {i}", text.format(i) + (zusatz if i == 1 else ""), "A", f"=LEN(C{i + 2})", str(i + 1)]
                        for i in range(1, 7)] + [["7", "Neu", neu_text, "B", None, None]]
    ergebnis = [kopf] + [[str(i), f"Titel {i}", text.format(i) + (zusatz if i == 1 else ""), "A", f"=LEN(C{i + 1})", str(i)]
                         for i in range(1, 7)] + [["7", "Neu", neu_text, "B", None, None]]
    res = messe_mit_fertigem_ergebnis(tmp_path, vorher, nachher, ergebnis)
    assert res["soll_zellen_formel"] == 6 and res["soll_zellen_formel_verfehlt"] == 6       # verschobene Formeln: verfehlt
    assert res["soll_zellen_unverwaltet_sonstige"] == 6 and res["soll_zellen_unverwaltet_sonstige_verfehlt"] == 6   # Lfd neu nummeriert
    assert res["soll_zellen_unverwaltet"] == 1                                                 # "B" der neuen Zeile
    assert res["soll_zellen_unverwaltet_aufschluesselung"] == {"neue_zeile": 1, "entfernte_zeile": 0}
    nur_haupt = messen.fasse_zusammen([dict(res, name="synth_9104")], "entwicklung", "1", False)["gesamt"]
    assert nur_haupt["soll_zellen_alle"] == res["soll_zellen"] + 6 + 6 + 1
