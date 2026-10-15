"""Selbsttest des Messwerkzeugs (src/messen.py): Es muss Fehler finden, die wir absichtlich einbauen.

Nur erfundene Daten (Regel 16). Alles in tmp_path; conftest.py setzt CDE_DATA_ROOT.
Die Paare werden hier selbst gebaut (pdf_writer + openpyxl), nicht mit dem Generator.
"""
import hashlib
import json
import re
import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

from src import cli, config, messen, runs
from src.models import PRUEFEN
from tests.fixtures.pdf_writer import baue_pdf
from tests.fixtures.szenario import baue_excel, zustaende

# Markante erfundene Woerter: duerfen nie in Bericht oder Konsole auftauchen.
MARKANT = ["quokka", "flamingo", "pinguin", "laub", "gras", "links", "rechts", "teich", "gefieder", "schatten",
           "entfällt", "zweiter absatz", "letzter absatz", "komplett", "ersatz"]

SAUBER = [
    ("h", "1 Eins"),
    ("p", "Der Quokka frisst {-Gras-} {+Laub+} am Morgen. Dann ruht er lange im Schatten."),
    ("p", "Zweiter Absatz bleibt unverändert hier stehen."),
    ("h", "2 Zwei"),
    ("p", "Ein Flamingo steht {-links-} {+rechts+} im Teich und putzt sein rosa Gefieder."),
    ("p", "Der Teich ist sehr tief."),
    ("h", "3 Drei"),
    ("p", "{-Dieser Absatz entfällt komplett im Dokument ohne Ersatz.-}"),
    ("p", "Letzter Absatz bleibt erhalten und ruhig."),
]
MIT_NEUER_SECTION = SAUBER + [("h", "{+4 Vier+}"), ("p", "{+Nur ein neuer Absatz mit einem Pinguin darin.+}")]


def baue_paar(eltern: Path, name: str, elemente=SAUBER, layout: str = "L1") -> Path:
    p = eltern / name
    p.mkdir(parents=True)
    baue_pdf(p / "aenderungen.pdf", elemente, "farbe")
    baue_excel(p / "vorher.xlsx", zustaende(elemente, False), layout)
    baue_excel(p / "nachher.xlsx", zustaende(elemente, True), layout)
    return p


def korpus_dir(korpus: str = "entwicklung") -> Path:
    d = config.pfad("paare_dir") / korpus
    d.mkdir(parents=True, exist_ok=True)
    return d


def sha(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def zellen_aendern(pfad: Path, aenderungen: dict[str, str]) -> None:
    wb = load_workbook(str(pfad))
    ws = wb["CDE"]
    for adresse, text in aenderungen.items():
        ws[adresse].value = text
    wb.save(str(pfad))


def sabotiere_ergebnis(monkeypatch, wirkung):
    """Der ERSTE uebernehmen()-Aufruf (Automatik) wird nachtraeglich veraendert: wirkung(ziel_pfad, aenderungen)."""
    original = messen.uebernehmen
    zaehler = {"n": 0}

    def falsch(aenderungen, quelle, ziel, freigaben=None, cfg=None):
        res = original(aenderungen, quelle, ziel, freigaben, cfg)
        zaehler["n"] += 1
        if zaehler["n"] == 1:
            wirkung(Path(ziel), aenderungen, Path(quelle))
        return res
    monkeypatch.setattr(messen, "uebernehmen", falsch)


def arten(res) -> set[str]:
    return {f["art"] for f in res["fehler"]}


# ---------------------------------------------------------------------------------------
# (1) Sauberes Paar
# ---------------------------------------------------------------------------------------
def test_sauberes_paar_hat_keine_stillen_fehler(tmp_path):
    p = baue_paar(tmp_path, "synth_0001_standard")
    res = messen.messe_paar(p)
    assert res["stille_fehler"] == 0 and res["fehler"] == []
    assert res["soll_zellen"] == 3                       # drei Zellen aendern sich laut nachher
    assert res["erreicht_auto"] == 3
    assert res["auto_gesamt"] == 3 and res["auto_richtig"] == 3
    assert res["kollateral"] == 0
    assert res["bilanz_stimmt"] and res["vollstaendig"]
    assert res["idempotent"] is True
    assert res["original_unveraendert"] is True
    assert res["lese_luecke_woerter"] == 0
    assert res["pruefen_zellen"] == 0 and res["nicht_uebernommen_zellen"] == 0
    assert res["vorschlaege_gesamt"] is None             # ohne --mit-vorschlaegen nicht gemessen


@pytest.mark.parametrize("layout", ["L1", "L2", "L3"])
def test_sauber_in_allen_layouts(tmp_path, layout):
    res = messen.messe_paar(baue_paar(tmp_path, "synth_0002", layout=layout))
    assert res["stille_fehler"] == 0 and res["fehler"] == []
    assert res["erreicht_auto"] == res["soll_zellen"] > 0


def test_neue_section_wird_als_neue_zeile_gemessen(tmp_path):
    res = messen.messe_paar(baue_paar(tmp_path, "synth_0003", MIT_NEUER_SECTION))
    assert res["stille_fehler"] == 0 and res["fehler"] == []
    assert res["erreicht_auto"] == res["soll_zellen"] and res["soll_zellen"] > 3


def test_original_wird_nie_veraendert(tmp_path):
    p = baue_paar(tmp_path, "synth_0004")
    vorher = {n: sha(p / n) for n in messen.DATEIEN}
    messen.messe_paar(p)
    assert {n: sha(p / n) for n in messen.DATEIEN} == vorher


def test_original_veraenderung_wird_gemeldet(tmp_path, monkeypatch):
    p = baue_paar(tmp_path, "synth_0005")
    original = messen.uebernehmen

    def boese(aenderungen, quelle, ziel, freigaben=None, cfg=None):
        Path(quelle).write_bytes(Path(quelle).read_bytes() + b"x")      # veraendert das Original
        return original(aenderungen, quelle, ziel, freigaben, cfg)
    monkeypatch.setattr(messen, "uebernehmen", boese)
    res = messen.messe_paar(p)
    assert res["original_unveraendert"] is False


# ---------------------------------------------------------------------------------------
# (2) Sabotage: das Messwerkzeug muss die eingebauten Fehler finden
# ---------------------------------------------------------------------------------------
def test_nichts_uebernommen_wird_als_luecke_erkannt(tmp_path, monkeypatch):
    p = baue_paar(tmp_path, "synth_1000")
    sabotiere_ergebnis(monkeypatch, lambda ziel, aend, quelle: shutil.copy2(quelle, ziel))   # E = vorher
    res = messen.messe_paar(p)
    assert res["erreicht_auto"] == 0 and res["soll_zellen"] == 3
    assert res["stille_fehler"] >= 3                                    # "uebernommen" gemeldet, Zelle unveraendert
    assert "keine_wirkung" in res["stille_fehler_nach_art"]
    assert "plan_falsch" in arten(res)                                  # Ziel stimmt, Ergebnis nicht


def test_manipulierte_zielzelle_ist_stiller_fehler(tmp_path, monkeypatch):
    p = baue_paar(tmp_path, "synth_1001")
    sabotiere_ergebnis(monkeypatch, lambda ziel, aend, q: zellen_aendern(ziel, {"C2": "Etwas ganz anderes steht hier."}))
    res = messen.messe_paar(p)
    assert res["stille_fehler"] == 1 and res["erreicht_auto"] == 2
    assert res["stille_fehler_nach_art"].get("uebernommen_aber_falsch") == 1
    assert res["auto_richtig"] == 2 and res["auto_gesamt"] == 3
    assert [f["art"] for f in res["fehler"]] == ["plan_falsch"]
    assert res["fehler"][0]["aenderung"].startswith("A") and res["fehler"][0]["stumm"] is True
    assert (res["fehler"][0]["blatt"], res["fehler"][0]["zeile"], res["fehler"][0]["spalte"]) == (2, 2, 3)   # Blatt 2 = "CDE"


def test_veraenderte_fremde_zelle_ist_kollateralschaden(tmp_path, monkeypatch):
    p = baue_paar(tmp_path, "synth_1002")
    sabotiere_ergebnis(monkeypatch, lambda ziel, aend, q: zellen_aendern(ziel, {"B3": "Fremder Titel"}))
    res = messen.messe_paar(p)
    assert res["kollateral"] == 1 and res["stille_fehler"] == 1
    assert res["erreicht_auto"] == 3                                    # die eigentlichen Aenderungen sind richtig
    assert res["auto_richtig"] == 3 and res["auto_gesamt"] == 4


def test_nachkontrolle_ok_trotz_falschem_inhalt(tmp_path, monkeypatch):
    p = baue_paar(tmp_path, "synth_1003")

    def wirkung(ziel, aend, quelle):
        zellen_aendern(ziel, {"C2": "Falscher Inhalt, aber das Tool sagt ok."})
        for a in aend:
            a.nachkontrolle = "ok"
    sabotiere_ergebnis(monkeypatch, wirkung)
    res = messen.messe_paar(p)
    assert res["stille_fehler_nach_art"].get("nachkontrolle_ok_aber_falsch") == 1
    assert res["stille_fehler"] == 1


def test_zweiter_lauf_der_etwas_aendert_wird_erkannt(tmp_path, monkeypatch):
    p = baue_paar(tmp_path, "synth_1010")
    original = messen.uebernehmen
    zaehler = {"n": 0}

    def unsauber(aenderungen, quelle, ziel, freigaben=None, cfg=None):
        res = original(aenderungen, quelle, ziel, freigaben, cfg)
        zaehler["n"] += 1
        if zaehler["n"] == 2:                      # zweiter Lauf (Idempotenz-Pruefung) aendert wieder etwas
            zellen_aendern(Path(ziel), {"C3": "Noch eine Aenderung beim zweiten Lauf."})
        return res
    monkeypatch.setattr(messen, "uebernehmen", unsauber)
    res = messen.messe_paar(p)
    assert res["idempotent"] is False and res["stille_fehler"] == 0


def test_bilanz_stimmt_false_ist_bilanz_kaputt(tmp_path, monkeypatch):
    p = baue_paar(tmp_path, "synth_1004")
    original = messen.bilanz
    monkeypatch.setattr(messen, "bilanz", lambda *a, **k: {**original(*a, **k), "stimmt": False})
    res = messen.messe_paar(p)
    assert res["bilanz_stimmt"] is False
    assert "bilanz_kaputt" in arten(res)
    assert res["stille_fehler_nach_art"].get("bilanz") == 1 and res["stille_fehler"] >= 1


def test_falsche_zuordnung_gleiche_spalte(tmp_path, monkeypatch):
    """Zwei Zellen in derselben Spalte vertauscht -> richtige Spalte, falsche Zeile."""
    p = baue_paar(tmp_path, "synth_1005")

    def wirkung(ziel, aend, quelle):
        ws = load_workbook(str(ziel))["CDE"]
        a, b = ws["C2"].value, ws["C3"].value
        zellen_aendern(ziel, {"C2": b, "C3": a})
    sabotiere_ergebnis(monkeypatch, wirkung)
    res = messen.messe_paar(p)
    assert res["erreicht_auto"] == 1
    assert arten(res) == {"zuordnung_falsch"}
    assert res["stille_fehler"] >= 2


def test_falsche_spalte_ist_layout_falsch(tmp_path, monkeypatch):
    p = baue_paar(tmp_path, "synth_1006")

    def wirkung(ziel, aend, quelle):
        ws = load_workbook(str(ziel))["CDE"]
        neu = ws["C2"].value
        zellen_aendern(ziel, {"D2": neu, "C2": load_workbook(str(quelle))["CDE"]["C2"].value})
    sabotiere_ergebnis(monkeypatch, wirkung)
    res = messen.messe_paar(p)
    assert "layout_falsch" in arten(res)


def test_falscher_plan_ziel_stimmt_ergebnis_nicht(tmp_path, monkeypatch):
    """Ein Wort doppelt eingetragen: Ziel und Woerter stimmen, das Ergebnis nicht -> plan_falsch."""
    p = baue_paar(tmp_path, "synth_1007")

    def wirkung(ziel, aend, quelle):
        ws = load_workbook(str(ziel))["CDE"]
        zellen_aendern(ziel, {"C2": ws["C2"].value.replace("Laub", "Laub Laub")})
    sabotiere_ergebnis(monkeypatch, wirkung)
    res = messen.messe_paar(p)
    assert arten(res) == {"plan_falsch"} and res["stille_fehler"] == 1


def test_pdf_falsch_gelesen(tmp_path, monkeypatch):
    p = baue_paar(tmp_path, "synth_1008")
    original = messen.analysiere_pdf

    def blind(pfad):
        pa = original(pfad)
        for b in pa.blocks:                                  # ein neues Wort "verschluckt" der Leser
            b.tokens = [t for t in b.tokens if not (t.status == "add" and t.new == "Laub")]
        return pa
    monkeypatch.setattr(messen, "analysiere_pdf", blind)
    res = messen.messe_paar(p)
    assert "pdf_falsch_gelesen" in arten(res)
    assert res["lese_luecke_woerter"] >= 1 and res["erreicht_auto"] < res["soll_zellen"]


def test_abgestuerzte_paare_zaehlen_als_unklar_ohne_text(tmp_path, monkeypatch):
    korpus = korpus_dir()
    baue_paar(korpus, "synth_2000")
    baue_paar(korpus, "synth_2001")
    original = messen.plane
    zaehler = {"n": 0}

    def kaputt(pdf, xl, cfg=None):
        zaehler["n"] += 1
        if zaehler["n"] == 1:
            raise RuntimeError("GEHEIMER-DOKUMENTTEXT-quokka")
        return original(pdf, xl, cfg)
    monkeypatch.setattr(messen, "plane", kaputt)
    assert messen.main(["--korpus", "entwicklung"]) == 0
    bericht = json.loads(next(config.pfad("messung_dir").glob("001_entwicklung/bericht.json")).read_text(encoding="utf-8"))
    assert bericht["fehlerarten"]["unklar"] == 1
    assert bericht["gesamt"]["paare"] == 2 and bericht["gesamt"]["paare_abgestuerzt"] == 1
    assert bericht["paare"][0]["fehler"][0]["grund"] == "RuntimeError"
    text = json.dumps(bericht)
    assert "GEHEIMER" not in text and "quokka" not in text.lower()


# ---------------------------------------------------------------------------------------
# Vorschlaege (Status "pruefen")
# ---------------------------------------------------------------------------------------
def test_vorschlaege_getrennt_ausgewiesen(tmp_path, monkeypatch):
    p = baue_paar(tmp_path, "synth_3000")
    original = messen.plane

    def mit_vorschlag(pdf, xl, cfg=None):
        plan = original(pdf, xl, cfg)
        ae = next(a for a in plan.aenderungen if a.ziel.endswith("C2"))
        ae.status, ae.konfidenz = PRUEFEN, "mittel"
        return plan
    monkeypatch.setattr(messen, "plane", mit_vorschlag)
    res = messen.messe_paar(p, mit_vorschlaegen=True)
    assert res["pruefen_zellen"] == 1 and res["erreicht_auto"] == 2
    assert res["fehler"] == [] and res["stille_fehler"] == 0              # "pruefen" ist kein Fehler
    assert res["auto_gesamt"] == 2 and res["auto_richtig"] == 2
    assert res["vorschlaege_gesamt"] == 1 and res["vorschlaege_richtig"] == 1


# ---------------------------------------------------------------------------------------
# erwartet.json (optional)
# ---------------------------------------------------------------------------------------
def test_erwartet_json_stoerer_und_luecken(tmp_path):
    korpus = korpus_dir()
    p = baue_paar(korpus, "synth_4000")
    (p / "erwartet.json").write_text(json.dumps({
        "format": 1, "profil": "standard", "stoerer": ["fussnote"], "woerter_markiert": 99,
        "luecken": [{"art": "scan_seite", "seite": 2}]}), encoding="utf-8")
    baue_paar(korpus, "synth_4001")                               # ohne erwartet.json: muss ebenfalls laufen
    assert messen.main(["--korpus", "entwicklung"]) == 0
    b = json.loads(next(config.pfad("messung_dir").glob("001_entwicklung/bericht.json")).read_text(encoding="utf-8"))
    assert b["je_stoerer"]["fussnote"]["paare"] == 1
    erst = b["paare"][0]
    assert erst["luecken_nicht_gemeldet"] == 1                    # erwartete Luecke nicht gemeldet
    assert erst["stille_fehler_nach_art"].get("luecke_nicht_gemeldet") == 1
    assert erst["lese_luecke_erwartet"] > 0
    assert b["paare"][1]["stille_fehler"] == 0


# ---------------------------------------------------------------------------------------
# (3) Keine Inhalte in Bericht und Konsole
# ---------------------------------------------------------------------------------------
def test_bericht_und_konsole_enthalten_keine_inhalte(tmp_path, monkeypatch, capsys):
    korpus = korpus_dir()
    baue_paar(korpus, "synth_5000", MIT_NEUER_SECTION)
    baue_paar(korpus, "synth_5001")
    # ein Paar mit Fehlern, damit auch die Fehlerpfade geprueft werden
    sabotiere_ergebnis(monkeypatch, lambda ziel, aend, q: zellen_aendern(ziel, {"C2": "Quokka Pinguin Flamingo anders"}))
    assert messen.main(["--korpus", "entwicklung", "--mit-vorschlaegen"]) == 0
    konsole = capsys.readouterr().out.lower()
    ordner = next(config.pfad("messung_dir").glob("001_entwicklung"))
    texte = [konsole, (ordner / "bericht.json").read_text(encoding="utf-8").lower(),
             (ordner / "bericht.md").read_text(encoding="utf-8").lower()]
    for text in texte:
        for wort in MARKANT:
            assert wort not in text, wort
    assert "stille" in konsole                                   # Hinweis bei stillen Fehlern


def test_bericht_nur_erlaubte_felder_in_fehlern(tmp_path):
    korpus = korpus_dir()
    baue_paar(korpus, "synth_5002")
    messen.main(["--korpus", "entwicklung"])
    b = json.loads(next(config.pfad("messung_dir").glob("001_entwicklung/bericht.json")).read_text(encoding="utf-8"))
    for p in b["paare"]:
        for f in p["fehler"]:
            assert set(f) <= {"blatt", "zeile", "spalte", "aenderung", "art", "stumm", "grund"}
            assert f["art"] in messen.FEHLERARTEN


# ---------------------------------------------------------------------------------------
# (4) Append-only
# ---------------------------------------------------------------------------------------
def test_bericht_ist_append_only(tmp_path):
    baue_paar(korpus_dir(), "synth_6000")
    assert messen.main(["--korpus", "entwicklung"]) == 0
    basis = config.pfad("messung_dir")
    erste = (basis / "001_entwicklung" / "bericht.json").read_bytes()
    assert messen.main(["--korpus", "entwicklung"]) == 0
    assert (basis / "002_entwicklung" / "bericht.json").exists() and (basis / "002_entwicklung" / "bericht.md").exists()
    assert (basis / "001_entwicklung" / "bericht.json").read_bytes() == erste


def test_nummer_zaehlt_ueber_korpora_und_weicht_belegten_aus(tmp_path):
    baue_paar(korpus_dir("test"), "synth_6001")
    basis = config.pfad("messung_dir")
    (basis / "007_entwicklung").mkdir(parents=True)
    assert messen.main(["--korpus", "test"]) == 0
    assert (basis / "008_test" / "bericht.md").exists()


def test_schreibe_bericht_ueberschreibt_nie(tmp_path):
    basis = config.pfad("messung_dir")
    (basis / "001_entwicklung").mkdir(parents=True)
    (basis / "001_entwicklung" / "bericht.json").write_text("alt", encoding="utf-8")
    ziel = messen.schreibe_bericht(messen.fasse_zusammen([], "entwicklung", "", False), "entwicklung")
    assert ziel.name == "002_entwicklung"
    assert (basis / "001_entwicklung" / "bericht.json").read_text(encoding="utf-8") == "alt"


# ---------------------------------------------------------------------------------------
# (5) Gleicher Weg wie ueber runs
# ---------------------------------------------------------------------------------------
@pytest.mark.parametrize("elemente", [SAUBER, MIT_NEUER_SECTION])
def test_weg_ueber_runs_ist_identisch(tmp_path, elemente):
    p = baue_paar(tmp_path / "paare", "synth_7000", elemente)
    kopie = tmp_path / "messen_ergebnis.xlsx"
    res = messen.messe_paar(p, ergebnis_kopie=kopie)

    run_id = runs.neuer_lauf(p / "aenderungen.pdf", p / "vorher.xlsx")
    daten = runs.analysiere(run_id)
    ergebnis = runs.uebernehmen_lauf(run_id, [])
    runs_excel = runs.lauf_dir(run_id) / ergebnis["ordner"] / ergebnis["excel"]

    assert messen._lese_excel(kopie) == messen._lese_excel(runs_excel)
    assert res["bilanz"] == {k: v for k, v in ergebnis["bilanz"].items() if isinstance(v, (int, bool))}
    assert res["aenderungen"] == daten["bilanz"]["aenderungen"]
    assert [res["aenderungen_nach_status"][s] for s in ("uebernommen", "pruefen", "nicht_uebernommen")] == \
        [ergebnis["bilanz"]["uebernommen"], ergebnis["bilanz"]["pruefen"], ergebnis["bilanz"]["nicht_uebernommen"]]


# ---------------------------------------------------------------------------------------
# Robustheit / Eingabe
# ---------------------------------------------------------------------------------------
def test_kaputte_eingabe_exit_2_ohne_rohtext(tmp_path, capsys):
    assert messen.main(["--korpus", "entwicklung"]) == 2          # Ordner fehlt
    korpus = korpus_dir()
    assert messen.main(["--korpus", "entwicklung"]) == 2          # Ordner leer
    out = capsys.readouterr().out
    assert "Traceback" not in out and "Error" not in out
    baue_paar(korpus, "synth_8000")
    assert messen.main(["--korpus", "entwicklung", "--paar", "gibt_es_nicht"]) == 2
    assert messen.main(["--korpus", "entwicklung", "--paar", "../synth_8000"]) == 2
    assert messen.main(["--ordner", str(tmp_path / "nirgends")]) == 2
    assert messen.main(["--unbekannt"]) == 2
    assert not config.pfad("messung_dir").exists() or not list(config.pfad("messung_dir").iterdir())


def test_kaputtes_paar_bricht_messung_nicht_ab(tmp_path, capsys):
    korpus = korpus_dir()
    baue_paar(korpus, "synth_8100")
    (korpus / "synth_8101").mkdir()                                # leerer Ordner: Dateien fehlen
    kaputt = baue_paar(korpus, "synth_8102")
    (kaputt / "vorher.xlsx").write_bytes(b"PK das ist keine Excel")
    kaputt2 = baue_paar(korpus, "synth_8103")
    (kaputt2 / "aenderungen.pdf").write_bytes(b"%PDF-1.4 kein echtes PDF")
    assert messen.main(["--korpus", "entwicklung"]) == 0
    b = json.loads(next(config.pfad("messung_dir").glob("001_entwicklung/bericht.json")).read_text(encoding="utf-8"))
    assert b["gesamt"]["paare"] == 4 and b["gesamt"]["paare_abgestuerzt"] == 3
    assert b["fehlerarten"]["unklar"] == 3
    grund = {p["name"]: p["fehler"][0].get("grund") for p in b["paare"] if p["fehler"]}
    assert grund["synth_8101"] == "FileNotFoundError"
    assert all(re.fullmatch(r"\w+", g) for g in grund.values())      # nur Klassennamen
    assert b["paare"][0]["fehler"] == [] and b["paare"][0]["soll_zellen"] == 3


def test_paar_ordner_und_korpus_hinweise(tmp_path, capsys):
    eigener = tmp_path / "irgendwo"
    baue_paar(eigener, "echt_0001")
    baue_paar(eigener, "synth_9")
    (eigener / "KORPUS.json").write_text(json.dumps({"generator_version": "7"}), encoding="utf-8")
    assert messen.main(["--korpus", "test", "--ordner", str(eigener), "--paar", "echt_0001"]) == 0
    out = capsys.readouterr().out
    assert "Regel 11" in out
    b = json.loads(next(config.pfad("messung_dir").glob("001_test/bericht.json")).read_text(encoding="utf-8"))
    assert [p["name"] for p in b["paare"]] == ["echt_0001"] and b["generator_version"] == "7"
    assert b["korpus"] == "test" and b["schema_version"] == 2 and b["format"] == 1
    assert {"paare", "soll_zellen", "bilanz_stimmig", "vollstaendig", "stille_fehler", "auto_richtig", "auto_gesamt",
            "abdeckung_ohne_nachfrage", "lese_luecke_woerter", "idempotent", "kollateral"} <= set(b["gesamt"])
    assert set(b["fehlerarten"]) == set(messen.FEHLERARTEN)


def test_ampel_und_kriterien_in_der_konsole(tmp_path, capsys):
    baue_paar(korpus_dir(), "synth_9100")
    assert cli.main(["messen", "--korpus", "entwicklung"]) == 0     # ueber die Weiche in cli.main
    out = capsys.readouterr().out
    assert "ok" in out and "Stille Fehler" in out and "Abdeckung ohne Nachfrage" in out
    assert "--mit-vorschlaegen" in out                              # Hinweis: Vorschlaege nicht gemessen
    assert "NICHT ok" not in out


# ---------------------------------------------------------------------------------------
# Ausrichtung (Soll/Ist) ohne Excel-Dateien
# ---------------------------------------------------------------------------------------
def _zeilen(*zeilen):
    return messen._zeilen([list(z) for z in zeilen])


def test_diff_zelle_zeile_neu_zeile_weg():
    v = _zeilen(("1", "Eins", "alpha beta"), ("2", "Zwei", "gamma delta epsilon"), ("3", "Drei", "zeta eta theta"))
    n = _zeilen(("1", "Eins", "alpha beta gamma"), ("2", "Zwei", "gamma delta epsilon"), ("2.5", "Neu", "iota kappa"),
                ("3", "Drei", "zeta eta theta"))
    d = messen._diff(v, n)
    assert d.zellen == {(1, 3): "alpha beta gamma"} and d.weg == set()
    assert d.neue == [(2, ("2.5", "Neu", "iota kappa"), ("2.5", "Neu", "iota kappa"))]
    d2 = messen._diff(v, _zeilen(("1", "Eins", "alpha beta"), ("3", "Drei", "zeta eta theta")))
    assert d2.weg == {2} and set(d2.zellen) == {(2, 1), (2, 2), (2, 3)} and d2.neue == []


def test_diff_leere_zeilen_zaehlen_nicht_und_geleerte_zeile_wird_gepaart():
    v = _zeilen(("1", "Eins", "alpha beta", "Rest"), ("", "", ""), ("2", "Zwei", "gamma delta", "Rest"))
    n = _zeilen(("1", "Eins", "alpha beta", "Rest"), ("", "", "", "Rest"))
    d = messen._diff(v, n)
    assert set(d.zellen) == {(3, 1), (3, 2), (3, 3)} and d.weg == set()          # Zeile 3 geleert, Rest bleibt
    assert d.zellen[(3, 3)] == ""


def test_zellvergleich_ignoriert_nur_leerraum():
    assert messen._k("  a  b\n c ") == "a b c"
    assert messen._k("A") != messen._k("a")


# ---------------------------------------------------------------------------------------
# (6) Bestehende CLI-Befehle bleiben unveraendert
# ---------------------------------------------------------------------------------------
def test_cli_analyse_und_lauf_unveraendert(tmp_path, capsys):
    p = baue_paar(tmp_path, "paar")
    excel = tmp_path / "meine_cde.xlsx"
    shutil.copy2(p / "vorher.xlsx", excel)
    vorher = sha(excel)

    assert cli.main(["analyse", "--pdf", str(p / "aenderungen.pdf"), "--excel", str(excel)]) == 0
    out = capsys.readouterr().out
    assert "Lauf:" in out and "Änderungen: 3" in out and "Bilanz stimmt: ja" in out
    lauf1 = runs.liste()
    assert len(lauf1) == 1 and lauf1[0]["ergebnisse"] == 0

    assert cli.main(["lauf", "--pdf", str(p / "aenderungen.pdf"), "--excel", str(excel)]) == 0
    out = capsys.readouterr().out
    assert "Ergebnis:" in out and "übernommen: 3" in out
    neu = [l for l in runs.liste() if l["run_id"] != lauf1[0]["run_id"]][0]
    assert neu["ergebnisse"] == 1
    ergebnis = next((runs.lauf_dir(neu["run_id"]) / "ergebnis_001").glob("*_aktualisiert.xlsx"))
    ws = load_workbook(str(ergebnis))["CDE"]
    assert "Laub" in ws["C2"].value and "Gras" not in ws["C2"].value
    assert sha(excel) == vorher                                    # Original unberuehrt

    with pytest.raises(SystemExit):
        cli.main(["unbekannt"])
    assert cli.main(["analyse", "--pdf", str(p / "aenderungen.pdf"), "--excel", str(excel), "--alle-vorschlaege"]) == 0

