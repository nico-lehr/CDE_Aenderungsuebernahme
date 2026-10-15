"""Neue Kommandozeilen-Befehle: `profil vorschau|bestaetigen|liste|analyse` und `farben` (src/cli.py, src/layouts.py).

Nur erfundene Daten (Regel 16). Die Ausgabe darf nie Zellinhalte oder Woerter aus dem PDF enthalten.
"""
import re

import pytest
import yaml

from src import cli, layouts, runs
from tests.fixtures.pdf_writer import baue_pdf
from tests.fixtures.szenario import baue_excel, zustaende

ELEMENTE = [
    ("h", "1 Tapir"),
    ("p", "Der Quokka frisst {-Gras-} {+Laub+} am Morgen. Dann ruht er lange im Schatten."),
    ("h", "2 Okapi"),
    ("p", "Ein Flamingo steht {-links-} {+rechts+} im Teich und putzt {@gn:rosa@} sein Gefieder."),
]
PDF_WOERTER = ["tapir", "quokka", "frisst", "gras", "laub", "morgen", "schatten", "okapi", "flamingo", "links",
               "rechts", "teich", "putzt", "rosa", "gefieder"]

PROFIL = {"name": "cli_typ", "aktiv": True, "blaetter": [
    {"blatt": "Deckblatt", "ignorieren": True},
    {"blatt": "CDE", "section_spalte": "A", "titel_spalte": "B", "inhalt_spalte": "C",
     "modell": "zeile_pro_section", "kopfzeile": 1}]}


@pytest.fixture
def dateien(tmp_path, monkeypatch):
    pdf, excel = tmp_path / "a.pdf", tmp_path / "vorher.xlsx"
    baue_pdf(pdf, ELEMENTE)
    baue_excel(excel, zustaende(ELEMENTE, False), "L1")
    yaml_datei = tmp_path / "layouts.yaml"
    yaml_datei.write_text(yaml.safe_dump({"profile": [PROFIL]}, allow_unicode=True), encoding="utf-8")
    monkeypatch.setattr(layouts, "profil_dateien", lambda: [yaml_datei, tmp_path / "layouts.local.yaml"])
    return pdf, excel


def lauf(capsys, *args):
    rc = cli.main(list(args))
    return rc, capsys.readouterr().out


# ---------------------------------------------------------------- profil
def test_profil_ablauf_vorschau_bestaetigen_analyse(capsys, dateien):
    pdf, excel = dateien
    rc, out = lauf(capsys, "profil", "liste")
    assert rc == 0 and "cli_typ" in out and "nicht bestätigt" in out

    rc, out = lauf(capsys, "profil", "analyse", "--pdf", str(pdf), "--excel", str(excel), "--profil", "cli_typ")
    assert rc == 2 and "noch nicht bestätigt" in out                       # nie direkt live

    rc, out = lauf(capsys, "profil", "vorschau", "--excel", str(excel), "--profil", "cli_typ")
    assert rc == 0 and "Section-Spalte A" in out and "Modell zeile_pro_section" in out
    vid = re.search(r"Vorschau-ID: (\S+)", out).group(1)
    assert f"profil bestaetigen --vorschau {vid}" in out

    rc, out = lauf(capsys, "profil", "bestaetigen", "--vorschau", vid)
    assert rc == 0 and "jetzt bestätigt" in out
    rc, out = lauf(capsys, "profil", "bestaetigen", "--vorschau", vid)
    assert rc == 0 and "schon bestätigt" in out

    rc, out = lauf(capsys, "profil", "liste")
    assert "| bestätigt |" in out

    rc, out = lauf(capsys, "profil", "analyse", "--pdf", str(pdf), "--excel", str(excel), "--profil", "cli_typ")
    assert rc == 0 and "Profil: cli_typ" in out and "Änderungen:" in out
    assert len(runs.liste()) == 1                                          # nur der erfolgreiche Lauf wurde angelegt


def test_vorschau_ohne_profil_speichert_nichts(capsys, dateien, tmp_path):
    _, excel = dateien
    rc, out = lauf(capsys, "profil", "vorschau", "--excel", str(excel))
    assert rc == 0 and "automatisch erkannt" in out and "Es wurde nichts gespeichert" in out
    assert not list(layouts.profile_dir().glob("*.json"))


def test_vorschau_ausgabe_ohne_zellinhalte(capsys, dateien):
    _, excel = dateien
    rc, out = lauf(capsys, "profil", "vorschau", "--excel", str(excel), "--profil", "cli_typ")
    low = out.casefold()
    for wort in PDF_WOERTER + ["abschnitt", "status offen", "beispielprojekt", "bemerkung", "inhalt:"]:
        assert wort not in low, wort


def test_fehler_in_einfachem_deutsch(capsys, dateien, tmp_path):
    _, excel = dateien
    rc, out = lauf(capsys, "profil", "vorschau", "--excel", str(tmp_path / "gibt_es_nicht.xlsx"))
    assert rc == 2 and "nicht gefunden" in out
    rc, out = lauf(capsys, "profil", "vorschau", "--excel", str(excel), "--profil", "gibt_es_nicht")
    assert rc == 2 and "kein Profil" in out and "Traceback" not in out
    rc, out = lauf(capsys, "profil", "bestaetigen", "--vorschau", "quatsch")
    assert rc == 2 and "ungültig" in out
    rc, out = lauf(capsys, "profil")
    assert rc == 2 and "Verwendung" in out


def test_liste_meldet_fehler_in_der_yaml(capsys, dateien, tmp_path):
    kaputt = dict(PROFIL, name="kaputt", farbe="rot")
    (tmp_path / "layouts.local.yaml").write_text(yaml.safe_dump({"profile": [kaputt]}), encoding="utf-8")
    rc, out = lauf(capsys, "profil", "liste")
    assert rc == 1 and "FEHLER" in out and "'farbe' ist unbekannt" in out and "cli_typ" in out


def test_alte_befehle_unveraendert(capsys, dateien):
    pdf, excel = dateien
    rc, out = lauf(capsys, "analyse", "--pdf", str(pdf), "--excel", str(excel))
    assert rc == 0 and "Lauf:" in out and "Änderungen:" in out and "Profil" not in out


# ---------------------------------------------------------------- farben
def test_farben_zeigt_zahlen_und_codes_aber_keine_woerter(capsys, dateien):
    pdf, _ = dateien
    rc, out = lauf(capsys, "farben", "--pdf", str(pdf))
    assert rc == 0
    low = out.casefold()
    for wort in PDF_WOERTER:
        assert not re.search(rf"\b{wort}\b", low), wort
    assert "Farben im PDF" in out and "Zeichen" in out
    assert re.search(r"#[0-9a-f]{6}", out)
    assert "gedeutet als neu" in out and "gedeutet als entfällt" in out
    assert "Wörter nach Status" in out and "neu: " in out and "entfällt: " in out
    assert "Unbekannte Farben" in out
    unbekannt = out.split("Unbekannte Farben")[1]
    assert re.search(r"#[0-9a-f]{6}: \d+ Zeichen", unbekannt)               # das gruene Wort bleibt unbekannt


def test_farben_ohne_unbekannte_farbe(capsys, tmp_path):
    pdf = tmp_path / "ohne.pdf"
    baue_pdf(pdf, [("h", "1 Eins"), ("p", "Ein {-alter-} {+neuer+} Satz bleibt.")])
    rc, out = lauf(capsys, "farben", "--pdf", str(pdf))
    assert rc == 0 and "  keine" in out.split("Unbekannte Farben")[1]
    assert not re.search(r"\b(alter|neuer|satz|bleibt)\b", out.casefold())


def test_farben_fehlermeldungen(capsys, tmp_path):
    rc, out = lauf(capsys, "farben", "--pdf", str(tmp_path / "gibt_es_nicht.pdf"))
    assert rc == 2 and "nicht gefunden" in out
    kaputt = tmp_path / "kaputt.pdf"
    kaputt.write_bytes(b"das ist kein pdf")
    rc, out = lauf(capsys, "farben", "--pdf", str(kaputt))
    assert rc == 2 and "nicht gelesen" in out and "Traceback" not in out
