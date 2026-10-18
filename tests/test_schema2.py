"""Schema 2 (ADR 0002): alte Laeufe bleiben lesbar, neue Felder haben Standardwerte, Luecken machen die Bilanz unvollstaendig.

Alle Daten erfunden (Regel 16).
"""
from __future__ import annotations

from src import config
from src.models import (GELESENE_SCHEMAS, LUECKEN_ARTEN, SCHEMA_VERSION, Aenderung, Operation, luecke)
from src.protokoll import aenderung_aus_dict, bilanz


def _aenderung() -> Aenderung:
    op = Operation(art="ersetzen", blatt="CDE", zeile=3, spalte=2, start=0, ende=4, text="neu", aenderung="a1", erwartet="alt")
    return Aenderung(id="a1", typ="text_ersetzt", seite=1, bbox=(0.0, 0.0, 1.0, 1.0), section="1.1",
                     alt="alt", neu="neu", woerter=2, status="uebernommen", konfidenz="hoch", ops=[op])


def test_schema_version_und_gelesene_versionen():
    assert SCHEMA_VERSION == 4          # Schema 4 (ADR 0006/0007): bereich/stuecke/grund_einfach; neue Felder nur mit Standardwert
    assert 1 in GELESENE_SCHEMAS and 3 in GELESENE_SCHEMAS and SCHEMA_VERSION in GELESENE_SCHEMAS


def test_schema4_felder_haben_standardwerte_und_schema3_bleibt_lesbar():
    from src.protokoll import aenderung_aus_dict
    d = _aenderung().zu_dict()
    for k in ("bereich", "stuecke", "grund_einfach"):
        d.pop(k)
    a = aenderung_aus_dict(d)
    assert a.bereich == "" and a.stuecke == [] and a.grund_einfach == ""
    a.stuecke = [["normal", "Die Pumpe "], ["alt", "alt"], ["neu", "neu"]]
    assert aenderung_aus_dict(a.zu_dict()).stuecke == a.stuecke


def test_schema1_aenderung_ohne_neue_felder_wird_gelesen():
    d = _aenderung().zu_dict()
    for neu in ("hinweise", "wirkung"):          # Schema-1-Plan kennt diese Felder nicht
        d.pop(neu)
    ae = aenderung_aus_dict(d)
    assert ae.hinweise == [] and ae.wirkung == ""
    assert ae.ops[0].erwartet == "alt"


def test_unbekannte_felder_aus_neuerem_schema_werden_ignoriert():
    d = _aenderung().zu_dict()
    d["zukunftsfeld"] = 1
    d["ops"][0]["noch_ein_feld"] = "x"
    ae = aenderung_aus_dict(d)
    assert ae.id == "a1"


def test_bilanz_vollstaendig_nur_ohne_luecken():
    ae = _aenderung()
    sauber = bilanz([ae], 2, [])
    assert sauber["stimmt"] and sauber["vollstaendig"] and sauber["luecken"] == 0
    mit = bilanz([ae], 2, [], [luecke("scan_seite", 2, "Seite 2 enthält keinen lesbaren Text.")])
    assert mit["stimmt"] and not mit["vollstaendig"] and mit["luecken"] == 1


def test_luecke_format_und_arten():
    l = luecke("tabelle", 3, "Auf Seite 3 steht eine Tabelle.", 12)
    assert set(l) == {"art", "seite", "beschreibung", "woerter"}
    assert l["art"] in LUECKEN_ARTEN


def test_neue_ordner_folgen_dem_daten_wurzelordner():
    runs = config.pfad("runs_dir")
    assert config.pfad("paare_dir") == runs.parent / "paare"
    assert config.pfad("messung_dir") == runs.parent / "messung"


def test_bilanz_satz_bei_unbekannter_vollstaendigkeit_ist_nie_gruen():
    """Aeltere Laeufe (Schema 1) haben nie auf Luecken geprueft: das Protokoll darf nicht 'nichts ging verloren' sagen."""
    from src.protokoll import bilanz_satz
    _, wert = bilanz_satz({"stimmt": True, "vollstaendig": None, "luecken": None})
    assert "nicht geprüft" in wert and "nichts ging still verloren" not in wert
    _, wert = bilanz_satz({"stimmt": True, "vollstaendig": True})
    assert "nichts ging still verloren" in wert
