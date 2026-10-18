"""Pruefcodes in einfachem Deutsch (src/codes.py, Schema 4 `Aenderung.grund_einfach`).

Jeder Code, den Planer, Abgleich, Applier, Gegenprobe oder PDF-Leser ausgeben koennen, braucht einen Satz. Die Codes werden
per Quelltext-Suche eingesammelt: kommt ein neuer Code dazu, faellt dieser Test, bis er einen Satz hat.
"""
from __future__ import annotations

import re
from pathlib import Path
from types import SimpleNamespace

import pytest

from src import abgleich, codes, config
from src.models import NICHT, PRUEFEN, UEBERNOMMEN

SRC = Path(__file__).resolve().parents[1] / "src"
_MUSTER = [
    r'_code\(\w+, "([a-z0-9_]+)"\)',                     # _code(ae, "...")
    r'codes\.append\("([a-z0-9_]+)"\)',
    r'probleme\.append\("([a-z0-9_]+)"\)',
    r'probleme\[\w+\] = \["([a-z0-9_]+)"\]',
    r'_deckel = \["([a-z0-9_]+)"\]',
    r'"(gegenprobe_(?!pflicht)[a-z_]+)"',
    r'hinweis\.add\("([a-z_]+)"\)',                       # Hinweis-Codes des PDF-Lesers
    r'\| \{"([a-z_]+)"\}',
]


def _quelltext_codes() -> set[str]:
    out: set[str] = set()
    for name in ("planner.py", "abgleich.py", "applier.py", "gegenprobe.py", "pdf_reader.py", "sicht.py", "matcher.py"):
        text = (SRC / name).read_text(encoding="utf-8")
        for m in _MUSTER:
            out.update(re.findall(m, text))
    # Status des Reissverschlusses (codes.append(z.status)), Abweichungsarten der Gegenprobe ("a1_" + art)
    out.update({abgleich.MEHRDEUTIG, abgleich.WIDERSPRUCH, abgleich.KEINER})
    g = (SRC / "gegenprobe.py").read_text(encoding="utf-8")
    arten = set(re.findall(r'return "([a-z_]+)"', g.split("def _art")[1].split("def ")[0]))
    arten |= set(re.findall(r'Vergleich\(False, abb, "([a-z_]+)"\)', g))
    out.update("a1_" + a for a in arten | {"abweichung"})
    # Absatz-Sperre und reine Hinweise aus der Konfiguration
    pl = config.lade("zuordnung")["planer"]
    out.update("sperre:" + str(c) for c in pl["unsichere_hinweise"])
    out.update(str(c) for c in pl["hinweis_codes"])
    out.add("sperre:kommentar_highlight")
    return out


def test_quelltext_suche_findet_die_codes():
    gefunden = _quelltext_codes()
    for c in ("nicht_gefunden", "schon_eingetragen", "naht_unklar", "gegenprobe_text", "ausserhalb_cde", "label_spalte",
              "blatt_ohne_anker", "farbe_ohne_linie", "hinterlegt", "a1_buchstabe", "sperre:tabelle", "mehrdeutig"):
        assert c in gefunden, c
    assert len(gefunden) >= 60


@pytest.mark.parametrize("code", sorted(_quelltext_codes()))
def test_jeder_code_hat_einen_satz(code):
    satz = codes.satz(code)
    assert satz and satz[0].isupper() and satz.rstrip().endswith((".", ")")), (code, satz)
    assert "_" not in satz, (code, satz)                    # einfaches Deutsch, keine Kennungen


def test_hinweis_codes_enthalten_die_neuen_leser_codes_ohne_sperre():
    pl = config.lade("zuordnung")["planer"]
    assert {"farbe_ohne_linie", "hinterlegt"} <= set(pl["hinweis_codes"])
    assert not {"farbe_ohne_linie", "hinterlegt"} & set(pl["unsichere_hinweise"])


def _ae(status, pruefcodes=(), grund=""):
    return SimpleNamespace(status=status, pruefcodes=list(pruefcodes), grund=grund)


def test_grund_einfach_wichtigster_grund_zuerst():
    assert codes.grund_einfach(_ae(PRUEFEN, ["kurz", "sperre:tabelle"])) == codes.satz("sperre:tabelle")
    assert codes.grund_einfach(_ae(NICHT, ["zu_unsicher", "ausserhalb_cde"])) == codes.satz("ausserhalb_cde")
    assert codes.grund_einfach(_ae(PRUEFEN, ["anker_teil", "a1_ziffer"])).startswith("Der alte Text steht in der Excel")
    assert codes.grund_einfach(_ae(UEBERNOMMEN)).startswith("Wird automatisch eingetragen")
    assert codes.grund_einfach(_ae(NICHT, [], "Der alte Text steht in der Excel anders; Bitte prüfen.")) == \
        "Der alte Text steht in der Excel anders"
    assert codes.grund_einfach(_ae(PRUEFEN)) and codes.grund_einfach(_ae(NICHT))


def test_planer_tabellen_kommen_aus_codes():
    from src import planner
    assert planner._GRUND is codes.SAETZE and planner._A1_SATZ is codes.A1_ART
