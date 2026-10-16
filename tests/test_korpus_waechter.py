"""Korpus-Waechter: 8 feste, kleine, erfundene Paare (Generator v2) werden erzeugt und gemessen (Regel 16: nur erfundene Daten).

Er schuetzt die wichtigsten Zusagen des Tools bei jeder Aenderung am Code -- schnell (< 60 s) und ohne die grossen Korpora:
  * Bilanz stimmt, 0 Regel-27-Verstoesse, 0 Kollateralschaden, zweiter Lauf aendert nichts (Idempotenz)
  * Paare, die beim Anlegen des Waechters (16.10.2026, Welle 1) 0 stille Fehler hatten, behalten 0 stille Fehler.
Verschiedene Stoerer-Familien; zwei Paare entstehen mit LibreOffice (fehlt soffice, werden nur diese uebersprungen).
Seeds 201-211 liegen in keinem Korpus (Regel 11: nur Entwicklungs-Seeds < 1000); alles entsteht in tmp_path.
"""
from __future__ import annotations

import time

import pytest

from src import messen
from tests.fixtures import pdf_lo
from tests.fixtures.generator import erzeuge_paar

KLEIN = {"basis": "standard", "kapitel": (2, 3), "gruppen": (3, 4), "min_aenderungen": 6}

# (Seed, Stoerer-Familien, Erzeuger, stille Fehler beim Anlegen des Waechters)
PAARE = [
    (201, [], "writer", 0),
    (202, ["fussnote", "silbentrennung"], "writer", 0),
    (203, ["aufzaehlung"], "writer", 0),
    (204, ["verschmolzen", "sonderzeichen"], "writer", 0),
    (205, ["mehrere_blaetter", "kommentar_spalte"], "writer", 0),
    (209, ["lfd_spalte", "doppelte_absaetze"], "writer", 0),
    (207, ["farbtoene", "hyperlink"], "libreoffice", 0),
    (211, ["aufzaehlung", "silbentrennung"], "libreoffice", 0),
]
ZEITLIMIT_S = 60


def test_korpus_waechter_8_feste_paare(tmp_path):
    start = time.monotonic()
    befunde: list[str] = []
    gemessen = 0
    for seed, familien, erzeuger, stille_heute in PAARE:
        if erzeuger == "libreoffice" and not pdf_lo.verfuegbar():
            continue                                       # ohne soffice: nur die LibreOffice-Paare entfallen
        profil = {**KLEIN, "name": f"waechter{seed}", "stoerer_liste": familien}
        paar = erzeuge_paar(seed, profil, tmp_path / "entwicklung", "entwicklung", erzeuger=erzeuger)
        r = messen.messe_paar(paar)
        gemessen += 1
        name = f"{seed}/{erzeuger}/{'+'.join(familien) or 'ohne'}"
        if r.get("abgestuerzt"):
            befunde.append(f"{name}: Messung abgestuerzt ({r['abgestuerzt']})")
            continue
        if r["bilanz_stimmt"] is not True:
            befunde.append(f"{name}: Bilanz stimmt nicht")
        if r["regel27_verstoesse"]:
            befunde.append(f"{name}: {r['regel27_verstoesse']} Regel-27-Verstoss/Verstoesse")
        if r["kollateral"]:
            befunde.append(f"{name}: {r['kollateral']} Kollateralschaden")
        if r["idempotent"] is not True:
            befunde.append(f"{name}: zweiter Lauf aendert etwas (Idempotenz)")
        if stille_heute == 0 and r["stille_fehler"]:
            arten = sorted(r["stille_fehler_nach_art"])
            befunde.append(f"{name}: {r['stille_fehler']} stille Fehler (vorher 0) -- Arten {arten}")
    dauer = time.monotonic() - start
    assert gemessen >= 6
    assert not befunde, "Korpus-Waechter: " + "; ".join(befunde)
    assert dauer < ZEITLIMIT_S, f"Korpus-Waechter zu langsam: {dauer:.0f} s"


@pytest.mark.skipif(not pdf_lo.verfuegbar(), reason="soffice (LibreOffice) fehlt")
def test_waechter_hat_libreoffice_paare():
    assert sum(1 for p in PAARE if p[2] == "libreoffice") in (1, 2)
    assert len(PAARE) == 8 and len({s for s, *_ in PAARE}) == 8
    familien = {f for _, fam, _, _ in PAARE for f in fam}
    assert len(familien) >= 8
