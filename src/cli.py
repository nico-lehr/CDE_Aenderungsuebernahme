"""Kommandozeile: python -m src.cli analyse|lauf --pdf ... --excel ...

analyse  : nur auswerten und die Bilanz anzeigen (schreibt nichts in die Excel)
lauf     : auswerten und alle sicheren Aenderungen eintragen
           (mit --alle-vorschlaege werden auch Vorschlaege ("pruefen") uebernommen)
"""
from __future__ import annotations

import argparse
import sys

from . import runs


def _zeige(daten: dict) -> None:
    b = daten["bilanz"]
    print(f"Änderungen: {b['aenderungen']}  |  übernommen: {b['uebernommen']}  |  zu prüfen: {b['pruefen']}  |  "
          f"nicht übernommen: {b['nicht_uebernommen']}")
    print(f"Wörter im PDF markiert: {b['woerter_im_pdf']}  |  Bilanz stimmt: {'ja' if b['stimmt'] else 'NEIN'}")
    for w in daten.get("warnungen", []):
        print("Hinweis:", w)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="cde", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("befehl", choices=["analyse", "lauf"])
    p.add_argument("--pdf", required=True)
    p.add_argument("--excel", required=True)
    p.add_argument("--alle-vorschlaege", action="store_true")
    a = p.parse_args(argv)
    run_id = runs.neuer_lauf(a.pdf, a.excel)
    daten = runs.analysiere(run_id)
    print("Lauf:", run_id)
    _zeige(daten)
    if a.befehl == "lauf":
        freig = [x["id"] for x in daten["aenderungen"] if x["status"] == "pruefen"] if a.alle_vorschlaege else []
        res = runs.uebernehmen_lauf(run_id, freig)
        print("Ergebnis:", runs.lauf_dir(run_id) / res["ordner"] / res["excel"])
        _zeige({"bilanz": res["bilanz"], "warnungen": res["warnungen"]})
    return 0


if __name__ == "__main__":
    sys.exit(main())
