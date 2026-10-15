"""Kommandozeile: python -m src.cli analyse|lauf --pdf ... --excel ...

analyse  : nur auswerten und die Bilanz anzeigen (schreibt nichts in die Excel)
lauf     : auswerten und alle sicheren Aenderungen eintragen
           (mit --alle-vorschlaege werden auch Vorschlaege ("pruefen") uebernommen)
messen   : Messwerkzeug fuer Paare (PDF + Excel vorher + nachher), siehe src/messen.py
farben   : zeigt, welche Farben im PDF vorkommen und wie sie gedeutet werden (nur Zahlen und Farbcodes, keine Woerter)
profil   : Layout-Profile fuer den Aufbau der Excel: vorschau | bestaetigen | liste | analyse (siehe docs/adr/0003)
"""
from __future__ import annotations

import argparse
import colorsys
import sys

from . import config, runs


def _zeige(daten: dict) -> None:
    b = daten["bilanz"]
    print(f"Änderungen: {b['aenderungen']}  |  übernommen: {b['uebernommen']}  |  zu prüfen: {b['pruefen']}  |  "
          f"nicht übernommen: {b['nicht_uebernommen']}")
    print(f"Wörter im PDF markiert: {b['woerter_im_pdf']}  |  Bilanz stimmt: {'ja' if b['stimmt'] else 'NEIN'}")
    for w in daten.get("warnungen", []):
        print("Hinweis:", w)


_STATUS_NAMEN = {"neu": "neu", "entfaellt": "entfällt", "ignorieren": "ignoriert"}


def _ist_farbig(hex_farbe: str, min_saettigung: float) -> bool:
    """Kraeftige Farbe (nicht grau/schwarz/weiss): so eine Farbe waere eine moegliche Markierung."""
    try:
        r, g, b = (int(hex_farbe[i:i + 2], 16) / 255 for i in (1, 3, 5))
    except (ValueError, IndexError):
        return False
    return colorsys.rgb_to_hsv(r, g, b)[1] >= min_saettigung


def farben_bericht(pa) -> list[str]:
    """Zeilen fuer `farben`: je Farbe Zeichenzahl, Klasse und gedeuteter Status; dazu Woerter je Status und unbekannte
    Farben. Enthaelt NUR Zahlen, Farbcodes und feste Begriffe, nie ein Wort aus dem PDF (Regel 16)."""
    cfg = config.lade("farben")
    legende = cfg.get("legende", {})
    z: list[str] = [f"Seiten: {pa.seiten}", "", "Farben im PDF (Zeichen = einzelne Buchstaben in dieser Farbe):"]
    if not pa.farben:
        z.append("  keine Farben gefunden (nur Schwarz/Grau oder kein auslesbarer Text)")
    for f in pa.farben:
        status = _STATUS_NAMEN.get(legende.get(f["klasse"], ""), "unverändert") if f["klasse"] != "-" else "unverändert"
        z.append(f"  {f['farbe']}  {f['art']}: {f['zeichen']} Zeichen | Gruppe {f['klasse']} | gedeutet als {status}")
    zaehl = {"add": 0, "del": 0, "mod": 0, "keep": 0}
    for b in pa.blocks:
        for t in b.tokens:
            zaehl[t.status] = zaehl.get(t.status, 0) + 1
    ign = sum(i["woerter"] for i in pa.ignoriert)
    z += ["", "Wörter nach Status:",
          f"  neu: {zaehl['add']} | entfällt: {zaehl['del']} | geändert: {zaehl['mod']} | unverändert: {zaehl['keep']} | "
          f"ignoriert (Kopf-/Fußzeile): {ign}"]
    sat = float(cfg.get("text", {}).get("min_saettigung", 0.35))
    unbekannt = [f for f in pa.farben if f["klasse"] == "-" and f["art"] == "Text" and _ist_farbig(f["farbe"], sat)]
    z += ["", "Unbekannte Farben (kräftig, aber nicht in der Legende, werden als unverändert behandelt):"]
    z += [f"  {f['farbe']}: {f['zeichen']} Zeichen" for f in unbekannt] or ["  keine"]
    luecken = [l for l in pa.luecken if l["art"] == "unbekannte_farbe"]
    if luecken:
        z.append("Als Lücke gemeldet: " + "; ".join(f"Seite {l['seite']}: {l['woerter']} Wörter" for l in luecken))
    z += ["", f"Lücken insgesamt: {len(pa.luecken)}"]
    for art in sorted({l["art"] for l in pa.luecken}):
        n = [l for l in pa.luecken if l["art"] == art]
        z.append(f"  {art}: {len(n)}x, zusammen {sum(l['woerter'] for l in n)} Wörter, Seiten {', '.join(str(s) for s in sorted({l['seite'] for l in n}))}")
    z += ["", "Wenn eine Farbe falsch gedeutet wird: Schwellen in config/farben.yaml (siehe docs/kalibrierung.md), nur an Entwicklungsbeispielen."]
    return z


def _farben(args: list[str]) -> int:
    ap = argparse.ArgumentParser(prog="cde farben", description="Zeigt die Farben eines Änderungs-PDFs (nur Zahlen und Farbcodes, keine Wörter).")
    ap.add_argument("--pdf", required=True)
    try:
        a = ap.parse_args(args)
    except SystemExit as e:
        return 2 if e.code else 0
    from pathlib import Path
    if not Path(a.pdf).is_file():
        print("Die PDF-Datei wurde nicht gefunden.")
        return 2
    from .pdf_reader import analysiere_pdf
    try:
        pa = analysiere_pdf(a.pdf)
    except Exception:                                  # kaputtes PDF: einfache Meldung, kein Rohtext
        print("Das PDF konnte nicht gelesen werden. Ist es eine unbeschädigte PDF-Datei?")
        return 2
    for zeile in farben_bericht(pa):
        print(zeile)
    return 0


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    if args and args[0] == "messen":
        from . import messen
        return messen.main(args[1:])
    if args and args[0] == "farben":
        return _farben(args[1:])
    if args and args[0] == "profil":
        from . import layouts
        return layouts.main(args[1:])
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
