"""Kommandozeile: python -m src.cli analyse|lauf --docx ... --excel ...   (oder --pdf statt --docx)

analyse  : nur auswerten und die Bilanz anzeigen (schreibt nichts in die Excel)
lauf     : auswerten und alle sicheren Aenderungen eintragen
           (--docx = Word-Dokument mit Aenderungsverfolgung, bevorzugt; --pdf = PDF mit Aenderungen; genau eins von beiden;
            mit --alle-vorschlaege werden auch Vorschlaege ("pruefen") uebernommen;
            mit --abwahl A0003,A0007 werden diese automatischen Aenderungen NICHT eingetragen;
            mit --spec NAME gilt die Spec NAME der CDE, sonst wird sie aus dem PDF erkannt)
messen   : Messwerkzeug fuer Paare (PDF + Excel vorher + nachher), siehe src/messen.py
steckbrief: Steckbrief eines Paars (nur Zahlen, Farbcodes und Formen, kein Text), siehe src/steckbrief.py und docs/anleitung_steckbrief.md
vertrauen: wertet die Freigabe-Protokolle aller Laeufe aus (je Pruefcode: Anzahl, Freigabe-/Ablehnquote, Empfehlung); schreibt nichts
farben   : zeigt, welche Farben im PDF vorkommen und wie sie gedeutet werden (nur Zahlen und Farbcodes, keine Woerter)
profil   : Layout-Profile fuer den Aufbau der Excel: vorschau | bestaetigen | liste | analyse (siehe docs/adr/0003)
"""
from __future__ import annotations

import argparse
import colorsys
import logging
import sys

from . import config, runs, vertrauen
from .fehler import NutzerFehler
from .protokoll import bilanz_satz


_LUECKE_ART = {"scan_seite": "Scan-Seite", "mehrspaltig": "Mehrere Spalten", "tabelle": "Tabelle", "bild": "Bild",
               "unbekannte_farbe": "Unbekannte Farbe", "nicht_lesbarer_text": "Nicht lesbarer Text", "sonstiges": "Sonstiges"}


def _zeige(daten: dict, luecken: list[dict] | None = None, gezeigt: set[str] | None = None) -> None:
    """Gibt die Bilanz aus. `luecken`: die Luecken-Liste des PDFs (wird bei unvollstaendiger Bilanz aufgelistet).
    `gezeigt`: schon ausgegebene Hinweise (jeder Hinweis erscheint nur einmal)."""
    b = daten["bilanz"]
    dok = "Word-Dokument" if daten.get("quelle_art") == "docx" else "PDF"          # Eingabe Word oder PDF
    print(f"Änderungen: {b['aenderungen']}  |  übernommen: {b['uebernommen']}  |  zu prüfen: {b['pruefen']}  |  "
          f"nicht übernommen: {b['nicht_uebernommen']}")
    print(f"Wörter im {dok} markiert: {b['woerter_im_pdf']}")
    if b.get("stimmt") and b.get("vollstaendig") is True:
        print(f"Bilanz stimmt: ja (das {dok} wurde vollständig gelesen)")
    elif b.get("stimmt") and b.get("vollstaendig") is None:
        print(f"Bilanz stimmt: ja, aber ob das {dok} vollständig gelesen wurde, ist nicht geprüft")
    else:                                              # UNVOLLSTÄNDIG (Lücken) oder NEIN
        label, wert = bilanz_satz(b)
        print(f"{label}: {wert}")
    if b.get("vollstaendig") is False and luecken:
        print(f"Diese Stellen im {dok} konnten nicht gelesen werden (hier kann eine Änderung fehlen, bitte im {dok} selbst prüfen):")
        for l in luecken:
            art = _LUECKE_ART.get(l.get("art", ""), l.get("art", "") or "Sonstiges")
            print(f"  Seite {l.get('seite', '?')} | {art} | {l.get('beschreibung', '')}")
    gezeigt = gezeigt if gezeigt is not None else set()
    for w in daten.get("warnungen", []):
        if w not in gezeigt:
            gezeigt.add(w)
            print("Hinweis:", w)


_STATUS_NAMEN = {"neu": "neu", "entfaellt": "entfällt", "ignorieren": "ignoriert"}
_LEGENDE_NAMEN = {"linie": "Linie entscheidet (unterstrichen = neu, durchgestrichen = entfällt)",
                  "farbe": "Farbe (rot oder blau = neu, durchgestrichen = entfällt)",
                  "alt": "alte Legende (blau = neu, rot = entfällt)"}


def _gedeutet(klasse: str, modus: str, legende: dict) -> str:
    """Bedeutung einer Farbgruppe in der GEWAEHLTEN Legende (ADR 0006), nicht in der eingestellten."""
    if klasse == "-":
        return "unverändert"
    if modus not in ("linie", "farbe"):                        # alt: wie in legende.blau/rot/durchgestrichen
        return _STATUS_NAMEN.get(legende.get(klasse, ""), "unverändert")
    if klasse == "durchgestrichen":
        return "entfällt"
    if modus == "linie":
        return "unverändert (es zählt nur die Linie)"
    return "neu (ohne Durchstreichung)" if klasse in ("rot", "blau") else "unverändert"


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
    # ADR 0006: es zaehlt die Legende, die der Leser fuer DIESES PDF gewaehlt hat (bei modus auto je Dokument entschieden)
    modus = str(getattr(pa, "legende_modus", "") or legende.get("modus", "alt"))
    z: list[str] = [f"Seiten: {pa.seiten}", f"Legende: {_LEGENDE_NAMEN.get(modus, modus)}"
                    + (f" ({pa.legende_grund})" if getattr(pa, "legende_grund", "") else ""),
                    "", "Farben im PDF (Zeichen = einzelne Buchstaben in dieser Farbe):"]
    if not pa.farben:
        z.append("  keine Farben gefunden (nur Schwarz/Grau oder kein auslesbarer Text)")
    for f in pa.farben:
        status = _gedeutet(f["klasse"], modus, legende)
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
    if Path(a.pdf).suffix.lower() == ".docx":
        print("Für Word-Dokumente gibt es keine Farb-Auswertung: Die Änderungen stehen dort direkt in der Änderungsverfolgung.")
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


def _vertrauen(args: list[str]) -> int:
    """`vertrauen`: liest alle Freigabe-Protokolle unter data/runs und zeigt je Pruefcode Zahlen und eine Empfehlung. Nur lesen."""
    ap = argparse.ArgumentParser(prog="cde vertrauen", description="Wertet die Freigabe-Protokolle (entscheidungen.jsonl) aller Läufe aus. "
                                 "Schreibt nichts und ändert nichts am Eintragen.")
    try:
        ap.parse_args(args)
    except SystemExit as e:
        return 2 if e.code else 0
    zeilen, laeufe, kaputt = vertrauen.lese_alle(config.pfad("runs_dir"))      # nur lesen: legt auch keinen Ordner an
    for zeile in vertrauen.bericht(zeilen, laeufe, kaputt):
        print(zeile)
    return 0


def _zeige_vertrauen(daten: dict) -> None:
    """Stichprobe und Dokument-Ampel des Laufs (nur Kennungen und feste Saetze)."""
    a = vertrauen.ampel(daten)
    print(f"Dokument-Ampel: {dict(gruen='grün', gelb='gelb', rot='rot')[a['stufe']]} – {a['satz']}")
    for g in a["gruende"]:
        print(f"  {g}")
    st = [x for x in daten.get("stichprobe") or [] if isinstance(x, str)]
    if st:
        print(f"Stichprobe (bitte kurz ansehen, im Dashboard mit „Passt“ / „Falsch“): {', '.join(st)}")


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    if args and args[0] == "messen":
        from . import messen
        return messen.main(args[1:])
    if args and args[0] == "steckbrief":
        from . import steckbrief
        return steckbrief.main(args[1:])
    if args and args[0] == "farben":
        return _farben(args[1:])
    if args and args[0] == "vertrauen":
        return _vertrauen(args[1:])
    if args and args[0] == "profil":
        from . import layouts
        return layouts.main(args[1:])
    p = argparse.ArgumentParser(prog="cde", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("befehl", choices=["analyse", "lauf"])
    p.add_argument("--docx", default=None, metavar="DATEI", help="Word-Dokument mit Änderungsverfolgung (bevorzugt)")
    p.add_argument("--pdf", default=None, metavar="DATEI", help="PDF mit Änderungen (statt --docx)")
    p.add_argument("--excel", required=True)
    p.add_argument("--alle-vorschlaege", action="store_true")
    p.add_argument("--abwahl", default="", metavar="ID,ID",
                   help="nur bei 'lauf': automatische Änderungen, die NICHT eingetragen werden (IDs mit Komma getrennt)")
    p.add_argument("--spec", default=None, metavar="NAME",
                   help="Spec der CDE, zu der das PDF gehört (z. B. \"11 22 33\"); sonst wird sie aus dem PDF erkannt")
    a = p.parse_args(argv)
    abwahl = [x.strip() for x in a.abwahl.split(",") if x.strip()]
    from pathlib import Path
    if bool(a.pdf) == bool(a.docx):
        print("Bitte genau eine Änderungsdatei angeben: --docx (Word-Dokument, bevorzugt) oder --pdf (PDF).")
        return 2
    dok = a.docx or a.pdf
    for pfad, name in ((dok, "Word" if a.docx else "PDF"), (a.excel, "Excel")):
        if not Path(pfad).is_file():
            print(f"Die {name}-Datei wurde nicht gefunden: {pfad}")
            return 2
    run_id = None
    analysiert = False
    try:
        spec = runs.pruefe_spec(a.spec)                        # vor dem Anlegen des Laufs: nur Zeichen und Laenge
        runs.pruefe_aenderungsdatei(dok)                       # Endung UND Inhalt (PDF oder Word), vor dem Anlegen des Laufs
        run_id = runs.neuer_lauf(dok, a.excel)
        daten = runs.analysiere(run_id, spec=spec)
        analysiert = True
        print("Lauf:", run_id)
        sp = daten.get("spec")
        if isinstance(sp, dict) and len(sp.get("kandidaten") or []) > 1:
            z = sp.get("zeilen")
            print("Spec: " + (f"{sp['erkannt']} (Zeilen {z[0]}–{z[1]})" if sp.get("erkannt") else
                              f"nicht eindeutig, die CDE enthält {len(sp['kandidaten'])} Specs ({', '.join(sp['kandidaten'])}); "
                              "bitte mit --spec wählen"))
        if a.befehl == "analyse":
            print(f"Gespeichert: Die Analyse liegt als Lauf in {runs.lauf_dir(run_id)}. "
                  "In die Excel wird bei „analyse“ nichts geschrieben.")
            if abwahl:
                print("Hinweis: --abwahl gilt nur für „lauf“ (bei „analyse“ wird nichts eingetragen).")
        gezeigt: set[str] = set()
        _zeige(daten, daten.get("luecken", []), gezeigt)
        _zeige_vertrauen(daten)
        if a.befehl == "lauf":
            freig = [x["id"] for x in daten["aenderungen"] if x["status"] == "pruefen"] if a.alle_vorschlaege else []
            res = runs.uebernehmen_lauf(run_id, freig, abwahl, pauschal=freig)      # pauschal: nicht einzeln angesehen (Freigabe-Protokoll)
            print("Ergebnis:", runs.lauf_dir(run_id) / res["ordner"] / res["excel"])
            if abwahl:
                print(f"Abgewählt (nicht eingetragen): {', '.join(res['abwahl'])}")
            _zeige({"bilanz": res["bilanz"], "warnungen": res["warnungen"], "quelle_art": daten.get("quelle_art")}, None, gezeigt)     # Lücken stehen schon oben
    except NutzerFehler as e:                                  # gewollte Meldung: schon einfaches Deutsch
        print(e.meldung)
        return 2
    except Exception:                                          # kaputte Datei o. ae.: einfache Meldung, kein Traceback/Rohtext
        logging.getLogger("cde.cli").debug("Verarbeitung fehlgeschlagen", exc_info=True)    # Einzelheiten nur mit Debug-Log
        meldung = ("Das Dokument oder die Excel konnte nicht verarbeitet werden. Sind es unbeschädigte Dateien "
                   "(Word .docx bzw. PDF und .xlsx/.xlsm)?")
        print(meldung)
        if run_id:
            print(f"Der Lauf {run_id} bleibt mit den Eingabedateien gespeichert.")
            if not analysiert:
                try:
                    runs.markiere_analyse_fehler(run_id, meldung)
                except Exception:                              # noqa: BLE001 -- die Meldung oben genuegt
                    pass
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
