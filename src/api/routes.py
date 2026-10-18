"""HTTP-Schnittstelle des Dashboards (nur Weiterleitung; die Logik steckt in src/runs.py).

Fehler: immer JSON {"fehler": "<einfacher deutscher Satz>"}. Nur NutzerFehler (src/fehler.py, auch LaufFehler) geben ihren
Text weiter; jede andere Ausnahme (auch ein ValueError) wird zur allgemeinen Meldung mit Status 500, der Rohtext geht nur
ins Server-Log. Die zentrale Behandlung steht in app.py.
"""
from __future__ import annotations

import logging
import tempfile
import threading
from pathlib import Path

from flask import Blueprint, Response, jsonify, request, send_file

from .. import runs
from ..fehler import NutzerFehler
from ..vorschau import ausschnitt
from .einstellungen import srv

api_bp = Blueprint("api", __name__)
log = logging.getLogger("cde.api")

# pdfium (pypdfium2, ueber pdfplumber.to_image in src/vorschau.py und in src/sicht.py bei der Analyse) ist NICHT thread-sicher:
# zwei gleichzeitige Aufrufe koennen den ganzen Server abstuerzen lassen (Speicherfehler). Der Server laeuft mit Threads und das
# Dashboard laedt mehrere PDF-Ausschnitte zugleich -- deshalb laeuft alles, was pdfium benutzt, nacheinander.
_PDFIUM = threading.Lock()

_FARBE = {"absatz_neu": "#0071E3", "section_neu": "#0071E3", "text_eingefuegt": "#0071E3",
          "text_ersetzt": "#B9770E"}

_MELDUNG_ANALYSE = ("Die Analyse ist fehlgeschlagen. Das PDF oder die Excel konnte nicht verarbeitet werden. "
                    "Die hochgeladenen Dateien bleiben unter diesem Lauf gespeichert.")


def _fehler(text: str, code: int = 400, **extra):
    return jsonify({"fehler": text, **extra}), code


def _fuers_frontend(aenderungen: list[dict]) -> list[dict]:
    """Je Aenderung ohne Operationen (gross, nur intern noetig), dafuer mit `kategorie` (wohin sie im Dashboard gehoert).
    Alle Felder des Plans (auch Schema 4: bereich, stuecke, grund_einfach) bleiben erhalten."""
    return [{**{k: v for k, v in a.items() if k != "ops"}, "ops_anzahl": len(a.get("ops", [])),
             "kategorie": runs.kategorie(a)} for a in aenderungen]


def _leicht(plan: dict) -> dict:
    """Plan fuers Frontend: Operationen und PDF-Absaetze (nur fuer die Gegenprobe) weglassen, Zusammenfassung dazu."""
    out = {k: v for k, v in plan.items() if k != "absaetze"}
    out["aenderungen"] = _fuers_frontend(plan["aenderungen"])
    out["zusammenfassung"] = runs.zusammenfassung(plan["aenderungen"])
    return out


def _id_liste(daten, name: str):
    """Liste von IDs aus dem JSON-Body (fehlt sie, ist sie leer); sonst None = ungueltig."""
    werte = (daten.get(name) or []) if isinstance(daten, dict) else None
    if not isinstance(werte, list) or not all(isinstance(x, str) for x in werte):
        return None
    return werte


@api_bp.get("/health")
def health():
    return jsonify({"ok": True, "projekt": "CDE-Änderungsübernahme", "port": srv()["port"]})


@api_bp.get("/lauf")
def laeufe():
    mit_archiv = request.args.get("archiviert", "").lower() in ("1", "true", "ja")
    return jsonify(runs.liste(mit_archiv))


@api_bp.post("/lauf")
def neu():
    pdf, xls = request.files.get("pdf"), request.files.get("excel")
    if not pdf or not xls:
        return _fehler("Bitte PDF und Excel hochladen.")
    if Path(pdf.filename or "").suffix.lower() != ".pdf":
        return _fehler("Die Änderungsdatei muss ein PDF sein.")
    if Path(xls.filename or "").suffix.lower() not in (".xlsx", ".xlsm"):
        return _fehler("Die Excel muss als .xlsx (oder .xlsm) vorliegen. Alte .xls-Dateien bitte in Excel als .xlsx speichern.")
    for f in (pdf, xls):
        try:
            runs.pruefe_dateiname(f.filename or "")
        except NutzerFehler as e:
            return _fehler(e.meldung, e.status)
    grenze_mb = int(srv()["upload_max_mb_datei"])
    for f in (pdf, xls):
        f.stream.seek(0, 2)
        groesse = f.stream.tell()
        f.stream.seek(0)
        if groesse > grenze_mb * 1024 * 1024:
            return _fehler(f"Die Datei „{Path(f.filename or '').name}“ ist zu groß. "
                           f"Erlaubt sind höchstens {grenze_mb} MB je Datei.", 413)
    with tempfile.TemporaryDirectory() as tmp:
        p, x = Path(tmp) / "a.pdf", Path(tmp) / "b.xlsx"      # feste Namen: der Originalname kann zu lang/ungueltig sein
        pdf.save(p)
        xls.save(x)
        with open(p, "rb") as fp, open(x, "rb") as fx:         # nur die ersten Bytes lesen, nicht die ganze Datei
            ok = fp.read(5) == b"%PDF-" and fx.read(2) == b"PK"
        if not ok:
            return _fehler("Eine der Dateien ist kein gültiges PDF bzw. keine gültige Excel-Datei.")
        run_id = runs.neuer_lauf(p, x, pdf.filename or "aenderungen.pdf", xls.filename or "cde.xlsx")
    profil = (request.form.get("profil") or "").strip() or None    # Layout-Profil: nur bestaetigte werden angenommen
    try:
        with _PDFIUM:                            # die Analyse liest das PDF ein zweites Mal mit pdfium (src/sicht.py)
            plan = runs.analysiere(run_id, profil)
    except Exception as e:                       # noqa: BLE001 -- der Lauf bleibt stehen (Regel 29), Meldung einfach halten
        if isinstance(e, NutzerFehler):          # gewollte Meldung (z. B. Profil nicht bestaetigt, Excel beschaedigt): schon einfaches Deutsch
            meldung, code = e.meldung, e.status
        else:                                    # alles andere ist unerwartet: nie der Rohtext, nur ins Server-Log
            log.exception("Analyse von Lauf %s fehlgeschlagen", run_id)
            meldung, code = _MELDUNG_ANALYSE, 500
        try:
            runs.markiere_analyse_fehler(run_id, meldung)
        except Exception:                        # noqa: BLE001
            log.exception("Fehlermarke fuer Lauf %s konnte nicht geschrieben werden", run_id)
        return _fehler(meldung, code, run_id=run_id, status="fehlgeschlagen")
    return jsonify(_leicht(plan))


@api_bp.get("/lauf/<run_id>")
def lauf(run_id: str):
    return jsonify(_leicht(runs.lade_plan(run_id)))


@api_bp.delete("/lauf/<run_id>")
def lauf_archivieren(run_id: str):
    """Regel 29: loescht NICHTS, sondern archiviert (Marker im Lauf-Ordner)."""
    runs.archiviere(run_id)
    return jsonify({"ok": True, "archiviert": True})


@api_bp.post("/lauf/<run_id>/wiederherstellen")
def lauf_wiederherstellen(run_id: str):
    runs.stelle_wieder_her(run_id)
    return jsonify({"ok": True, "archiviert": False})


@api_bp.post("/lauf/<run_id>/uebernehmen")
def uebernehmen(run_id: str):
    daten = request.get_json(silent=True) or {}
    freig, abwahl = _id_liste(daten, "freigaben"), _id_liste(daten, "abwahl")
    if freig is None:
        return _fehler("freigaben muss eine Liste von IDs sein.")
    if abwahl is None:
        return _fehler("abwahl muss eine Liste von IDs sein.")
    try:
        res = runs.uebernehmen_lauf(run_id, freig, abwahl)
    except NutzerFehler:
        raise                                     # zentrale Behandlung (app.py): klare Meldung + Status
    except Exception:                             # noqa: BLE001
        log.exception("Uebernahme fuer Lauf %s fehlgeschlagen", run_id)
        return _fehler("Das Übernehmen ist fehlgeschlagen. Es wurde nichts verändert und kein Ergebnis angelegt. "
                       "Bitte noch einmal versuchen.", 500)
    res["zusammenfassung"] = runs.zusammenfassung(res["aenderungen"])
    res["aenderungen"] = _fuers_frontend(res["aenderungen"])
    return jsonify(res)


@api_bp.get("/lauf/<run_id>/vorschau/<aid>.png")
def vorschau(run_id: str, aid: str):
    plan = runs.lade_plan(run_id)
    a = next((x for x in plan["aenderungen"] if x["id"] == aid), None)
    if a is None:
        return _fehler("Änderung nicht gefunden.", 404)
    pdf = runs.lauf_dir(run_id) / "eingabe" / plan["pdf_name"]
    farbe = "#D93025" if "geloescht" in a["typ"] else _FARBE.get(a["typ"], "#0071E3")
    with _PDFIUM:                                # nie zwei Ausschnitte gleichzeitig rendern (siehe _PDFIUM)
        png = ausschnitt(pdf, a["seite"], tuple(a["bbox"]), farbe)
    return Response(png, mimetype="image/png", headers={"Cache-Control": "private, max-age=600"})


@api_bp.get("/lauf/<run_id>/download/<ordner>/<name>")
def download(run_id: str, ordner: str, name: str):
    p = runs.datei(run_id, ordner, name)
    return send_file(p, as_attachment=True, download_name=p.name)
