"""HTTP-Schnittstelle des Dashboards (nur Weiterleitung; die Logik steckt in src/runs.py).

Fehler: immer JSON {"fehler": "<einfacher deutscher Satz>"}; der Rohtext einer Ausnahme geht nur ins Server-Log.
"""
from __future__ import annotations

import logging
import tempfile
from pathlib import Path

from flask import Blueprint, Response, jsonify, request, send_file

from .. import runs
from ..vorschau import ausschnitt
from .einstellungen import srv

api_bp = Blueprint("api", __name__)
log = logging.getLogger("cde.api")

_FARBE = {"absatz_neu": "#0071E3", "section_neu": "#0071E3", "text_eingefuegt": "#0071E3",
          "text_ersetzt": "#B9770E"}

_MELDUNG_ANALYSE = ("Die Analyse ist fehlgeschlagen. Das PDF oder die Excel konnte nicht verarbeitet werden. "
                    "Die hochgeladenen Dateien bleiben unter diesem Lauf gespeichert.")


def _fehler(text: str, code: int = 400, **extra):
    return jsonify({"fehler": text, **extra}), code


def _leicht(plan: dict) -> dict:
    """Plan fuers Frontend: Operationen weglassen (gross, nur intern noetig)."""
    out = dict(plan)
    out["aenderungen"] = [{**{k: v for k, v in a.items() if k != "ops"}, "ops_anzahl": len(a.get("ops", []))}
                          for a in plan["aenderungen"]]
    return out


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
        except ValueError as e:
            return _fehler(str(e))
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
    try:
        plan = runs.analysiere(run_id)
    except Exception as e:                       # noqa: BLE001 -- der Lauf bleibt stehen (Regel 29), Meldung einfach halten
        if isinstance(e, ValueError):            # unsere eigenen Meldungen (z. B. "Excel beschaedigt") sind schon einfaches Deutsch
            meldung, code = str(e), 400
        else:
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
    try:
        return jsonify(_leicht(runs.lade_plan(run_id)))
    except (ValueError, FileNotFoundError) as e:
        return _fehler(str(e), 404)


@api_bp.delete("/lauf/<run_id>")
def lauf_archivieren(run_id: str):
    """Regel 29: loescht NICHTS, sondern archiviert (Marker im Lauf-Ordner)."""
    try:
        runs.archiviere(run_id)
    except (ValueError, FileNotFoundError) as e:
        return _fehler(str(e), 404)
    return jsonify({"ok": True, "archiviert": True})


@api_bp.post("/lauf/<run_id>/wiederherstellen")
def lauf_wiederherstellen(run_id: str):
    try:
        runs.stelle_wieder_her(run_id)
    except (ValueError, FileNotFoundError) as e:
        return _fehler(str(e), 404)
    return jsonify({"ok": True, "archiviert": False})


@api_bp.post("/lauf/<run_id>/uebernehmen")
def uebernehmen(run_id: str):
    daten = request.get_json(silent=True) or {}
    freig = (daten.get("freigaben") or []) if isinstance(daten, dict) else None
    if not isinstance(freig, list) or not all(isinstance(x, str) for x in freig):
        return _fehler("freigaben muss eine Liste von IDs sein.")
    try:
        res = runs.uebernehmen_lauf(run_id, freig)
    except (ValueError, FileNotFoundError) as e:
        return _fehler(str(e), 404)
    except runs.LaufFehler:
        raise                                     # zentrale Behandlung: klare Meldung + Status
    except Exception:                             # noqa: BLE001
        log.exception("Uebernahme fuer Lauf %s fehlgeschlagen", run_id)
        return _fehler("Das Übernehmen ist fehlgeschlagen. Es wurde nichts verändert und kein Ergebnis angelegt. "
                       "Bitte noch einmal versuchen.", 500)
    res["aenderungen"] = [{k: v for k, v in a.items() if k != "ops"} for a in res["aenderungen"]]
    return jsonify(res)


@api_bp.get("/lauf/<run_id>/vorschau/<aid>.png")
def vorschau(run_id: str, aid: str):
    try:
        plan = runs.lade_plan(run_id)
        a = next((x for x in plan["aenderungen"] if x["id"] == aid), None)
        if a is None:
            return _fehler("Änderung nicht gefunden.", 404)
        pdf = runs.lauf_dir(run_id) / "eingabe" / plan["pdf_name"]
        farbe = "#D93025" if "geloescht" in a["typ"] else _FARBE.get(a["typ"], "#0071E3")
        png = ausschnitt(pdf, a["seite"], tuple(a["bbox"]), farbe)
    except (ValueError, FileNotFoundError) as e:
        return _fehler(str(e), 404)
    return Response(png, mimetype="image/png", headers={"Cache-Control": "private, max-age=600"})


@api_bp.get("/lauf/<run_id>/download/<ordner>/<name>")
def download(run_id: str, ordner: str, name: str):
    try:
        p = runs.datei(run_id, ordner, name)
    except (ValueError, FileNotFoundError) as e:
        return _fehler(str(e), 404)
    return send_file(p, as_attachment=True, download_name=p.name)
