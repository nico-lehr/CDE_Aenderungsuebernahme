"""HTTP-Schnittstelle des Dashboards (nur Weiterleitung; die Logik steckt in src/runs.py)."""
from __future__ import annotations

import tempfile
from pathlib import Path

from flask import Blueprint, Response, jsonify, request, send_file

from .. import config, runs
from ..vorschau import ausschnitt

api_bp = Blueprint("api", __name__)

_FARBE = {"absatz_neu": "#0071E3", "section_neu": "#0071E3", "text_eingefuegt": "#0071E3",
          "text_ersetzt": "#B9770E"}


def _fehler(text: str, code: int = 400):
    return jsonify({"fehler": text}), code


def _leicht(plan: dict) -> dict:
    """Plan fuers Frontend: Operationen weglassen (gross, nur intern noetig)."""
    out = dict(plan)
    out["aenderungen"] = [{**{k: v for k, v in a.items() if k != "ops"}, "ops_anzahl": len(a.get("ops", []))}
                          for a in plan["aenderungen"]]
    return out


@api_bp.get("/health")
def health():
    srv = config.lade("server")
    return jsonify({"ok": True, "projekt": "CDE-Änderungsübernahme", "port": srv["port"]})


@api_bp.get("/lauf")
def laeufe():
    return jsonify(runs.liste())


@api_bp.post("/lauf")
def neu():
    pdf, xls = request.files.get("pdf"), request.files.get("excel")
    if not pdf or not xls:
        return _fehler("Bitte PDF und Excel hochladen.")
    if Path(pdf.filename or "").suffix.lower() != ".pdf":
        return _fehler("Die Änderungsdatei muss ein PDF sein.")
    if Path(xls.filename or "").suffix.lower() not in (".xlsx", ".xlsm"):
        return _fehler("Die Excel muss als .xlsx (oder .xlsm) vorliegen. Alte .xls-Dateien bitte in Excel als .xlsx speichern.")
    with tempfile.TemporaryDirectory() as tmp:
        p, x = Path(tmp) / "a.pdf", Path(tmp) / Path(xls.filename).name
        pdf.save(p)
        xls.save(x)
        if p.read_bytes()[:5] != b"%PDF-" or x.read_bytes()[:2] != b"PK":
            return _fehler("Eine der Dateien ist kein gültiges PDF bzw. keine gültige Excel-Datei.")
        run_id = runs.neuer_lauf(p, x, pdf.filename or "aenderungen.pdf", xls.filename or "cde.xlsx")
    try:
        plan = runs.analysiere(run_id)
    except ValueError as e:
        runs.loeschen(run_id)
        return _fehler(str(e))
    except Exception as e:                       # noqa: BLE001 -- dem Nutzer verstaendlich melden
        runs.loeschen(run_id)
        return _fehler(f"Die Analyse ist fehlgeschlagen: {e}", 500)
    return jsonify(_leicht(plan))


@api_bp.get("/lauf/<run_id>")
def lauf(run_id: str):
    try:
        return jsonify(_leicht(runs.lade_plan(run_id)))
    except (ValueError, FileNotFoundError) as e:
        return _fehler(str(e), 404)


@api_bp.delete("/lauf/<run_id>")
def lauf_loeschen(run_id: str):
    try:
        runs.loeschen(run_id)
    except (ValueError, FileNotFoundError) as e:
        return _fehler(str(e), 404)
    return jsonify({"ok": True})


@api_bp.post("/lauf/<run_id>/uebernehmen")
def uebernehmen(run_id: str):
    daten = request.get_json(silent=True) or {}
    freig = daten.get("freigaben") or []
    if not isinstance(freig, list) or not all(isinstance(x, str) for x in freig):
        return _fehler("freigaben muss eine Liste von IDs sein.")
    try:
        res = runs.uebernehmen_lauf(run_id, freig)
    except (ValueError, FileNotFoundError) as e:
        return _fehler(str(e), 404)
    except Exception as e:                        # noqa: BLE001
        return _fehler(f"Das Übernehmen ist fehlgeschlagen: {e}", 500)
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
