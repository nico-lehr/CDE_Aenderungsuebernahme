"""HTTP-Schnittstelle des Dashboards (nur Weiterleitung; die Logik steckt in src/runs.py).

Fehler: immer JSON {"fehler": "<einfacher deutscher Satz>"}. Nur NutzerFehler (src/fehler.py, auch LaufFehler) geben ihren
Text weiter; jede andere Ausnahme (auch ein ValueError) wird zur allgemeinen Meldung mit Status 500, der Rohtext geht nur
ins Server-Log. Die zentrale Behandlung steht in app.py.
"""
from __future__ import annotations

import logging
import tempfile
from pathlib import Path

from flask import Blueprint, Response, jsonify, request, send_file

from .. import runs, vertrauen
from ..fehler import NutzerFehler
from ..vorschau import PDFIUM_SPERRE, ausschnitt
from .einstellungen import srv

api_bp = Blueprint("api", __name__)
log = logging.getLogger("cde.api")

# pdfium ist nicht thread-sicher (siehe src/vorschau.py): alles, was pdfium benutzt (Analyse, PDF-Ausschnitte), laeuft unter
# der gemeinsamen, wiedereintrittsfaehigen Sperre vorschau.PDFIUM_SPERRE.
_PDFIUM = PDFIUM_SPERRE

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
    """Plan fuers Frontend: Operationen und PDF-Absaetze (nur fuer die Gegenprobe) weglassen, Zusammenfassung und Dokument-Ampel dazu."""
    out = {k: v for k, v in plan.items() if k != "absaetze"}
    out["aenderungen"] = _fuers_frontend(plan["aenderungen"])
    out["zusammenfassung"] = runs.zusammenfassung(plan["aenderungen"])
    out["ampel"] = vertrauen.ampel(plan)                    # Ausgangslage: noch nichts entschieden, Stichprobe nicht angesehen
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
    # Das Dokument mit den Aenderungen: Word (.docx) oder PDF. Feld "datei" (neu), "docx" oder "pdf" (frueher); die Endung entscheidet.
    pdf = request.files.get("datei") or request.files.get("docx") or request.files.get("pdf")
    xls = request.files.get("excel")
    if not pdf or not xls:
        return _fehler("Bitte das Word-Dokument (oder PDF) mit den Änderungen und die Excel hochladen.")
    try:
        endung = runs.pruefe_aenderungsendung(pdf.filename or "")
    except NutzerFehler as e:
        return _fehler(e.meldung, e.status)
    if Path(xls.filename or "").suffix.lower() not in (".xlsx", ".xlsm"):
        return _fehler("Die Excel muss als .xlsx (oder .xlsm) vorliegen. Alte .xls-Dateien bitte in Excel als .xlsx speichern.")
    for f in (pdf, xls):
        try:
            runs.pruefe_dateiname(f.filename or "")
        except NutzerFehler as e:
            return _fehler(e.meldung, e.status)
    try:
        spec = runs.pruefe_spec(request.form.get("spec"))       # nur Zeichen und Laenge; vor dem Anlegen des Laufs
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
        p, x = Path(tmp) / f"a{endung}", Path(tmp) / "b.xlsx"      # feste Namen: der Originalname kann zu lang/ungueltig sein
        pdf.save(p)
        xls.save(x)
        try:
            runs.pruefe_aenderungsdatei(p, p.name)             # Inhalt: PDF beginnt mit %PDF-, Word ist ein ZIP mit word/document.xml
        except NutzerFehler as e:
            return _fehler(e.meldung, e.status)
        with open(x, "rb") as fx:                              # nur die ersten Bytes lesen, nicht die ganze Datei
            if fx.read(2) != b"PK":
                return _fehler("Die Datei ist keine gültige Excel-Datei (.xlsx).")
        run_id = runs.neuer_lauf(p, x, pdf.filename or f"aenderungen{endung}", xls.filename or "cde.xlsx")
    profil = (request.form.get("profil") or "").strip() or None    # Layout-Profil: nur bestaetigte werden angenommen
    return _analysiere_antwort(run_id, profil, spec)


def _analysiere_antwort(run_id: str, profil: str | None, spec: str | None):
    """Analyse eines frisch angelegten Laufs; bei einem Fehler bleibt der Lauf mit seinen Eingaben stehen (Regel 29)."""
    try:
        with _PDFIUM:                            # die Analyse liest das PDF ein zweites Mal mit pdfium (src/sicht.py)
            plan = runs.analysiere(run_id, profil, spec)
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


def _zahl_arg(name: str) -> int | None:
    """Ganze Zahl >= 0 aus der Adresse; fehlt sie, None (= alles noch offen)."""
    roh = request.args.get(name)
    if roh is None:
        return None
    if not roh.isascii() or not roh.isdigit() or len(roh) > 6:
        raise NutzerFehler(f"{name} muss eine Zahl sein.", 400)
    return int(roh)


@api_bp.get("/lauf/<run_id>/ampel")
def ampel(run_id: str):
    """Dokument-Ampel zum aktuellen Stand der Entscheidungen. ?offen=<unentschiedene Vorschlaege>&stichprobe_offen=<ungeprueft>.
    Die Regeln stehen in src/vertrauen.py und config/vertrauen.yaml; das Dashboard rechnet nichts nach."""
    return jsonify(vertrauen.ampel(runs.lade_plan(run_id), _zahl_arg("offen"), _zahl_arg("stichprobe_offen")))


@api_bp.get("/lauf/<run_id>/specs")
def specs(run_id: str):
    """Specs, die die CDE dieses Laufs enthaelt (Name, Zeilenzahl, von, bis), und die gewaehlte Spec des PDFs.
    Leer, wenn die Excel keine Spalte mit Spezifikationsnummern hat."""
    plan = runs.lade_plan(run_id)
    return jsonify({"specs": plan["specs"], "spec": plan["spec"]})


@api_bp.post("/lauf/<run_id>/neu-analysieren")
def neu_analysieren(run_id: str):
    """Neuer Lauf mit denselben Eingabedateien und einer anderen Spec ({"spec": "11 22 33"}). Der bisherige Lauf bleibt
    unveraendert (Regel 29); die Dateien muessen nicht erneut hochgeladen werden."""
    daten = request.get_json(silent=True)
    if not isinstance(daten, dict) or not isinstance(daten.get("spec"), str) or not daten["spec"].strip():
        return _fehler("Bitte eine Spec angeben.")
    try:
        spec = runs.pruefe_spec(daten["spec"])
        with _PDFIUM:
            _neu, plan = runs.neu_analysieren(run_id, spec)
    except NutzerFehler:
        raise                                     # zentrale Behandlung (app.py)
    except Exception as e:                        # noqa: BLE001
        neu = getattr(e, "run_id", None)
        log.exception("Neue Analyse zu Lauf %s fehlgeschlagen", run_id)
        if neu:
            try:
                runs.markiere_analyse_fehler(neu, _MELDUNG_ANALYSE)
            except Exception:                     # noqa: BLE001
                log.exception("Fehlermarke fuer Lauf %s konnte nicht geschrieben werden", neu)
        return _fehler(_MELDUNG_ANALYSE, 500, run_id=neu, status="fehlgeschlagen")
    return jsonify(_leicht(plan))


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
    # Vertrauensmodus (alles freiwillig): abgelehnt/pauschal = Listen von IDs, stichprobe = {"passt": [...], "falsch": [...]}
    abgelehnt, pauschal = _id_liste(daten, "abgelehnt"), _id_liste(daten, "pauschal")
    if abgelehnt is None or pauschal is None:
        return _fehler("abgelehnt und pauschal müssen Listen von IDs sein.")
    sp = daten.get("stichprobe") or {}
    passt, falsch = _id_liste(sp, "passt"), _id_liste(sp, "falsch")
    if passt is None or falsch is None:
        return _fehler("stichprobe muss die Listen passt und falsch mit IDs enthalten.")
    try:
        res = runs.uebernehmen_lauf(run_id, freig, abwahl, abgelehnt=abgelehnt, pauschal=pauschal,
                                    stichprobe_passt=passt, stichprobe_falsch=falsch)
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
    if plan.get("quelle_art") == "docx":                      # Word hat keine Seiten: es gibt nichts auszuschneiden
        return _fehler("Zu einem Word-Dokument gibt es keinen PDF-Ausschnitt.", 404)
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
