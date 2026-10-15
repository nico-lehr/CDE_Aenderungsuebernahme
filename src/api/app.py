"""Flask-Server der CDE-Aenderungsuebernahme (eigener Port, siehe config/server.yaml).

Bewusste Unterschiede zur bess-pipeline, damit beide parallel laufen koennen:
  * anderer Port (5100 statt 5000), nur auf Loopback (127.0.0.1 / localhost / ::1)
  * kein CORS, keine Cookies/Sessions (Browser teilen Cookies zwischen Ports)
  * eigener Datenordner (data/), eigene .venv, eigenes Frontend
Sicherheit: nur eigener Host und eigener Origin werden bedient (403 sonst); alle Fehler kommen als
JSON {"fehler": "<einfacher deutscher Satz>"} -- nie als HTML-Seite, nie mit dem Rohtext einer Ausnahme.
"""
import logging
import os
from urllib.parse import urlsplit

from flask import Flask, jsonify, request
from werkzeug.exceptions import HTTPException

from ..fehler import NutzerFehler
from .einstellungen import srv
from .routes import api_bp
from .static_routes import static_bp

RELOADER_ENABLED = os.environ.get("CDE_FLASK_RELOAD") == "1"
LOOPBACK = {"127.0.0.1", "localhost", "::1"}

log = logging.getLogger("cde.api")

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = int(srv()["upload_max_mb"]) * 1024 * 1024
app.json.ensure_ascii = False
app.register_blueprint(api_bp, url_prefix="/api")
app.register_blueprint(static_bp)


# ---------------------------------------------------------------------------------------------
# Host-/Origin-Schutz
# ---------------------------------------------------------------------------------------------
def _name_port(wert: str) -> tuple[str | None, int | None]:
    try:
        t = urlsplit("//" + wert)
        return (t.hostname or None), t.port
    except ValueError:
        return None, None


def _eigener_rechner(name: str | None, port: int | None, port_pflicht: bool) -> bool:
    s = srv()
    erlaubt = {str(h).strip("[]").lower() for h in s["erlaubte_hosts"]}
    if name is None or name.lower() not in erlaubt:
        return False
    if port is None:
        return not port_pflicht
    return port == int(s["port"])


def _json_fehler(text: str, code: int, **extra):
    antwort = jsonify({"fehler": text, **extra})
    antwort.status_code = code
    return antwort


@app.before_request
def schutz_host_und_origin():
    name, port = _name_port(request.host)
    if not _eigener_rechner(name, port, port_pflicht=False):
        return _json_fehler("Diese Anfrage ist nicht erlaubt: Das Dashboard antwortet nur auf diesem Rechner.", 403)
    origin = request.headers.get("Origin")
    if origin is not None:
        try:
            t = urlsplit(origin)
            ok = t.scheme == "http" and _eigener_rechner(t.hostname, t.port, port_pflicht=True)
        except ValueError:
            ok = False
        if not ok:
            return _json_fehler("Diese Anfrage ist nicht erlaubt: Sie kommt nicht vom Dashboard selbst.", 403)
    return None


# ---------------------------------------------------------------------------------------------
# Zentrale Fehlerbehandlung: immer JSON, einfaches Deutsch
# ---------------------------------------------------------------------------------------------
def _http_text(code: int) -> str:
    if code == 413:
        return f"Die hochgeladenen Dateien sind zu groß. Erlaubt sind insgesamt höchstens {srv()['upload_max_mb']} MB."
    return {
        400: "Die Anfrage ist ungültig.",
        403: "Diese Anfrage ist nicht erlaubt.",
        404: "Diese Seite oder Adresse gibt es nicht.",
        405: "Diese Anfrageart ist hier nicht erlaubt.",
        408: "Die Anfrage hat zu lange gedauert. Bitte noch einmal versuchen.",
        415: "Dieses Dateiformat wird hier nicht angenommen.",
        421: "Diese Anfrage ist nicht erlaubt.",
        500: "Es ist ein unerwarteter Fehler aufgetreten. Ihre Daten sind unverändert. Bitte noch einmal versuchen.",
    }.get(code, "Die Anfrage konnte nicht verarbeitet werden.")


@app.errorhandler(HTTPException)
def http_fehler(e: HTTPException):
    code = e.code or 500
    if code < 400:                                   # Weiterleitungen (z. B. Schraegstrich) bleiben wie sie sind
        return e
    antwort = _json_fehler(_http_text(code), code)
    for k, v in e.get_headers():                      # 'Allow' bei 405 behalten
        if k.lower() == "allow":
            antwort.headers[k] = v
    return antwort


@app.errorhandler(NutzerFehler)
def nutzer_fehler(e: NutzerFehler):
    """Gewollte Fehler (auch runs.LaufFehler): ihre Meldung ist schon einfaches Deutsch."""
    return _json_fehler(e.meldung, e.status, status=e.art)


@app.errorhandler(Exception)
def unerwartet(e: Exception):
    log.exception("Unerwarteter Fehler bei %s %s", request.method, request.path)   # Einzelheiten nur im Server-Log
    return _json_fehler(_http_text(500), 500)


@app.after_request
def kopfzeilen(antwort):
    antwort.headers.setdefault("X-Content-Type-Options", "nosniff")
    antwort.headers.setdefault("Referrer-Policy", "no-referrer")
    return antwort


def pruefe_bindung(host: str) -> None:
    """Der Server bindet nur an Loopback; alles andere waere aus dem Netzwerk erreichbar."""
    if str(host).strip("[]").lower() not in LOOPBACK:
        raise SystemExit(f"Der Server darf nur auf diesem Rechner lauschen (127.0.0.1, localhost oder ::1), "
                         f"in config/server.yaml steht aber '{host}'. Bitte den Eintrag 'host' korrigieren.")


def start() -> None:
    import socket
    from werkzeug.serving import run_simple

    s = srv()
    pruefe_bindung(s["host"])
    with socket.socket() as sock:
        if sock.connect_ex((s["host"], s["port"])) == 0:
            raise SystemExit(f"Port {s['port']} ist bereits belegt (läuft das Dashboard schon?). "
                             "Der Port steht in config/server.yaml.")
    print(f"CDE-Änderungsübernahme: http://{s['host']}:{s['port']}/")
    run_simple(s["host"], s["port"], app, use_debugger=False, use_reloader=RELOADER_ENABLED, threaded=True)


if __name__ == "__main__":
    start()
