"""Flask-Server der CDE-Aenderungsuebernahme (eigener Port, siehe config/server.yaml).

Bewusste Unterschiede zur bess-pipeline, damit beide parallel laufen koennen:
  * anderer Port (5100 statt 5000), nur auf 127.0.0.1
  * kein CORS, keine Cookies/Sessions (Browser teilen Cookies zwischen Ports)
  * eigener Datenordner (data/), eigene .venv, eigenes Frontend
"""
import os

from flask import Flask

from .. import config
from .routes import api_bp
from .static_routes import static_bp

RELOADER_ENABLED = os.environ.get("CDE_FLASK_RELOAD") == "1"

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024
app.json.ensure_ascii = False
app.register_blueprint(api_bp, url_prefix="/api")
app.register_blueprint(static_bp)


@app.after_request
def kopfzeilen(antwort):
    antwort.headers.setdefault("X-Content-Type-Options", "nosniff")
    antwort.headers.setdefault("Referrer-Policy", "no-referrer")
    return antwort


def start() -> None:
    import socket
    from werkzeug.serving import run_simple

    srv = config.lade("server")
    with socket.socket() as s:
        if s.connect_ex((srv["host"], srv["port"])) == 0:
            raise SystemExit(f"Port {srv['port']} ist bereits belegt (läuft das Dashboard schon?). "
                             "Der Port steht in config/server.yaml.")
    print(f"CDE-Änderungsübernahme: http://{srv['host']}:{srv['port']}/")
    run_simple(srv["host"], srv["port"], app, use_debugger=False, use_reloader=RELOADER_ENABLED, threaded=True)


if __name__ == "__main__":
    start()
