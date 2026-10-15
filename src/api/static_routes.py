"""Auslieferung der Frontend-Dateien (Positivliste wie in der bess-pipeline)."""
from pathlib import Path

from flask import Blueprint, abort, send_from_directory

from .. import config

static_bp = Blueprint("frontend_assets", __name__)

ASSET_MIMETYPES = {
    ".js": "application/javascript",
    ".css": "text/css",
    ".png": "image/png",
    ".ico": "image/x-icon",
    ".svg": "image/svg+xml",
}


@static_bp.route("/")
def index():
    return send_from_directory(str(config.pfad("frontend_dir")), "index.html", mimetype="text/html")


@static_bp.route("/<path:filename>")
def serve_frontend_asset(filename: str):
    suffix = Path(filename).suffix.lower()
    if suffix not in ASSET_MIMETYPES:
        abort(404)
    return send_from_directory(str(config.pfad("frontend_dir")), filename, mimetype=ASSET_MIMETYPES[suffix])
