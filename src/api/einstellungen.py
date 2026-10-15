"""Server-Einstellungen aus config/server.yaml (Regel 28). Fehlende Schluessel einer aelteren
server.yaml bekommen einen Standardwert, damit das Tool nicht an einer Kleinigkeit scheitert."""
from __future__ import annotations

from typing import Any

from .. import config

_STANDARD: dict[str, Any] = {
    "host": "127.0.0.1",
    "port": 5100,
    "erlaubte_hosts": ["127.0.0.1", "localhost", "::1"],
    "upload_max_mb": 200,
    "upload_max_mb_datei": 100,
    "dateiname_max_laenge": 255,
}


def srv() -> dict[str, Any]:
    return {**_STANDARD, **config.lade("server")}
