"""Server-Einstellungen aus config/server.yaml (Regel 28).

`host` und `port` stehen NUR in der YAML-Datei (keine festen Werte im Code): fehlen sie, bricht der Start mit einer
klaren Meldung ab. Fehlende Upload-Grenzen und die Liste der erlaubten Rechnernamen einer aelteren server.yaml
bekommen einen Standardwert, damit das Tool nicht an einer Kleinigkeit scheitert."""
from __future__ import annotations

from typing import Any

from .. import config

_ERFORDERLICH = ("host", "port")

_STANDARD: dict[str, Any] = {
    "erlaubte_hosts": ["127.0.0.1", "localhost", "::1"],
    "upload_max_mb": 200,
    "upload_max_mb_datei": 100,
    "dateiname_max_laenge": 255,
}


def srv() -> dict[str, Any]:
    werte = config.lade("server")
    fehlend = [k for k in _ERFORDERLICH if werte.get(k) in (None, "")]
    if fehlend:
        raise SystemExit("Der Server kann nicht starten: In config/server.yaml fehlt der Eintrag "
                         + " und ".join(f"'{k}'" for k in fehlend)
                         + ". Bitte ergänzen, zum Beispiel 'host: 127.0.0.1' und 'port: 5100'.")
    return {**_STANDARD, **werte}
