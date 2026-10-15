"""Laedt die YAML-Konfiguration (Regel 28: keine hartcodierten Pfade)."""
from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path
from typing import Any

import yaml

PROJEKT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = PROJEKT_ROOT / "config"
# Ordner, die bei gesetztem CDE_DATA_ROOT unter diesem Ordner liegen (z.B. fuer Tests)
_UNTER_DATA_ROOT = {"runs_dir": "runs", "paare_dir": "paare", "messung_dir": "messung", "profile_dir": "profile"}


@lru_cache(maxsize=None)
def lade(name: str) -> dict[str, Any]:
    """Liest config/<name>.yaml. Wird gecacht; Tests duerfen lade.cache_clear() rufen."""
    pfad = CONFIG_DIR / f"{name}.yaml"
    with open(pfad, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def pfad(schluessel: str) -> Path:
    """Loest einen Eintrag aus paths.yaml gegen den Projektordner auf."""
    werte = lade("paths")
    roh = Path(werte[schluessel])
    wurzel = os.environ.get("CDE_DATA_ROOT")
    if wurzel and schluessel == "data_dir":
        roh = Path(wurzel)
    if wurzel and schluessel in _UNTER_DATA_ROOT:
        return Path(wurzel) / _UNTER_DATA_ROOT[schluessel]
    return roh if roh.is_absolute() else (PROJEKT_ROOT / roh).resolve()
