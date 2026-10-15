"""Laedt die YAML-Konfiguration (Regel 28: keine hartcodierten Pfade)."""
from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path
from typing import Any

import yaml

PROJEKT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = PROJEKT_ROOT / "config"


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
    if schluessel == "data_dir" and os.environ.get("CDE_DATA_ROOT"):
        roh = Path(os.environ["CDE_DATA_ROOT"])
    if schluessel == "runs_dir" and os.environ.get("CDE_DATA_ROOT"):
        return Path(os.environ["CDE_DATA_ROOT"]) / "runs"
    return roh if roh.is_absolute() else (PROJEKT_ROOT / roh).resolve()
