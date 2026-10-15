"""Rendert den Ausschnitt einer PDF-Seite mit Markierung der Aenderung (fuer das Dashboard)."""
from __future__ import annotations

import io
from functools import lru_cache
from pathlib import Path

import pdfplumber
from PIL import ImageDraw

_DPI = 110


@lru_cache(maxsize=24)
def _seite(pfad: str, nr: int):
    with pdfplumber.open(pfad) as pdf:
        return pdf.pages[nr - 1].to_image(resolution=_DPI).original.convert("RGB")


def ausschnitt(pdf: str | Path, seite: int, bbox: tuple[float, float, float, float], farbe: str = "#0071E3",
               rand_pt: float = 70.0) -> bytes:
    bild = _seite(str(pdf), seite).copy()
    f = _DPI / 72.0
    x0, top, x1, bot = (v * f for v in bbox)
    d = ImageDraw.Draw(bild, "RGBA")
    d.rectangle((x0 - 3, top - 3, x1 + 3, bot + 3), outline=farbe, width=3)
    rand = rand_pt * f
    ausschnitt_ = bild.crop((0, max(0, int(top - rand)), bild.width, min(bild.height, int(bot + rand))))
    out = io.BytesIO()
    ausschnitt_.save(out, "PNG")
    return out.getvalue()
