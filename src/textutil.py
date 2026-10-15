"""Kleine Text-Helfer: Normalisieren und Zerlegen in Woerter (Tokens)."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

_RAND = "\"'„“”‚‘’«»()[]{}<>.,;:!?*•·"
_UNSICHTBAR = dict.fromkeys(map(ord, "­​‌‍﻿"), None)
_WS = re.compile(r"\s+")
_NUMMER = re.compile(r"^(\d{1,3}(?:\.\d{1,3}){0,6})\.?$")
_KOPF = re.compile(r"^(\d{1,3}(?:\.\d{1,3}){0,6})\.?\s+(\S.*)$", re.S)


def norm(text: str) -> str:
    """Vergleichsform eines Wortes: klein, ohne Randzeichen, ohne unsichtbare Zeichen."""
    t = unicodedata.normalize("NFKC", text).translate(_UNSICHTBAR).casefold()
    t = t.replace("–", "-").replace("−", "-").replace(" ", " ")
    return t.strip(_RAND + " ")


def norm_text(text: str) -> str:
    return " ".join(norm(w) for w in _WS.split(text.strip()) if norm(w))


@dataclass(frozen=True)
class Wort:
    text: str
    start: int
    end: int

    @property
    def norm(self) -> str:
        return norm(self.text)


def woerter(text: str) -> list[Wort]:
    """Zerlegt einen Text in Woerter mit Zeichenposition (Wort = Block ohne Leerraum)."""
    return [Wort(m.group(), m.start(), m.end()) for m in re.finditer(r"\S+", text)]


def section_nummer(text: str) -> str | None:
    """'3.2' oder '3.2.' -> '3.2'; sonst None."""
    m = _NUMMER.match(text.strip())
    return m.group(1) if m else None


def kopf_zerlegen(text: str) -> tuple[str, str] | None:
    """'3.2 Titel' -> ('3.2', 'Titel'); sonst None."""
    m = _KOPF.match(text.strip())
    return (m.group(1), m.group(2).strip()) if m else None


def tiefe(nummer: str | None) -> int:
    return nummer.count(".") + 1 if nummer else 0
