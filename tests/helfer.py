"""Gemeinsame Hilfen fuer die Tests (nur erfundene Daten, Regel 16)."""
import re
from pathlib import Path

from src.applier import uebernehmen
from src.excel_model import ExcelModell
from src.pdf_reader import analysiere_pdf
from src.planner import plane
from src.protokoll import bilanz
from tests.fixtures.pdf_writer import baue_pdf
from tests.fixtures.szenario import baue_excel, inhalt_zeilen, zustaende


def norm(rows):
    return [tuple(re.sub(r"\s+", " ", c).strip() for c in r) for r in rows]


def durchlauf(ordner: Path, elemente, layout="L1", stil="farbe", alt_hook=None):
    """PDF + alte Excel bauen, Tool laufen lassen, Ergebnis mit der Loesung vergleichen."""
    ordner.mkdir(parents=True, exist_ok=True)
    baue_pdf(ordner / "a.pdf", elemente, stil)
    alt, neu = zustaende(elemente, False), zustaende(elemente, True)
    if alt_hook:
        alt = alt_hook(alt)
    baue_excel(ordner / "alt.xlsx", alt, layout)
    baue_excel(ordner / "soll.xlsx", neu, layout)
    pdf = analysiere_pdf(ordner / "a.pdf")
    plan = plane(pdf, ExcelModell(ordner / "alt.xlsx"))
    uebernehmen(plan.aenderungen, ordner / "alt.xlsx", ordner / "out.xlsx")
    bil = bilanz(plan.aenderungen, plan.markierte_woerter, plan.ignoriert)
    gleich = norm(inhalt_zeilen(ordner / "out.xlsx", layout)) == norm(inhalt_zeilen(ordner / "soll.xlsx", layout))
    return plan, bil, gleich
