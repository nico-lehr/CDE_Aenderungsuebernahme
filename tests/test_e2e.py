"""Ende-zu-Ende: PDF-Stile x Excel-Layouts -> Ergebnis muss exakt der Loesung entsprechen."""
import pytest

from src.applier import uebernehmen
from src.excel_model import ExcelModell
from src.pdf_reader import analysiere_pdf
from src.planner import plane
from tests.fixtures.szenario import erzeuge, inhalt_zeilen
from tests.helfer import norm


@pytest.mark.parametrize("stil", ["farbe", "hintergrund", "strich", "annot"])
@pytest.mark.parametrize("layout", ["L1", "L2", "L3"])
def test_ergebnis_gleich_loesung(tmp_path, stil, layout):
    d = erzeuge(tmp_path, stil, layout)
    plan = plane(analysiere_pdf(d["pdf"]), ExcelModell(d["alt"]))
    out = tmp_path / "ergebnis.xlsx"
    uebernehmen(plan.aenderungen, d["alt"], out)
    assert norm(inhalt_zeilen(out, layout)) == norm(inhalt_zeilen(d["soll"], layout))
    assert all(a.status == "uebernommen" for a in plan.aenderungen)
    assert all(a.nachkontrolle == "ok" for a in plan.aenderungen)
