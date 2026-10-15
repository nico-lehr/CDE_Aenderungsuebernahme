"""Randfaelle: keine Aenderungen, kaputte Dateien, alte Excel-Formate."""
import pytest

from src.excel_model import ExcelModell
from src.pdf_reader import analysiere_pdf
from src.planner import plane
from tests.fixtures.pdf_writer import baue_pdf
from tests.fixtures.szenario import baue_excel, zustaende


def test_pdf_ohne_aenderungen(tmp_path):
    el = [("h", "1 Eins"), ("p", "Alles bleibt wie es ist.")]
    baue_pdf(tmp_path / "a.pdf", el, "farbe")
    baue_excel(tmp_path / "x.xlsx", zustaende(el, False), "L1")
    plan = plane(analysiere_pdf(tmp_path / "a.pdf"), ExcelModell(tmp_path / "x.xlsx"))
    assert plan.aenderungen == []


def test_kaputtes_pdf(tmp_path):
    (tmp_path / "kaputt.pdf").write_bytes(b"%PDF-1.4 das ist kein echtes PDF")
    with pytest.raises(Exception):
        analysiere_pdf(tmp_path / "kaputt.pdf")


def test_kaputte_excel(tmp_path):
    (tmp_path / "kaputt.xlsx").write_bytes(b"PK nicht wirklich")
    with pytest.raises(Exception):
        ExcelModell(tmp_path / "kaputt.xlsx")
