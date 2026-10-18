"""Gemeinsame Test-Einstellungen: Laeufe landen in einem Temp-Ordner, nie im echten data/."""
import pytest

from src import config


@pytest.fixture(autouse=True)
def _daten_ordner(tmp_path, monkeypatch):
    monkeypatch.setenv("CDE_DATA_ROOT", str(tmp_path / "data"))
    yield


@pytest.fixture(autouse=True)
def _markierung_aus(request, monkeypatch):
    """Die frueheren Tests pruefen die Reinigungs-Logik (alter Text wird ersetzt). Die farbige Markierung (ADR 0005,
    Standard in der Produktion) ist dort aus; Tests mit @pytest.mark.markierung_an laufen mit dem Standard."""
    if request.node.get_closest_marker("markierung_an") is None:
        monkeypatch.setitem(config.lade("zuordnung")["markierung"], "aenderungen", "keine")
    yield


@pytest.fixture(autouse=True)
def _legende_alt(request, monkeypatch):
    """Die frueheren Tests und Korpora nutzen die alte Legende (blau = neu, rot = entfaellt). Tests mit
    @pytest.mark.legende_neu laufen mit dem Standard aus config/farben.yaml (ADR 0006: Linie entscheidet)."""
    if request.node.get_closest_marker("legende_neu") is None:
        monkeypatch.setitem(config.lade("farben")["legende"], "modus", "alt")
    yield


@pytest.fixture(autouse=True)
def _schreiber(request):
    """ADR 0007: Standard ist der verlustfreie Patch-Schreiber. Tests mit @pytest.mark.schreiber_openpyxl pruefen das
    fruehere Speichern mit openpyxl (Kommentar-/Gelb-Markierung, Zeilen entfernen, alte Einfuege-Risiken, keine farbige
    Markierung). Eigenes MonkeyPatch-Objekt: ein monkeypatch.undo() im Test setzt diese Weiche nicht zurueck."""
    mp = pytest.MonkeyPatch()
    if request.node.get_closest_marker("schreiber_openpyxl") is not None:
        mp.setitem(config.lade("zuordnung")["anwenden"], "schreiber", "openpyxl")
    yield
    mp.undo()


# Tests in Dateien ausserhalb von Auftrag 003 / SCHREIBEN, die fachlich das fruehere Speichern pruefen. Die Weiche steht hier,
# bis der Marker im Test selbst gesetzt wird (Hinweis an den Orchestrator).
_SCHREIBER_OPENPYXL = {
    # messen erwartet, dass die Vorschlags-Excel Bereiche NICHT mitzieht (openpyxl) und der Plan das vorab meldet
    "tests/test_messen.py::test_ende_zu_ende_nicht_mitgezogenes_wird_in_der_vorschlags_excel_als_angekuendigt_gezaehlt",
    # nach monkeypatch.undo() gilt im Test der Produktions-Standard (Markierung farbe); messen liest markierte Zellen
    # noch nicht als Neu-Ansicht -- mit openpyxl wird wie frueher ohne Markierung ersetzt
    "tests/test_messen.py::test_zahl_wird_text_an_unveraenderter_zelle_in_verwalteter_spalte_ist_stiller_fehler",
}


def pytest_collection_modifyitems(items):
    for item in items:
        if item.nodeid in _SCHREIBER_OPENPYXL:
            item.add_marker(pytest.mark.schreiber_openpyxl)
