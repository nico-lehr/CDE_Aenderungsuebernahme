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
