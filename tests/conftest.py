"""Gemeinsame Test-Einstellungen: Laeufe landen in einem Temp-Ordner, nie im echten data/."""
import pytest


@pytest.fixture(autouse=True)
def _daten_ordner(tmp_path, monkeypatch):
    monkeypatch.setenv("CDE_DATA_ROOT", str(tmp_path / "data"))
    yield
