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
