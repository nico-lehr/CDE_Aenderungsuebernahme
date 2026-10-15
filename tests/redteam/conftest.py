"""Registriert die Marker der Red-Team-Tests (pyproject.toml bleibt unveraendert)."""


def pytest_configure(config):
    config.addinivalue_line("markers", "redteam: Red-Team-Test, der einen bekannten Fehler des Tools belegt (erwartet: xfail)")
