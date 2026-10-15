"""Gewollte Fehler mit einer Meldung, die schon einfaches Deutsch ist und dem Nutzer gezeigt werden darf.

Nur diese Klassen duerfen ihren Text an den Client (API, CLI, analyse_fehler.json) weitergeben. Jede andere
Ausnahme (auch ein ValueError aus einer Bibliothek) gilt als unerwartet: der Nutzer sieht nur eine allgemeine Meldung,
der Rohtext geht ausschliesslich ins Server-Log.

Beide Klassen sind auch ValueError bzw. FileNotFoundError, damit aelterer Code (`except ValueError`) weiter greift.
"""
from __future__ import annotations


class NutzerFehler(ValueError):
    """Ein gewollter Fehler. `meldung` ist ein einfacher deutscher Satz fuer den Nutzer, `status` der HTTP-Code."""

    def __init__(self, meldung: str, status: int = 400, art: str = "fehler"):
        super().__init__(meldung)
        self.meldung = meldung
        self.status = status
        self.art = art


class NichtGefunden(NutzerFehler, FileNotFoundError):
    """Ein Lauf oder eine Datei gibt es nicht (404). Ist zugleich ein FileNotFoundError."""

    def __init__(self, meldung: str, status: int = 404, art: str = "nicht_gefunden"):
        super().__init__(meldung, status, art)


class DateiVorhanden(NutzerFehler, FileExistsError):
    """Eine Zieldatei gibt es schon und wird nie ueberschrieben (Regel 29). Ist zugleich ein FileExistsError."""

    def __init__(self, meldung: str, status: int = 409, art: str = "vorhanden"):
        super().__init__(meldung, status, art)
