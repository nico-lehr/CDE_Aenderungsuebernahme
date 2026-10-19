"""src/vorschau.py: pdfium ist nicht thread-sicher -- das Modul rendert nie zwei Seiten gleichzeitig (auch ohne Server)."""
from __future__ import annotations

import threading
import time

from PIL import Image

from src import vorschau


class _FalschesPdf:
    """Steht fuer pdfplumber.open(...): zaehlt, wie viele Aufrufe gleichzeitig im Rendern sind."""
    lock = threading.Lock()
    drin = 0
    max_drin = 0

    def __init__(self, pfad):
        self.pages = [self]

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def to_image(self, resolution):
        cls = type(self)
        with cls.lock:
            cls.drin += 1
            cls.max_drin = max(cls.max_drin, cls.drin)
        time.sleep(0.02)
        with cls.lock:
            cls.drin -= 1
        self.original = Image.new("RGB", (200, 300), "white")
        return self


def test_ausschnitte_aus_mehreren_threads_rendern_nacheinander(monkeypatch):
    monkeypatch.setattr(vorschau.pdfplumber, "open", _FalschesPdf)
    vorschau._seite.cache_clear()
    fehler: list[str] = []

    def hole(i):
        try:
            for k in range(3):
                png = vorschau.ausschnitt(f"erfunden_{i}_{k}.pdf", 1, (10, 20, 50, 30))
                assert png[:4] == b"\x89PNG"
        except Exception as e:                              # noqa: BLE001
            fehler.append(repr(e))
    ts = [threading.Thread(target=hole, args=(i,)) for i in range(6)]
    for t in ts:
        t.start()
    for t in ts:
        t.join(60)
    vorschau._seite.cache_clear()
    assert not fehler and _FalschesPdf.max_drin == 1
    with vorschau.PDFIUM_SPERRE:                            # von aussen gehalten: kein Verklemmen (RLock)
        assert vorschau.ausschnitt("erfunden_x.pdf", 1, (10, 20, 50, 30))[:4] == b"\x89PNG"
    vorschau._seite.cache_clear()
