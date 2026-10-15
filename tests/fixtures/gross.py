"""Erzeugt ein langes, erfundenes Dokument (mehrere Seiten) mit verstreuten Aenderungen."""
from __future__ import annotations

import random

WOERTER = ("Anlage Betrieb Daten Leistung Spannung Kapazität System Modul Rack Zelle Sicherheit Prüfung Lieferung Wartung "
           "Dokument Anforderung Schnittstelle Steuerung Messung Temperatur Kühlung Gehäuse Wandler Schutz Netz Anschluss "
           "Kabel Software Protokoll Meldung Alarm Zugriff Standort Montage Transport Verpackung Garantie Dauer Frist "
           "muss soll kann darf wird werden ist sind für mit ohne nach vor bei durch gegen über unter zwischen "
           "jeder alle eine ein der die das dem den des und oder aber sowie bzw").split()


def satz(rng: random.Random, n: int) -> str:
    w = [rng.choice(WOERTER) for _ in range(n)]
    w[0] = w[0].capitalize()
    return " ".join(w) + "."


def erzeuge_gross(seed: int = 7, sections: int = 12, absaetze: int = 5, aenderungen: int = 14):
    rng = random.Random(seed)
    el: list[tuple[str, str]] = []
    absatz_ids: list[int] = []
    for s in range(1, sections + 1):
        el.append(("h", f"{s} Abschnitt {rng.choice(WOERTER)}{s}"))
        for a in range(absaetze):
            text = " ".join(satz(rng, rng.randint(8, 14)) for _ in range(rng.randint(2, 3)))
            # eindeutige Marke je Absatz, damit Absaetze sich garantiert unterscheiden
            text = f"Punkt{s}x{a} " + text
            absatz_ids.append(len(el))
            el.append(("p", text))
    # Aenderungen verstreuen
    ziele = rng.sample(absatz_ids, aenderungen)
    for i, idx in enumerate(ziele):
        art, text = el[idx]
        worte = text.split()
        k = i % 4
        if k == 0:       # Wort ersetzen
            j = rng.randint(3, len(worte) - 3)
            worte[j] = "{-" + worte[j] + "-} {+" + rng.choice(WOERTER) + "X" + "+}"
        elif k == 1:     # Satz einfuegen am Ende
            worte.append("{+" + satz(rng, 7) + "+}")
        elif k == 2:     # Satzteil loeschen
            j = rng.randint(3, len(worte) - 6)
            worte[j:j + 3] = ["{-" + " ".join(worte[j:j + 3]) + "-}"]
        else:            # ganzer Absatz geloescht
            worte = ["{-" + " ".join(worte) + "-}"]
        el[idx] = (art, " ".join(worte))
    return el
