"""Gegenprobe (src/gegenprobe.py, ADR 0004): Zustand eines Absatzes, Normalform, Vergleich, A3-Pruefung.
Nur erfundene Woerter."""
import random

from src.gegenprobe import gleich, neue_spanne, normalform, pruefe_absatz, vergleiche, zustand, zustand_mit_optional
from src.models import Absatz, Lauf, Operation


def L(spec: str) -> list[Lauf]:
    """Kurzschrift: Woerter mit Leerzeichen; {-x-} entfaellt (Aenderung D), {+x+} neu (Aenderung N); '|' klebt ohne Trenner."""
    out: list[Lauf] = []
    for i, teil in enumerate(spec.split(" ")):
        if i:
            out.append(Lauf("s", " "))
        for stueck in teil.split("|"):
            if stueck.startswith("{-"):
                out.append(Lauf("d", stueck[2:-2], "D"))
            elif stueck.startswith("{+"):
                out.append(Lauf("a", stueck[2:-2], "N"))
            elif stueck:
                out.append(Lauf("k", stueck))
    return out


def test_zustand_satzzeichen_bleibt_beim_einfuegen_und_loeschen():
    l = L("Die Anlage {+insgesamt+}|.")
    assert normalform(zustand(l))[0] == "Die Anlage."
    assert zustand(l, {"N"}) == "Die Anlage insgesamt."
    l = L("Die Anlage {-insgesamt-}|.")
    assert normalform(zustand(l, {"D"}))[0] == "Die Anlage."
    assert zustand(l) == "Die Anlage insgesamt."


def test_zustand_ersetzung_rand_und_verschmolzen():
    assert zustand(L("darf {-12-} {+7+} ms"), {"D", "N"}) == "darf 7 ms"
    assert zustand(L("darf {-12-} {+7+} ms")) == "darf 12 ms"
    assert zustand(L("{-Zudem-} gilt"), {"D"}) == "gilt"
    assert zustand(L("x {-alt-}|{+neu+} y"), {"D", "N"}) == "x neu y"
    assert zustand(L("Haus|{-tuer-}|schloss"), {"D"}) == "Hausschloss"
    assert zustand(L("Wort|{-,-} weiter"), {"D"}) == "Wort weiter"


def test_trenner_rang_geschuetztes_leerzeichen_bleibt():
    l = [Lauf("k", "20"), Lauf("s", " "), Lauf("d", "000", "D"), Lauf("s", " "), Lauf("k", "kWh")]
    assert zustand(l, {"D"}) == "20 kWh"


def test_normalform_streng_aber_fair():
    assert normalform("„Hallo“  Welt –  ok .")[0] == '"Hallo" Welt - ok.'
    assert normalform("Text­ mit\nUmbruch und ( Klammer )")[0] == "Text mit Umbruch und (Klammer)"
    assert normalform("• Punkt eins")[0] == "Punkt eins"
    assert normalform("Fläche m² ist groß¹")[0] == "Fläche m ist groß"
    assert normalform("Gross")[0] != normalform("gross")[0]                 # Gross/Klein zaehlt
    s, m = normalform("a  b")
    assert s == "a b" and m == [0, 1, 3]


def test_vergleiche_meldet_art_der_abweichung_ohne_inhalt():
    assert vergleiche("Messwert drei", "Messwert vier").art == "buchstabe"
    assert vergleiche("Wert 3", "Wert 4").art == "ziffer"
    assert vergleiche("Wert,", "Wert;").art == "satzzeichen"
    assert vergleiche("Wert", "wert").art == "gross_klein"
    assert vergleiche("Wert eins", "Wert").art == "pdf_laenger"
    assert vergleiche("Wert", "Wert eins").art == "excel_laenger"


def test_optionaler_trennstrich():
    l = [Lauf("k", "E"), Lauf("k", "-", optional=True), Lauf("k", "Mail")]
    text, opt = zustand_mit_optional(l)
    assert text == "E-Mail" and opt == frozenset({1})
    assert gleich(text, "EMail", opt) and gleich(text, "E-Mail", opt)
    assert not gleich(text, "E Mail")


def _op(start, ende, text, ae):
    return Operation("ersetzen", "CDE", 2, 3, start, ende, text, aenderung=ae)


def test_pruefe_absatz_a3_ok_und_falsch():
    orig = "Erster Satz.\nDie Anlage arbeitet.\nLetzter Satz."
    s = orig.index("Die")
    e = orig.index("Letzter") - 1
    ab = Absatz("P1", "CDE", [(2, 3, s, e)], L("Die Anlage arbeitet {+leise+}|."))
    pos = orig.index(".", s)
    op = _op(pos, pos, " leise", "N")
    gesp = orig[:pos] + " leise" + orig[pos:]
    assert pruefe_absatz(ab, {(2, 3): (orig, gesp, [op])}, {"N"}) is None
    falsch = orig[:pos] + ". leise" + orig[pos + 1:]          # "…arbeitet. leise" -- der Fehler aus der Analyse
    assert pruefe_absatz(ab, {(2, 3): (orig, falsch, [op])}, {"N"}) == "gegenprobe_text"
    # nichts angewendet: der alte Absatz muss unveraendert dastehen
    assert pruefe_absatz(ab, {(2, 3): (orig, orig, [])}, set()) is None


def test_pruefe_absatz_aussen_veraendert_wird_erkannt():
    orig = "Erster Satz.\nDie Anlage arbeitet."
    s = orig.index("Die")
    ab = Absatz("P1", "CDE", [(2, 3, s, len(orig))], L("Die {-Anlage-} arbeitet."))
    a = orig.index("Anlage")
    op = _op(a - 1, a + len("Anlage"), "", "D")
    gesp = "Erster Satz!\nDie arbeitet."                      # ausserhalb des Absatzes veraendert
    assert pruefe_absatz(ab, {(2, 3): (orig, gesp, [op])}, {"D"}) == "gegenprobe_aussen"
    gut = "Erster Satz.\nDie arbeitet."
    assert pruefe_absatz(ab, {(2, 3): (orig, gut, [op])}, {"D"}) is None


def test_neue_spanne_fremde_operation_davor_verschiebt():
    ops = [_op(0, 5, "Ein", "X"), _op(12, 12, " neu", "N")]
    assert neue_spanne(10, 20, ops, {"N"}) == (8, 22)


def test_absatz_rundreise_dict():
    ab = Absatz("P1", "CDE", [(2, 3, 0, 9)], L("a {-b-} {+c+}") + [Lauf("k", "-", optional=True)])
    zurueck = Absatz.aus_dict(ab.zu_dict())
    assert zurueck == ab


def test_zufall_jeder_teilzustand_stimmt_mit_naiver_berechnung():
    """Zufaellige Laeufe: zustand(S) == naiv gebauter Text (Woerter nach Status, ein Leerzeichen dazwischen) in Normalform."""
    rnd = random.Random(7)
    woerter = ["Ventil", "Pumpe", "Druck", "Wert", "neu", "alt", "Anlage", "Prüfung"]
    for _ in range(400):
        laeufe: list[Lauf] = []
        eintraege: list[tuple[str, str, str]] = []           # (art, text, id)
        for i in range(rnd.randint(1, 9)):
            art = rnd.choice("kkkda")
            ae = f"A{rnd.randint(1, 3)}" if art in "da" else ""
            if laeufe:
                laeufe.append(Lauf("s", rnd.choice([" ", " ", " "])))
            w = rnd.choice(woerter)
            laeufe.append(Lauf(art, w, ae))
            eintraege.append((art, w, ae))
        ids = {e[2] for e in eintraege if e[2]}
        for k in range(len(ids) + 1):
            aktiv = set(rnd.sample(sorted(ids), k)) if ids else set()
            naiv = " ".join(w for art, w, ae in eintraege
                            if art == "k" or (art == "d" and ae not in aktiv) or (art == "a" and ae in aktiv))
            assert normalform(zustand(laeufe, aktiv))[0] == normalform(naiv)[0]
