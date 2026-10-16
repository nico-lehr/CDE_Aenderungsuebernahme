"""Red-Team R4, Teil PDF/PLAN: Lesefehler, die heute still in die Excel durchschlagen (erfundene Daten, Regel 16).

Jeder Test laeuft Ende-zu-Ende (analysiere_pdf -> ExcelModell -> plane -> uebernehmen(..., absaetze=plan.absaetze)).
Pruefregel: "richtig ODER sichtbar unsicher, nie still falsch". Faellt ein Test heute, ist er strikter xfail.
"""
import pytest

from tests.redteam._bau_r4 import (GROESSE, LINKS, PdfDokument, RohDok, aufbau, breite, ende_zu_ende, erkannt_oder_gemeldet,
                                   excel_sections, richtig_oder_unsicher, scheitern, zelle)

Z2 = "Zweiter Absatz mit genug Text darin für den Test."


def _excel(tmp_path, absaetze):
    return excel_sections(tmp_path, [("1", "Eins", absaetze), ("2", "Zwei", [Z2])])


def _elemente(markup):
    return [("h", "1 Eins"), ("p", markup), ("h", "2 Zwei"), ("p", Z2)]


# ---------------------------------------------------------------------------------------------
# RT-200 / RT-201: farbiges Wort + schwarzer Punkt direkt dahinter
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
def test_rt200_blaues_wort_vor_schwarzem_punkt(tmp_path):
    vorher = "Die Steuerung überwacht die Anlage."
    alt = _excel(tmp_path, [vorher])
    l = ende_zu_ende(tmp_path, _elemente("Die Steuerung überwacht die Anlage {+insgesamt+}."), alt)
    betroffen = [a for a in l.aen if "insgesamt" in a.neu]
    richtig_oder_unsicher(l, betroffen, 2, 3, "Die Steuerung überwacht die Anlage insgesamt.", vorher, "RT-200")


@pytest.mark.redteam
def test_rt201_rotes_wort_vor_schwarzem_punkt(tmp_path):
    vorher = "Die Steuerung überwacht die Anlage insgesamt."
    alt = _excel(tmp_path, [vorher])
    l = ende_zu_ende(tmp_path, _elemente("Die Steuerung überwacht die Anlage {-insgesamt-}."), alt)
    betroffen = [a for a in l.aen if "insgesamt" in a.alt]
    richtig_oder_unsicher(l, betroffen, 2, 3, "Die Steuerung überwacht die Anlage.", vorher, "RT-201")


# ---------------------------------------------------------------------------------------------
# RT-202: schwarze Durchstreichlinie ueber ganze Zeilen eines Absatzes (>= 70 % der Seitenbreite)
# ---------------------------------------------------------------------------------------------
ERSTER = "Der erste Absatz bleibt unverändert stehen und beschreibt die Anlage."
ZEILE1 = "Die Wartung der Lüftung erfolgt zweimal im Jahr durch den Betreiber der gesamten Anlage"
ZEILE2 = "und wird im Wartungsbuch mit Datum, Uhrzeit und Unterschrift sorgfältig dokumentiert."


@pytest.mark.redteam
def test_rt202_strich_ueber_ganze_zeilen(tmp_path):
    d = RohDok()
    d.ueberschrift("1 Eins")
    d.absatz([[(ERSTER, "k")]])
    kaesten = d.absatz([[(ZEILE1, "k")], [(ZEILE2, "k")]])
    for x0, top, x1, _ in kaesten:
        aufbau(x1 - x0 >= 0.7 * 595, f"Linie zu kurz ({x1 - x0:.0f} pt)")
        d.linie(x0, x1, top + GROESSE * 0.5)          # Zeichenmitte, schwarz, 0,8 pt
    d.ueberschrift("2 Zwei")
    d.absatz([[(Z2, "k")]])
    vorher = ERSTER + "\n" + ZEILE1 + " " + ZEILE2
    alt = _excel(tmp_path, [ERSTER, ZEILE1 + " " + ZEILE2])
    aufbau(zelle(alt, 2, 3) == vorher, "Excel-Aufbau")
    l = ende_zu_ende(tmp_path, d, alt)
    erkannt_oder_gemeldet(l, 2, 3, ERSTER, vorher, "RT-202", passt=lambda a: "Wartung" in a.alt)


# ---------------------------------------------------------------------------------------------
# RT-206: blauer neuer Satz ueber einen Seitenumbruch
# ---------------------------------------------------------------------------------------------
ALT7 = "Die Anlage wird regelmäßig geprüft und alle Ergebnisse werden schriftlich festgehalten."
NEU7 = "Zusätzlich wird jede einzelne Prüfung im System vermerkt und vom Betreiber unterschrieben."


@pytest.mark.redteam
def test_rt206_neuer_satz_ueber_seitenumbruch(tmp_path):
    d = PdfDokument(variante="farbe", kopf_fuss=False)
    d.ueberschrift("1 Eins")
    d.block(ALT7 + " {+" + NEU7 + "+}", umbruch_nach_zeile=1)
    aufbau(d.letzter_block == (1, 2), f"Absatz muss ueber zwei Seiten laufen: {d.letzter_block}")
    d.ueberschrift("2 Zwei")
    d.block(Z2)
    alt = _excel(tmp_path, [ALT7])
    l = ende_zu_ende(tmp_path, d, alt)
    seiten = {t.page for b in l.pdf.blocks for t in b.tokens if t.new and t.new in NEU7.split() and t.status == "add"}
    aufbau(seiten == {1, 2}, f"der blaue Satz muss auf beiden Seiten stehen: {seiten}")
    betroffen = [a for a in l.aen if a.neu and set(a.neu.split()) <= set(NEU7.split())]
    richtig_oder_unsicher(l, betroffen, 2, 3, ALT7 + " " + NEU7, ALT7, "RT-206")


# ---------------------------------------------------------------------------------------------
# RT-214: Fussnoten-Hochzahl am geaenderten Wort
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
def test_rt214_hochzahl_am_geaenderten_wort(tmp_path):
    d = RohDok()
    d.ueberschrift("1 Eins")
    d.absatz([[("Die Steuerung meldet die Werte an die ", "k"), ("Leitstelle", "1 0 0 rg"), (" ", "k"),
               ("Zentrale", "0 0 1 rg"), ("1", "0 0 1 rg", {"groesse": 7, "hoch": 4}), (" weiter.", "k")]])
    d.ueberschrift("2 Zwei")
    d.absatz([[(Z2, "k")]])
    d.absatz([[("1 Die Zentrale ist erfunden.", "k", {"groesse": 8})]])     # Fussnote
    vorher = "Die Steuerung meldet die Werte an die Leitstelle weiter."
    alt = _excel(tmp_path, [vorher])
    l = ende_zu_ende(tmp_path, d, alt)
    betroffen = [a for a in l.aen if "Zentrale" in a.neu or "Leitstelle" in a.alt]
    richtig_oder_unsicher(l, betroffen, 2, 3, "Die Steuerung meldet die Werte an die Zentrale weiter.", vorher, "RT-214")


# ---------------------------------------------------------------------------------------------
# RT-215: Silbentrennung im NEUEN Text: "E-" am Zeilenende, "Mail" in der naechsten Zeile
# ---------------------------------------------------------------------------------------------
ANFANG = "Die Steuerung meldet alle Störungen der gesamten Anlage sofort und zuverlässig "


@pytest.mark.redteam
def test_rt215_trennung_e_mail_im_neuen_text(tmp_path):
    d = RohDok()
    d.ueberschrift("1 Eins")
    kaesten = d.absatz([[(ANFANG, "k"), ("per E-", "0 0 1 rg")], [("Mail", "0 0 1 rg"), (" an die Leitstelle weiter.", "k")]])
    aufbau(kaesten[0][2] - LINKS > 0.7 * 595, "erste Zeile muss fast voll sein (sonst ist sie kein Fliesstext)")
    d.ueberschrift("2 Zwei")
    d.absatz([[(Z2, "k")]])
    vorher = ANFANG + "an die Leitstelle weiter."
    alt = _excel(tmp_path, [vorher])
    l = ende_zu_ende(tmp_path, d, alt)
    betroffen = [a for a in l.aen if "Mail" in a.neu]
    ergebnis = richtig_oder_unsicher(l, betroffen, 2, 3, ANFANG + "per E-Mail an die Leitstelle weiter.", vorher, "RT-215")
    if ergebnis == "richtig" and "EMail" in str(zelle(l.out, 2, 3)):
        scheitern("RT-215: Bindestrich verloren ('EMail')")
