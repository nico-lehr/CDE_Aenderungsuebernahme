"""Red-Team R3 (Analyse der stillen Fehler im Entwicklungskorpus), Teil 2: MESSARTEFAKTE.

16 der 43 "stillen Fehler" sind KEINE Fehler des Tools, sondern Zaehlfehler des Messwerkzeugs (src/messen.py). Dazu kommen
zwei Verzerrungen der Lese-Luecke (RT-131/132). Jeder Test hat den gleichen Aufbau:

  1. `aufbau(...)` beweist (als RuntimeError bei Abweichung), dass das TOOL in diesem Fall nichts Falsches getan hat, und
  2. ein AssertionError zeigt, was die Messung falsch zaehlt. Besitzer ist deshalb der Mess-Bauer (messen.py).

Vorschlaege, wie `messen.py` richtig zaehlen muesste, stehen in tests/redteam/README_korpus.md (KEIN Eingriff in src/).
Alles ERFUNDEN (Regel 16); Seeds < 1000 (Regel 11).
"""
from __future__ import annotations

from collections import Counter
from unittest import mock

import pytest

from src import messen
from src.models import PRUEFEN, UEBERNOMMEN
from tests.fixtures.pdf_writer import baue_pdf
from tests.fixtures.szenario import baue_excel, zustaende
from tests.redteam._bau import aufbau, scheitern
from tests.redteam._bau_korpus import (KOPF, auto_lauf, baue_basis_paar, messe_mit_fertigem_ergebnis, messung,
                                       nur_durch_offene_aenderungen_verschieden, stumm_bei, zeile_im_ergebnis,
                                       zelle_vorher_soll_ist, paar)

# ==================================================================================================
# MA-1 (8 Zellen): In einer Zelle steht neben einer SICHEREN, richtig angewendeten Aenderung eine OFFENE Aenderung
#      ("pruefen"). Die Zelle entspricht dem Soll erst, wenn der Mensch die offene Aenderung freigibt. Die Messung vergleicht
#      die GANZE Zelle mit dem Soll und zaehlt die richtig angewendete Aenderung als "uebernommen, aber falsch".
# ==================================================================================================
_MA1 = [("h", "1 Eins"), ("p", "Der Nachweis erfolgt durch ein Messprotokoll der Abnahme."),
        ("h", "2 Zwei"), ("p", "Die Reaktionszeit darf 150 ms nicht ueberschreiten."),
        ("h", "3 Drei"), ("p", "Die Wartung erfolgt {-jaehrlich-} {+monatlich+} durch den Betreiber und {-den-} {+das+} Personal."),
        ("h", "4 Vier"), ("p", "Alle Steckverbindungen sind eindeutig beschriftet.")]


@pytest.mark.redteam
def test_rt125_offene_aenderung_in_derselben_zelle_ist_kein_stiller_fehler(tmp_path):
    baue_pdf(tmp_path / "aenderungen.pdf", _MA1, "farbe")
    baue_excel(tmp_path / "vorher.xlsx", zustaende(_MA1, False), "L1", True)
    baue_excel(tmp_path / "nachher.xlsx", zustaende(_MA1, True), "L1", True)
    # Kontrolle: ohne Eingriff macht das Tool beide Aenderungen richtig, die Messung findet nichts
    sauber = messen.messe_paar(tmp_path, False)
    aufbau(sauber["stille_fehler"] == 0 and sauber["auto_richtig"] == sauber["auto_gesamt"] == 1,
           f"Ausgangslage nicht sauber: {sauber['stille_fehler']} stille Fehler")
    # Eingriff: die zweite Aenderung der Zelle bleibt offen, genau wie es das Tool bei unsicheren Treffern tut
    echt = messen.plane

    def plane_mit_offener(pdf, xl, *a, **k):
        p = echt(pdf, xl, *a, **k)
        p.aenderungen[-1].status, p.aenderungen[-1].konfidenz = PRUEFEN, "mittel"
        return p

    with mock.patch.object(messen, "plane", plane_mit_offener):
        res = messen.messe_paar(tmp_path, True)
    aufbau(res["aenderungen_nach_status"].get(PRUEFEN, 0) == 1 and res["aenderungen_nach_status"].get(UEBERNOMMEN, 0) == 1,
           "Erwartet: eine uebernommene und eine offene Aenderung")
    aufbau(res["vorschlaege_gesamt"] == 1 and res["vorschlaege_richtig"] == 1,
           "Mit Freigabe der offenen Aenderung muss die Zelle dem Soll entsprechen")
    if res["stille_fehler"]:
        scheitern(f"{res['stille_fehler']} stiller Fehler {res['stille_fehler_nach_art']}, obwohl die angewendete Aenderung "
                  f"richtig ist und die offene Aenderung sichtbar auf 'pruefen' steht (Treffergenauigkeit "
                  f"{res['auto_richtig']} von {res['auto_gesamt']})")


# --- RT-126: derselbe Fall mit anderen Arten der offenen Aenderung (synthetisch, unabhaengig vom Zustand des Korpus) ---------
def _offene_variante(ende: str) -> list:
    return [("h", "1 Eins"), ("p", "Der Nachweis erfolgt durch ein Messprotokoll der Abnahme."),
            ("h", "2 Zwei"), ("p", "Die Reaktionszeit darf 150 ms nicht ueberschreiten."),
            ("h", "3 Drei"), ("p", "Die Wartung erfolgt {-jaehrlich-} {+monatlich+} durch den Betreiber. " + ende),
            ("h", "4 Vier"), ("p", "Alle Steckverbindungen sind eindeutig beschriftet.")]


VARIANTEN_OFFEN = {
    "ersetzt": "Es gilt {-der-} {+jeder+} Standort.",
    "eingefuegt": "{+Die Pruefung findet jaehrlich statt.+}",
    "geloescht": "{-Ersatzteile werden fuenf Jahre vorgehalten.-}",
}


@pytest.mark.redteam
@pytest.mark.parametrize("art", sorted(VARIANTEN_OFFEN))
def test_rt126_synthetisch_offene_aenderung_in_derselben_zelle(tmp_path, art):
    """Wie RT-125, mit der offenen Aenderung als Ersetzung, Einfuegung oder Streichung (je Aenderung bewerten)."""
    elemente = _offene_variante(VARIANTEN_OFFEN[art])
    baue_pdf(tmp_path / "aenderungen.pdf", elemente, "farbe")
    baue_excel(tmp_path / "vorher.xlsx", zustaende(elemente, False), "L1", True)
    baue_excel(tmp_path / "nachher.xlsx", zustaende(elemente, True), "L1", True)
    sauber = messen.messe_paar(tmp_path, False)
    aufbau(sauber["stille_fehler"] == 0 and sauber["auto_richtig"] == sauber["auto_gesamt"] == 1,
           f"Ausgangslage nicht sauber: {sauber['stille_fehler']} stille Fehler, {sauber['auto_richtig']}/{sauber['auto_gesamt']}")
    echt = messen.plane

    def plane_mit_offener(pdf, xl, *a, **k):
        p = echt(pdf, xl, *a, **k)
        p.aenderungen[-1].status, p.aenderungen[-1].konfidenz = PRUEFEN, "mittel"
        return p

    with mock.patch.object(messen, "plane", plane_mit_offener):
        res = messen.messe_paar(tmp_path, True)
    aufbau(res["aenderungen_nach_status"].get(PRUEFEN, 0) == 1 and res["aenderungen_nach_status"].get(UEBERNOMMEN, 0) >= 1,
           "Erwartet: mindestens eine uebernommene und genau eine offene Aenderung")
    aufbau(res["vorschlaege_gesamt"] == 1 and res["vorschlaege_richtig"] == 1,
           "Mit Freigabe der offenen Aenderung muss die Zelle dem Soll entsprechen")
    if res["stille_fehler"] or res.get("aenderungen_uebernommen_falsch"):
        scheitern(f"{res['stille_fehler']} stille Fehler {res['stille_fehler_nach_art']}, {res.get('aenderungen_uebernommen_falsch')} "
                  f"Aenderung(en) 'falsch', obwohl die angewendete Aenderung richtig ist ({art}) und die offene sichtbar auf "
                  f"'pruefen' steht")
    aufbau(res.get("zellen_warten_auf_freigabe") == 1 and res["pruefen_zellen"] == 1,
           "Die Zelle muss als 'wartet auf Freigabe' und 'zu pruefen' gezaehlt werden")


# (Seed, Profil, Blatt 0-basiert, Zeile, Spalte). Die frueheren Faelle Seed 22 (2x), 37 und 42 treten im heutigen Korpus nicht
# mehr so auf (PDF-Reparaturen bzw. zusaetzlicher Leerraum-Fehler, siehe README_korpus.md); dafuer gilt der synthetische Test oben.
MA1_ZELLEN = [(11, "standard", 2, 3, 2), (49, "standard", 1, 8, 2), (57, "schwer", 1, 11, 3), (58, "standard", 2, 2, 2)]


@pytest.mark.redteam
@pytest.mark.parametrize("seed,profil,blatt,zeile,spalte", MA1_ZELLEN,
                         ids=[f"seed{s}_B{b + 1}Z{z}S{c}" for s, _, b, z, c in MA1_ZELLEN])
def test_rt126_korpus_zelle_weicht_nur_um_offene_aenderung_ab(seed, profil, blatt, zeile, spalte):
    lf = auto_lauf(seed, profil)
    _, soll, ist = zelle_vorher_soll_ist(lf.ordner, lf.ergebnis, blatt, zeile, spalte)
    aend = lf.aenderungen_bei(blatt, zeile, spalte)
    offene = [a for a in aend if a.status != UEBERNOMMEN]
    if aend and not offene and ist == soll:
        # Welle 1 (ADR 0004): die frueher offene Aenderung ist jetzt sicher (exakte Stelle) und die Zelle stimmt -- der
        # Messartefakt-Fall entsteht hier nicht mehr (der synthetische Test oben haelt ihn fest)
        return
    aufbau(any(a.status == UEBERNOMMEN for a in aend) and offene, "Erwartet: uebernommene UND offene Aenderung in der Zelle")
    aufbau(nur_durch_offene_aenderungen_verschieden(ist, soll, offene),
           "Die Zelle weicht auch AUSSERHALB der offenen Aenderungen vom Soll ab (dann waere es ein echter Fehler)")
    gezaehlt = stumm_bei(messung(seed, profil), blatt + 1, zeile, spalte)
    if gezaehlt:
        scheitern(f"Seed {seed}, Zelle B{blatt + 1}Z{zeile}S{spalte} zaehlt als stiller Fehler ({gezaehlt[0]['aenderung']}), "
                  f"obwohl nur {', '.join(a.id for a in offene)} ({'/'.join(sorted({a.status for a in offene}))}) fehlt")

# ==================================================================================================
# MA-2 (4 Eintraege, "Kollateral"): Eine Zeile wird nur TEILWEISE geleert (Nummer/Titel weg, ein Absatz bleibt). Die
#      Zeilen-Ausrichtung der Messung (ZEILEN_AEHNLICH = 0,6) erkennt die Zeile nicht wieder und meldet eine "neue Zeile
#      aus dem Nichts" = Kollateralschaden an Zellen, die gar nicht angefasst wurden.
# ==================================================================================================
@pytest.mark.redteam
def test_rt127_teilweise_geleerte_zeile_ist_kein_kollateral(tmp_path):
    satz = "Satz eins ist hier. Satz zwei ist hier. Satz drei ist hier."
    vorher = [KOPF, ["1", "Eins", "Absatz A.", "Status offen"], ["2", "Zwei", satz + "\nAbsatz zwei ist hier.", "Status offen"],
              ["3", "Drei", "Absatz C.", "Status offen"]]
    nachher = [KOPF, ["1", "Eins", "Absatz A.", "Status offen"], ["3", "Drei", "Absatz C.", "Status offen"]]
    ergebnis = [KOPF, ["1", "Eins", "Absatz A.", "Status offen"], [None, None, satz, "Status offen"],
                ["3", "Drei", "Absatz C.", "Status offen"]]
    res = messe_mit_fertigem_ergebnis(tmp_path, vorher, nachher, ergebnis)
    aufbau(res["aenderungen"] == 0, "Das Tool darf hier keine Aenderung erkennen (leeres PDF)")
    # Die Zellen der Zeilen 2 und 4 (Absatz A / Absatz C) sind in Vorher, Soll und Ergebnis gleich: wer sie als Kollateral zaehlt
    # (Phantom einer "neuen Zeile aus dem Nichts"), zaehlt falsch. Die teilweise geleerte Zeile 3 selbst hat dieses Ergebnis
    # tatsaechlich ohne Aenderung verkuerzt (handgebaut) und darf als Kollateral gelten.
    phantom = [(f["blatt"], f["zeile"], f["spalte"]) for f in res["fehler"] if f["stumm"] and f["zeile"] != 3]
    if phantom:
        scheitern(f"{len(phantom)} Phantom-Kollateralschaeden an unveraenderten Zellen (Positionen: {phantom})")
    if res["kollateral"] > 1:
        scheitern(f"{res['kollateral']} Kollateralschaeden, erlaubt ist hoechstens die verkuerzte Zelle der Zeile 3")


# --- RT-128: Phantom-Kollateral an UNVERAENDERTEN Zellen, synthetisch aus dem logischen Modell aufgebaut ---------------------
# Die Situation wird nicht mehr vom Zustand des Korpus abhaengig gemacht: Das Tool arbeitet echt, nur der Status der
# Absatz-Loeschungen wird (wie bei unsicheren Treffern) gesetzt. Eine Section wird geloescht (Nummer/Titel laufen automatisch),
# ein Absatz ihres Inhalts bleibt offen ("pruefen"), die anderen werden gestrichen: zurueck bleibt eine TEILWEISE geleerte Zeile
# (Nummer und Titel weg, ein Absatz und die Bemerkung bleiben). Das Soll hat die ganze Zeile entfernt. Die Faelle Seed 2 und 37
# (je Zeile mit Zelle 3 und 4) waren die urspruenglichen Funde.
def _elemente_geloeschte_section(absaetze: int) -> list:
    inhalt = [("p", "{-%s-}" % t) for t in ("Die Reaktionszeit darf 150 ms nicht ueberschreiten.",
                                           "Die Pruefung erfolgt einmal je Jahr durch den Betreiber.",
                                           "Das Protokoll wird zehn Jahre aufbewahrt.")[:absaetze]]
    return ([("h", "1 Eins"), ("p", "Der Nachweis erfolgt durch ein Messprotokoll der Abnahme."), ("h", "{-2 Zwei-}")] + inhalt +
            [("h", "3 Drei"), ("p", "Alle Steckverbindungen sind eindeutig beschriftet."),
             ("h", "4 Vier"), ("p", "Die Lieferung umfasst alles Zubehoer fuer die Montage.")])


@pytest.mark.redteam
@pytest.mark.parametrize("absaetze,offen", [(2, 1), (2, 0), (3, 1), (3, 2)],
                         ids=["2_Absaetze_zweiter_offen", "2_Absaetze_erster_offen", "3_Absaetze_mittlerer_offen",
                              "3_Absaetze_letzter_offen"])
def test_rt128_synthetisch_phantom_kollateral_an_unveraenderter_zelle(tmp_path, absaetze, offen):
    elemente = _elemente_geloeschte_section(absaetze)
    baue_pdf(tmp_path / "aenderungen.pdf", elemente, "farbe")
    baue_excel(tmp_path / "vorher.xlsx", zustaende(elemente, False), "L1", True)
    baue_excel(tmp_path / "nachher.xlsx", zustaende(elemente, True), "L1", True)
    echt = messen.plane

    def plane_ein_absatz_offen(pdf, xl, *a, **k):
        p = echt(pdf, xl, *a, **k)
        absatz = [a_ for a_ in p.aenderungen if a_.typ == "absatz_geloescht"]
        for a_ in p.aenderungen:
            if a_.typ == "section_geloescht":
                a_.status, a_.konfidenz = UEBERNOMMEN, "hoch"             # laeuft automatisch ...
        for i, a_ in enumerate(absatz):                                   # ... genau ein Absatz bleibt offen
            a_.status, a_.konfidenz = (PRUEFEN, "mittel") if i == offen else (UEBERNOMMEN, "hoch")
        return p

    with mock.patch.object(messen, "plane", plane_ein_absatz_offen):
        res = messen.messe_paar(tmp_path, False, tmp_path / "ergebnis.xlsx")
    ist = messen._lese_excel(tmp_path / "ergebnis.xlsx")[1]
    # Aufbau: eine Zeile ist teilweise geleert (Nummer/Titel leer, Inhalt steht noch, Bemerkung steht noch)
    rest = [z for z in ist if not (z[0] or z[1]) and z[2] and len(z) > 3 and z[3]]
    aufbau(len(rest) == 1, f"Erwartet: genau eine teilweise geleerte Zeile, gefunden {len(rest)}")
    aufbau(res["aenderungen_nach_status"].get(UEBERNOMMEN, 0) >= 2 and res["aenderungen_nach_status"].get(PRUEFEN, 0) == 1,
           "Erwartet: uebernommene Aenderungen und genau eine offene")
    if res["kollateral"] or res["stille_fehler"]:
        scheitern(f"Phantom: {res['kollateral']} Kollateral / {res['stille_fehler']} stille Fehler "
                  f"{res['stille_fehler_nach_art']} an Zellen, die in Vorher, Soll und Ergebnis gleich sind "
                  f"(Positionen: {[(f['blatt'], f['zeile'], f['spalte']) for f in res['fehler'] if f['stumm']]})")

# ==================================================================================================
# MA-3 (4 Zellen, "uebernommen, aber falsch" an NEUEN Zeilen): Das Tool leert eine geloeschte Section (Nummer/Titel/Inhalt),
#      die Zeile bleibt mit der Bemerkung stehen; die neue Zeile steht dahinter. Das Soll hat die alte Zeile GANZ entfernt
#      und die neue davor angesetzt. Inhalt und Reihenfolge sind gleich, aber der "Anker" ist ein anderer.
# ==================================================================================================
@pytest.mark.redteam
def test_rt129_geleerte_statt_geloeschte_zeile_verschiebt_den_anker(tmp_path):
    vorher = [KOPF, ["1", "Eins", "Absatz A.", "Status offen"], ["2", "Zwei", "Absatz B ist hier.", "Status offen"],
              ["3", "Drei", "Absatz C.", "Status offen"]]
    neue = ["7", "Leistung", "Voellig anderer Text hier ohne Bezug."]
    nachher = [KOPF, ["1", "Eins", "Absatz A.", "Status offen"], neue + ["Status offen"], ["3", "Drei", "Absatz C.", "Status offen"]]
    ergebnis = [KOPF, ["1", "Eins", "Absatz A.", "Status offen"], [None, None, None, "Status offen"], neue + [None],
                ["3", "Drei", "Absatz C.", "Status offen"]]
    res = messe_mit_fertigem_ergebnis(tmp_path, vorher, nachher, ergebnis)
    inhalte = lambda zeilen: [tuple(z[:3]) for z in zeilen if any(z[:3])]          # verwaltete Spalten, ohne Leerzeilen
    aufbau(inhalte(nachher) == inhalte(ergebnis), "Inhalt/Reihenfolge der verwalteten Spalten muessen gleich sein")
    if res["stille_fehler"]:
        scheitern(f"{res['stille_fehler']} stille Fehler {res['stille_fehler_nach_art']} bei gleichem Inhalt und gleicher "
                  f"Reihenfolge (nur eine leere Restzeile mit Bemerkung bleibt stehen)")


MA3_ZELLEN = [(43, "standard", 2, 4, 3), (57, "schwer", 1, 9, 1), (57, "schwer", 1, 9, 2)]


@pytest.mark.redteam
@pytest.mark.parametrize("seed,profil,blatt,zeile,spalte", MA3_ZELLEN,
                         ids=[f"seed{s}_B{b + 1}Z{z}S{c}" for s, _, b, z, c in MA3_ZELLEN])
def test_rt130_korpus_anker_nach_geleerter_zeile(seed, profil, blatt, zeile, spalte):
    lf = auto_lauf(seed, profil)
    rest = zeile_im_ergebnis(lf.ergebnis, blatt, zeile)
    aufbau(len(rest) >= 4 and not any(rest[:3]) and rest[3] != "",
           "In der Zeile ist nicht nur die Bemerkung uebrig geblieben (geleerte Zeile erwartet)")
    gezaehlt = stumm_bei(messung(seed, profil), blatt + 1, zeile, spalte)
    if gezaehlt:
        scheitern(f"Seed {seed}: Zelle B{blatt + 1}Z{zeile}S{spalte} der NEUEN Zeile zaehlt als stiller Fehler "
                  f"({gezaehlt[0]['aenderung']}), nur weil die alte Zeile geleert statt geloescht ist (Rest: {rest[3]!r})")


# ==================================================================================================
# MA-4 / MA-5 (Kriterium 5, "Lese-Luecke"): Woerter, die gar nicht im PDF stehen, werden als "vom Leser nicht erkannt"
#      gezaehlt. Macht die Lese-Luecke schlechter, als sie ist (Zielwert 2 %; gemessen 2,35 %).
# ==================================================================================================
_EL_L3 = [("h", "1 Eins"), ("p", "Der Nachweis erfolgt durch ein Messprotokoll der Abnahme."),
          ("h", "2 Zwei"),
          ("p", "Alle Lueftungen sind frei. {+Die Anlage muss bei jedem Einsatz ohne Werkzeug und ohne fremde Hilfe "
                "erreichbar sein, sofern der Zugang gesichert ist.+}"),
          ("p", "Die Reaktionszeit darf 150 ms nicht ueberschreiten."),
          ("h", "3 Drei"), ("p", "Alle Steckverbindungen sind eindeutig beschriftet.")]


@pytest.mark.redteam
def test_rt131_kurzer_absatz_mit_langem_zusatz_ist_keine_lese_luecke(tmp_path):
    baue_basis_paar(tmp_path / "l3", "L3", _EL_L3)
    baue_basis_paar(tmp_path / "l1", "L1", _EL_L3)
    l1 = messen.messe_paar(tmp_path / "l1", False)
    l3 = messen.messe_paar(tmp_path / "l3", False)
    for r in (l1, l3):
        aufbau(r["stille_fehler"] == 0 and r["auto_richtig"] == r["auto_gesamt"], "Das Tool muss hier alles richtig machen")
    aufbau(l1["lese_luecke_woerter"] == 0, "Im Layout 'eine Zeile je Section' muss die Lese-Luecke 0 sein")
    if l3["lese_luecke_woerter"]:
        scheitern(f"Lese-Luecke {l3['lese_luecke_woerter']} Woerter von {l3['soll_woerter']} im Layout 'eine Zeile je Absatz' "
                  f"(derselbe Inhalt im Layout 'eine Zeile je Section': {l1['lese_luecke_woerter']} von {l1['soll_woerter']}); "
                  f"Tool und Leser sind in beiden fehlerfrei")


def _fehlende_soll_woerter(ordner) -> Counter:
    """Die Woerter, die die Messung als 'im PDF nicht erkannt' zaehlt (die Rechnung selbst: `_Auswertung.lese_luecke_fehlend`)."""
    gesammelt: dict[str, Counter] = {}
    original = messen._Auswertung.auswerten

    def haken(self):
        gesammelt["f"] = self.lese_luecke_fehlend()
        return original(self)

    with mock.patch.object(messen._Auswertung, "auswerten", haken):
        messen.messe_paar(ordner, False)
    return gesammelt["f"]


@pytest.mark.redteam
def test_rt132_korpus_formelzellen_sind_keine_pdf_woerter_seed60():
    fehlend = _fehlende_soll_woerter(paar(60, "schwer"))
    formeln = {w: n for w, n in fehlend.items() if w.startswith("=")}
    aufbau(sum(fehlend.values()) > 0, "Die Messung zaehlt in Seed 60 keine fehlenden Woerter")
    if formeln:
        scheitern(f"{sum(formeln.values())} von {sum(fehlend.values())} 'nicht erkannten' Woertern in Seed 60 sind Formeln "
                  f"(z. B. {list(formeln)[:3]}); die Lese-Luecke ist zu schlecht")
