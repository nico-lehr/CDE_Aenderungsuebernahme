"""Red-Team R3 (Analyse der stillen Fehler im Entwicklungskorpus), Teil 1: ECHTE TOOL-FEHLER.

Ausgangslage: `python -m src.cli messen --korpus entwicklung` meldet 43 "stille Fehler" (60 Paare, Seeds 1-60).
Hier stehen die Ursachen, die wirklich beim Tool liegen (27 der 43). Fuer jede Ursache gibt es

  * einen KLEINEN Fall, im Speicher mit den Hilfen aus tests/fixtures gebaut (zeigt die Ursache isoliert), und
  * einen Korpus-Fall (Seeds < 1000, Regel 11), der zeigt, dass es dieselbe Ursache im Korpus ist.

Jeder Fehlertest traegt `@pytest.mark.redteam` und `@pytest.mark.xfail(strict=True, raises=AssertionError,
reason="RT-nn ...; Besitzer ...")`. Die FALSCHE Wirkung wird als AssertionError gemeldet (`scheitern`), Fehler im
Testaufbau als RuntimeError (`aufbau`). Alles ERFUNDEN (Regel 16); nichts in src/, config/ oder tests/fixtures/ geaendert.

Die Messartefakte (16 der 43) stehen in test_korpus_messartefakte.py, die Schwaechen der Messung selbst in
test_korpus_messluecken.py. Uebersicht mit Zahlen: tests/redteam/README_korpus.md.
"""
from __future__ import annotations

from collections import Counter

import pytest

from src.models import UEBERNOMMEN
from src.pdf_reader import analysiere_pdf
from tests.redteam._bau import aufbau, baue_alt, lauf, scheitern, zelle, zeile_mit
from tests.redteam._bau import analysiere_mit_pdf, uebernimm
from tests.redteam._bau_korpus import (auto_lauf, erwartet, paar, zelle_vorher_soll_ist)
from tests.redteam._bau_pdf import dok

# ==================================================================================================
# G1: Der PDF-Leser meldet NIE eine Luecke (15 Paare, 17 Luecken). Einzelfaelle: RT-95..RT-99, RT-109 (R2).
#     Hier: die Korpus-Zaehlung je Luecken-Art.
# ==================================================================================================
SCAN = [(6, "schwer"), (12, "schwer"), (27, "schwer")]
SPALTEN = [(3, "schwer"), (9, "schwer"), (18, "schwer"), (42, "schwer"), (48, "schwer")]
TABELLE = [(33, "schwer"), (39, "schwer"), (45, "schwer")]
BILD = [(6, "schwer"), (21, "schwer"), (48, "schwer"), (51, "schwer"), (54, "schwer"), (60, "schwer")]


def _fehlende_luecken(faelle: list[tuple[int, str]], art: str) -> tuple[list[str], int]:
    """-> (Paare, in denen die erwartete Luecke `art` nicht gemeldet ist, Zahl der erwarteten Luecken)."""
    fehlt: list[str] = []
    erwartet_n = 0
    for seed, profil in faelle:
        soll = [l for l in erwartet(seed, profil)["luecken"] if l["art"] == art]
        aufbau(bool(soll), f"Seed {seed}: der Generator erwartet keine Luecke '{art}'")
        erwartet_n += len(soll)
        gemeldet = Counter(l["art"] for l in analysiere_pdf(paar(seed, profil) / "aenderungen.pdf").luecken)
        if gemeldet.get(art, 0) < len(soll):
            fehlt.append(f"synth_{seed:04d}")
    return fehlt, erwartet_n


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-110 Korpus: Scan-Seiten (Seeds 6, 12, 27) werden nicht als Luecke 'scan_seite' gemeldet; "
                          "3 stille Fehler 'Luecke nicht gemeldet' (mit Wortverlust); Besitzer K-PDF")
def test_rt110_korpus_scan_seiten_ohne_luecke():
    fehlt, n = _fehlende_luecken(SCAN, "scan_seite")
    if fehlt:
        scheitern(f"Luecke 'scan_seite' nicht gemeldet in {len(fehlt)} von {len(SCAN)} Paaren ({', '.join(fehlt)}); "
                  f"erwartet: {n}; die Bilanz sagt 'vollstaendig'")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-111 Korpus: Zwei-Spalten-Abschnitte (Seeds 3, 9, 18, 42, 48) werden nicht als Luecke "
                          "'mehrspaltig' gemeldet; 5 stille Fehler (mit Wortverlust); Besitzer K-PDF")
def test_rt111_korpus_zwei_spalten_ohne_luecke():
    fehlt, n = _fehlende_luecken(SPALTEN, "mehrspaltig")
    if fehlt:
        scheitern(f"Luecke 'mehrspaltig' nicht gemeldet in {len(fehlt)} von {len(SPALTEN)} Paaren ({', '.join(fehlt)}); "
                  f"erwartet: {n}; die Bilanz sagt 'vollstaendig'")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-112 Korpus: Tabellen (Seeds 33, 39, 45) werden nicht als Luecke 'tabelle' gemeldet; "
                          "3 stille Fehler (2 davon mit Wortverlust); Besitzer K-PDF")
def test_rt112_korpus_tabellen_ohne_luecke():
    fehlt, n = _fehlende_luecken(TABELLE, "tabelle")
    if fehlt:
        scheitern(f"Luecke 'tabelle' nicht gemeldet in {len(fehlt)} von {len(TABELLE)} Paaren ({', '.join(fehlt)}); "
                  f"erwartet: {n}; die Bilanz sagt 'vollstaendig'")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-113 Korpus: Bilder im Fliesstext (Seeds 6, 21, 48, 51, 54, 60) werden nicht als Luecke 'bild' "
                          "gemeldet; 4 stille Fehler OHNE Wortverlust (Bild ohne Text); Besitzer K-PDF")
def test_rt113_korpus_bilder_ohne_luecke():
    fehlt, n = _fehlende_luecken(BILD, "bild")
    if fehlt:
        scheitern(f"Luecke 'bild' nicht gemeldet in {len(fehlt)} von {len(BILD)} Paaren ({', '.join(fehlt)}); "
                  f"erwartet: {n}; die Bilanz sagt 'vollstaendig'")


# ==================================================================================================
# G2: Silbentrennung in einem BLAUEN (neuen) Wort: das Tool traegt das Trennstueck mit Strich/Leerzeichen ein,
#     "uebernommen / hoch / Nachkontrolle ok". (7 Eintraege: Seeds 3, 9 (3), 51, 57 (2). Reader-Ebene: RT-63/64/108.)
# ==================================================================================================
_KOPF = [("h", "1 Eins"), ("p", "Der Nachweis erfolgt durch ein Messprotokoll der Abnahme."),
         ("h", "2 Zwei"), ("p", "Die Reaktionszeit darf 150 ms nicht ueberschreiten."),
         ("h", "3 Drei"), ("p", "Alle Steckverbindungen sind eindeutig beschriftet.")]
_FUELL = "Bei jeder Pruefung durch den Betreiber wird die Anlage im Betrieb ueberwacht und danach"


def _trennstuecke(ist: str, soll: str) -> list[str]:
    """Woerter, die im Soll ganz stehen, im Ist aber in zwei Stuecken ('Wort-' 'Rest' oder 'Wo' 'rt')."""
    ist_w = ist.split()
    fehlen = Counter(soll.split()) - Counter(ist_w)
    return [f"{a} {b}" for w in fehlen for a, b in zip(ist_w, ist_w[1:]) if a.rstrip("-") + b == w]


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-114 Ein neues BLAUES Wort, das am Zeilenende getrennt ist ('Netzanschlussbe-' 'dingungen'), "
                          "wird mit Trennstrich bzw. Leerzeichen in die Excel geschrieben (uebernommen/hoch/ok); "
                          "Ursache im Leser (RT-63/64); Besitzer K-PDF")
@pytest.mark.parametrize("modus", ["ascii", "weich"])
def test_rt114_blaues_getrenntes_wort_landet_zerrissen_in_der_excel(tmp_path, modus):
    ziel = ("p", f"{' '.join(_FUELL.split()[:6])} {{-regelmaessig-}} {{+Sicherheitsarchitektur Netzanschlussbedingungen+}} "
                 "geprueft und dokumentiert, sofern der Zugang gesichert ist.")
    el = _KOPF + [("h", "4 Vier"), ziel, ("h", "5 Fuenf"), ("p", "Die Lieferung umfasst alles Zubehoer fuer die Montage.")]
    d = dok(trennung=modus)
    for art, markup in el:
        d.ueberschrift(markup) if art == "h" else d.block(markup)
    aufbau(any(w == "Netzanschlussbedingungen" and stil == ("n",) for w, stil, _ in d.getrennt),
           "Das blaue Wort wurde im Test-PDF nicht getrennt")
    pdf = tmp_path / "a.pdf"
    d.speichern(pdf)
    alt = baue_alt(tmp_path, el, "L1")
    l = uebernimm(analysiere_mit_pdf(pdf, alt))
    aufbau(len(l.aen) >= 1, "Das Tool hat keine Aenderung erkannt")
    text = zelle(l.out, zeile_mit(l.out, "geprueft und dokumentiert"), 3)
    ae = l.aen[0]
    if "Netzanschlussbedingungen" not in text and ae.status == UEBERNOMMEN:
        scheitern(f"Wort zerrissen in der Excel, Status '{ae.status}'/{ae.konfidenz}, Nachkontrolle '{ae.nachkontrolle}': "
                  + repr(text[text.index("Sicherheitsarchitektur"):][:70]))


# Korpus-Zellen: (Seed, Profil, Blatt 0-basiert, Zeile, Spalte)
SILBEN_ZELLEN = [(3, "schwer", 1, 39, 3), (9, "schwer", 2, 8, 3), (9, "schwer", 2, 12, 3), (51, "schwer", 1, 11, 4),
                 (57, "schwer", 1, 3, 3)]


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-115 Korpus: in 5 Zellen (Seeds 3, 9, 9, 51, 57) steht ein neues blaues Wort zerrissen in der "
                          "Excel (uebernommen/hoch/ok); dazu 2 neue Zeilen in Seed 9 und 57 (7 stille Fehler); Besitzer K-PDF")
@pytest.mark.parametrize("seed,profil,blatt,zeile,spalte", SILBEN_ZELLEN,
                         ids=[f"seed{s}_B{b + 1}Z{z}S{c}" for s, _, b, z, c in SILBEN_ZELLEN])
def test_rt115_korpus_zerrissenes_wort_in_der_excel(seed, profil, blatt, zeile, spalte):
    lf = auto_lauf(seed, profil)
    _, soll, ist = zelle_vorher_soll_ist(lf.ordner, lf.ergebnis, blatt, zeile, spalte)
    aenderungen = lf.aenderungen_bei(blatt, zeile, spalte)
    aufbau(any(a.status == UEBERNOMMEN for a in aenderungen), "In dieser Zelle ist keine Aenderung uebernommen")
    stuecke = _trennstuecke(ist, soll)
    if stuecke:
        a = next((a for a in aenderungen if a.status == UEBERNOMMEN and any(s.split()[0] in a.neu for s in stuecke)),
                 next(a for a in aenderungen if a.status == UEBERNOMMEN))
        scheitern(f"Seed {seed}, Zelle B{blatt + 1}Z{zeile}S{spalte}: Wort in zwei Stuecken in der Excel ({stuecke[0]!r}); "
                  f"{a.id} {a.status}/{a.konfidenz}, Nachkontrolle '{a.nachkontrolle}'")


# ==================================================================================================
# G3: Tabelle mit blauem Kopf / Baenderung (unveraendert!) wird als "neu" gelesen: der Tabellenkopf wird in die
#     Zelle eingefuegt (uebernommen/hoch/ok). 3 Zellen: Seeds 33, 39, 45. Reader-Ebene: RT-76/77.
# ==================================================================================================
_TAB = [["Pos.", "Bauteil", "Menge", "Einheit"], ["1", "Netzteil", "24", "Stueck"], ["2", "Sicherung", "8", "Stueck"],
        ["3", "Schuetz", "12", "Stueck"]]
_TAB_TXT = "\n".join(" | ".join(z) for z in _TAB)
_KOPFWOERTER = "Pos. Bauteil Menge Einheit"


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-116 Eine UNVERAENDERTE Tabelle mit blauem Kopf und hellblauer Baenderung: das Tool fuegt die "
                          "Kopf- und Datenwoerter ('Pos. Bauteil Menge Einheit ...') in die Excel-Zelle ein "
                          "(uebernommen/hoch/ok); Ursache im Leser (RT-76/77); Besitzer K-PDF")
def test_rt116_unveraenderte_farbtabelle_wird_in_die_zelle_geschrieben(tmp_path):
    el_excel = _KOPF + [("h", "4 Vier"), ("p", "Die Anlage wird regelmaessig geprueft und dokumentiert."), ("p", _TAB_TXT),
                        ("h", "5 Fuenf"), ("p", "Die Lieferung umfasst alles Zubehoer fuer die Montage.")]
    d = dok()
    for art, markup in _KOPF + [("h", "4 Vier"), ("p", "Die Anlage wird regelmaessig geprueft und dokumentiert.")]:
        d.ueberschrift(markup) if art == "h" else d.block(markup)
    d.tabelle(_TAB, kopf_farbe="4472C4", band_farbe="B4C6E7")
    d.ueberschrift("5 Fuenf")
    d.block("Die Lieferung umfasst alles Zubehoer fuer die Montage.")
    pdf = tmp_path / "a.pdf"
    d.speichern(pdf)
    alt = baue_alt(tmp_path, el_excel, "L1")
    aufbau(_KOPFWOERTER not in zelle(alt, zeile_mit(alt, "Anlage wird regelmaessig"), 3), "Excel enthaelt den Kopf schon")
    l = uebernimm(analysiere_mit_pdf(pdf, alt))
    text = zelle(l.out, zeile_mit(l.out, "Anlage wird regelmaessig"), 3)
    if _KOPFWOERTER in text:
        ae = next(a for a in l.aen if _KOPFWOERTER in a.neu)
        scheitern(f"Unveraenderte Tabelle in die Zelle geschrieben: {ae.id} {ae.status}/{ae.konfidenz}, Nachkontrolle "
                  f"'{ae.nachkontrolle}', Zelle jetzt: {text.split(chr(10), 1)[-1][:85]!r}...; Luecken: {l.plan.luecken}")


TABELLEN_ZELLEN = [(33, "schwer", 2, 13, 2), (39, "schwer", 2, 25, 3), (45, "schwer", 1, 27, 3)]


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-117 Korpus: in 3 Zellen (Seeds 33, 39, 45) steht nach dem Lauf der Tabellenkopf als Fliesstext "
                          "mitten in der Tabelle der Excel (uebernommen/hoch/ok); 3 stille Fehler; Besitzer K-PDF")
@pytest.mark.parametrize("seed,profil,blatt,zeile,spalte", TABELLEN_ZELLEN,
                         ids=[f"seed{s}_B{b + 1}Z{z}S{c}" for s, _, b, z, c in TABELLEN_ZELLEN])
def test_rt117_korpus_tabellenkopf_in_der_excel(seed, profil, blatt, zeile, spalte):
    lf = auto_lauf(seed, profil)
    _, soll, ist = zelle_vorher_soll_ist(lf.ordner, lf.ergebnis, blatt, zeile, spalte)
    aufbau(" | " in soll, "Die Soll-Zelle ist keine Tabellen-Zelle")
    aendg = [a for a in lf.aenderungen_bei(blatt, zeile, spalte) if a.status == UEBERNOMMEN and _KOPFWOERTER in a.neu]
    if _KOPFWOERTER in ist and _KOPFWOERTER not in soll:
        a = aendg[0] if aendg else None
        scheitern(f"Seed {seed}: Tabellenkopf als Fliesstext in der Excel-Zelle B{blatt + 1}Z{zeile}S{spalte}"
                  + (f" ({a.id} {a.status}/{a.konfidenz}, Nachkontrolle '{a.nachkontrolle}')" if a else ""))


# ==================================================================================================
# G4: Zwei-Spalten-Text wird Zeile fuer Zeile ueber beide Spalten gelesen: ein neuer Satz der rechten Spalte wird in
#     zwei Stuecken MITTEN in Saetze der linken Spalte eingetragen (uebernommen/hoch/ok). 1 Zelle: Seed 42.
# ==================================================================================================
_LINKS = ["Der Nennwert von 20 Monate gilt bei 40 Grad Umgebungstemperatur. Die Betriebsspannung darf 120 V nicht "
          "ueberschreiten.", "Bei Luefterausfall wird ein Bericht an die Leitstelle uebermittelt."]
_RECHTS = "Bei Ueberstrom wird ein Alarm an die Leitstelle uebermittelt."


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-118 Zwei-Spalten-Text: ein neuer Satz der rechten Spalte wird in zwei Stuecken mitten in Saetze "
                          "der linken Spalte geschrieben (uebernommen/hoch/ok), ohne Luecke 'mehrspaltig'; Besitzer K-PDF")
def test_rt118_zwei_spalten_satz_wird_zerrissen_eingetragen(tmp_path):
    el_excel = _KOPF + [("h", "4 Vier"), ("p", _LINKS[0]), ("p", _LINKS[1]), ("h", "5 Fuenf"),
                        ("p", "Die Lieferung umfasst alles Zubehoer fuer die Montage.")]
    d = dok()
    for art, markup in _KOPF + [("h", "4 Vier")]:
        d.ueberschrift(markup) if art == "h" else d.block(markup)
    d.zwei_spalten(_LINKS, ["{+" + _RECHTS + "+}"])
    d.ueberschrift("5 Fuenf")
    d.block("Die Lieferung umfasst alles Zubehoer fuer die Montage.")
    pdf = tmp_path / "a.pdf"
    d.speichern(pdf)
    alt = baue_alt(tmp_path, el_excel, "L1")
    l = uebernimm(analysiere_mit_pdf(pdf, alt))
    aufbau(len(l.aen) >= 1, "Das Tool hat keine Aenderung erkannt")
    text = zelle(l.out, zeile_mit(l.out, "Der Nennwert von 20 Monate"), 3)
    auto = [a for a in l.aen if a.status == UEBERNOMMEN]
    if auto and _RECHTS not in text and not l.plan.luecken:
        scheitern(f"Satz der rechten Spalte zerrissen in der Excel ({len(auto)} Teile uebernommen/hoch), keine Luecke: "
                  + repr(text[:150]))


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-119 Korpus Seed 42 (Stoerer 'zwei_spalten'): ein neuer Satz steht in zwei Stuecken mitten in "
                          "anderen Saetzen der Excel-Zelle (uebernommen/hoch/ok); 1 stiller Fehler; Besitzer K-PDF")
def test_rt119_korpus_zwei_spalten_satz_zerrissen_seed42():
    lf = auto_lauf(42, "schwer")
    _, soll, ist = zelle_vorher_soll_ist(lf.ordner, lf.ergebnis, 1, 3, 3)
    satz = "Bei Überstrom wird ein Alarm an die Leitstelle übermittelt."
    aufbau(satz in soll, "Der Soll-Zelle fehlt der erwartete Satz (Generator geaendert?)")
    if satz not in ist and "Bei Überstrom wird ein Alarm an die Leitstelle" in ist:
        scheitern("Seed 42: der neue Satz steht zerrissen in der Excel-Zelle B2Z3S3: "
                  + repr(ist[ist.index("40 °C") - 20:][:130] if "40 °C" in ist else ist[:130]))


# ==================================================================================================
# G5: Neuer Listenpunkt "4. Sicherung tauschen" (Nummer schwarz, Text blau): nur die Nummer "4." wird zum Suchen benutzt,
#     sie trifft die Nummern-Zelle "4" der Section 4; der Text wird IN DIE NUMMERNSPALTE geschrieben (uebernommen/hoch/ok).
#     1 Zelle: Seed 30.
# ==================================================================================================
_LISTE_PDF = [
    ("h", "3 Betrieb"), ("p", "Der Betrieb erfolgt jederzeit durch geschultes Personal vor Ort."),
    ("h", "4 Stoermeldungen"), ("p", "Die Reaktionszeit darf 150 ms nicht ueberschreiten."),
    ("h", "4.1 Geltungsbereich"), ("p", "1. Kabel pruefen"), ("p", "2. Tuerkontakt testen"), ("p", "3. Protokoll ausfuellen"),
    ("p", "4. {+Sicherung tauschen+}"),
    ("h", "4.2 Normen"), ("p", "Die Norm gilt fuer alle Anlagen im Betrieb."),
]
_LISTE_EXCEL = [e for e in _LISTE_PDF if "Sicherung" not in e[1]]


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-120 Neuer Listenpunkt '4. {+Sicherung tauschen+}': der Text wird in die NUMMERNSPALTE der "
                          "Section 4 geschrieben ('4 Sicherung tauschen'), uebernommen/hoch/ok (die Suche nutzt nur das "
                          "alte Wort '4.'); Besitzer K-XL (matcher: Nummern-Zelle als Absatz-Kandidat)")
def test_rt120_listenpunkt_text_landet_in_der_nummernspalte(tmp_path):
    alt = baue_alt(tmp_path, _LISTE_EXCEL, "L1")
    l = lauf(tmp_path, _LISTE_PDF, "L1", alt=alt)
    aufbau(len(l.aen) == 1 and "Sicherung" in l.aen[0].neu, "Erwartet genau eine Aenderung (der neue Listenpunkt)")
    ae = l.aen[0]
    zeile = zeile_mit(l.out, "Stoermeldungen")
    nr = zelle(l.out, zeile, 1)
    if nr != "4" and ae.status == UEBERNOMMEN:
        scheitern(f"Der Text steht in der Nummernspalte: {nr!r} ({ae.id} {ae.status}/{ae.konfidenz}, Nachkontrolle "
                  f"'{ae.nachkontrolle}', Operation auf Spalte {ae.ops[0].spalte})")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-121 Korpus Seed 30 (Stoerer 'nummerierte_liste'): '4 Sicherung tauschen' steht in der "
                          "Nummernspalte (Blatt 2, Zeile 18, Spalte 1); 1 stiller Fehler; Besitzer K-XL")
def test_rt121_korpus_listenpunkt_in_nummernspalte_seed30():
    lf = auto_lauf(30, "schwer")
    vorher, soll, ist = zelle_vorher_soll_ist(lf.ordner, lf.ergebnis, 1, 18, 1)
    aenderungen = lf.aenderungen_bei(1, 18, 1)
    aufbau(vorher == "4" and soll == "4", "Die Nummern-Zelle ist im Soll nicht unveraendert '4'")
    if ist != vorher and any(a.status == UEBERNOMMEN for a in aenderungen):
        a = aenderungen[0]
        scheitern(f"Seed 30: Nummern-Zelle B2Z18S1 ist jetzt {ist!r} statt '4' ({a.id} {a.status}/{a.konfidenz}, "
                  f"Nachkontrolle '{a.nachkontrolle}')")


# ==================================================================================================
# G6 (kein stiller Fehler, aber Ursache von 2 der 4 Phantom-Kollateral und von verwaisten Zeilen):
#     "Section geloescht" laeuft automatisch, der geloeschte INHALT der Section bleibt offen ("pruefen").
# ==================================================================================================
@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-122 Section-Loeschung (Nummer + Titel) laeuft automatisch (hoch), die Loeschung des Inhalts "
                          "bleibt auf 'pruefen': zurueck bleibt eine verwaiste Zeile ohne Nummer/Titel mit altem Inhalt "
                          "(Seeds 2, 37, 60 im Korpus); Besitzer K-PL (planner: Abhaengigkeit Kopf <-> Inhalt, wie RT-01/02)")
def test_rt122_section_geloescht_automatisch_inhalt_offen_laesst_verwaiste_zeile(tmp_path):
    pdf_el = [("h", "1 Eins"), ("p", "Der Nachweis erfolgt durch ein Messprotokoll der Abnahme."),
              ("h", "{-2 Zwei-}"), ("p", "{-Die Reaktionszeit darf 150 ms nicht ueberschreiten.-}"),
              ("h", "3 Drei"), ("p", "Alle Steckverbindungen sind eindeutig beschriftet."),
              ("h", "4 Vier"), ("p", "Die Lieferung umfasst alles Zubehoer fuer die Montage.")]
    xl_el = [e if "Reaktionszeit" not in e[1] else ("p", "Die Reaktionszeit darf 150 ms nicht ueberschreiten. "
                                                     "Die Messung erfolgt bei Nennlast.") for e in pdf_el]
    xl_el = [(a, m.replace("{-", "").replace("-}", "")) for a, m in xl_el]       # Excel: alter Stand, ungestrichen
    alt = baue_alt(tmp_path, xl_el, "L1")
    l = lauf(tmp_path, pdf_el, "L1", alt=alt)
    kopf = next((a for a in l.aen if a.typ == "section_geloescht"), None)
    inhalt = next((a for a in l.aen if a.typ == "absatz_geloescht"), None)
    aufbau(kopf is not None and inhalt is not None, "Erwartet: eine Section- und eine Absatz-Loeschung")
    aufbau(inhalt.status != UEBERNOMMEN, "Der Inhalt muss offen ('pruefen') sein, sonst gilt dieser Fall nicht")
    zeile = zeile_mit(l.out, "Die Reaktionszeit")
    if kopf.status == UEBERNOMMEN and zelle(l.out, zeile, 1) == "" and zelle(l.out, zeile, 2) == "":
        scheitern(f"Verwaiste Zeile {zeile}: Nummer und Titel sind geleert ({kopf.id} {kopf.status}/{kopf.konfidenz}), der "
                  f"Inhalt steht noch da, weil {inhalt.id} auf '{inhalt.status}' steht")
