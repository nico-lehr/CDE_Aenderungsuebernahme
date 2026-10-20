"""Leerzeichen am Zeilenumbruch neben einer Ersetzung (Auftrag ZUORDNEN, Welle 6): Am Zeilen- oder Seitenumbruch zeichnet das PDF
kein Leerzeichen, also sieht man nicht, ob es markiert ist. Steht es hinter einer Ersetzung ("12" gestrichen, "15" neu, dann der
Umbruch, dann "mm"), kann es dreierlei sein: gemeinsam (alt und neu), nur neu (Teil der Einfuegung "15 ") oder nur alt (Teil der
Streichung). Die Excel entscheidet nur bei echter Mehrdeutigkeit: Steht der alte Absatz dort OHNE Leerzeichen ("12mm"), kann es nur
zur Einfuegung gehoeren; steht er MIT Leerzeichen, bleibt es beim gemeinsamen (bisher) -- die Excel kann "gemeinsam" und "nur alt"
nicht unterscheiden. Ohne Umbruch (sichtbarer, unmarkierter Leerraum) und ohne passenden Excel-Absatz aendert sich nichts.

Nur erfundene Daten (Regel 16). Die Mappe und das Word-PDF entstehen im Test (tests/test_register.py, tests/_word_pdf.py).
"""
from __future__ import annotations

import hashlib
import json
import pdfplumber
import pytest
from openpyxl import load_workbook

from src import markierung as mk
from src import messen
from src.applier import uebernehmen
from src.excel_model import ExcelModell
from src.models import PRUEFEN, UEBERNOMMEN
from src.pdf_reader import analysiere_pdf
from src.planner import plane
from tests._word_pdf import EBENEN, RECHTS, WordDok, breite
from tests.fixtures import word_markup as wm
from tests.test_register import GRUND, G, _EBENE, _label_art, register_mappe, stuecke

VOKABULAR = "Das Gehaeuse steht auf einem Sockel aus Beton und ist gegen Feuchte und Frost gut geschuetzt".split()
EBENE_TEXT = 3                                  # Text eines Listenpunkts "1." steht auf Einzugsebene 3 (Nummer auf 2)
ZIEL = 4                                        # Index in GRUND: der Listenpunkt "1." unter "B."
ZEILE_B = 8 + ZIEL                              # Zeile in der Mappe (Daten ab Zeile 8)
REST = "mm im Betrieb und darf nicht ueberschritten werden."


def _fuelltext(platz: float) -> str:
    """Fuelltext mit Leerzeichen dahinter, der in `platz` pt passt (Zeile fast voll: der Leser haelt die naechste Zeile dann fuer
    die Fortsetzung desselben Absatzes)."""
    text, i = "", 0
    while breite(text + VOKABULAR[i % len(VOKABULAR)] + " ") <= platz - 1.0 and i < 60:
        text += VOKABULAR[i % len(VOKABULAR)] + " "
        i += 1
    return text


def _fall(tmp_path, vorn: str, mitte: list[tuple], rest: str, excel_mitte: str, umbruch: bool = True,
          excel_rest: str | None = None, zeile2: list[tuple] | None = None):
    """Das kleine Register GRUND (tests/test_register.py); der Listenpunkt "1." unter "B." (Zeile 12) ist der Zielabsatz mit
    Umbruch hinter der Aenderung `mitte` (Zeile 1 endet mit ihr, Zeile 2 beginnt mit `zeile2` und `rest`).
    -> (Plan, ExcelModell, Excel-Datei, PdfAnalyse, Excel-Text des Listenpunkts ohne Nummer)"""
    fuell = _fuelltext(RECHTS - EBENEN[EBENE_TEXT] - breite(vorn) - sum(breite(st[0]) for st in mitte)) if umbruch else ""
    zeile1 = [(fuell + vorn, "k")] + mitte
    zeilen = [zeile1, (zeile2 or []) + [(rest, "k")]] if umbruch else [zeile1 + [(rest, "k")]]
    excel_b = fuell + vorn + excel_mitte + (rest if excel_rest is None else excel_rest)
    d = WordDok()
    for i, (art, label, text) in enumerate(GRUND):
        if d.top > 600:
            d.neue_seite()
        if art == "part":
            d.kapitel(stuecke(f"PART {label} ") + stuecke(text))
        elif art == "section":
            d.ueberschrift(stuecke(label), stuecke(text))
        elif art in ("sub", "sub_titel"):
            d.listenpunkt(1, stuecke(label), [stuecke(text)], abstand=0)
        elif i == ZIEL:
            d.listenpunkt(2, stuecke(label), zeilen, abstand=0)
        else:
            d.listenpunkt(_EBENE[_label_art(label)], stuecke(label), [stuecke(text)], abstand=0)
    pdf = analysiere_pdf(d.speichern(tmp_path / "w.pdf"))
    absaetze = [(art, label.rstrip(".") if art in ("sub", "sub_titel") else label, excel_b if i == ZIEL else text)
                for i, (art, label, text) in enumerate(GRUND)]
    xlsx = register_mappe(tmp_path / "alt.xlsx", absaetze)
    xl = ExcelModell(xlsx)
    return plane(pdf, xl), xl, xlsx, pdf, excel_b


def _zelle(excel_b: str) -> str:
    """Zellinhalt der Zeile: Listennummer, Tabulator (Polster der Mappe) und Text."""
    return "1.\t" + excel_b


def _ops(ae):
    return [(o.start, o.ende, o.text) for o in ae.ops]


def _a1(ae) -> list[str]:
    return [c for c in ae.pruefcodes if c.startswith("a1_")]


@pytest.mark.legende_neu
def test_ersetzung_am_zeilenende_excel_ohne_leerzeichen_wird_exakt(tmp_path):
    """Das Leerzeichen hinter "15" gehoert zur Einfuegung (Excel: "12mm"): der alte Absatz stimmt Zeichen fuer Zeichen, die
    Aenderung laeuft automatisch (vorher: A1 `a1_leerraum` + `excel_weicht_ab`, nur ein Vorschlag)."""
    plan, xl, xlsx, pdf, excel_b = _fall(tmp_path, "betraegt ", [("12", "d"), ("15", "n")], REST, "12")
    (ae,) = plan.aenderungen
    assert ae.status == UEBERNOMMEN and not _a1(ae) and "excel_weicht_ab" not in ae.pruefcodes, (ae.status, ae.pruefcodes)
    s = _zelle(excel_b).index("12")
    assert _ops(ae) == [(s, s + 2, "15 ")]                       # "12mm" -> "15 mm": das Leerzeichen kommt mit der Einfuegung


@pytest.mark.markierung_an
@pytest.mark.legende_neu
def test_ersetzung_am_zeilenende_ende_zu_ende_mit_markierung_und_zweitem_pruefweg(tmp_path):
    """Excel-Absatz im alten Zustand unveraendert, neuer Zustand = PDF neu, das fehlende Leerzeichen ist neu (rot); die Gegenprobe
    (A3, `wirkung`) und der zweite Pruefweg (Konsens-Text) sind gruen."""
    plan, xl, xlsx, pdf, excel_b = _fall(tmp_path, "betraegt ", [("12", "d"), ("15", "n")], REST, "12")
    out = tmp_path / "out.xlsx"
    uebernehmen(plan.aenderungen, xlsx, out, xl=xl, absaetze=plan.absaetze, pdf=pdf)
    (ae,) = plan.aenderungen
    assert ae.status == UEBERNOMMEN and ae.wirkung == "ok", (ae.status, ae.wirkung, ae.pruefcodes)
    assert not any(c.startswith(("gegenprobe", "konsens")) for c in ae.pruefcodes)
    st = mk.stuecke_aus_zelle(load_workbook(str(out), rich_text=True)["Abgleich"].cell(ZEILE_B, G))
    assert mk.ansicht(st, "alt") == _zelle(excel_b)                            # alter Zustand unveraendert
    assert mk.ansicht(st, "neu") == _zelle(excel_b).replace("12mm", "15 mm")     # neuer Zustand = PDF neu
    assert [[s.format, s.text] for s in st if s.format != mk.NORMAL] == [["alt", "12"], ["neu", "15 "]]   # Leerzeichen: neu (rot)
    assert load_workbook(str(xlsx))["Abgleich"].cell(ZEILE_B, G).value == _zelle(excel_b)    # die Original-Excel bleibt
    assert ae.stuecke and ["neu", "15 "] in ae.stuecke                           # Vorschau = spaeteres Ergebnis


@pytest.mark.markierung_an
@pytest.mark.legende_neu
def test_ersetzung_mit_markiertem_leerzeichen_vorn_und_neuer_wortgruppe(tmp_path):
    """Form aus dem Befund: gestrichen " 12" (beginnt mit dem markierten Leerzeichen), neu " ca. 15" (mit markiertem Leerzeichen
    in der Mitte), dann der Umbruch und "mm" (die Excel: "12mm")."""
    plan, xl, xlsx, pdf, excel_b = _fall(tmp_path, "betraegt", [(" 12", "d"), (" ca. 15", "n")], REST, " 12")
    (ae,) = plan.aenderungen
    assert ae.status == UEBERNOMMEN and not _a1(ae) and "excel_weicht_ab" not in ae.pruefcodes, (ae.status, ae.pruefcodes)
    out = tmp_path / "out.xlsx"
    uebernehmen(plan.aenderungen, xlsx, out, xl=xl, absaetze=plan.absaetze, pdf=pdf)
    assert ae.status == UEBERNOMMEN and ae.wirkung == "ok", (ae.status, ae.wirkung, ae.pruefcodes)
    st = mk.stuecke_aus_zelle(load_workbook(str(out), rich_text=True)["Abgleich"].cell(ZEILE_B, G))
    assert mk.ansicht(st, "alt") == _zelle(excel_b)
    assert mk.ansicht(st, "neu") == _zelle(excel_b).replace(" 12mm", " ca. 15 mm")


@pytest.mark.legende_neu
def test_excel_mit_leerzeichen_bleibt_gemeinsam_wie_bisher(tmp_path):
    """Steht das Leerzeichen in der Excel ("12 mm"), gehoert es wie bisher beiden Fassungen: nur "12" -> "15". Die Excel kann
    "gemeinsam" und "nur alt" nicht unterscheiden -- dann aendert sich nichts."""
    plan, xl, xlsx, pdf, excel_b = _fall(tmp_path, "betraegt ", [("12", "d"), ("15", "n")], REST, "12 ")
    (ae,) = plan.aenderungen
    assert ae.status == UEBERNOMMEN and not _a1(ae) and "excel_weicht_ab" not in ae.pruefcodes, (ae.status, ae.pruefcodes)
    s = _zelle(excel_b).index("12")
    assert _ops(ae) == [(s, s + 2, "15")]


@pytest.mark.legende_neu
def test_ohne_umbruch_bleibt_fehlendes_leerzeichen_eine_abweichung(tmp_path):
    """Steht das Leerzeichen sichtbar (unmarkiert) hinter der Ersetzung, gehoert es beiden Fassungen. Hat die Excel es nicht, ist
    das eine Abweichung der Excel (kein Umbruch, also keine Mehrdeutigkeit): nur ein Vorschlag, nie automatisch."""
    plan, xl, xlsx, pdf, excel_b = _fall(tmp_path, "betraegt ", [("12", "d"), ("15", "n")], " mm im Betrieb.", "12", umbruch=False,
                                         excel_rest="mm im Betrieb.")
    (ae,) = plan.aenderungen
    assert ae.status == PRUEFEN and "excel_weicht_ab" in ae.pruefcodes and "a1_leerraum" in ae.pruefcodes, (ae.status, ae.pruefcodes)


# ================================================================================================ Spiegelfall
@pytest.mark.markierung_an
@pytest.mark.legende_neu
def test_spiegelfall_gestrichenes_leerzeichen_am_umbruch_wird_nie_still_falsch(tmp_path):
    """Spiegelbild: Excel "15 mm", Word streicht "15 " (das Leerzeichen am Umbruch gehoert zur Streichung) und schreibt "12" direkt
    vor "mm" -> neu "12mm". Die Excel hat das Leerzeichen, "gemeinsam" und "nur alt" geben denselben alten Absatz: die Excel kann
    sie nicht unterscheiden, also bleibt es bei "gemeinsam" (die Auslegung "nur neu" passt nicht). Der Ablauf bleibt wie bisher:
    entweder steht das Ergebnis richtig da, oder es ist ein Fall zum Pruefen (die Gegenprobe haelt die Datei unveraendert). Nie
    automatisch uebernommen und falsch."""
    plan, xl, xlsx, pdf, excel_b = _fall(tmp_path, "betraegt ", [("15", "d")], "mm im Betrieb.", "15 ", zeile2=[("12", "n")])
    out = tmp_path / "out.xlsx"
    uebernehmen(plan.aenderungen, xlsx, out, xl=xl, absaetze=plan.absaetze, pdf=pdf)
    (ae,) = plan.aenderungen
    st = mk.stuecke_aus_zelle(load_workbook(str(out), rich_text=True)["Abgleich"].cell(ZEILE_B, G))
    if ae.status == UEBERNOMMEN:
        assert mk.ansicht(st, "alt") == _zelle(excel_b) and mk.ansicht(st, "neu") == _zelle(excel_b).replace("15 mm", "12mm")
    else:
        assert ae.status == PRUEFEN and any(c.startswith("gegenprobe") for c in ae.pruefcodes), (ae.status, ae.pruefcodes)
        assert load_workbook(str(out), rich_text=True)["Abgleich"].cell(ZEILE_B, G).value == _zelle(excel_b)    # Zelle unveraendert


def test_mehrdeutige_trenner_und_auslegung_einzeln():
    """Bausteine: `_Aufbau.mehrdeutig` nennt den unmarkierten Trenner am Umbruch neben der Ersetzung; `auslegung_waehlen` nimmt
    "nur neu" nur, wenn der alte Absatz damit (und nur damit) passt."""
    from src.abgleich import Spur, _aufbau, auslegung_waehlen
    from src.models import ADD, DEL, KEEP, MOD, PdfBlock, PdfToken

    def tok(old, new, status, x, y, teile=None):
        return PdfToken(old, new, status, 1, (x, y, x + 20, y + 8), trenner=" ", teile=teile or [])
    toks = [tok("Wert", "Wert", KEEP, 100, 100), tok("12", "15", MOD, 130, 100, [(DEL, "12"), (ADD, "15")]),
            tok("mm", "mm", KEEP, 100, 110)]                            # "mm" steht in der naechsten Zeile
    block = PdfBlock(0, 1, "absatz", toks, (0, 0, 0, 0))
    sp = Spur(1, [block])
    assert sp.aufbau.mehrdeutig == [1] and sp.alt == "Wert 12 mm"
    assert not auslegung_waehlen(sp, lambda s: s.alt == "Wert 12 mm")                  # "gemeinsam" passt: bleibt
    assert sp.alt == "Wert 12 mm" and not sp.auslegung
    assert auslegung_waehlen(sp, lambda s: s.alt == "Wert 12mm")                       # nur "nur neu" passt
    assert sp.alt == "Wert 12mm" and sp.neu == "Wert 15 mm" and sp.auslegung == {1: "a"}
    sp2 = Spur(2, [block])
    assert not auslegung_waehlen(sp2, lambda s: True)                                  # mehrere passen: keine Wahl
    assert not auslegung_waehlen(sp2, lambda s: False)                                 # keine passt: keine Wahl
    assert sp2.auslegung == {} and sp2.alt == "Wert 12 mm"
    toks3 = [tok("Wert", "Wert", KEEP, 100, 100), tok("12", "15", MOD, 130, 100, [(DEL, "12"), (ADD, "15")]),
             tok("mm", "mm", KEEP, 160, 100)]                                          # gleiche Zeile: nichts auszulegen
    assert _aufbau(toks3).mehrdeutig == []


# ================================================================================================ Generator v3: Stoerer
KLEIN = {"basis": "word_markup", "name": "klein", "zeilen": (45, 60)}
STOERER = "umbruch_leerzeichen_ersetzung"
# Regel 29: Paare der Seeds < 401 aendern sich durch den neuen Stoerer nie. Hashes des LOGISCHEN MODELLS (Absaetze, Texte, Aenderungen,
# Stoerer) einiger Seeds, vor dem Stoerer erzeugt -- unabhaengig von Bibliotheksversionen (anders als die Bytes der PDF-/Excel-Dateien;
# "scripts/erzeuge_paare.py --seeds 301-360" erzeugt bytegleiche Paare, einmalig geprueft).
MODELL_HASHES = {(301, False): "8edf9de69737dd980fb1da2e8cba8298d356dd726712e42f57c4c864cac0f03f",
                 (330, False): "8c69cdadee5978b0178ccf3f3cc16798726700feb690659dcf1540e98e218311",
                 (360, False): "5b04d1c6f9ae4d10c830d6cedd447e9c0596353b2b28382303d25a8ce2ee7a57",
                 (371, True): "48e9527237dd2a01faa253eff50dfff1af40c961a0ea32925f6b10f4eb7156c6",
                 (385, True): "ed76287960f34542c4c4b7a062d9daf264540f157058a54b2aef9d654bb31f6d"}


def _modell_hash(seed: int, klein: bool) -> str:
    ctx = wm.baue_kontext(seed, KLEIN if klein else None)
    daten = [[a.id, a.art, a.text, [(e.start, e.ende, e.neu, e.autor, e.art) for e in a.edits], a.excel_alt, sorted(a.stoerer),
              a.label_alt, a.label_neu, a.status] for a in ctx.absaetze]
    return hashlib.sha256(json.dumps([daten, sorted(ctx.stoerer)], ensure_ascii=False).encode("utf-8")).hexdigest()


def test_stoerer_wirkt_nur_ab_seed_401_im_korpus_nachbau_word():
    assert wm.UMBRUCH_LEERZEICHEN_AB == 401 and wm.STOERER_UMBRUCH_LEERZEICHEN == STOERER
    for seed in (301, 360, 371):
        assert STOERER not in wm.baue_kontext(seed, KLEIN).stoerer                 # bestehende Seeds: nie
    assert STOERER not in wm.baue_kontext(405, KLEIN, "nachbau_test").stoerer       # andere Korpora: nie
    gefunden = [s for s in range(401, 412) if STOERER in wm.baue_kontext(s, KLEIN).stoerer]
    assert len(gefunden) >= 8, gefunden                                              # ab 401 fast immer (2 bis 4 Absaetze je Paar)
    assert STOERER in wm.baue_kontext(371, {**KLEIN, "umbruch_leerzeichen": True}).stoerer      # nur fuer Tests des Generators


def test_modell_der_seeds_unter_401_bleibt_unveraendert():
    """Waechter (Regel 29): der Stoerer aendert das Modell bestehender Seeds nicht (also auch nicht ihre Paare)."""
    for (seed, klein), h in MODELL_HASHES.items():
        assert _modell_hash(seed, klein) == h, seed


def test_stoerer_paar_aufbau_und_wahrheit(tmp_path):
    from tests.test_nachbau_word import _laeufe
    cfg = {**KLEIN, "umbruch_leerzeichen": True}
    p = wm.erzeuge(371, tmp_path / "fam", "nachbau_word", cfg, name="fam")
    e = json.loads((p / "erwartet.json").read_text(encoding="utf-8"))
    assert STOERER in e["stoerer"]
    mit = [x for x in e["aenderungen"] if STOERER in x["stoerer"]]
    assert 2 <= len(mit) <= 4
    assert all(x["soll_status"] == "uebernommen" and x["typ"] == "text_ersetzt" and x["woerter"] == 2 for x in mit)
    absaetze = [a for a in wm.baue_kontext(371, cfg).absaetze if STOERER in a.stoerer]
    assert len(absaetze) == len(mit)
    vorher = load_workbook(p / "vorher.xlsx")[e["hauptblatt"]]
    nachher = load_workbook(p / "nachher.xlsx", rich_text=True)[e["hauptblatt"]]
    pos = {x["absatz"]: (x["zeile"], x["nachher_zeile"]) for x in mit}
    with pdfplumber.open(p / "aenderungen.pdf") as pdf:
        woerter = [w for pg in pdf.pages for w in pg.extract_words(keep_blank_chars=False)]
    for a in absaetze:
        (ed,) = a.edits
        wert, einheit = a.text[ed.start:ed.ende], a.text[ed.ende:].split()[0]
        zv, zn = pos[a.id]
        zelle = vorher.cell(zv, 7).value
        assert wert + einheit in zelle and (wert + " " + einheit) not in zelle                                 # Excel: "12mm"
        assert ed.neu == ed.neu.strip() + " " and ed.neu.strip().isdigit()                                    # Einfuegung "15 "
        # Umbruch dahinter: die Einheit steht in der Zeile darunter, weiter links (am Zeilenanfang)
        ende = [w for w in woerter if w["text"].endswith(wert + ed.neu.strip())]
        assert any(abs(u["top"] - (w["top"] + 10.0)) < 3 and u["x0"] < w["x0"] and u["text"].startswith(einheit)
                   for w in ende for u in woerter), (wert, einheit)
        laeufe = _laeufe(nachher.cell(zn, 7))
        assert ("alt", wert) in laeufe and ("neu", ed.neu) in laeufe


@pytest.mark.legende_neu
def test_stoerer_paar_wird_gemessen_ohne_stille_fehler_und_automatisch(tmp_path):
    """Echtes Tool am Paar mit dem Stoerer: keine stillen Fehler, kein Kollateral, alle Absaetze mit Stoerer laufen automatisch."""
    ziel = tmp_path / "nachbau_word"
    p = wm.erzeuge(371, ziel, "nachbau_word", {**KLEIN, "umbruch_leerzeichen": True}, name="fam")
    (ziel / "KORPUS.json").write_text(json.dumps({"generator_version": "3", "korpus": "nachbau_word", "seeds": [371],
                                                  "legende": "auto", "nachher_format": "markiert"}), encoding="utf-8")
    r = messen.messe_paar(p)
    assert "abgestuerzt" not in r and r["stille_fehler"] == 0 and r["kollateral"] == 0 and r["bilanz_stimmt"] is True
    assert not any(c.startswith("a1_leerraum") for c in r.get("pruefcodes", {})), r.get("pruefcodes")
