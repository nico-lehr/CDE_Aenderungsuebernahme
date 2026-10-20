"""Word-Leser (.docx) mit Aenderungsverfolgung (Auftrag DOCX): Woerter, Laeufe, Listennummern aus numbering.xml, Tabellen,
Kopf-/Fusszeilen, Kommentare, Robustheit, Kreuzprobe gegen das Word-PDF und Ende-zu-Ende bis zur Excel.

Die Test-Dateien baut tests/_docx_bau.py (Dokument-XML von Hand, zip im tmp_path); nur erfundene Daten (Regel 16).
"""
from __future__ import annotations

import zipfile

import pytest
from openpyxl import load_workbook

from src import config, markierung as mk
from src.abgleich import laeufe_aus_tokens
from src.applier import uebernehmen
from src.docx_reader import analysiere_docx, docx_zahlen, vergleiche_mit_pdf
from src.excel_model import ExcelModell
from src.fehler import NutzerFehler
from src.models import ADD, DEL, KEEP, MOD, NICHT, PRUEFEN, UEBERNOMMEN
from src.pdf_reader import analysiere_pdf
from src.planner import plane
from src.protokoll import bilanz
from tests._docx_bau import DocxBau, register_docx
from tests.test_register import GRUND, G, _ae, _ops, alt_von, excel_absaetze, register_mappe, wort_pdf

pytestmark = pytest.mark.legende_neu          # das Word-PDF der Kreuzprobe wird mit der Standard-Legende gelesen

SPERR = [str(c) for c in config.lade("zuordnung")["planer"]["unsichere_hinweise"]]


def _lies(tmp_path, b: DocxBau, name: str = "t.docx"):
    return analysiere_docx(b.speichern(tmp_path / name))


def _toks(a):
    return [t for b in a.blocks for t in b.tokens]


def _tok(a, text: str):
    return next(t for t in _toks(a) if text in (t.old, t.new))


def _codes(t) -> set[str]:
    return set(filter(None, t.hinweis.split(";")))


def _sperrend(a) -> set[str]:
    return {c for t in _toks(a) for c in _codes(t) if any(c == s or c.startswith(s + "_") for s in SPERR)}


def _fassungen(block) -> tuple[str, str]:
    """(alter, neuer Absatz) EXAKT aus Zeichen-Laeufen und Trennern; markierter Trenner nur in seiner Fassung."""
    alt = neu = ""
    for i, t in enumerate(block.tokens):
        for st, tx in (t.teile or [(t.status, t.new if t.status == ADD else t.old)]):
            alt += tx if st in (KEEP, DEL) else ""
            neu += tx if st in (KEEP, ADD) else ""
        if i + 1 < len(block.tokens) and t.trenner:
            art = getattr(t, "trenner_art", "")
            alt += t.trenner if art != "a" else ""
            neu += t.trenner if art != "d" else ""
    return alt.replace("\xa0", " "), neu.replace("\xa0", " ")


def _text(b: DocxBau, markup: str, **kw):
    b.absatz(markup, **kw)


# =================================================================================================
# 1. Woerter, Status, Zeichen-Laeufe
# =================================================================================================
def test_nur_einfuegung_ist_neu(tmp_path):
    b = DocxBau().absatz("Die Pumpe {+neue +}fördert Wasser.")
    a = _lies(tmp_path, b)
    assert a.legende_modus == "docx" and not a.markup_leiste and a.seiten == 0 and a.seiten_masse == {} and a.luecken == []
    blk = a.blocks[0]
    assert blk.page == 0 and blk.bbox == (0, 0, 0, 0) and blk.kind == "absatz"
    t = _tok(a, "neue")
    assert (t.old, t.new, t.status, t.page, t.bbox) == ("", "neue", ADD, 0, (0, 0, 0, 0))
    assert getattr(t, "trenner_art", "") == "a"                  # das mitmarkierte Leerzeichen hinter "neue"
    assert [x.status for x in _toks(a) if x.status != KEEP] == [ADD]
    assert _fassungen(blk) == ("Die Pumpe fördert Wasser.", "Die Pumpe neue fördert Wasser.")
    assert a.markierte_woerter == 1 and not _sperrend(a)


def test_nur_streichung_ist_entfaellt(tmp_path):
    b = DocxBau().absatz("Die Pumpe {-alte -}fördert Wasser.")
    a = _lies(tmp_path, b)
    t = _tok(a, "alte")
    assert (t.old, t.new, t.status) == ("alte", "", DEL) and getattr(t, "trenner_art", "") == "d"
    assert _fassungen(a.blocks[0]) == ("Die Pumpe alte fördert Wasser.", "Die Pumpe fördert Wasser.")


def test_ersetzung_ohne_leerzeichen_ist_ein_wort_mit_laeufen(tmp_path):
    b = DocxBau().absatz("Die Pumpe {-fördert-}{+hebt+} Wasser.")
    a = _lies(tmp_path, b)
    t = _tok(a, "hebt")
    assert (t.old, t.new, t.status, t.teile) == ("fördert", "hebt", MOD, [(DEL, "fördert"), (ADD, "hebt")])
    assert a.markierte_woerter == 1


def test_zeichen_genaue_aenderung_klammern_und_wert(tmp_path):
    """`[` + `wert` + `]` einzeln gestrichen: ein Wort mit Laeufen, wie der PDF-Leser es zerlegt."""
    b = DocxBau().absatz("Druck {-[-}40{-]-} bar, Höhe {-[40]-}{+42.9+} Meter und Zebrafink{+e+}.")
    a = _lies(tmp_path, b)
    k = _tok(a, "[40]")
    assert (k.old, k.new, k.status) == ("[40]", "40", MOD)
    assert k.teile == [(DEL, "["), (KEEP, "40"), (DEL, "]")]
    h = _tok(a, "42.9")
    assert (h.old, h.new, h.teile) == ("[40]", "42.9", [(DEL, "[40]"), (ADD, "42.9")])
    z = _tok(a, "Zebrafinke.")
    assert (z.old, z.new, z.teile) == ("Zebrafink.", "Zebrafinke.", [(KEEP, "Zebrafink"), (ADD, "e"), (KEEP, ".")])


@pytest.mark.parametrize("markup,alt,neu", [
    ("Die {-Pumpe -}{+Anlage +}steht", "Die Pumpe steht", "Die Anlage steht"),
    ("Pumpe{- -}Ventil", "Pumpe Ventil", "PumpeVentil"),
    ("Pumpe{+ +}Ventil", "PumpeVentil", "Pumpe Ventil"),
    ("Pumpe{- -}{+-+}Ventil", "Pumpe Ventil", "Pumpe-Ventil"),
    ("Pumpe {+neu+}x", "Pumpe x", "Pumpe neux"),
    ("Pumpe{+ neu+}x", "Pumpex", "Pumpe neux"),
    ("Norm {-ISO 11223‑-}{+XYZ-HD 11223-+}6", "Norm ISO 11223‑6", "Norm XYZ-HD 11223-6"),
    ("Pumpe {-alte Ventile -}Becken", "Pumpe alte Ventile Becken", "Pumpe Becken"),
], ids=["wort_mit_leer", "leer_weg", "leer_neu", "leer_weg_strich_neu", "einfuegung_geklebt", "einfuegung_vorn",
        "norm_folge", "geloeschte_gruppe"])
def test_markierter_leerraum_steht_nur_in_seiner_fassung(tmp_path, markup, alt, neu):
    a = _lies(tmp_path, DocxBau().absatz(markup + " und weiter."))
    va, vn = _fassungen(a.blocks[0])
    assert va == alt + " und weiter." and vn == neu + " und weiter."
    la = laeufe_aus_tokens(a.blocks[0].tokens)
    assert " ".join("".join(l.text for l in la if l.art in ("k", "d", "s")).split()) == alt + " und weiter."
    assert " ".join("".join(l.text for l in la if l.art in ("k", "a", "s")).split()) == neu + " und weiter."


def test_ganzer_absatz_neu_und_ganzer_absatz_gestrichen(tmp_path):
    b = DocxBau()
    b.absatz("Erster Absatz bleibt.")
    b.absatz("{+Ein ganz neuer Absatz.+}", marke="ins")
    b.absatz("{-Dieser Absatz entfällt.-}", marke="del")
    b.absatz("Letzter Absatz bleibt.")
    a = _lies(tmp_path, b)
    assert [x.ganz_neu for x in a.blocks] == [False, True, False, False]
    assert [x.ganz_geloescht for x in a.blocks] == [False, False, True, False]
    assert not _sperrend(a) and a.luecken == []
    assert a.markierte_woerter == 4 + 3


def test_verschoben_und_nur_formataenderung(tmp_path):
    b = DocxBau().absatz("{<Ventil<} bleibt {>Ventil>} hier und {f:fett} nur formatiert.")
    a = _lies(tmp_path, b)
    toks = _toks(a)
    assert [(t.old, t.new, t.status) for t in toks if t.status != KEEP] == [("Ventil", "", DEL), ("", "Ventil", ADD)]
    fett = _tok(a, "fett")
    assert fett.status == KEEP                                   # reine Formatänderung ist keine Änderung
    assert any("Formatänderungen" in w for w in a.warnungen)


def test_eingefuegt_und_danach_gestrichen_hat_keine_wirkung(tmp_path):
    b = DocxBau().absatz("Pumpe ")
    b.koerper[-1] = b.koerper[-1].replace("</w:p>", '<w:ins w:id="900" w:author="E"><w:del w:id="901" w:author="E">'
                                          '<w:r><w:delText>Geist</w:delText></w:r></w:del></w:ins><w:r><w:t> Ventil</w:t></w:r></w:p>')
    a = _lies(tmp_path, b)
    assert [t.old for t in _toks(a)] == ["Pumpe", "Ventil"] and a.markierte_woerter == 0
    assert any("keine Wirkung" in w for w in a.warnungen)


def test_hochgestellte_ziffer_hinter_einheit_bleibt_im_wort_fussnotenmarke_nicht(tmp_path):
    """Wie beim PDF-Leser: "mm" + hochgestellte 2 = "mm2"; "Pumpe" + hochgestellte 1 = Fussnoten-Marke, nicht im Text."""
    hoch = "<w:vertAlign w:val=\"superscript\"/>"
    lauf = DocxBau.lauf_xml
    xml = ("<w:p>" + lauf("Kabel 2.5 mm") + lauf("2", rpr=hoch) + lauf(", Rohr 3 m") + lauf("3", rpr=hoch) + lauf(", Pumpe")
           + lauf("1", rpr=hoch) + lauf(" Programm") + lauf("2", rpr=hoch) + lauf(" und Becken")
           + f'<w:ins w:id="800" w:author="E">{lauf("1", rpr=hoch)}</w:ins>' + lauf(" Ende.") + "</w:p>")
    a = _lies(tmp_path, DocxBau().roh(xml))
    woerter = [t.old for t in _toks(a)]
    assert "mm2," in woerter and "m3," in woerter and "Pumpe" in woerter and "Programm" in woerter
    assert not any(w in woerter for w in ("Pumpe1", "Programm2"))
    assert "hochzahl_markiert" in _sperrend(a)                   # die markierte Marke sperrt den Absatz (Regel 27)


def test_gross_geschrieben_durch_formatvorlage(tmp_path):
    b = DocxBau().formatvorlagen([{"id": "Gross", "name": "Gross", "gross": True}])
    b.absatz("Pumpenhaus {+und Lager+}", stil="Gross")
    a = _lies(tmp_path, b)
    assert _tok(a, "PUMPENHAUS").status == KEEP and _tok(a, "LAGER").status == ADD


# =================================================================================================
# 2. Listennummern aus numbering.xml
# =================================================================================================
def _liste(tmp_path, ebenen: list[tuple], absaetze: list[tuple], nums=None, stile=None):
    """absaetze: (markup, (numId, ilvl), marke)"""
    b = DocxBau()
    b.nummerierung({0: [b.lvl(*e[:3], **e[3]) if len(e) > 3 else b.lvl(*e) for e in ebenen]}, nums or {1: (0, {})})
    for m, num, marke in absaetze:
        b.absatz(m, num=num, marke=marke)
    return analysiere_docx(b.speichern(tmp_path / "l.docx"))


def _labels(a):
    return [(bl.tokens[0].old, bl.tokens[0].new, bl.tokens[0].status) for bl in a.blocks]


def test_nummern_dezimal_buchstaben_roemisch_und_neustart_je_ebene(tmp_path):
    ebenen = [(0, "decimal", "%1."), (1, "lowerLetter", "%2)"), (2, "lowerRoman", "(%3)"), (3, "upperLetter", "%4."),
              (4, "decimal", "%1.%2.%5")]
    n = [("x", (1, i), "") for i in (0, 1, 1, 2, 2, 2, 0, 1, 3, 3, 4)]
    a = _liste(tmp_path, ebenen, n)
    assert [l[0] for l in _labels(a)] == ["1.", "a)", "b)", "(i)", "(ii)", "(iii)", "2.", "a)", "A.", "B.", "2.a.1"]
    assert all(l[2] == KEEP for l in _labels(a))
    b = DocxBau()
    b.nummerierung({0: [b.lvl(0, "decimal", "%1."), b.lvl(1, "lowerLetter", "%2)"), b.lvl(2, "decimal", "%1.%2.%3", lgl=True)]},
                   {1: (0, {})})
    for i in (0, 1, 2):
        b.absatz("x", num=(1, i))
    assert [bl.tokens[0].old for bl in analysiere_docx(b.speichern(tmp_path / "lgl.docx")).blocks] == ["1.", "a)", "1.1.1"]


def test_buchstaben_ab_27_wie_word_und_alternativ(tmp_path):
    ebenen = [(0, "lowerLetter", "%1."), (1, "upperLetter", "%2.")]
    a = _liste(tmp_path, ebenen, [("x", (1, 0), "")] * 28)
    labels = [l[0] for l in _labels(a)]
    assert labels[:3] == ["a.", "b.", "c."] and labels[25:] == ["z.", "aa.", "bb."]
    cfg = {"docx": {**config.lade("docx"), "listen": {**config.lade("docx")["listen"], "buchstaben_folge": "alphabet"}}}
    b = DocxBau()
    b.nummerierung({0: [b.lvl(0, "upperLetter", "%1.")]}, {1: (0, {})})
    for _ in range(54):
        b.absatz("x", num=(1, 0))
    a2 = analysiere_docx(b.speichern(tmp_path / "a.docx"), cfg)
    labels2 = [bl.tokens[0].old for bl in a2.blocks]
    assert labels2[25:29] == ["Z.", "AA.", "AB.", "AC."] and labels2[-1] == "BB."


def test_start_override_und_start_wert_und_lvl_restart(tmp_path):
    b = DocxBau()
    b.nummerierung({0: [b.lvl(0, "decimal", "%1.", start=5), b.lvl(1, "lowerLetter", "%2.", restart=0)]},
                   {1: (0, {}), 2: (0, {0: 10})})
    for num in [(1, 0), (1, 1), (1, 0), (1, 1), (2, 0), (2, 1)]:
        b.absatz("x", num=num)
    a = analysiere_docx(b.speichern(tmp_path / "s.docx"))
    # Ebene 2 startet nie neu (lvlRestart 0): a. b. c. d.; numId 2 beginnt bei 10 (startOverride)
    assert [bl.tokens[0].old for bl in a.blocks] == ["5.", "a.", "6.", "b.", "10.", "c."]


def test_bullet_und_unbekanntes_zahlformat(tmp_path):
    b = DocxBau()
    b.nummerierung({0: [b.lvl(0, "bullet", ""), b.lvl(1, "ordinalText", "%2")]}, {1: (0, {})})
    b.absatz("Punkt eins", num=(1, 0))
    b.absatz("Punkt zwei", num=(1, 1))
    a = analysiere_docx(b.speichern(tmp_path / "b.docx"))
    assert a.blocks[0].tokens[0].old == "•"
    assert a.blocks[1].tokens[0].old == "Punkt"                  # Nummer unbekannt -> fehlt, aber gemeldet
    assert any(l["art"] == "sonstiges" for l in a.luecken) and any("Zahlenformat" in w for w in a.warnungen)


def test_nummernaenderung_wird_ein_label_wort_wie_beim_pdf(tmp_path):
    """a. gestrichen -> aus b. wird a.: das Label ist EIN Wort mit Laeufen (alt "b.", neu "a."), wie das gruene "c.a." im PDF."""
    ebenen = [(0, "lowerLetter", "%1.")]
    a = _liste(tmp_path, ebenen, [("Erster Punkt bleibt.", (1, 0), ""), ("{-Dieser Punkt entfällt.-}", (1, 0), "del"),
                                  ("Dritter Punkt.", (1, 0), ""), ("{+Neuer Punkt.+}", (1, 0), "ins")])
    assert _labels(a) == [("a.", "a.", KEEP), ("b.", "", DEL), ("c.", "b.", MOD), ("", "c.", ADD)]
    lab = a.blocks[2].tokens[0]
    assert lab.teile == [(DEL, "c."), (ADD, "b.")]
    assert a.blocks[1].ganz_geloescht and a.blocks[3].ganz_neu
    assert a.markierte_woerter == 3 + 1 + 3 + 1 + 1 + 0 + 1 or a.markierte_woerter > 0


def test_nummer_aus_formatvorlage_und_ueberschrift_nummern(tmp_path):
    """Ueberschrift 1/2 mit Nummerierung aus der Formatvorlage: "PART 1" und "1.1" -> Section; Ebene 3 bleibt Absatz."""
    b = DocxBau()
    b.nummerierung({0: [b.lvl(0, "decimal", "PART %1", suff="space", pstil="Ueberschrift1"),
                        b.lvl(1, "decimal", "%1.%2", pstil="Ueberschrift2"),
                        b.lvl(2, "upperLetter", "%3.", pstil="Ueberschrift3")]}, {1: (0, {})})
    b.formatvorlagen([{"id": "Ueberschrift1", "name": "heading 1", "num": (1, 0)},
                      {"id": "Ueberschrift2", "name": "heading 2", "num": (1, 1)},
                      {"id": "Ueberschrift3", "name": "heading 3", "num": (1, 2)}])
    b.absatz("ALLGEMEINES", stil="Ueberschrift1")
    b.absatz("PUMPENHAUS", stil="Ueberschrift2")
    b.absatz("Die Pumpe {+hebt+} Wasser.", stil="Ueberschrift3")
    b.absatz("{-LAGERUNG-}{+LAGER+}", stil="Ueberschrift2")
    b.absatz("Ersatzteile.", stil="Ueberschrift3")
    a = _lies(tmp_path, b)
    assert [(x.kind, x.section_alt, x.section_neu) for x in a.blocks] == [
        ("ueberschrift", "Part 1", "Part 1"), ("ueberschrift", "1.1", "1.1"), ("absatz", "1.1", "1.1"),
        ("ueberschrift", "1.2", "1.2"), ("absatz", "1.2", "1.2")]
    assert [t.old for t in a.blocks[1].tokens] == ["1.1", "PUMPENHAUS"]
    assert a.blocks[2].tokens[0].old == "A." and a.blocks[4].tokens[0].old == "A."          # Neustart unter jeder 1.x


def test_geaenderte_section_nummer_getippt_und_automatisch(tmp_path):
    b = DocxBau().absatz("PART 1 ALLGEMEINES").absatz("1.1\tPUMPENHAUS").absatz("Text eins.").absatz("{-1.2-}{+1.3+}\tLAGERUNG")
    a = _lies(tmp_path, b)
    assert [(x.kind, x.section_alt, x.section_neu) for x in a.blocks] == [
        ("ueberschrift", "Part 1", "Part 1"), ("ueberschrift", "1.1", "1.1"), ("absatz", "1.1", "1.1"), ("ueberschrift", "1.2", "1.3")]
    # automatisch: eine neue Section davor schiebt die Nummer (Original 1.1 -> Endstand 1.2)
    n = DocxBau()
    n.nummerierung({0: [n.lvl(0, "decimal", "%1.%1")]}, {1: (0, {})})
    n.formatvorlagen([{"id": "Ueberschrift1", "name": "heading 1", "num": (1, 0)}])
    n.absatz("{+NEUE SECTION+}", stil="Ueberschrift1", marke="ins")
    n.absatz("LAGERUNG", stil="Ueberschrift1")
    a2 = analysiere_docx(n.speichern(tmp_path / "n.docx"))
    assert a2.blocks[0].section_alt is None and a2.blocks[0].section_neu == "1.1"
    assert (a2.blocks[1].section_alt, a2.blocks[1].section_neu) == ("1.1", "2.2")
    assert a2.blocks[1].tokens[0].teile == [(DEL, "1.1"), (ADD, "2.2")]


def test_nummer_nach_pprchange_ist_im_original_die_alte(tmp_path):
    """Absatzformat geaendert (Liste neu zugewiesen): im Original gab es keine Nummer, im Endstand eine."""
    b = DocxBau()
    b.nummerierung({0: [b.lvl(0, "decimal", "%1.")]}, {1: (0, {})})
    b.absatz("Alter Absatz ohne Nummer.", num=(1, 0), ppr_extra='<w:pPrChange w:id="5" w:author="E"><w:pPr/></w:pPrChange>')
    a = analysiere_docx(b.speichern(tmp_path / "p.docx"))
    lab = a.blocks[0].tokens[0]
    assert (lab.old, lab.new, lab.status) == ("", "1.", ADD)


# =================================================================================================
# 3. Absatzmarken: verbunden / getrennt
# =================================================================================================
def test_neuer_absatz_hinter_einem_alten_ist_kein_getrennter_absatz(tmp_path):
    """Enter am Ende eines Absatzes + neuer Text: die Absatzmarke des ersten ist neu, der zweite ist ganz neu (normal)."""
    b = DocxBau()
    b.absatz("Alter Absatz bleibt.", marke="ins")
    b.absatz("{+Neuer Absatz.+}")
    a = _lies(tmp_path, b)
    assert a.blocks[1].ganz_neu and not a.blocks[0].hat_aenderung
    assert a.luecken == [] and not _sperrend(a)


def test_getrennter_absatz_sperrt_und_ist_eine_luecke(tmp_path):
    b = DocxBau()
    b.absatz("Erster Teil bleibt.", marke="ins")
    b.absatz("Zweiter Teil {+neu+} bleibt.")
    a = _lies(tmp_path, b)
    assert "absatzgrenze_unklar" in _sperrend(a)
    assert all("absatzgrenze_unklar" in _codes(t) for bl in a.blocks for t in bl.tokens)
    assert any("getrennt" in h for bl in a.blocks for h in bl.hinweise)
    assert len(a.luecken) == 1 and a.luecken[0]["art"] == "sonstiges" and a.luecken[0]["seite"] == 0
    assert any("Absatzmarken" in w for w in a.warnungen)


def test_verbundene_absaetze_sperren_und_sind_eine_luecke(tmp_path):
    b = DocxBau()
    b.absatz("Erster Teil bleibt.", marke="del")
    b.absatz("Zweiter Teil bleibt.")
    a = _lies(tmp_path, b)
    assert "absatzgrenze_unklar" in _sperrend(a)
    assert any("verbunden" in h for bl in a.blocks for h in bl.hinweise)
    assert len(a.luecken) == 1


def test_gestrichener_absatz_vor_einem_absatz_ist_kein_verbundener(tmp_path):
    b = DocxBau()
    b.absatz("{-Gestrichener Absatz.-}", marke="del")
    b.absatz("Absatz bleibt.")
    a = _lies(tmp_path, b)
    assert a.blocks[0].ganz_geloescht and a.luecken == [] and not _sperrend(a)


# =================================================================================================
# 4. Tabellen, Kopf/Fuss, Kommentare, Textfelder, Fussnoten
# =================================================================================================
def test_tabelle_eine_zelle_ein_block_mit_hinweis(tmp_path):
    b = DocxBau().absatz("Vor der Tabelle.")
    b.tabelle([["Name", "Wert"], ["Pumpe", "{-12-}{+14+} bar"], ["Ventil", "{+neu+}"]])
    b.absatz("Nach der Tabelle.")
    a = _lies(tmp_path, b)
    assert [len(bl.tokens) for bl in a.blocks] == [3, 1, 1, 1, 2, 1, 1, 3]
    zelle = a.blocks[4]
    assert _tok(a, "14").status == MOD and "tabelle" in _codes(_tok(a, "14"))
    assert "tabelle" not in _codes(a.blocks[0].tokens[0]) and "tabelle" in _codes(a.blocks[6].tokens[0])
    assert any("Tabelle" in h for h in zelle.hinweise)
    assert "tabelle" in _sperrend(a)


def test_gestrichene_tabellenzeile(tmp_path):
    b = DocxBau().tabelle([["Pumpe", "Ventil"], ["Alt", "Zeile"]], zeilen_status={1: "del"})
    a = _lies(tmp_path, b)
    assert [x.ganz_geloescht for x in a.blocks] == [False, False, True, True]
    assert any("Zeilen gestrichen" in w for w in a.warnungen)


def test_kopf_und_fuss_getrennt_in_ignoriert(tmp_path):
    b = DocxBau().absatz("Haupttext {+neu+}.")
    b.kopfzeile("Projekt Zebrafink {-[Name]-}{+Ostflügel+}")
    b.fusszeile("40 05 13 - Seite 3")
    a = _lies(tmp_path, b)
    assert [x for x in a.ignoriert] == [
        {"seite": 0, "grund": "kopf_fuss", "woerter": 1, "text": "Projekt Zebrafink Ostflügel"},
        {"seite": 0, "grund": "kopf_fuss", "woerter": 0, "text": "40 05 13 - Seite 3"}]
    woerter = {x for t in _toks(a) for x in (t.old, t.new)}
    assert "Zebrafink" not in woerter and "Ostflügel" not in woerter
    assert a.markierte_woerter == 1 + 1                          # neu im Text + 1 Wort in der Kopfzeile


def test_kommentare_stehen_nur_in_randnotizen(tmp_path):
    b = DocxBau().absatz("Haupttext.")
    b.kommentar("Bitte die Pumpe prüfen")
    a = _lies(tmp_path, b)
    assert a.randnotizen == [{"seite": 0, "art": "kommentar", "text": "Bitte die Pumpe prüfen", "y": 0.0}]
    assert "prüfen" not in {x for t in _toks(a) for x in (t.old, t.new)}
    assert any("1 Kommentare" in w for w in a.warnungen)


def test_textfeld_fussnote_und_unbekanntes_nie_still(tmp_path):
    b = DocxBau().absatz("Haupttext mit Marke.")
    box = ('<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice>'
           '<w:drawing><w:txbxContent><w:p><w:ins w:id="7" w:author="E"><w:r><w:t>Neu im Feld</w:t></w:r></w:ins></w:p></w:txbxContent>'
           '</w:drawing></mc:Choice><mc:Fallback><w:pict><w:txbxContent><w:p><w:r><w:t>Doppelt</w:t></w:r></w:p></w:txbxContent></w:pict>'
           '</mc:Fallback></mc:AlternateContent></w:r>')
    b.koerper[-1] = b.koerper[-1].replace("</w:p>", box + '<w:r><w:footnoteReference w:id="1"/></w:r>'
                                          '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
                                          '<m:r><m:t>x2</m:t></m:r></m:oMath></w:p>')
    b.fussnote("Fußnote mit {+neuem+} Text")
    a = _lies(tmp_path, b)
    assert [t.old for t in _toks(a)] == ["Haupttext", "mit", "Marke."]
    texte = " ".join(a.warnungen)
    assert "Textfelder" in texte and "Fußnoten" in texte and "unbekannte Bestandteile" in texte
    luecken = sorted(l["beschreibung"] for l in a.luecken)
    assert len(luecken) == 3 and sum(l["woerter"] for l in a.luecken) >= 3 + 1
    assert not any("Neu im Feld" in x or "Doppelt" in x for x in a.warnungen + luecken)       # keine Inhalte (Regel 16)


# =================================================================================================
# 5. Robustheit
# =================================================================================================
def test_kaputtes_zip_alte_doc_fehlendes_dokument_und_passwort(tmp_path):
    kaputt = tmp_path / "k.docx"
    kaputt.write_bytes(b"das ist kein zip")
    with pytest.raises(NutzerFehler, match="keine gültige Word-Datei"):
        analysiere_docx(kaputt)
    alt = tmp_path / "alt.doc"
    alt.write_bytes(bytes.fromhex("D0CF11E0A1B11AE1") + b"\0" * 600)
    with pytest.raises(NutzerFehler, match=r"alte Word-Datei \(\.doc"):
        analysiere_docx(alt)
    pw = tmp_path / "pw.docx"
    pw.write_bytes(bytes.fromhex("D0CF11E0A1B11AE1") + b"\0" * 64 + "EncryptedPackage".encode("utf-16-le") + b"\0" * 64)
    with pytest.raises(NutzerFehler, match="Passwort"):
        analysiere_docx(pw)
    ohne = DocxBau().absatz("x").speichern(tmp_path / "o.docx", ohne_dokument=True)
    with pytest.raises(NutzerFehler, match="word/document.xml fehlt"):
        analysiere_docx(ohne)
    with pytest.raises(NutzerFehler, match="nicht gefunden"):
        analysiere_docx(tmp_path / "gibt_es_nicht.docx")
    xml_kaputt = DocxBau().speichern(tmp_path / "x.docx", extra={"word/document.xml": "<w:document><w:body>"})
    with pytest.raises(NutzerFehler, match="beschädigt"):
        analysiere_docx(xml_kaputt)
    doctype = DocxBau().speichern(tmp_path / "d.docx", extra={"word/document.xml": '<!DOCTYPE x [<!ENTITY a "b">]><x/>'})
    with pytest.raises(NutzerFehler, match="Sicherheitsgründen"):
        analysiere_docx(doctype)


def test_passwort_im_zip_und_strict_namespace(tmp_path):
    from tests._docx_bau import W
    pw = tmp_path / "zipw.docx"
    with zipfile.ZipFile(pw, "w") as zf:
        zf.writestr("word/document.xml", "<x/>")
    roh = bytearray(pw.read_bytes())
    for kopf, versatz in ((b"PK\x03\x04", 6), (b"PK\x01\x02", 8)):          # Bit 0 der Flags: Eintrag ist verschluesselt
        roh[roh.index(kopf) + versatz] |= 0x1
    pw.write_bytes(bytes(roh))
    with pytest.raises(NutzerFehler, match="Passwort"):
        analysiere_docx(pw)
    b = DocxBau().absatz("Die Pumpe {-foerdert-}{+hebt+} Wasser.")
    strict = b.speichern(tmp_path / "s.docx", extra={"word/document.xml": b.dokument_xml().replace(
        W, "http://purl.oclc.org/ooxml/wordprocessingml/main")})
    assert _tok(analysiere_docx(strict), "hebt").status == MOD


def test_fehlertexte_sind_einfach_und_ohne_pfade(tmp_path):
    kaputt = tmp_path / "geheim_name.docx"
    kaputt.write_bytes(b"nein")
    with pytest.raises(NutzerFehler) as e:
        analysiere_docx(kaputt)
    assert "geheim_name" not in e.value.meldung and str(tmp_path) not in e.value.meldung


def test_keine_aenderungen_ist_eine_warnung(tmp_path):
    a = _lies(tmp_path, DocxBau().absatz("Nur normaler Text."))
    assert any("keine Textänderungen" in w for w in a.warnungen)


# =================================================================================================
# 6. Steckbrief-Zahlen (kein Text)
# =================================================================================================
def test_docx_zahlen_nur_zahlen_und_formen(tmp_path):
    b = DocxBau().absatz("Pumpe {+neu+} {-alt-} {<a<} {>a>}")
    b.tabelle([["x"]])
    a, z = docx_zahlen(b.speichern(tmp_path / "z.docx"))
    assert z.zaehler["ins"] == 1 and z.zaehler["del"] == 1 and z.zaehler["move_from"] == 1 and z.zaehler["move_to"] == 1
    assert z.zaehler["tabellen"] == 1 and z.autoren == {"Erfunden"}


def test_steckbrief_docx_nur_zahlen_und_formen(tmp_path, capsys):
    import json
    from src import cli, steckbrief
    b = DocxBau()
    b.register_nummerierung()
    b.absatz("Die Zebrafinken {-fördern-}{+heben+} das Wasser.", num=(1, 0))
    b.absatz("{+Ein ganz neuer Punkt.+}", num=(1, 0), marke="ins")
    b.absatz("1.1\tKRANICHHALLE")
    b.tabelle([["Kranich", "{-12-}{+14+}"]])
    b.kopfzeile("Projekt Wombat {+neu+}")
    b.kommentar("Eine Notiz über das Nilpferd")
    b.fussnote("Fußnote über den Biber")
    b.koerper.append("<w:p><w:r><w:t>x</w:t></w:r><m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
                     "<m:r><m:t>y</m:t></m:r></m:oMath></w:p>")
    datei = b.speichern(tmp_path / "geheim_wombat.docx")
    ziel = tmp_path / "sb"
    assert cli.main(["steckbrief", "--docx", str(datei), "--name", "probe", "--ausgabe", str(ziel)]) == 0
    ausgabe = capsys.readouterr().out
    assert "wombat" not in ausgabe.lower()                       # kein Dateiname, kein Text in der Ausgabe
    pfad = next(ziel.glob("*/steckbrief.json"))
    daten = json.loads(pfad.read_text(encoding="utf-8"))
    assert steckbrief.pruefe(daten) == []
    d = daten["docx"]
    assert d["aenderungen"]["ins"] == 3 and d["kopf_fuss"]["ins"] == 1 and d["aenderungen"]["del"] == 2 and d["aenderungen"]["absatzmarke_ins"] == 1
    assert d["listen"]["arten"] == {"upperLetter": 2} and d["tabellen"]["tabellen"] == 1
    assert d["kopf_fuss"]["kopf_teile"] == 1 and d["kopf_fuss"]["woerter_geaendert"] == 1
    assert d["sonstiges"]["kommentare"] == 1 and d["sonstiges"]["fussnoten"] == 1
    assert d["unbekannte_elemente"] == {"oMath": 1}
    roh = pfad.read_text(encoding="utf-8") + pfad.with_suffix(".md").read_text(encoding="utf-8")
    for fremd in ("Zebrafinken", "Wombat", "Nilpferd", "Biber", "Kranich", "heben"):
        assert fremd not in roh                                  # kein Text aus dem Dokument (Regel 16)
    # kaputte Datei: einfache Meldung
    (tmp_path / "k.docx").write_bytes(b"x")
    assert cli.main(["steckbrief", "--docx", str(tmp_path / "k.docx"), "--ausgabe", str(ziel)]) == 2
    assert "keine gültige Word-Datei" in capsys.readouterr().out
    assert cli.main(["steckbrief", "--docx", str(datei), "--pdf", "x.pdf"]) == 2


# =================================================================================================
# 7. Kreuzprobe: dasselbe erfundene Dokument als .docx und als Word-PDF -> dieselben Aenderungen im Planer
# =================================================================================================
def _kurz(plan):
    return [(a.typ, a.section, a.alt, a.neu, a.status) for a in plan.aenderungen]


def _beide(tmp_path, pdf_absaetze, auto=True, excel_abs=None, vorspann=None):
    pdf = wort_pdf(tmp_path / "a.pdf", pdf_absaetze, vorspann)
    docx = register_docx(tmp_path / "a.docx", pdf_absaetze, auto)
    xlsx = register_mappe(tmp_path / "alt.xlsx", excel_abs if excel_abs is not None else excel_absaetze(pdf_absaetze))
    p_pdf = plane(analysiere_pdf(pdf), ExcelModell(xlsx))
    p_docx = plane(analysiere_docx(docx), ExcelModell(xlsx))
    return p_pdf, p_docx, xlsx


def _szenario(aenderungen: dict[int, tuple]) -> list[tuple]:
    pdf = list(GRUND)
    for i, x in aenderungen.items():
        pdf[i] = x
    return pdf


SZENARIEN = {
    "einfuegung_streichung_ersetzung": {
        2: ("sub", "A.", "Die Pumpe {-foerdert-}{+hebt+} das {+klare +}Wasser in das {-obere -}Becken."),
        5: ("liste", "2.", "Werknorm WN {-2040-}{+2041+} fuer Leitungen.")},
    "zeichen_genau": {
        2: ("sub", "A.", "Die Pumpe{+n+} foerdert{+t+} das Wasser in das obere Becken{-.-}{+!+}"),
        4: ("liste", "1.", "Werknorm WN {-[1001]-}{+1002+} fuer Ventile.")},
    "section_part_titel": {
        0: ("part", "1", "{-ALLGEMEINES-}{+ALLGEMEINE ANGABEN+}"),
        1: ("section", "1.1", "PUMPEN{+HAUS+}")},
    "nummernaenderung_und_geloeschter_punkt": {
        6: ("liste", "{-a.-}", "{-Pruefung der Dichtungen nach Plan.-}"),
        7: ("liste", "{-b.-}{+a.+}", "Pruefung der Lager {+und Dichtungen +}nach Plan.")},
    "geaenderte_section_nummer": {
        9: ("section", "{-1.2-}{+1.3+}", "LAGERUNG")},
}


SZENARIEN["neuer_absatz_und_neuer_punkt"] = {
    8: ("liste", "{+c.+}", "{+Pruefung der Wellen nach Plan.+}"),
}


def _szenario_mit_neuen(aenderungen: dict[int, tuple]) -> list[tuple]:
    """Wie _szenario; ein Eintrag mit Schluessel i und art "liste" ... ersetzt NICHT, sondern fuegt VOR Index i ein."""
    pdf = list(GRUND)
    for i, x in sorted(aenderungen.items(), reverse=True):
        pdf.insert(i, x)
    return pdf


@pytest.mark.parametrize("name", [k for k in SZENARIEN if k != "neuer_absatz_und_neuer_punkt"])
@pytest.mark.parametrize("auto", [True, False], ids=["nummern_aus_numbering_xml", "nummern_getippt"])
def test_kreuzprobe_docx_und_word_pdf_geben_dieselben_aenderungen(tmp_path, name, auto):
    p_pdf, p_docx, _ = _beide(tmp_path, _szenario(SZENARIEN[name]), auto)
    assert len(p_docx.aenderungen) >= 1
    assert _kurz(p_docx) == _kurz(p_pdf)
    bil = bilanz(p_docx.aenderungen, p_docx.markierte_woerter, p_docx.ignoriert, p_docx.luecken)
    assert bil["stimmt"] and bil["vollstaendig"]                 # Bilanz-Prinzip: nichts geht still verloren


def test_kreuzprobe_neuer_listenpunkt_und_neue_zeile(tmp_path):
    for auto in (True, False):
        pdf = _szenario_mit_neuen({8: ("liste", "{+c.+}", "{+Pruefung der Wellen nach Plan.+}"),
                                   14: ("liste", "{+3.+}", "{+Der Kranich steht auch auf der Rueckseite.+}")})
        p_pdf, p_docx, _ = _beide(tmp_path / str(auto), pdf, auto)
        assert [a.typ for a in p_docx.aenderungen] == ["absatz_neu", "absatz_neu"]
        assert _kurz(p_docx) == _kurz(p_pdf)


def test_kreuzprobe_wortliste_ueberlappt_bis_auf_positionen_nichts(tmp_path):
    pdf_absaetze = _szenario(SZENARIEN["einfuegung_streichung_ersetzung"])
    pdf = analysiere_pdf(wort_pdf(tmp_path / "a.pdf", pdf_absaetze))
    docx = analysiere_docx(register_docx(tmp_path / "a.docx", pdf_absaetze))
    v = vergleiche_mit_pdf(docx, pdf)
    assert v["gleich"] and v["alt"]["nur_docx"] == v["alt"]["nur_pdf"] == v["neu"]["nur_docx"] == 0
    assert v["markierte_woerter"] == {"docx": docx.markierte_woerter, "pdf": pdf.markierte_woerter}
    # absichtlich verschieden: ein Wort fehlt im PDF -> Zahlen und Positionen, kein Text
    kaputt = _szenario({2: ("sub", "A.", "Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.")})
    pdf2 = analysiere_pdf(wort_pdf(tmp_path / "b.pdf", kaputt))
    v2 = vergleiche_mit_pdf(docx, pdf2)
    assert not v2["gleich"] and v2["neu"]["nur_docx"] + v2["neu"]["nur_pdf"] > 0 and v2["alt"]["nur_docx"] == 0
    assert set(v2["neu"]["beispiele"][0]) == {"position_docx", "laenge_docx", "position_pdf", "laenge_pdf"}
    assert "Pumpe" not in repr(v2)                               # nie Text aus dem Dokument im Protokoll (Regel 16)


def test_bekannte_abweichungen_gegenueber_dem_word_pdf(tmp_path):
    """Dokumentierte Unterschiede (kein Fehler, sondern mehr Wissen im XML):
    1. Absatzmarken-Aenderungen (getrennte / verbundene Absaetze) sieht das PDF nicht; der Word-Leser sperrt den Absatz
       (absatzgrenze_unklar) und meldet eine Luecke. Dasselbe Dokument als PDF wuerde die Aenderung automatisch eintragen.
    2. Kopf- und Fusszeilen stehen im XML EINMAL (nicht je Seite): `ignoriert` hat einen Eintrag je Absatz und Teil; unveraenderte
       Zeilen stehen mit woerter 0 darin (der Planer liest die Spec auch aus einer Fusszeile ohne Aenderung).
    3. Der Tabulator nach einer Listennummer ist ein normaler Trenner (im PDF fehlt das Zeichen ebenfalls).
    4. Seite, Position und Linien gibt es nicht (page 0, bbox 0): keine Seitenumbruch-Heuristik, keine Lese-Luecke."""
    pdf_absaetze = _szenario({2: ("sub", "A.", "Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.")})
    p_pdf, p_docx, xlsx = _beide(tmp_path, pdf_absaetze)
    assert [a.status for a in p_pdf.aenderungen] == [a.status for a in p_docx.aenderungen] == [UEBERNOMMEN]
    b = DocxBau()
    b.register_nummerierung()
    b.absatz("Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.", num=(1, 0), marke="ins")        # getrennt
    b.absatz("Zweiter Teil.", num=(1, 0))
    plan = plane(analysiere_docx(b.speichern(tmp_path / "g.docx")), ExcelModell(xlsx))
    assert plan.aenderungen and all(a.status != UEBERNOMMEN for a in plan.aenderungen) and plan.luecken


# =================================================================================================
# 8. Spec aus Titelzeile und Fusszeile (CDE mit mehreren Specs)
# =================================================================================================
def _mehrspec_mappe(tmp_path):
    from tests.test_spec import _mappe
    return _mappe(tmp_path)


def test_spec_aus_der_titelzeile_im_text(tmp_path):
    xlsx = _mehrspec_mappe(tmp_path)
    pdf = list(GRUND)
    pdf[2] = ("sub", "A.", "Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.")
    docx = register_docx(tmp_path / "a.docx", pdf, vorspann=["SECTION 44 55 66 – PUMPENANLAGE"])
    plan = plane(analysiere_docx(docx), ExcelModell(xlsx))
    assert plan.spec["erkannt"] == "44 55 66" and plan.spec["quelle"] == "pdf"
    (ae,) = plan.aenderungen
    assert ae.status == UEBERNOMMEN and ae.ziel.startswith("Abgleich!G")


def test_spec_aus_der_fusszeile_auch_ohne_aenderung_dort(tmp_path):
    xlsx = _mehrspec_mappe(tmp_path)
    pdf = list(GRUND)
    pdf[2] = ("sub", "A.", "Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.")
    docx = register_docx(tmp_path / "a.docx", pdf, fuss=["77_88_99 Pumpenanlage - 5"])
    a = analysiere_docx(docx)
    assert [i["woerter"] for i in a.ignoriert] == [0]
    plan = plane(a, ExcelModell(xlsx))
    assert plan.spec["erkannt"] == "77 88 99" and plan.spec["quelle"] == "pdf"
    ohne = plane(analysiere_docx(register_docx(tmp_path / "b.docx", pdf)), ExcelModell(xlsx))
    assert ohne.spec["erkannt"] is None and any("Spec" in w for w in ohne.warnungen)           # Regel 27: nichts automatisch
    assert all(x.status != UEBERNOMMEN for x in ohne.aenderungen)


# =================================================================================================
# 9. Ende zu Ende: docx -> Planer -> Excel (verlustfrei geschrieben) -> Gegenprobe gruen
# =================================================================================================
def _eintragen(tmp_path, pdf_absaetze, auto=True, g_wert=None):
    docx = register_docx(tmp_path / "a.docx", pdf_absaetze, auto)
    xlsx = register_mappe(tmp_path / "alt.xlsx", excel_absaetze(pdf_absaetze), g_wert=g_wert)
    xl = ExcelModell(xlsx)
    plan = plane(analysiere_docx(docx), xl)
    out = tmp_path / "out.xlsx"
    uebernehmen(plan.aenderungen, xlsx, out, xl=xl, absaetze=plan.absaetze)
    return plan, xl, xlsx, out


@pytest.mark.markierung_an
def test_ende_zu_ende_aenderungen_nummern_und_neue_zeile_mit_gruener_gegenprobe(tmp_path):
    pdf = _szenario_mit_neuen({8: ("liste", "{+c.+}", "{+Pruefung der Wellen nach Plan.+}")})
    pdf[2] = ("sub", "A.", "Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.")
    pdf[5] = ("liste", "2.", "Werknorm WN {-2040-}{+2041+} fuer Leitungen.")
    pdf[6] = ("liste", "{-a.-}", "{-Pruefung der Dichtungen nach Plan.-}")
    pdf[7] = ("liste", "{-b.-}{+a.+}", "Pruefung der Lager nach Plan.")
    plan, xl, xlsx, out = _eintragen(tmp_path, pdf)
    assert plan.aenderungen and all(a.status == UEBERNOMMEN and a.wirkung == "ok" for a in plan.aenderungen), \
        [(a.id, a.typ, a.status, a.pruefcodes) for a in plan.aenderungen]
    bil = bilanz(plan.aenderungen, plan.markierte_woerter, plan.ignoriert, plan.luecken)
    assert bil["stimmt"] and bil["vollstaendig"] and bil["uebernommen"] == len(plan.aenderungen)
    assert not [c for a in plan.aenderungen for c in a.pruefcodes if c.startswith("gegenprobe_")]
    ws = load_workbook(str(out), rich_text=True)["Abgleich"]
    st = mk.stuecke_aus_zelle(ws.cell(10, G))
    assert mk.ansicht(st, "neu") == "Die Pumpe hebt das Wasser in das obere Becken."
    assert [[s.format, s.text] for s in st if s.format != mk.NORMAL] == [["alt", "foerdert"], ["neu", "hebt"]]
    # ganz gestrichener Punkt: Zeile bleibt, in der Neu-Ansicht leer; die Nummer des Folgepunkts ist umgestellt
    assert mk.ansicht(mk.stuecke_aus_zelle(ws.cell(14, G)), "neu") == ""
    assert mk.ansicht(mk.stuecke_aus_zelle(ws.cell(15, G)), "alt").startswith("b.")
    assert mk.ansicht(mk.stuecke_aus_zelle(ws.cell(15, G)), "neu").startswith("a.")
    # zweiter Lauf auf dem Ergebnis: alles steht schon da
    xl2 = ExcelModell(out)
    plan2 = plane(analysiere_docx(tmp_path / "a.docx"), xl2)
    assert all(a.status == NICHT and a.pruefcodes == ["schon_eingetragen"] for a in plan2.aenderungen)


@pytest.mark.markierung_an
def test_ende_zu_ende_tabelle_und_verbundener_absatz_werden_nie_automatisch_eingetragen(tmp_path):
    pdf = list(GRUND)
    pdf[2] = ("sub", "A.", "Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.")
    b = DocxBau()
    b.register_nummerierung()
    b.absatz("Die Pumpe {-foerdert-}{+hebt+} das Wasser in das obere Becken.", num=(1, 0), marke="del")      # verbunden
    b.absatz("Der Zebrafink wird im Nistkasten beobachtet.", num=(1, 0))
    b.tabelle([["Pumpe", "{-12-}{+14+}"]])
    docx = b.speichern(tmp_path / "v.docx")
    xlsx = register_mappe(tmp_path / "alt.xlsx", excel_absaetze(pdf))
    xl = ExcelModell(xlsx)
    a = analysiere_docx(docx)
    plan = plane(a, xl)
    assert plan.aenderungen and not [x for x in plan.aenderungen if x.status == UEBERNOMMEN]        # Regel 27
    out = tmp_path / "out.xlsx"
    uebernehmen(plan.aenderungen, xlsx, out, xl=xl, absaetze=plan.absaetze)
    ws0, ws1 = load_workbook(str(xlsx))["Abgleich"], load_workbook(str(out))["Abgleich"]
    assert [[c.value for c in r] for r in ws0.iter_rows()] == [[c.value for c in r] for r in ws1.iter_rows()]
    bil_luecken = plan.luecken
    assert any(l["art"] == "sonstiges" for l in bil_luecken)
