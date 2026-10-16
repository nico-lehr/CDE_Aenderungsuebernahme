"""Register der Stoer-Familien (erfundene Testdaten, Regel 16).

Eine Stoer-Familie ist ein Sachverhalt, der in echten Dokumenten vorkommt und ein Tool aus dem Tritt
bringen kann (Fussnote, Silbentrennung, Scan-Seite, Kommentarspalte ...). Jede Familie hat:

  name          Kennung, steht in erwartet.json unter "stoerer"
  ebene         pdf | excel | beides
  beschreibung  was sie im Modell bzw. im PDF/in der Excel bewirkt
  soll_status   was ein sauberes Tool mit den betroffenen Aenderungen tun soll:
                  uebernommen        eindeutig, automatisch
                  pruefen            absichtlich mehrdeutig -> nur nach Freigabe
                  nicht_uebernommen  absichtlich nicht machbar (Scan, Zwei-Spalten, Tabelle ...)
  luecke        Art der Luecke (models.LUECKEN_ARTEN), die das Tool melden muss, oder None
  nur_test      True = kommt nie im Entwicklungskorpus vor (Regel 11), siehe NUR_TEST
  vorbereiten   Hook VOR dem Bau des Dokuments (setzt Optionen: Layout, PDF-Einstellungen, Textarten)
  anwenden      Hook NACH Struktur und Aenderungen (veraendert das Modell); False = nicht anwendbar
  garantie      Pruefung nach dem Zeichnen des PDFs (z. B. "es wurde wirklich getrennt"); schlaegt sie fehl,
                baut der Generator das Paar mit einem anderen (aber festen) Zufallsstrom neu

Die Hooks bekommen den Kontext aus generator.py (ctx.rng, ctx.opt, ctx.sektionen, ctx.layout ...).
Dieses Modul importiert nichts aus generator.py (kein Kreis).
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Callable

from .pdf_writer import zustand

# Arten, die das Tool als Luecke melden kann (Kopie von models.LUECKEN_ARTEN; ein Test vergleicht beide)
LUECKEN_ARTEN = ("scan_seite", "mehrspaltig", "tabelle", "bild", "unbekannte_farbe",
                 "nicht_lesbarer_text", "unerklaerte_markierung", "sonstiges")

BLAUTOENE = ["1F4E9F", "2E75B6", "0070C0", "2F5597", "4472C4"]
ROTTOENE = ["C00000", "E06666", "CC0000", "FF3333"]

# Stil -> strengster soll_status; Stil -> Familie (fuer die Markierung am Eintrag)
STIL_SOLL = {"bs": "pruefen", "ru": "pruefen", "gn": "nicht_uebernommen"}
STIL_FAMILIE = {"bu": "mischmarkierung", "bs": "mischmarkierung", "ru": "mischmarkierung",
                "gn": "gruene_markierung"}


RANG = {"uebernommen": 0, "pruefen": 1, "nicht_uebernommen": 2}


def schlimmer(a: str, b: str | None) -> str:
    """Der strengere von zwei soll_status."""
    return b if b and RANG[b] > RANG[a] else a


def wrap_stil(stil: str, text: str) -> str:
    if stil == "n":
        return "{+" + text + "+}"
    if stil == "g":
        return "{-" + text + "-}"
    return "{@" + stil + ":" + text + "@}"


@dataclass(frozen=True)
class Stoerer:
    name: str
    ebene: str
    beschreibung: str
    soll_status: str
    luecke: str | None = None
    nur_test: bool = False
    mild: bool = False                      # darf im Profil "standard" vorkommen
    reihenfolge: int = 0
    vorbereiten: Callable[[Any], Any] | None = None
    anwenden: Callable[[Any], Any] | None = None
    garantie: Callable[[Any, Any], bool] | None = None   # (ctx, pdf) -> Besonderheit steckt wirklich im PDF; sonst neuer Versuch


# ====================================================================== Hilfen
def _text_absaetze(ctx, ohne_marker: bool = True):
    return [(s, a) for s, a in ctx.alle_absaetze()
            if s.status is None and not s.scan and not s.zwei_spalten and a.art == "text"
            and not (ohne_marker and a.marker)]


_AUFZ_MARKER = ["•", "–", "1.", "a)"]


def _marker(art: str, i: int) -> str:
    if art == "1.":
        return f"{i + 1}."
    if art == "a)":
        return f"{chr(97 + i)})"
    return art


# ====================================================================== Hooks: PDF-Optionen
def _silben_vor(ctx) -> None:
    ctx.pdfkw["trennung"] = ctx.rng.choice(["ascii", "weich", "gemischt"])
    ctx.opt["lange_woerter"] = True


def _farbtoene_vor(ctx) -> None:
    ctx.pdfkw["variante"] = "farbe"
    ctx.pdfkw["farbe_neu"] = ctx.rng.choice(BLAUTOENE)
    ctx.pdfkw["farbe_weg"] = ctx.rng.choice(ROTTOENE)
    ctx.opt["farben_pro_absatz"] = ctx.rng.random() < 0.5


def _cmyk_vor(ctx) -> None:
    ctx.pdfkw.update(variante="farbe", farbraum="cmyk", grundfarbraum="gray")


def _kopf_vor(ctx) -> None:
    v = ctx.rng.choice(["nur_erste", "blau", "nummern"])
    if v == "nur_erste":
        ctx.pdfkw["kopf_fuss"] = {1}
    elif v == "blau":
        ctx.pdfkw["kopf_farbe"] = (0.12, 0.31, 0.62)
    else:
        ctx.pdfkw.update(kopf_text="Anlage 2.3 - Technische Spezifikation Rev. 4 (erfunden)",
                         fuss_vorlage="Dokument CDE-0815 - Blatt {nr}/{n}")


def _misch_vor(ctx) -> None:
    ctx.pdfkw["variante"] = "farbe"
    ctx.opt["stil_neu"] = ["n", "bu"]
    ctx.opt["stil_weg"] = ["g", "bs", "ru"]
    ctx.opt["misch_p"] = 0.45


def _gruen_vor(ctx) -> None:
    ctx.pdfkw["variante"] = "farbe"
    ctx.opt["stil_neu"] = ["n", "gn"]
    ctx.opt["misch_p"] = 0.4


def _link_vor(ctx) -> None:
    ctx.pdfkw["variante"] = "farbe"


def _verschmolzen_vor(ctx) -> None:
    ctx.pdfkw["variante"] = "farbe"
    ctx.opt["verschmolzen"] = True


def _sonder_vor(ctx) -> None:
    ctx.opt["sonder"] = True


def _rotiere_vor(ctx) -> None:
    ctx.opt["rotiere"] = True


# ====================================================================== Hooks: Excel-Layout
def _komm_vor(ctx) -> None:
    ctx.layout["kommentar"] = True
    ctx.layout["kommentar_vor"] = ctx.rng.random() < 0.5


def _lfd_vor(ctx) -> None:
    ctx.layout["lfd"] = True


def _zahl_vor(ctx) -> None:
    ctx.layout["zahl"] = True
    ctx.opt["sub_nummern"] = {2: [1, 2, 10], 3: [1, 10]}      # 2.1 und 2.10 sind als Zahl nicht unterscheidbar
    ctx.opt["geschuetzt"] = {"2.1", "2.2", "2.10", "3.1", "3.10"}


def _blaetter_vor(ctx) -> None:
    ctx.blaetter = "mehrere"


def _formeln_vor(ctx) -> None:
    ctx.layout.update(formeln=True, merges=True, banner=True)


# ====================================================================== Hooks: Modell
def _fussnote(ctx) -> bool:
    kand = [(s, a) for s, a in _text_absaetze(ctx) if a.markup[:1].isalpha()]
    if not kand:
        return False
    for _, a in ctx.rng.sample(kand, min(len(kand), ctx.rng.randint(2, 4))):
        a.fussnote = ctx.rng.choice(ctx.FUSSNOTEN)
        ctx.markiere(a, "fussnote")
    return True


def _liste_bauen(ctx, marker_art: str, kurz: bool, name: str) -> bool:
    r = ctx.rng
    kand = [(s, a) for s, a in _text_absaetze(ctx) if a.typ is None and not a.fussnote]
    if kand:
        s, a = r.choice(kand)
    else:
        s = ctx.kleine_section(False, 1)
        a = s.absaetze[0]
    n = r.randint(3, 5)
    items = [ctx.neuer_absatz(ctx.kurzer_punkt(kurz), marker=_marker(marker_art, i)) for i in range(n)]
    pos = s.absaetze.index(a) + 1
    s.absaetze[pos:pos] = items
    neu = ctx.neuer_absatz(ctx.kurzer_punkt(kurz), marker=_marker(marker_art, n))
    ctx.aendere(neu, "absatz_neu")
    ctx.markiere(neu, name)
    s.absaetze.insert(pos + n, neu)
    if r.random() < 0.5:
        ctx.aendere(items[r.randrange(n)], "text_ersetzt")
    for it in items:
        if it.typ:
            ctx.markiere(it, name)
    if r.random() < 0.4:
        weg = items[r.randrange(1, n)]
        if weg.typ is None and ctx.aendere(weg, "absatz_geloescht"):
            ctx.markiere(weg, name)
    return True


def _aufzaehlung(ctx) -> bool:
    return _liste_bauen(ctx, ctx.rng.choice(_AUFZ_MARKER), False, "aufzaehlung")


def _nummeriert(ctx) -> bool:
    return _liste_bauen(ctx, "1.", True, "nummerierte_liste")


def _doppelt(ctx) -> bool:
    r = ctx.rng
    kand = [(s, a) for s, a in _text_absaetze(ctx)
            if a.typ in ("text_ersetzt", "text_eingefuegt", "text_geloescht") and not a.fussnote]
    if not kand:
        return False
    s, a = r.choice(kand)
    modus = r.choice(["andere_section", "gleiche_section", "gleiche_section", "fast_gleich"])
    alt = zustand(a.markup, False)
    kopie = ctx.fast_gleich(alt) if modus == "fast_gleich" else alt
    d = ctx.neuer_absatz(kopie)
    ctx.benutzt.add(kopie)
    if modus == "andere_section":
        andere = [t for t in ctx.sektionen if t is not s and t.status is None and not t.scan and not t.zwei_spalten]
        if not andere:
            modus = "gleiche_section"
        else:
            t = r.choice(andere)
            t.absaetze.insert(r.randint(0, len(t.absaetze)), d)
    if modus != "andere_section":
        i = s.absaetze.index(a)
        s.absaetze.insert(i + r.choice([0, 1]), d)
        if modus == "gleiche_section":
            ctx.markiere(a, "doppelte_absaetze", "pruefen")
    ctx.markiere(a, "doppelte_absaetze")
    ctx.markiere(d, "doppelte_absaetze")
    return True


def _zwei_spalten(ctx) -> bool:
    kand = [s for s in ctx.sektionen
            if s.status is None and not s.scan and not s.zwei_spalten
            and sum(1 for a in s.absaetze if a.art == "text") >= 2
            and all(a.art == "text" and not a.marker and not a.fussnote for a in s.absaetze)
            and any(a.typ for a in s.absaetze) and ctx.section_zeilen(s) <= 30]
    s = ctx.rng.choice(kand) if kand else ctx.kleine_section(True, 3)
    s.zwei_spalten = True
    ctx.markiere(s, "zwei_spalten")
    return True


def _scan(ctx) -> bool:
    kand = [s for s in ctx.sektionen
            if not s.scan and not s.zwei_spalten and s.absaetze
            and all(a.art == "text" and not a.marker and not a.fussnote for a in s.absaetze)
            and (s.status or any(a.typ for a in s.absaetze)) and ctx.section_zeilen(s) <= 26]
    s = ctx.rng.choice(kand) if kand else ctx.kleine_section(True, 3)
    s.scan = True
    ctx.markiere(s, "scan_seite")
    return True


def _tabelle(ctx) -> bool:
    r = ctx.rng
    kand = [s for s in ctx.sektionen if s.status is None and not s.scan and not s.zwei_spalten]
    s = r.choice(kand)
    zeilen = [["Pos.", "Bauteil", "Menge", "Einheit"]]
    for i in range(r.randint(3, 5)):
        zeilen.append([str(i + 1), r.choice(ctx.BAUTEILE), str(r.choice([2, 4, 6, 8, 12, 24])), "Stück"])
    stil = r.choice([("4472C4", "B4C6E7"), ("4472C4", None), ("1F4E9F", "B4C6E7"), (None, "B4C6E7")])
    a = ctx.neuer_absatz("", art="tabelle", tabelle=zeilen, tabelle_stil=stil)
    if r.random() < 0.7:
        i = r.randrange(1, len(zeilen))
        alt = zeilen[i][2]
        zeilen[i][2] = ctx.tausch(a, alt, ctx.ersatz(alt))
        a.typ = "text_ersetzt"
        ctx.markiere(a, "tabelle_mit_farbkopf")
    s.absaetze.insert(r.randint(0, len(s.absaetze)), a)
    return True


def _bild(ctx) -> bool:
    kand = [s for s in ctx.sektionen if s.status is None and not s.scan and not s.zwei_spalten]
    s = ctx.rng.choice(kand)
    a = ctx.neuer_absatz("", art="bild", bild_seed=ctx.rng.randint(0, 10 ** 6))
    s.absaetze.insert(ctx.rng.randint(0, len(s.absaetze)), a)
    return True


def _hyperlink(ctx) -> bool:
    kand = [(s, a) for s, a in _text_absaetze(ctx)
            if a.typ in (None, "text_ersetzt", "text_eingefuegt", "text_geloescht") and "{@" not in a.markup]
    if not kand:
        return False
    namen = ["Datenblatt DB-07", "Norm EN-1234", "Merkblatt MB-15", "Prüfvorschrift PV-02"]
    for _, a in ctx.rng.sample(kand, min(len(kand), 2)):
        a.markup += " Siehe {@l:" + ctx.rng.choice(namen) + "@}."
        ctx.markiere(a, "hyperlink")
    return True


_ALT_NEU = re.compile(r"(\{-.*?-\}|\{@(?:bs|ru|r|s):.*?@\}) (\{\+.*?\+\}|\{@(?:bu|gn):.*?@\})", re.S)


def _verschmolzen(ctx) -> bool:
    """Stellt sicher, dass mindestens eine Ersetzung ohne Leerzeichen dazwischen steht."""
    if any(a.art == "text" and ("-}{+" in a.markup or "-}{@" in a.markup or "@}{+" in a.markup or "@}{@" in a.markup)
           for _, a in ctx.alle_absaetze()):
        return True
    for _, a in ctx.alle_absaetze():
        if a.art == "text" and _ALT_NEU.search(a.markup):
            a.markup = _ALT_NEU.sub(lambda m: m.group(1) + m.group(2), a.markup, count=1)
            ctx.markiere(a, "verschmolzen")
            return True
    return False


def _umbruch(ctx) -> bool:
    """Absatz mit Aenderung bricht mitten im Text auf die naechste Seite um.

    In den ersten Versuchen reicht der natuerliche Seitenumbruch (die Garantie prueft, ob einer vorkommt);
    erst danach wird ein Umbruch erzwungen (dann bleibt unten auf der Seite Platz frei)."""
    if ctx.versuch < 4:
        return True
    r = ctx.rng
    kand = [(s, a) for s, a in _text_absaetze(ctx)
            if a.typ in (None, "text_ersetzt", "text_eingefuegt", "text_geloescht") and not a.fussnote]
    mit = [(s, a) for s, a in kand if a.typ]
    if not mit:
        return False
    lang = [(s, a) for s, a in mit if ctx.zeilen(a) >= 4]
    s, a = r.choice(lang or mit)
    a.umbruch = r.choice([2, 3])
    ctx.verlaengere(a, 6)
    ctx.markiere(a, "seitenumbruch_im_absatz")
    return True


def _versteckt(ctx) -> bool:
    r = ctx.rng
    ctx.layout["versteckt"] = True
    absatzmodus = ctx.layout["base"] == "L3"
    kand = [(s, a) for s, a in _text_absaetze(ctx)
            if a.typ in ("text_ersetzt", "text_eingefuegt", "text_geloescht")]
    if not kand:
        return False
    for s, a in r.sample(kand, min(len(kand), 2)):
        if absatzmodus:
            a.decoy_text = ctx.fast_gleich(zustand(a.markup, False))
        elif not s.decoy_text:
            alt = "\n".join(t for t in (x.excel_text(False) for x in s.absaetze) if t)
            s.decoy_text = ctx.fast_gleich(alt)
        ctx.markiere(a, "versteckte_zeilen")
    return True


_STIL_RE = {"neu": re.compile(r"\{\+(.*?)\+\}", re.S), "weg": re.compile(r"\{-(.*?)-\}", re.S)}


def _stil_erzwingen(ctx, art: str, stil: str) -> bool:
    for _, a in _text_absaetze(ctx):
        if a.typ and a.art == "text" and _STIL_RE[art].search(a.markup):
            a.markup = _STIL_RE[art].sub(lambda m: wrap_stil(stil, m.group(1)), a.markup, count=1)
            a.soll = schlimmer(a.soll, STIL_SOLL.get(stil))
            ctx.markiere(a, STIL_FAMILIE[stil])
            return True
    return False


def _misch(ctx) -> bool:
    if any("mischmarkierung" in a.stoerer for _, a in ctx.alle_absaetze()):
        return True
    return _stil_erzwingen(ctx, "weg", "bs") or _stil_erzwingen(ctx, "neu", "bu")


def _gruen(ctx) -> bool:
    if any("gruene_markierung" in a.stoerer for _, a in ctx.alle_absaetze()):
        return True
    return _stil_erzwingen(ctx, "neu", "gn")


_NBSP = re.compile(r"(?:(?<=\d)|(?<=\+\})|(?<=-\})) (?=(?:kW|kWh|V|kg|A|Wochen|Monate|Stunden|dB|ms)\b)")


def _sonder(ctx) -> bool:
    for _, a in ctx.alle_absaetze():
        if a.art == "text":
            a.markup = _NBSP.sub(" ", a.markup)
    return True


# ====================================================================== Garantien (nach dem Zeichnen)
def _g_silben(ctx, pdf) -> bool:
    return pdf.stats["trennungen"] >= 1


def _g_umbruch(ctx, pdf) -> bool:
    return any(a.typ and a.art == "text" and not s.scan and a.seite_ende > a.seite for s, a in ctx.alle_absaetze())


def _g_fussnote(ctx, pdf) -> bool:
    return sum(len(sx.fussnoten) for sx in pdf.seiten) >= 1


def _g_link(ctx, pdf) -> bool:
    return pdf.stats["links"] >= 1


def _g_gedreht(ctx, pdf) -> bool:
    return bool(ctx.opt.get("rotiert_seite"))


# ====================================================================== Register
def _mk(*familien: Stoerer) -> dict[str, Stoerer]:
    return {f.name: f for f in familien}


REGISTER: dict[str, Stoerer] = _mk(
    Stoerer("fussnote", "pdf", "Fussnotentext (8 pt) unten auf der Seite, Fussnotenzeichen am Wort; steht nicht in der Excel.",
            "uebernommen", mild=True, reihenfolge=50, anwenden=_fussnote, garantie=_g_fussnote),
    Stoerer("silbentrennung", "pdf", "Woerter am Zeilenende getrennt (ASCII-Bindestrich oder weicher Trennstrich), auch in "
            "farbigen Woertern und am Seitenende. Bei weichem Trennstrich in markiertem Wort: pruefen.",
            "uebernommen", mild=True, vorbereiten=_silben_vor, garantie=_g_silben),
    Stoerer("aufzaehlung", "beides", "Aufzaehlung mit Zeichen (•, –, 1., a)); die Zeichen stehen auch im Excel-Text. Neuer, "
            "geaenderter oder entfallener Punkt.", "uebernommen", mild=True, reihenfolge=10, anwenden=_aufzaehlung),
    Stoerer("nummerierte_liste", "beides", "Kurze nummerierte Punkte (1., 2., ...) im Text -- sehen wie Section-Nummern aus.",
            "uebernommen", reihenfolge=11, anwenden=_nummeriert),
    Stoerer("zwei_spalten", "pdf", "Ein Abschnitt steht in zwei Textspalten mit rechtem Rand; Aenderungen darin sind nicht sicher "
            "zuzuordnen.", "nicht_uebernommen", luecke="mehrspaltig", reihenfolge=30, anwenden=_zwei_spalten),
    Stoerer("tabelle_mit_farbkopf", "pdf", "Tabelle mit Linien, blauem Kopf (4472C4, weisse Schrift) und Baenderung (B4C6E7); "
            "meist mit einer geaenderten Zelle.", "nicht_uebernommen", luecke="tabelle", reihenfolge=32, anwenden=_tabelle),
    Stoerer("bild", "pdf", "Eingebettetes Bild (Diagramm mit blauen Flaechen) im Textfluss, ohne Aenderung.",
            "uebernommen", luecke="bild", reihenfolge=33, anwenden=_bild),
    Stoerer("scan_seite", "pdf", "Ein Abschnitt steht auf einer reinen Bildseite (kein Text, kein OCR).",
            "nicht_uebernommen", luecke="scan_seite", reihenfolge=31, anwenden=_scan),
    Stoerer("farbtoene", "pdf", "Andere Rot-/Blautoene als 0000FF/FF0000 (1F4E9F, 2E75B6, C00000, E06666 ...), teils je Absatz anders.",
            "uebernommen", mild=True, vorbereiten=_farbtoene_vor),
    Stoerer("cmyk_grau", "pdf", "Farben im Farbraum DeviceCMYK, normaler Text in DeviceGray.",
            "uebernommen", vorbereiten=_cmyk_vor),
    Stoerer("mischmarkierung", "pdf", "Mischformen: blau+unterstrichen (neu), blau+durchgestrichen und rot+unterstrichen "
            "(widerspruechlich -> pruefen).", "pruefen", reihenfolge=70, vorbereiten=_misch_vor, anwenden=_misch),
    Stoerer("seitenumbruch_im_absatz", "pdf", "Ein Absatz mit Aenderungen bricht mitten im Text auf die naechste Seite um.",
            "uebernommen", mild=True, reihenfolge=41, anwenden=_umbruch, garantie=_g_umbruch),
    Stoerer("kopf_fusszeile", "pdf", "Kopf-/Fusszeile nur auf Seite 1, in Blau oder mit section-aehnlichen Nummern.",
            "uebernommen", mild=True, vorbereiten=_kopf_vor),
    Stoerer("verschmolzen", "pdf", "Ersetzung ohne Leerzeichen dazwischen: {-100-}{+120+} kWh.",
            "uebernommen", mild=True, reihenfolge=80, vorbereiten=_verschmolzen_vor, anwenden=_verschmolzen),
    Stoerer("doppelte_absaetze", "beides", "Gleicher oder fast gleicher Absatz zweimal (andere Section, derselben Section/Zelle). "
            "Identisch in derselben Section: pruefen.", "pruefen", reihenfolge=20, anwenden=_doppelt),
    Stoerer("kommentar_spalte", "excel", "Kommentarspalte mit MEHR Text als die Anforderungsspalte (zitiert Textstuecke).",
            "uebernommen", vorbereiten=_komm_vor),
    Stoerer("lfd_spalte", "excel", "Laufende-Nummer-Spalte 'Lfd' vor der Section-Nummer.",
            "uebernommen", vorbereiten=_lfd_vor),
    Stoerer("zahl_sections", "excel", "Section-Nummern als Zahl in der Excel (2.1 und 2.10 sind beide 2,1).",
            "uebernommen", vorbereiten=_zahl_vor),
    Stoerer("mehrere_blaetter", "excel", "Mehrere Blaetter (Deckblatt, Historie mit altem Text, CDE, Glossar).",
            "uebernommen", mild=True, vorbereiten=_blaetter_vor),
    Stoerer("formeln_merges", "excel", "Titelzeile verbunden, Formelspalte und Summenzeile, verbundene Prio-Zellen. "
            "Zeilen einfuegen/loeschen: pruefen.", "pruefen", vorbereiten=_formeln_vor),
    Stoerer("versteckte_zeilen", "excel", "Versteckte Zeilen mit fast gleichem Text wie ein geaenderter Absatz.",
            "uebernommen", reihenfolge=60, anwenden=_versteckt),
    Stoerer("sonderzeichen", "beides", "ss, Umlaute, Anfuehrungszeichen, Gedankenstrich, geschuetztes Leerzeichen.",
            "uebernommen", mild=True, reihenfolge=90, vorbereiten=_sonder_vor, anwenden=_sonder),
    Stoerer("hyperlink", "pdf", "Blauer, unterstrichener Link mit Link-Annotation (0563C1) -- unveraenderter Text, keine Aenderung.",
            "uebernommen", reihenfolge=40, vorbereiten=_link_vor, anwenden=_hyperlink, garantie=_g_link),
    # ---- nur im Testkorpus (Regel 11) ---------------------------------------------------------
    Stoerer("gedrehte_seite", "pdf", "Eine Textseite traegt /Rotate 90.", "pruefen", nur_test=True,
            vorbereiten=_rotiere_vor, garantie=_g_gedreht),
    Stoerer("gruene_markierung", "pdf", "Neuer Text in Gruen (unbekannte Markierungsfarbe).", "nicht_uebernommen",
            luecke="unbekannte_farbe", nur_test=True, reihenfolge=71, vorbereiten=_gruen_vor, anwenden=_gruen),
)

NUR_TEST: tuple[str, ...] = tuple(sorted(n for n, f in REGISTER.items() if f.nur_test))
# Im Testkorpus erzwungen: Seed % modul == rest -> Familie kommt vor (in 30 Seeds je 6 Paare)
TEST_ZWANG: dict[str, tuple[int, int]] = {"gedrehte_seite": (5, 0), "gruene_markierung": (5, 2)}


def pool(art: str, korpus: str = "entwicklung") -> list[str]:
    """Familien, aus denen zufaellig gezogen wird. NUR_TEST-Familien sind nie dabei (sie kommen ueber TEST_ZWANG)."""
    return [n for n, f in sorted(REGISTER.items()) if not f.nur_test and (art != "mild" or f.mild)]
