"""Red-Team R2 (PDF-Leser), Teil 1: Layout und Woerter (Faelle a, b, c, e, i, j, l, n).

Alles ERFUNDEN (Regel 16), nichts in src/ geaendert. Jeder Fehlertest traegt
`@pytest.mark.redteam` und `@pytest.mark.xfail(strict=True, raises=AssertionError, reason="RT-nn ...; Besitzer K-PDF")`.
Die FALSCHE Wirkung wird als AssertionError gemeldet (`scheitern`), Fehler im Testaufbau als RuntimeError (`aufbau`).

Massstab ("Bilanz-Prinzip"): Der Leser darf entweder KORREKT lesen oder muss die Unsicherheit SICHTBAR melden
(PdfToken.hinweis, PdfBlock.hinweise, PdfAnalyse.luecken, Warnung). Nie still falsch.
"""
import pytest

from src.models import ADD, DEL
from tests.redteam._bau import aufbau, scheitern
from tests.redteam._bau_pdf import (RohSeite, alle_tokens, eindeutiger_text as _eindeutiger_text, bloecke_mit, dok, formen, geaendert, hinweis_bei, kopf_fuss_zone,
                                    kurz, lese, meldungen, roh_pdf, token, tokens_mit, zeile)
from tests.fixtures.pdf_writer import PdfDokument

# ==================================================================================================
# a) Verschmolzene Ersetzung
# ==================================================================================================
@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-60 Verschmolzene Ersetzung {-Netzleistung-}{+Leistung+}: wird EIN gestrichenes Wort "
                          "'NetzleistungLeistung', der neue Text geht verloren; Besitzer K-PDF")
def test_rt60_verschmolzene_ersetzung_ungleich_lang(tmp_path):
    d = dok()
    d.block("Die {-Netzleistung-}{+Leistung+} betraegt 5 kW.")
    a = lese(tmp_path, d)
    aufbau(token(a, "kW.") is not None, "Satz wurde nicht gelesen")
    b = bloecke_mit(a, "kW.")[0]
    ts = tokens_mit(a, "eistung")
    aufbau(bool(ts), "Das Wort mit 'leistung' wurde gar nicht gelesen")
    alt_ok = "Netzleistung" in b.alt_text.split()
    neu_ok = "Leistung" in b.neu_text.split()
    if not (alt_ok and neu_ok) and not hinweis_bei(a, ts):
        scheitern("Verschmolzene Ersetzung falsch gelesen und ohne Hinweis: " + "; ".join(kurz(t) for t in ts)
                  + " -- erwartet: alt='Netzleistung', neu='Leistung'")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-61 'ab{+gesichert+}': das unveraenderte 'ab' fehlt im alten Stand (old=''), "
                          "obwohl es im PDF schwarz steht; Besitzer K-PDF")
def test_rt61_unveraenderter_wortteil_bleibt_im_alten_stand(tmp_path):
    d = dok()
    d.block("Die Anlage ist ab{+gesichert+} und sicher.")
    a = lese(tmp_path, d)
    aufbau(token(a, "sicher.") is not None, "Satz wurde nicht gelesen")
    t = token(a, "abgesichert")
    if t is None:
        scheitern("Das Wort 'abgesichert' wird nicht als ein Wort gelesen: "
                  + "; ".join(kurz(x) for x in tokens_mit(a, "gesichert")))
    if t.old != "ab" and not hinweis_bei(a, [t]):
        scheitern(f"Das unveraenderte 'ab' fehlt im alten Stand: {kurz(t)} -- erwartet old='ab', ohne Hinweis")


# ==================================================================================================
# b) Fussnote druckt die Zeilenabstand-Schaetzung
# ==================================================================================================
@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-62 Fussnoten (8 pt, 3 Zeilen) druecken die Zeilenabstand-Schaetzung: jede Zeile des "
                          "Fliesstexts wird ein eigener Block; Besitzer K-PDF")
def test_rt62_fussnote_zerreisst_den_fliesstext(tmp_path):
    d = dok()
    d.block(_eindeutiger_text(62, 38, "Erstwort", "Endwort", "{+neuwort+}"))
    d.block(_eindeutiger_text(63, 30, "Zweitwort", "Schlusswort"))
    for i in range(3):
        aufbau(d.fussnote("Dies ist eine erfundene Fussnote mit etwas Text zur Erlaeuterung.", marke=str(i + 1)),
               "Fussnote passte nicht mehr auf die Seite")
    aufbau(sum(len(z) for z in d.seiten[0].fussnoten) >= 3, "weniger als 3 Fussnotenzeilen")
    a = lese(tmp_path, d)
    aufbau(token(a, "Erstwort") is not None and token(a, "Endwort") is not None, "Absatz nicht gelesen")
    if not bloecke_mit(a, "Erstwort", "Endwort") and not any(b.hinweise for b in a.blocks):
        n = len([b for b in a.blocks if b.kind == "absatz"])
        scheitern(f"Der dreizeilige Absatz ist zerrissen: Anfang und Ende stehen in verschiedenen Bloecken "
                  f"({n} Absatz-Bloecke im ganzen Dokument), ohne Hinweis")


# ==================================================================================================
# c) Silbentrennung in farbigen Woertern
# ==================================================================================================
_TEXT_C = ("Wir pruefen heute ausfuehrlich die Anforderung ob die {+Kapazitaetserweiterung Sicherheitsarchitektur "
           "Netzanbindung Verfuegbarkeitsgarantie Kapazitaetsplanung Leistungsfaehigkeit Netzanschlussbedingungen "
           "Verantwortlichkeiten Wartungsintervalle+} gilt und sonst nichts weiter.")


def _wort_zusammen_oder_hinweis(a, ganz: str, anfang: str) -> bool:
    """True, wenn das Wort als Ganzes gelesen wurde ODER ein Hinweis an den Teilstuecken haengt."""
    if any(ganz in (t.old, t.new) for t in alle_tokens(a)):
        return True
    return hinweis_bei(a, tokens_mit(a, anfang))


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-63 Silbentrennung ('-') in einem BLAUEN Wort: 'Sicherheitsarchi-' und 'tektur' bleiben "
                          "zwei neue Woerter (nur schwarze Woerter werden zusammengefuegt); Besitzer K-PDF")
def test_rt63_silbentrennung_blaues_wort_mit_strich(tmp_path):
    d = dok(trennung="ascii")
    d.block(_TEXT_C)
    aufbau(("Sicherheitsarchitektur", ("n",), "-") in d.getrennt, "Wort wurde im Test-PDF nicht getrennt")
    a = lese(tmp_path, d)
    if not _wort_zusammen_oder_hinweis(a, "Sicherheitsarchitektur", "Sicherheitsarchi"):
        scheitern("Blaues getrenntes Wort nicht zusammengefuegt und ohne Hinweis: "
                  + ", ".join(repr(t.new) for t in tokens_mit(a, "Sicherheitsarchi") + tokens_mit(a, "tektur")))


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-64 Weicher Trennstrich (U+00AD) in einem BLAUEN Wort: 'Sicherheitsarchi' und 'tektur' "
                          "bleiben zwei neue Woerter, der Strich fehlt, kein Hinweis; Besitzer K-PDF")
def test_rt64_silbentrennung_blaues_wort_weicher_strich(tmp_path):
    d = dok(trennung="weich")
    d.block(_TEXT_C)
    aufbau(("Sicherheitsarchitektur", ("n",), "­") in d.getrennt, "Wort wurde im Test-PDF nicht getrennt")
    a = lese(tmp_path, d)
    if not _wort_zusammen_oder_hinweis(a, "Sicherheitsarchitektur", "Sicherheitsarchi"):
        scheitern("Blaues Wort mit weichem Trennstrich zerrissen und ohne Hinweis: "
                  + ", ".join(repr(t.new) for t in tokens_mit(a, "Sicherheitsarchi") + tokens_mit(a, "tektur")))


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-65 Weicher Trennstrich (U+00AD) in einem SCHWARZEN Wort: pdfplumber laesst den Strich weg, "
                          "'Ueberga' und 'beprotokolle' bleiben zwei Woerter (der Zusammenfueger sucht nur '-'); Besitzer K-PDF")
def test_rt65_silbentrennung_schwarzes_wort_weicher_strich(tmp_path):
    d = dok(trennung="weich")
    d.block("Der Betreiber stellt die Speicherkapazitaet bereit und dokumentiert Verantwortlichkeiten "
            "Uebergabeprotokolle sowie Wartungsintervalle.")
    aufbau(any(w == "Uebergabeprotokolle" and "­" == z for w, _, z in d.getrennt), "Wort wurde nicht getrennt")
    a = lese(tmp_path, d)
    if not _wort_zusammen_oder_hinweis(a, "Uebergabeprotokolle", "Ueberga"):
        scheitern("Schwarzes Wort mit weichem Trennstrich zerrissen und ohne Hinweis: "
                  + ", ".join(repr(t.old) for t in tokens_mit(a, "Ueberga") + tokens_mit(a, "beprotokolle")))


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-66 Silbentrennung genau am Seitenende: 'Sicherheitsarchi-' steht am Ende von Seite 1, "
                          "'tektur' am Anfang von Seite 2; die Teile werden nie zusammengefuegt; Besitzer K-PDF")
def test_rt66_silbentrennung_am_seitenende(tmp_path):
    d = dok(trennung="ascii")
    d.block("Wort Wort Wort Wort zuerst die {+Verfuegbarkeitsgarantie Kapazitaetserweiterung Sicherheitsarchitektur+} "
            "und danach noch viel mehr Text der auf der naechsten Seite weitergeht.", umbruch_nach_zeile=1)
    aufbau(d.stats["trennung_seitenende"] == 1 and d.seitenzahl == 2, "Trennung am Seitenende nicht erzeugt")
    a = lese(tmp_path, d)
    if not _wort_zusammen_oder_hinweis(a, "Sicherheitsarchitektur", "Sicherheitsarchi"):
        seiten = {t.new: t.page for t in tokens_mit(a, "Sicherheitsarchi") + tokens_mit(a, "tektur")}
        scheitern(f"Wort ueber den Seitenumbruch getrennt, nicht zusammengefuegt, kein Hinweis: {seiten}")


# ==================================================================================================
# e) Absatz ueber den Seitenumbruch
# ==================================================================================================
@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-67 Absatz laeuft ueber den Seitenumbruch: er wird immer in zwei Bloecke geteilt, "
                          "ohne Hinweis, dass der zweite die Fortsetzung ist; Besitzer K-PDF")
def test_rt67_absatz_ueber_seitenumbruch_bleibt_ein_absatz(tmp_path):
    d = dok()
    d.block("Dies ist ein Absatz der ueber einen Seitenumbruch geht und mehrere Zeilen hat damit das Ganze testbar ist "
            "und noch weiter geht fuer eine dritte Zeile {+neu am Ende+}", umbruch_nach_zeile=1)
    aufbau(d.letzter_block == (1, 2), "Absatz geht nicht ueber zwei Seiten")
    a = lese(tmp_path, d)
    aufbau(token(a, "Seitenumbruch") is not None and token(a, "testbar") is not None, "Absatz nicht gelesen")
    if not bloecke_mit(a, "Seitenumbruch", "testbar") and not any(b.hinweise for b in a.blocks):
        seiten = sorted({b.page for b in a.blocks})
        scheitern(f"Ein Absatz ist in {len(a.blocks)} Bloecke auf den Seiten {seiten} zerrissen, ohne Fortsetzungs-Hinweis")


# ==================================================================================================
# i) Aufzaehlungen, Symbolzeichen, Gedankenstrich
# ==================================================================================================
@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-68 Enge Liste '1.' / '2.' / '3.' (Zeilenabstand wie im Fliesstext): alle Punkte werden "
                          "EIN Block (nur '1)', 'a)' und Punkte wie '-' werden erkannt); Besitzer K-PDF")
def test_rt68_enge_liste_mit_nummer_punkt(tmp_path):
    d = dok()
    d.block("Einleitung vor der Liste ohne Neuerung.")
    for i, wort in enumerate(("Alpha", "Beta", "Gamma"), 1):
        d.block(f"Das {wort} Listenelement ist hier eingetragen und lang genug.", marker=f"{i}.", abstand_danach=0)
    a = lese(tmp_path, d)
    aufbau(all(token(a, w) for w in ("Alpha", "Beta", "Gamma")), "Liste nicht gelesen")
    if bloecke_mit(a, "Alpha", "Beta") and not any(b.hinweise for b in a.blocks):
        scheitern("Die Listenpunkte 1., 2. und 3. stehen in EINEM Block: "
                  + str([len(b.tokens) for b in a.blocks]) + " Woerter je Block")


_SYMBOL_SCHRIFT = ("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Type /Encoding "
                   "/BaseEncoding /WinAnsiEncoding /Differences [128 /uniF0B7 /filledbox] >> >>")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-69 Enge Liste mit 'A)' / 'i.' / Symbolzeichen (Wingdings U+F0B7, Quadrat): "
                          "die Punkte werden EIN Block (das Muster kennt nur Kleinbuchstaben 'a)' und Ziffern); Besitzer K-PDF")
@pytest.mark.parametrize("marker,schrift", [(["A)", "B)", "C)"], ""), (["i.", "ii.", "iii."], ""),
                                            (["\x80"] * 3, _SYMBOL_SCHRIFT), (["\x81"] * 3, _SYMBOL_SCHRIFT)],
                         ids=["A)", "i.", "U+F0B7", "Quadrat"])
def test_rt69_enge_liste_symbolzeichen_und_grossbuchstaben(tmp_path, marker, schrift):
    inhalt = zeile(56, 64, "Einleitung vor der Liste ohne Neuerung.")
    for i, (m, wort) in enumerate(zip(marker, ("Alpha", "Beta", "Gamma"))):
        top = 79 + 15 * i
        inhalt += zeile(56, top, m, kodierung="latin-1") + zeile(72, top, f"Das {wort} Listenelement ist hier und lang genug.")
    a = lese(tmp_path, roh_pdf([RohSeite(inhalt=inhalt, schrift=schrift)]))
    aufbau(all(token(a, w) for w in ("Alpha", "Beta", "Gamma")), "Liste nicht gelesen")
    if bloecke_mit(a, "Alpha", "Beta") and not any(b.hinweise for b in a.blocks):
        scheitern("Die drei Listenpunkte stehen in EINEM Block, ohne Hinweis: " + str([len(b.tokens) for b in a.blocks]))


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-70 Gedankenstrich '–' steht zufaellig am Zeilenanfang im Fliesstext: der Absatz wird "
                          "dort geteilt (gilt als Aufzaehlungszeichen); Besitzer K-PDF")
def test_rt70_gedankenstrich_am_zeilenanfang_teilt_keinen_absatz(tmp_path):
    d = dok()
    d.block("Wort Wort Wort Wort Wort Wort Der Betrieb laeuft – stabil und wird dokumentiert und die Anlage wird "
            "regelmaessig gewartet damit alles in Ordnung bleibt und {+nichts ausfaellt+}.", spalte_bis=300)
    a = lese(tmp_path, d)
    strich = token(a, "–")
    aufbau(strich is not None and abs(strich.bbox[0] - 56.0) < 0.6, "Der Gedankenstrich steht nicht am Zeilenanfang")
    if not bloecke_mit(a, "laeuft", "stabil") and not any(b.hinweise for b in a.blocks):
        scheitern(f"Ein Satz mit Gedankenstrich am Zeilenanfang ist in {len(a.blocks)} Bloecke zerrissen")


# ==================================================================================================
# j) Kopf-/Fusszeile ohne 'Seite #'-Muster
# ==================================================================================================
@pytest.mark.redteam
def test_rt71_kopf_fusszeile_einseitig_ohne_seite_muster(tmp_path):
    d = PdfDokument(kopf_text="Anlage 3 - Technische Spezifikation", fuss_vorlage="Dokument CDE-0815 Blatt {nr}")
    d.block("Normaler Text der Seite.")
    a = lese(tmp_path, d)
    zone = kopf_fuss_zone(a)
    aufbau(token(a, "Normaler") is not None, "Fliesstext nicht gelesen")
    schlimm = [b for b in zone if not b.ignoriert and not b.hinweise]
    if schlimm:
        scheitern(f"{len(schlimm)} Block/Bloecke aus Kopf-/Fusszeile werden als normaler Text behandelt "
                  f"(y={[round(b.bbox[1]) for b in schlimm]})")


@pytest.mark.redteam
@pytest.mark.parametrize("fall", ["eine_seite", "zwei_seiten_kopf_nur_auf_seite_1"])
def test_rt72_blaue_kopfzeile_zaehlt_als_neu(tmp_path, fall):
    kw = dict(kopf_text="Anlage 3 - Technische Spezifikation", kopf_farbe=(0.0, 0.0, 1.0),
              fuss_vorlage="Dokument CDE-0815 Blatt {nr}")
    if fall == "zwei_seiten_kopf_nur_auf_seite_1":
        kw.update(kopf_fuss={1}, fuss_vorlage="Seite {nr} von {n}")
    d = PdfDokument(**kw)
    d.block("Normaler Text der Seite.")
    if fall != "eine_seite":
        d.seitenumbruch()
        d.block("Normaler Text der zweiten Seite.")
        aufbau(d.seitenzahl == 2, "zweite Seite fehlt")
    a = lese(tmp_path, d)
    aufbau(token(a, "Normaler") is not None, "Fliesstext nicht gelesen")
    falsch = [t for b in kopf_fuss_zone(a) if not b.hinweise for t in b.tokens if t.status == ADD and not t.hinweis]
    if falsch:
        scheitern(f"{len(falsch)} Woerter der blauen Kopfzeile zaehlen still als 'neu': "
                  + " ".join(t.new for t in falsch))


# ==================================================================================================
# l) Linie als Durchstreichung
# ==================================================================================================
@pytest.mark.redteam
def test_rt73_zierlinie_hinter_ueberschrift_ist_keine_durchstreichung(tmp_path):
    def bau(mit_linie: bool):
        d = dok()
        d.ueberschrift("1.1 Geltungsbereich Anforderungen")
        d.block("Normaler Text fuer die Anforderung hier.")
        if mit_linie:        # graue Haarlinie hinter der Ueberschrift (Mitte der 18-pt-Zeile ab y=70)
            d._s.inhalt.append((d._linie_op(56, 300, 70 + 9, (0.7, 0.7, 0.7), 0.5) + "\n").encode("latin-1"))
        return d
    ohne = lese(tmp_path / "ohne", bau(False))
    aufbau(not any(geaendert(t) for t in alle_tokens(ohne)), "Ohne Linie darf nichts markiert sein")
    a = lese(tmp_path / "mit", bau(True))
    kopf = [b for b in a.blocks if b.kind == "ueberschrift"]
    aufbau(bool(kopf), "Ueberschrift nicht erkannt")
    falsch = [t for t in kopf[0].tokens if t.status == DEL and not t.hinweis]
    if falsch and not kopf[0].hinweise:
        scheitern(f"Eine graue Zierlinie macht {len(falsch)} Woerter der Ueberschrift zu 'entfaellt': "
                  + " ".join(t.old for t in falsch))


# ==================================================================================================
# n) Ueberschriften
# ==================================================================================================
@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-74 Zweizeilige Ueberschrift: die zweite Zeile wird ein eigener Absatz-Block "
                          "(und damit Teil des Section-Textes); Besitzer K-PDF")
def test_rt74_zweizeilige_ueberschrift_bleibt_ein_block(tmp_path):
    d = dok()
    d.ueberschrift("2.4 Anforderungen an die Verfuegbarkeit und die Kapazitaet der gesamten Netzanbindung im Normalbetrieb")
    d.block("Folgetext der Section hier im Satz und {+neu+}.")
    a = lese(tmp_path, d)
    koepfe = [b for b in a.blocks if b.kind == "ueberschrift"]
    aufbau(len(koepfe) >= 1 and any("Folgetext" in formen(b) for b in a.blocks), "Aufbau gelesen")
    ganz = any("Normalbetrieb" in " ".join(formen(b)) for b in koepfe)
    if not ganz and not any(b.hinweise for b in a.blocks):
        rest = [b for b in a.blocks if b.kind == "absatz" and any("Normalbetrieb" in x for x in formen(b))]
        scheitern(f"Die zweite Zeile der Ueberschrift steht in einem {rest[0].kind if rest else '?'}-Block, "
                  f"nicht in der Ueberschrift (Ueberschrift-Bloecke: {len(koepfe)})")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-75 Umnummerierte Ueberschrift '{-2.3-} {+2.2+}': section_neu wird nur aus dem ERSTEN "
                          "Wort gelesen und bleibt leer (auch fuer den Folgeabsatz); Besitzer K-PDF")
def test_rt75_umnummerierte_ueberschrift_section_neu(tmp_path):
    d = dok()
    d.ueberschrift("{-2.3-} {+2.2+} Geltungsbereich der Anforderungen")
    d.block("Folgetext der Section hier im Satz.")
    a = lese(tmp_path, d)
    kopf = next((b for b in a.blocks if b.kind == "ueberschrift"), None)
    aufbau(kopf is not None and kopf.section_alt == "2.3", "Ueberschrift/alte Nummer nicht erkannt")
    folge = next((b for b in a.blocks if "Folgetext" in formen(b)), None)
    aufbau(folge is not None, "Folgeabsatz nicht gelesen")
    if kopf.section_neu != "2.2" and not kopf.hinweise:
        scheitern(f"section_neu der Ueberschrift ist {kopf.section_neu!r} (erwartet '2.2'), "
                  f"des Folgeabsatzes {folge.section_neu!r}; section_alt {kopf.section_alt!r}")
