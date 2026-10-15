"""Red-Team R2 (PDF-Leser), Teil 4: Faelle aus dem ENTWICKLUNGSKORPUS mit Fehlerart `pdf_falsch_gelesen`.

Die Paare werden hier mit dem festen Generator (tests/fixtures/generator.py, Version 1) aus dem Seed neu gebaut --
nichts wird gespeichert, nichts liegt im Repo. Nur Seeds unter 1000 (Entwicklungskorpus, Regel 11); der Testkorpus
bleibt unberuehrt. Alles ERFUNDEN (Regel 16).

Gefunden mit `python -m src.cli messen --korpus entwicklung` (Baseline: 99 Stellen `pdf_falsch_gelesen`) und einem
Wort-fuer-Wort-Vergleich: Was der Leser als neu/entfaellt meldet, gegen das, was im Markup des Generators steht.
Der Leser darf abweichen, wenn er die Abweichung SICHTBAR meldet (Hinweis oder Luecke); sonst ist es ein stiller Fehler.
"""
import json
import re
from collections import Counter

import pytest

from src.models import ADD, DEL, MOD
from tests.fixtures import generator as g
from tests.fixtures.pdf_writer import parse_markup
from tests.redteam._bau import aufbau, scheitern
from tests.redteam._bau_pdf import alle_tokens, lese, meldungen

# Zaehlzeichen von Aufzaehlungen ("4.", "d)", "-") sind Teil der Liste, nicht des Markups: nicht mitzaehlen
_MARKER = re.compile(r"^(\(?\d{1,2}[.)]|[a-zA-Z][.)]|[•\-–·▪●○*])$")


def _paar(seed: int, profil: str):
    aufbau(g.GENERATOR_VERSION == "1", "Generator-Version geaendert: Seeds neu pruefen (Tests an Version 1 gebunden)")
    aufbau(seed < 1000, "Nur Entwicklungskorpus (Regel 11)")
    name, cfg = g._profil(profil)
    ctx, pdf = g._baue(seed, name, cfg, "entwicklung")
    return ctx, pdf


def _soll(ctx) -> tuple[Counter, Counter]:
    """Woerter, die im Markup als 'neu' bzw. 'entfaellt' stehen (Hyperlinks und gruene Woerter zaehlen NICHT)."""
    neu: Counter = Counter()
    weg: Counter = Counter()

    def fuettere(markup: str) -> None:
        for wort in parse_markup(markup):
            for text, stil in wort:
                for stueck in text.split("\u00a0"):          # geschuetztes Leerzeichen trennt Woerter (wie im Leser)
                    if not stueck or _MARKER.match(stueck):
                        continue
                    if stil in ("n", "bu"):
                        neu[stueck] += 1
                    elif stil in ("g", "ru", "r", "s", "bs"):
                        weg[stueck] += 1

    for s in ctx.sektionen:
        fuettere(s.kopf_markup)
        for a in s.absaetze:
            if a.art == "tabelle":
                for zeile in a.tabelle or []:
                    for zelle in zeile:
                        fuettere(zelle)
            elif a.art == "text":
                fuettere(a.markup)
    return neu, weg


def _ist(a) -> tuple[Counter, Counter]:
    neu: Counter = Counter()
    weg: Counter = Counter()
    for t in alle_tokens(a):
        if t.status in (ADD, MOD) and t.new and not _MARKER.match(t.new):
            neu[t.new] += 1
        if t.status in (DEL, MOD) and t.old and not _MARKER.match(t.old):
            weg[t.old] += 1
    return neu, weg


def _abweichung(tmp_path, seed: int, profil: str):
    ctx, pdf = _paar(seed, profil)
    a = lese(tmp_path, pdf)
    sn, sw = _soll(ctx)
    ist_n, ist_w = _ist(a)
    return a, {"fehlt_neu": sn - ist_n, "falsch_neu": ist_n - sn, "fehlt_weg": sw - ist_w, "falsch_weg": ist_w - sw}


def _pruefe_stille_abweichung(a, ab: dict, was: str) -> None:
    summe = sum(sum(c.values()) for c in ab.values())
    if summe and not meldungen(a):
        teile = [f"{k}: {dict(list(c.items())[:3])}" for k, c in ab.items() if c]
        scheitern(f"{was}: {summe} Woerter weichen vom Soll ab, ohne Hinweis oder Luecke -- " + "; ".join(teile))


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-105 Korpus synth_0024_schwer (Stoerer 'verschmolzen'): eine Ersetzung ohne Leerzeichen "
                          "wird ein einziges gestrichenes Wort, das neue Wort geht verloren (wie RT-60, hier im "
                          "Korpus); Besitzer K-PDF")
def test_rt105_korpus_verschmolzene_ersetzung_seed24(tmp_path):
    a, ab = _abweichung(tmp_path, 24, "schwer")
    _pruefe_stille_abweichung(a, ab, "Paar 24")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-106 Korpus synth_0015_schwer (Stoerer 'hyperlink'): vier unveraenderte Link-Woerter in "
                          "Link-Blau werden als NEU gelesen (wie RT-78, hier im Korpus); Besitzer K-PDF")
def test_rt106_korpus_hyperlinks_zaehlen_als_neu_seed15(tmp_path):
    a, ab = _abweichung(tmp_path, 15, "schwer")
    _pruefe_stille_abweichung(a, ab, "Paar 15")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-107 Korpus synth_0037_standard (Stoerer 'silbentrennung'): ein GESTRICHENES Wort mit "
                          "weichem Trennstrich bleibt in zwei Stuecke zerrissen (Zusammenfueger arbeitet nur bei "
                          "schwarzem Text); Besitzer K-PDF")
def test_rt107_korpus_getrenntes_gestrichenes_wort_seed37(tmp_path):
    a, ab = _abweichung(tmp_path, 37, "standard")
    _pruefe_stille_abweichung(a, ab, "Paar 37")


@pytest.mark.redteam
@pytest.mark.xfail(strict=True, raises=AssertionError,
                   reason="RT-108 Korpus synth_0057_schwer (Stoerer 'silbentrennung'): drei BLAUE neue Woerter mit "
                          "Trennstrich bleiben in je zwei Stuecke zerrissen (wie RT-63/64, hier im Korpus); Besitzer K-PDF")
def test_rt108_korpus_getrennte_blaue_woerter_seed57(tmp_path):
    a, ab = _abweichung(tmp_path, 57, "schwer")
    _pruefe_stille_abweichung(a, ab, "Paar 57")


@pytest.mark.redteam
def test_rt109_korpus_scan_seite_wird_als_luecke_gemeldet_seed12(tmp_path):
    aufbau(g.GENERATOR_VERSION == "1", "Generator-Version geaendert")
    ordner = g.erzeuge_paar(12, "schwer", tmp_path / "paare", "entwicklung")
    erwartet = json.loads((ordner / "erwartet.json").read_text(encoding="utf-8"))
    soll = [l for l in erwartet["luecken"] if l["art"] == "scan_seite"]
    aufbau(bool(soll), "Der Generator hat keine Scan-Seite erwartet")
    from src.pdf_reader import analysiere_pdf
    a = analysiere_pdf(ordner / "aenderungen.pdf")
    gemeldet = {(l["art"], l["seite"]) for l in a.luecken}
    fehlt = [(l["art"], l["seite"], l["woerter"]) for l in soll if (l["art"], l["seite"]) not in gemeldet]
    if fehlt:
        scheitern(f"Erwartete Luecken nicht gemeldet (Art, Seite, Soll-Woerter): {fehlt}; gemeldet: {sorted(gemeldet)}; "
                  f"nur Warnungen: {[w[:40] for w in a.warnungen]}")
