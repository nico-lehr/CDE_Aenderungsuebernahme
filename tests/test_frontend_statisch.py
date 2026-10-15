"""Statische Prüfung des Dashboards (frontend/): ohne Browser, nur Dateien lesen.

Prüft das, was sich ohne Browser prüfen lässt: HTML parsebar, IDs aus app.js vorhanden, Rollen für Tastatur und
Screenreader, Farben nur als Token, Fokus-Ring, keine festen Adressen, kein "Löschen"-Text, Syntax von app.js (wenn node da ist).
Was nur ein Browser zeigt (Layout, Darstellung), steht nicht hier.
"""
from __future__ import annotations

import re
import shutil
import subprocess
from html.parser import HTMLParser
from pathlib import Path

import pytest

FRONTEND = Path(__file__).resolve().parents[1] / "frontend"
HTML = (FRONTEND / "index.html").read_text(encoding="utf-8")
JS = (FRONTEND / "app.js").read_text(encoding="utf-8")
CSS = (FRONTEND / "style.css").read_text(encoding="utf-8")


# ----------------------------------------------------------------------------------------------
# Hilfen
# ----------------------------------------------------------------------------------------------
class _Seite(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.elemente: list[tuple[str, dict[str, str | None]]] = []
        self.offen: list[str] = []
        self.fehler: list[str] = []

    _LEER = {"meta", "link", "img", "input", "br", "hr", "source", "area", "base", "col", "embed", "param", "track", "wbr"}

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if "label" in self.offen:
            a["_in_label"] = "1"                       # Eingabefeld steckt in einem <label>
        self.elemente.append((tag, a))
        if tag not in self._LEER:
            self.offen.append(tag)

    def handle_endtag(self, tag):
        if tag in self._LEER:
            return
        if not self.offen or self.offen[-1] != tag:
            self.fehler.append(f"</{tag}> passt nicht zu {self.offen[-1:] or 'nichts'}")
            if tag in self.offen:
                while self.offen and self.offen.pop() != tag:
                    pass
            return
        self.offen.pop()


def _parse_html() -> _Seite:
    p = _Seite()
    p.feed(HTML)
    p.close()
    return p


SEITE = _parse_html()
HTML_IDS = [a["id"] for _, a in SEITE.elemente if a.get("id")]


def _css_regeln(css: str):
    """Zerlegt CSS in (Kontext, Kopf, Körper). Kontext = Liste der umgebenden @media/@supports-Köpfe.
    @keyframes zählt als ein Blatt (Kopf = '@keyframes ...')."""
    css = re.sub(r"/\*.*?\*/", "", css, flags=re.S)
    aus: list[tuple[tuple[str, ...], str, str]] = []

    def lauf(text: str, ctx: tuple[str, ...]) -> None:
        i = 0
        while True:
            j = text.find("{", i)
            if j < 0:
                return
            kopf = text[i:j].strip()
            tiefe, p = 1, j + 1
            while tiefe and p < len(text):
                tiefe += {"{": 1, "}": -1}.get(text[p], 0)
                p += 1
            koerper = text[j + 1:p - 1]
            if kopf.startswith(("@media", "@supports")):
                lauf(koerper, ctx + (kopf,))
            else:
                aus.append((ctx, kopf, koerper))
            i = p

    lauf(css, ())
    return aus


REGELN = _css_regeln(CSS)
_DUNKEL_MEDIA = "@media (prefers-color-scheme: dark)"


def _ist_token_block(ctx: tuple[str, ...], kopf: str) -> bool:
    if any(c.startswith(_DUNKEL_MEDIA) for c in ctx):
        return True
    return kopf in (":root", ':root[data-theme="dark"]', ':root[data-theme="light"]')


def _deklarationen(koerper: str) -> dict[str, str]:
    d: dict[str, str] = {}
    for teil in koerper.split(";"):
        if ":" in teil:
            k, v = teil.split(":", 1)
            d[k.strip()] = v.strip()
    return d


def _tokens(kopf: str, in_media: bool) -> dict[str, str]:
    for ctx, k, koerper in REGELN:
        if k == kopf and (any(c.startswith(_DUNKEL_MEDIA) for c in ctx) == in_media):
            return {n: v for n, v in _deklarationen(koerper).items() if n.startswith("--")}
    raise AssertionError(f"Token-Block {kopf!r} fehlt")


def _luminanz(hexfarbe: str) -> float:
    h = hexfarbe.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4  # noqa: E731
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _kontrast(a: str, b: str) -> float:
    la, lb = sorted((_luminanz(a), _luminanz(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


# ----------------------------------------------------------------------------------------------
# HTML
# ----------------------------------------------------------------------------------------------
def test_html_parsebar_und_sauber_verschachtelt():
    assert not SEITE.fehler, SEITE.fehler
    assert not SEITE.offen, f"nicht geschlossene Elemente: {SEITE.offen}"
    tags = {t for t, _ in SEITE.elemente}
    assert {"html", "head", "body", "main", "header", "nav", "footer"} <= tags
    html = next(a for t, a in SEITE.elemente if t == "html")
    assert html.get("lang") == "de"


def test_ids_sind_eindeutig():
    doppelt = {i for i in HTML_IDS if HTML_IDS.count(i) > 1}
    assert not doppelt, f"doppelte IDs: {sorted(doppelt)}"


def test_alle_ids_aus_app_js_gibt_es_im_html():
    muster = [
        r"\$\(\s*'([A-Za-z][\w-]*)'\s*\)",                        # $('id')
        r"getElementById\(\s*'([A-Za-z][\w-]*)'\s*\)",
        r"querySelector\(\s*'#([A-Za-z][\w-]*)'\s*\)",
        r"fuelleSelect\(\s*'([A-Za-z][\w-]*)'",                   # Hilfsfunktionen, die eine ID bekommen
        r"dropzone\(\s*'([A-Za-z][\w-]*)'\s*,\s*'([A-Za-z][\w-]*)'",
    ]
    benutzt: set[str] = set()
    for m in muster:
        for treffer in re.findall(m, JS):
            benutzt.update(treffer if isinstance(treffer, tuple) else (treffer,))
    for liste in re.findall(r"for \(const id of \[([^\]]+)\]\)", JS):      # Listen von IDs
        benutzt.update(re.findall(r"'([\w-]+)'", liste))
    for schluessel in re.findall(r"'(f-[\w-]+)'\s*:", JS):                   # Zuordnung Feld-ID -> Filter
        benutzt.add(schluessel)
    assert len(benutzt) > 30, "Die Erkennung der IDs in app.js scheint nicht zu greifen"
    fehlen = sorted(benutzt - set(HTML_IDS))
    assert not fehlen, f"IDs in app.js, die index.html nicht hat: {fehlen}"


def test_labels_und_aria_verweise_zeigen_auf_vorhandene_ids():
    ids = set(HTML_IDS)
    for tag, a in SEITE.elemente:
        for attr in ("aria-controls", "aria-labelledby", "aria-describedby"):
            for ziel in (a.get(attr) or "").split():
                assert ziel in ids, f"<{tag}> {attr}={ziel!r} zeigt ins Leere"
        if tag == "label" and a.get("for"):
            assert a["for"] in ids, f"<label for={a['for']!r}> zeigt ins Leere"
        if tag == "a" and (a.get("href") or "").startswith("#"):
            assert a["href"][1:] in ids, f"Sprunglink {a['href']!r} zeigt ins Leere"


def test_tabs_haben_rollen_und_gehoeren_zu_ihrem_bereich():
    rollen = [(t, a) for t, a in SEITE.elemente if a.get("role")]
    tablists = [a for _, a in rollen if a["role"] == "tablist"]
    tabs = [a for _, a in rollen if a["role"] == "tab"]
    panels = {a["id"]: a for _, a in rollen if a["role"] == "tabpanel"}
    assert tablists and len(tabs) == 4 and len(panels) == 4
    for tab in tabs:
        assert tab.get("aria-selected") in ("true", "false")
        assert tab.get("aria-controls") in panels
        assert panels[tab["aria-controls"]].get("aria-labelledby") == tab["id"]
    assert sum(t["aria-selected"] == "true" for t in tabs) == 1
    # Nur der aktive Tab ist per Tab-Taste erreichbar (Pfeiltasten wechseln)
    assert sorted(t.get("tabindex") for t in tabs) == ["-1", "-1", "-1", "0"]


def test_pfeiltasten_und_aria_in_app_js():
    for taste in ("ArrowRight", "ArrowLeft", "Home", "End"):
        assert f"'{taste}'" in JS, f"Tabs: Taste {taste} fehlt"
    for taste in ("ArrowDown", "ArrowUp", "'Enter'", "' '"):
        assert taste in JS, f"Tabellenzeilen: {taste} fehlt"
    assert "'aria-pressed'" in JS and "aria-selected" in JS and "aria-expanded" in JS
    assert "tabindex: '0'" in JS, "Tabellenzeilen müssen per Tastatur erreichbar sein"


def test_wartebildschirm_toast_und_spaltenkopf_sind_zugaenglich():
    def finde(**kw):
        return [a for _, a in SEITE.elemente if all(a.get(k) == v for k, v in kw.items())]

    busy = finde(id="busy-card")[0]
    assert busy.get("role") == "alertdialog" and busy.get("aria-labelledby") and busy.get("tabindex") == "-1"
    assert finde(id="toast")[0].get("role") == "status"
    assert finde(id="toast-fehler")[0].get("role") == "alert"
    assert finde(id="chk-alle")[0].get("aria-label"), "Checkbox im Spaltenkopf braucht eine Beschriftung"
    assert finde(id="theme-toggle")[0].get("aria-label") and finde(id="theme-toggle")[0].get("aria-pressed")
    # Icons sind für Screenreader unsichtbar
    for t, a in SEITE.elemente:
        if t == "svg":
            assert a.get("aria-hidden") == "true", "SVG-Icon ohne aria-hidden"
        if t == "img":
            assert "alt" in a, "Bild ohne alt"
    # Jedes Eingabefeld hat ein Label oder eine Beschriftung
    labels = {a.get("for") for t, a in SEITE.elemente if t == "label"}
    for t, a in SEITE.elemente:
        if t in ("input", "select") and a.get("type") != "file":
            assert a.get("aria-label") or a.get("_in_label") or a.get("id") in labels, f"<{t} id={a.get('id')!r}> ohne Beschriftung"
    # Info-Icons: Name, Popup-Hinweis
    infos = [a for t, a in SEITE.elemente if t == "button" and "info-icon" in (a.get("class") or "")]
    assert infos and all(a.get("aria-label") and a.get("data-info") for a in infos)
    # Jede Info, auf die das HTML verweist, hat in app.js einen Text
    for a in infos:
        assert re.search(rf"\b{re.escape(a['data-info'])}:\s*\{{\s*titel", JS), f"Kein Erklärtext für {a['data-info']!r}"


def test_keine_externen_quellen_und_keine_festen_adressen():
    for tag, a in SEITE.elemente:
        for attr in ("src", "href"):
            wert = a.get(attr) or ""
            assert not re.match(r"^(https?:)?//", wert), f"<{tag}> {attr}={wert!r} zeigt nach außen"
    for name, quelle in (("index.html", HTML), ("app.js", JS), ("style.css", CSS)):
        assert not re.search(r"https?://", quelle), f"{name}: feste Internetadresse"
        assert not re.search(r"localhost|127\.0\.0\.1|:5100", quelle), f"{name}: fester Rechner oder Port"
        assert "@import" not in quelle and "url(" not in quelle.replace("url(#", ""), f"{name}: externe Ressource"
    skripte = [a for t, a in SEITE.elemente if t == "script"]
    assert skripte and all(not (a.get("src") or "").startswith(("http", "//")) for a in skripte)


def test_jede_fetch_adresse_beginnt_mit_api():
    aufrufe = [m.start() for m in re.finditer(r"\bfetch\(", JS)]
    assert aufrufe, "fetch nicht gefunden"
    for pos in aufrufe:
        arg = JS[pos + len("fetch("):pos + len("fetch(") + 40].lstrip()
        assert re.match(r"""['"`]/api(/|['"`])""", arg), f"fetch ohne /api/-Pfad: {arg!r}"
    for muster in re.findall(r"""(?:href|src)\s*[:=]\s*[`'"]([^`'"]+)""", JS):
        if muster.startswith("/"):
            assert muster.startswith("/api/"), f"Adresse außerhalb von /api/: {muster!r}"


def test_kein_loeschen_text_und_kein_innerhtml():
    for name, quelle in (("index.html", HTML), ("app.js", JS)):
        assert not re.search(r"[Ll]ösch|Loesch", quelle), f"{name}: 'Löschen' darf nirgends mehr stehen (Archivieren statt Löschen)"
    assert "Archivieren" in HTML + JS and "Wiederherstellen" in HTML + JS and "Archivierte anzeigen" in HTML
    assert "innerHTML" not in JS and "outerHTML" not in JS and "insertAdjacentHTML" not in JS, "Texte nur über textContent setzen"
    assert "eval(" not in JS and "document.write" not in JS


def test_fehlertexte_einfach():
    assert "Der Server antwortet nicht. Läuft das Dashboard noch?" in JS
    assert not re.search(r"Fehler \$\{", JS), "Rohe Fehlercodes dürfen nicht angezeigt werden"
    assert "daten.fehler" in JS, "Der Satz aus der API muss angezeigt werden"


def test_wichtige_aussagen_stehen_im_dashboard():
    for wort in ("Unvollständig gelesen", "Die Bilanz stimmt nicht", "Dort steht jetzt", "Im PDF: alt → neu"):
        assert wort in JS, f"Fehlt: {wort}"
    for feld in ("vollstaendig", "luecken", "hinweise", "wirkung", "nachkontrolle", "excel_text", "ergebnisse"):
        assert feld in JS, f"app.js liest das Feld {feld!r} nicht"
    assert "Noch kein Lauf geöffnet" in HTML.split('id="tab-farben"')[1], "Tab 'Farben im PDF' braucht einen Hinweis ohne Lauf"


# ----------------------------------------------------------------------------------------------
# CSS
# ----------------------------------------------------------------------------------------------
def test_css_hat_keine_farbwerte_ausser_in_token_bloecken():
    assert len(REGELN) > 80, "CSS-Zerlegung hat nichts gefunden"
    fehler = []
    for ctx, kopf, koerper in REGELN:
        if _ist_token_block(ctx, kopf):
            continue
        for m in re.finditer(r"#[0-9a-fA-F]{3,8}\b|\b(?:rgba?|hsla?)\(", koerper):
            fehler.append(f"{kopf}: {m.group(0)}")
    assert not fehler, "Farbwerte im Komponenten-CSS (nur als Token erlaubt): " + "; ".join(fehler[:10])


def test_css_tokens_hell_dunkel_system_sind_stimmig():
    hell = _tokens(":root", False)
    dunkel = _tokens(':root[data-theme="dark"]', False)
    system = _tokens(':root:not([data-theme="light"])', True)
    assert dunkel == system, "Dunkel-Schalter und Dunkel-System müssen gleich sein"
    for name in dunkel:
        assert name in hell, f"Dunkel-Token {name} hat kein Gegenstück in :root"
    for pflicht in ("--rr-blue", "--rr-blue-dark", "--rr-blue-light", "--green", "--amber", "--red", "--white", "--grey-50", "--grey-100",
                    "--grey-150", "--grey-300", "--grey-500", "--ink", "--shadow-xs", "--shadow-sm", "--shadow-md", "--shadow-lg", "--shadow-xl"):
        assert pflicht in hell, f"Token {pflicht} fehlt"
    for farbe in ("--green", "--amber", "--red", "--rr-blue-ink", "--green-wash", "--amber-wash", "--red-wash", "--white", "--ink", "--grey-500"):
        assert farbe in dunkel, f"Dark-Mode-Wert für {farbe} fehlt"
    assert "--blue:" not in CSS and "var(--blue)" not in CSS, "Alter Token --blue statt --rr-blue"


def test_css_alle_verwendeten_variablen_sind_definiert():
    definiert = set(_tokens(":root", False))
    verwendet = set(re.findall(r"var\((--[\w-]+)", CSS))
    assert not (verwendet - definiert), f"Nicht definierte Variablen: {sorted(verwendet - definiert)}"


@pytest.mark.parametrize("thema", ["hell", "dunkel"])
def test_css_kontraste_text_auf_flaeche(thema):
    t = dict(_tokens(":root", False))
    if thema == "dunkel":
        t.update(_tokens(':root[data-theme="dark"]', False))
    paare = [
        ("--ink", "--white"), ("--ink", "--grey-100"), ("--ink", "--grey-150"), ("--ink", "--page-bg"),
        ("--grey-500", "--white"), ("--grey-500", "--grey-100"), ("--grey-500", "--page-bg"),
        ("--rr-blue-ink", "--white"), ("--rr-blue-ink", "--grey-100"), ("--rr-blue-ink", "--rr-blue-light"),
        ("--green", "--white"), ("--green", "--green-wash"), ("--amber", "--white"), ("--amber", "--amber-wash"),
        ("--red", "--white"), ("--red", "--red-wash"), ("--on-accent", "--rr-blue"), ("--on-accent", "--rr-blue-dark"),
        ("--on-accent", "--red-solid"), ("--toast-ink", "--toast-bg"),
    ]
    schlecht = []
    for vorn, hinten in paare:
        k = _kontrast(t[vorn], t[hinten])
        if k < 4.5:
            schlecht.append(f"{vorn} auf {hinten}: {k:.2f}")
    assert not schlecht, f"Kontrast unter 4,5:1 ({thema}): {schlecht}"


def test_css_fokus_und_kein_outline_none():
    for ctx, kopf, koerper in REGELN:
        if re.search(r"outline\s*:\s*(none|0)\b", koerper):
            assert ":focus:not(:focus-visible)" in kopf, f"outline:none ohne Ersatz in {kopf!r}"
    fokus = [kopf for _, kopf, _ in REGELN if ":focus-visible" in kopf]
    assert fokus, "Keine :focus-visible-Regel"
    alle = " ".join(fokus)
    for element in ("a:focus-visible", "button:focus-visible", "input:focus-visible", "select:focus-visible", "summary:focus-visible", "[tabindex]:focus-visible"):
        assert element in alle, f"Fokus-Ring fehlt für {element}"
    assert re.search(r"outline:\s*2px solid", CSS)


def test_css_bewegung_reduzieren_und_dark_mode_system():
    assert any("prefers-reduced-motion: reduce" in c for ctx, _, _ in REGELN for c in ctx)
    assert any(c.startswith(_DUNKEL_MEDIA) for ctx, _, _ in REGELN for c in ctx)
    assert ':root[data-theme="dark"]' in CSS


def test_css_touchziele_mindestens_32px():
    hoehen = {}
    for ctx, kopf, koerper in REGELN:
        if ctx:
            continue
        d = _deklarationen(koerper)
        for sel in kopf.split(","):
            hoehen[sel.strip()] = d
    for sel in (".btn", ".btn-small", ".chip", ".page-tab", ".eingabe", ".status-pill"):
        mh = hoehen[sel].get("min-height", "0px")
        assert int(re.match(r"(\d+)", mh).group(1)) >= 32, f"{sel}: min-height {mh} < 32px"
    for sel in (".icon-button", ".info-popup-close"):
        assert int(hoehen[sel]["width"][:-2]) >= 32 and int(hoehen[sel]["height"][:-2]) >= 32, sel
    assert re.search(r"\.chk-ziel\s*\{[^}]*width:\s*32px[^}]*height:\s*32px", CSS)


# ----------------------------------------------------------------------------------------------
# JavaScript-Syntax (nur wenn node vorhanden ist)
# ----------------------------------------------------------------------------------------------
def _node() -> str | None:
    return shutil.which("node") or (str(p) if (p := Path("/opt/node22/bin/node")).exists() else None)


@pytest.mark.skipif(_node() is None, reason="node nicht vorhanden")
def test_app_js_syntax_mit_node():
    r = subprocess.run([_node(), "--check", str(FRONTEND / "app.js")], capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
