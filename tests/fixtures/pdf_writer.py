"""Winziger PDF-Schreiber NUR fuer Tests (erfundene Beispieldokumente, Regel 16).

Erzeugt PDFs mit Standard-Schrift Helvetica, farbigem Text, Hintergrund-
Markierungen, Durchstreichungen und PDF-Kommentaren (Annotationen). Keine
zusaetzliche Bibliothek noetig -- die PDF-Struktur wird von Hand geschrieben.
(Nur fuer Bilder wird Pillow benutzt, das ohnehin installiert ist.)

Markup fuer Absaetze:
    {+neuer Text+}   hinzugekommen (blau)
    {-alter Text-}   entfaellt (rot + durchgestrichen)
    Rest             unveraendert
Beispiel: "Die Kapazitaet betraegt {-100 kWh-} {+120 kWh+} je Rack."

Mischformen (zusaetzlich, alles optional):  {@stil:Text@}
    bu  blau + unterstrichen      (gilt als NEU)
    bs  blau + durchgestrichen    (gilt als ENTFAELLT, widerspruechlich -> "pruefen")
    ru  rot + unterstrichen       (gilt als ENTFAELLT)
    r   nur rot                   (ENTFAELLT)
    s   nur durchgestrichen       (ENTFAELLT)
    l   Hyperlink (blau 0563C1, unterstrichen, Link-Annotation), unveraendert
    gn  gruen (unbekannte Farbe), gilt als NEU
Verschmolzene Ersetzung ohne Leerzeichen: "{-100-}{+120+} kWh".

Neue Faehigkeiten von PdfDokument (alle optional, Standardverhalten bleibt
byte-gleich): Silbentrennung ("-" oder weicher Trennstrich), Fussnoten (8 pt),
Aufzaehlungen, zwei Spalten, Tabellen (Farbkopf, Baenderung), Bilder,
Scan-Seiten, Farbraeume (RGB/CMYK/Gray), Rot-/Blautoene, Hyperlinks,
Seitenumbruch mitten im Absatz, Kopf-/Fusszeile nur auf bestimmten Seiten,
Sonderzeichen, Seitendrehung (/Rotate).
"""
from __future__ import annotations

import re
import zlib

SEITE_B, SEITE_H = 595.0, 842.0
RAND_L, RAND_R, RAND_O, RAND_U = 56.0, 56.0, 64.0, 70.0

_BREITEN = {
    " ": 278, "a": 556, "b": 556, "c": 500, "d": 556, "e": 556, "f": 278, "g": 556,
    "h": 556, "i": 222, "j": 222, "k": 500, "l": 222, "m": 833, "n": 556, "o": 556,
    "p": 556, "q": 556, "r": 333, "s": 500, "t": 278, "u": 556, "v": 500, "w": 722,
    "x": 500, "y": 500, "z": 500, "A": 667, "B": 667, "C": 722, "D": 722, "E": 667,
    "F": 611, "G": 778, "H": 722, "I": 278, "J": 500, "K": 667, "L": 556, "M": 833,
    "N": 722, "O": 778, "P": 667, "Q": 778, "R": 722, "S": 667, "T": 611, "U": 722,
    "V": 667, "W": 944, "X": 667, "Y": 667, "Z": 611, ".": 278, ",": 278, "-": 333,
    ":": 278, ";": 278, "(": 333, ")": 333, "/": 278, "ä": 556, "ö": 556, "ü": 556,
    "Ä": 667, "Ö": 778, "Ü": 722, "ß": 611, "%": 889, "•": 350,
    # neu (nur fuer neue Zeichen; bisherige Zeichen behalten ihre Breite)
    "­": 333, " ": 278, "„": 333, "“": 333, "”": 333, "¹": 333, "²": 333, "³": 333,
    "|": 260, "°": 400, "*": 389,
}

# Farben der Mischformen und Farbtoene
LINK_BLAU = (0.0196, 0.3882, 0.7569)      # 0563C1
GRUEN = (0.0, 0.6, 0.0)                   # 009900 (unbekannte Farbe)

# Stil -> (Farbschluessel, Linie). Farbschluessel: n=neu-Farbe, g=weg-Farbe, link, gruen, None=Grundfarbe
_MISCH = {
    "bs": ("n", "strich"),
    "bu": ("n", "unter"),
    "ru": ("g", "unter"),
    "r": ("g", None),
    "s": (None, "strich"),
    "l": ("link", "unter"),
    "gn": ("gruen", None),
}
# Bedeutung der Stile: im alten / im neuen Stand sichtbar
_STIL_ALT = {"k", "g", "bs", "ru", "r", "s", "l"}
_STIL_NEU = {"k", "n", "bu", "gn", "l"}

_VOKALE = set("aeiouäöüAEIOUÄÖÜ")


def hex_farbe(h: str) -> tuple[float, float, float]:
    """'1F4E9F' -> (0.1216, 0.3059, 0.6235)."""
    h = h.lstrip("#")
    return tuple(round(int(h[i:i + 2], 16) / 255.0, 4) for i in (0, 2, 4))  # type: ignore[return-value]


def breite(text: str, groesse: float, fett: bool = False) -> float:
    faktor = 1.05 if fett else 1.0
    return sum(_BREITEN.get(c, 556) for c in text) * groesse / 1000.0 * faktor


def _esc(text: str) -> bytes:
    return (text.encode("cp1252", errors="replace")
            .replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)"))


_MARKE = re.compile(r"\{\+(.*?)\+\}|\{-(.*?)-\}|\{@(\w+):(.*?)@\}", re.S)
_LEER = r"[^\S ]+"        # Leerraum, aber das geschuetzte Leerzeichen trennt keine Woerter


def _teile(markup: str) -> list[tuple[str, str]]:
    teile: list[tuple[str, str]] = []
    pos = 0
    for m in _MARKE.finditer(markup):
        if m.start() > pos:
            teile.append((markup[pos:m.start()], "k"))
        if m.group(1) is not None:
            teile.append((m.group(1), "n"))
        elif m.group(2) is not None:
            teile.append((m.group(2), "g"))
        else:
            teile.append((m.group(4), m.group(3)))
        pos = m.end()
    if pos < len(markup):
        teile.append((markup[pos:], "k"))
    return teile


def parse_markup(markup: str) -> list[list[tuple[str, str]]]:
    """Markup -> Liste von Woertern; ein Wort ist eine Liste von (Text, Stil), Stil: k|n|g|bu|bs|ru|r|s|l|gn."""
    woerter: list[list[tuple[str, str]]] = []
    aktuell: list[tuple[str, str]] = []
    for text, stil in _teile(markup):
        for stueck in re.split(f"({_LEER})", text):
            if not stueck:
                continue
            if re.fullmatch(_LEER, stueck):
                if aktuell:
                    woerter.append(aktuell)
                    aktuell = []
            else:
                aktuell.append((stueck, stil))
    if aktuell:
        woerter.append(aktuell)
    return woerter


def zustand(markup: str, neu: bool) -> str:
    """Text im alten (neu=False) oder neuen Stand, ohne Markup."""
    s = _MARKE.sub(lambda m: (m.group(4) if (m.group(3) in (_STIL_NEU if neu else _STIL_ALT)) else "")
                   if m.group(3) is not None else m.group(0), markup)
    if neu:
        s = re.sub(r"\{-.*?-\}", "", s, flags=re.S)
        s = re.sub(r"\{\+(.*?)\+\}", r"\1", s, flags=re.S)
    else:
        s = re.sub(r"\{\+.*?\+\}", "", s, flags=re.S)
        s = re.sub(r"\{-(.*?)-\}", r"\1", s, flags=re.S)
    return re.sub(_LEER, " ", s).strip(" \t\r\n\f\v")


def markierte_woerter(markup: str) -> int:
    """Anzahl der Woerter, die im PDF als Aenderung markiert sind (je Stil-Abschnitt gezaehlt)."""
    n = 0
    for text, stil in _teile(markup):
        if stil not in ("k", "l"):
            n += len([w for w in re.split(_LEER, text) if w])
    return n


# ---------------------------------------------------------------- Bilder (Pillow)
def muster_bild(breite_px: int = 240, hoehe_px: int = 140, seed: int = 0):
    """Erfundenes Diagramm-Bild (Balken, Linien; enthaelt auch blaue Flaechen)."""
    import random

    from PIL import Image, ImageDraw
    rng = random.Random(seed)
    img = Image.new("RGB", (breite_px, hoehe_px), (246, 246, 246))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, breite_px - 1, hoehe_px - 1], outline=(120, 120, 120))
    n = rng.randint(4, 7)
    bw = (breite_px - 30) // n
    for i in range(n):
        h = rng.randint(20, hoehe_px - 30)
        farbe = rng.choice([(68, 114, 196), (46, 117, 182), (180, 198, 231), (192, 0, 0), (120, 120, 120)])
        x0 = 15 + i * bw
        d.rectangle([x0, hoehe_px - 10 - h, x0 + bw - 6, hoehe_px - 10], fill=farbe)
    d.line([10, hoehe_px - 10, breite_px - 10, hoehe_px - 10], fill=(0, 0, 0), width=1)
    return img


_UMLAUT = {"ä": "a", "ö": "o", "ü": "u", "Ä": "A", "Ö": "O", "Ü": "U"}


def _pil_text(d, font, x: float, y: float, text: str, fill) -> float:
    """Zeichnet Text als Pixel. Die eingebaute Pillow-Schrift hat keine Umlaute: sie werden aus Buchstabe + Punkten gebaut."""
    for c in text:
        ersatz = "ss" if c == "ß" else _UMLAUT.get(c, c)
        d.text((x, y), ersatz, fill=fill, font=font)
        w = font.getlength(ersatz)
        if c in _UMLAUT:
            r = max(1.0, font.size * 0.055)
            oben = y + font.size * (0.02 if c.isupper() else 0.10)
            for dx in (0.28, 0.68):
                d.ellipse([x + w * dx - r, oben, x + w * dx + r, oben + 2 * r], fill=fill)
        x += w
    return x


def _pil_breite(font, text: str) -> float:
    return font.getlength(text.replace("ß", "ss").translate(str.maketrans(_UMLAUT)))


def scan_bild(markups: list[tuple[str, str]], farbe_neu=(0, 0, 255), farbe_weg=(255, 0, 0),
              faktor: float = 1.5, seed: int = 0):
    """Zeichnet Text als Pixel auf eine Seite (wie ein Scan). markups: [('h'|'p', markup)].

    Blau = neu, rot + Strich = entfaellt. Keine Textebene -- nur Pixel.
    """
    import random

    from PIL import Image, ImageDraw, ImageFont
    rng = random.Random(seed)
    b, h = int(SEITE_B * faktor), int(SEITE_H * faktor)
    img = Image.new("RGB", (b, h), (250, 250, 247))
    d = ImageDraw.Draw(img)
    f_norm = ImageFont.load_default(size=max(8, round(11 * faktor)))
    f_gross = ImageFont.load_default(size=max(9, round(14 * faktor)))
    y = RAND_O * faktor
    for art, markup in markups:
        gross = art == "h"
        font = f_gross if gross else f_norm
        zeile_h = (18 if gross else 15) * faktor
        x = RAND_L * faktor
        max_x = (SEITE_B - RAND_R) * faktor
        platz = _pil_breite(font, " ")
        for wort in parse_markup(markup):
            text = "".join(t for t, _ in wort)
            w = _pil_breite(font, text)
            if x > RAND_L * faktor and x + w > max_x:
                y += zeile_h
                x = RAND_L * faktor
            for t, stil in wort:
                tw = _pil_breite(font, t)
                farbe = (20, 20, 20)
                strich = False
                if stil in ("n", "bu", "bs"):
                    farbe = farbe_neu
                elif stil in ("g", "ru", "r"):
                    farbe, strich = farbe_weg, stil == "g"
                elif stil == "gn":
                    farbe = (0, 150, 0)
                elif stil == "l":
                    farbe = (5, 99, 193)
                _pil_text(d, font, x, y, t, farbe)
                if gross:
                    _pil_text(d, font, x + 0.7, y, t, farbe)
                if strich or stil in ("bs", "s"):
                    d.line([x, y + zeile_h * 0.45, x + tw, y + zeile_h * 0.45], fill=farbe, width=max(1, round(faktor)))
                x += tw
            x += platz
        y += zeile_h + (8 if gross else 9) * faktor
    for _ in range(60):     # ein paar Staubkoernchen
        px, py = rng.randrange(b), rng.randrange(h)
        d.point((px, py), fill=(200, 200, 196))
    return img


class _Seite:
    def __init__(self) -> None:
        self.inhalt: list[bytes] = []
        self.annots: list[dict] = []
        self.links: list[dict] = []
        self.fussnoten: list[list[str]] = []     # je Fussnote die Zeilen
        self.fuss_hoehe: float = 0.0
        self.bilder: list[int] = []
        self.nur_bild: bool = False
        self.rotate: int = 0


class PdfDokument:
    """variante: farbe | hintergrund | strich | annot.

    Optionen (alle optional):
      farbe_neu / farbe_weg   Hex-Farben (z. B. "1F4E9F", "C00000") statt 0000FF / FF0000
      farbraum                "rgb" | "cmyk": Farbraum der bunten Farben
      grundfarbraum           "rgb" | "gray" | "cmyk": Farbraum von Schwarz/Grau
      kopf_fuss               True/False oder Menge von Seitennummern (1 = erste), auf denen Kopf/Fuss steht
      kopf_text, kopf_farbe, fuss_vorlage   Text, Farbe (RGB 0..1) der Kopfzeile, Vorlage der Fusszeile
      trennung                None | "ascii" | "weich" | "gemischt": Silbentrennung am Zeilenende
    """

    def __init__(self, variante: str = "farbe", kopf_fuss=True, *, farbe_neu: str | None = None,
                 farbe_weg: str | None = None, farbraum: str = "rgb", grundfarbraum: str = "rgb",
                 kopf_text: str = "Beispieldokument - erfunden", kopf_farbe=(0.4, 0.4, 0.4),
                 fuss_vorlage: str = "Seite {nr} von {n}", trennung: str | None = None,
                 link_url: str = "https://beispiel.example/dokument") -> None:
        self.variante = variante
        self.kopf_fuss = kopf_fuss
        self.farbe_neu = hex_farbe(farbe_neu) if farbe_neu else (0, 0, 1)
        self.farbe_weg = hex_farbe(farbe_weg) if farbe_weg else (1, 0, 0)
        self.farbraum = farbraum
        self.grundfarbraum = grundfarbraum
        self.kopf_text = kopf_text
        self.kopf_farbe = tuple(kopf_farbe)
        self.fuss_vorlage = fuss_vorlage
        self.trennung = trennung
        self.link_url = link_url
        self.seiten: list[_Seite] = []
        self._bilder: list[tuple[int, int, bytes]] = []
        self._seite_voll = False
        self.stats = {"trennungen": 0, "trennung_seitenende": 0, "links": 0}
        self.getrennt: list[tuple[str, tuple[str, ...], str]] = []   # (Wort, Stile, Trennzeichen)
        self.letzter_block: tuple[int, int] = (1, 1)
        self.y = RAND_O
        self._neue_seite()

    # ---- Seiten --------------------------------------------------------
    def _neue_seite(self) -> None:
        self.seiten.append(_Seite())
        self.y = RAND_O
        self._seite_voll = False

    def seitenumbruch(self) -> None:
        """Beginnt sicher eine neue Seite (nur wenn auf der aktuellen schon etwas steht)."""
        s = self._s
        if s.inhalt or s.nur_bild or s.fussnoten:
            self._neue_seite()

    def _bereit(self) -> None:
        if self._seite_voll:
            self._neue_seite()

    @property
    def _s(self) -> _Seite:
        return self.seiten[-1]

    @property
    def seitenzahl(self) -> int:
        return len(self.seiten)

    def _unten(self) -> float:
        return SEITE_H - RAND_U - self._s.fuss_hoehe

    def seite_rotieren(self, grad: int, seite: int | None = None) -> None:
        """Setzt /Rotate (90, 180, 270) auf eine Seite (1 = erste; Standard: aktuelle Seite)."""
        (self.seiten[-1] if seite is None else self.seiten[seite - 1]).rotate = grad

    # ---- Farben --------------------------------------------------------
    def _fb(self, farbe, strich: bool = False) -> str:
        r, g, b = farbe
        neutral = r == g == b
        raum = self.grundfarbraum if neutral else self.farbraum
        if raum == "gray":
            wert = r if neutral else round(0.299 * r + 0.587 * g + 0.114 * b, 4)
            return f"{wert} {'G' if strich else 'g'}"
        if raum == "cmyk":
            k = round(1 - max(r, g, b), 4)
            if k >= 1:
                c = m = y = 0
            else:
                c, m, y = (round((1 - v - k) / (1 - k), 4) for v in (r, g, b))
            return f"{c} {m} {y} {k} {'K' if strich else 'k'}"
        return f"{r} {g} {b} {'RG' if strich else 'rg'}"

    # ---- Zeichnen ------------------------------------------------------
    def _op(self, s: str) -> None:
        self._s.inhalt.append(s.encode("latin-1") + b"\n")

    def _text_bytes(self, x: float, top: float, text: str, groesse: float, fett: bool, farbe) -> bytes:
        basis = SEITE_H - (top + groesse * 0.8)
        return (f"BT /{'F2' if fett else 'F1'} {groesse} Tf {self._fb(farbe)} {x:.2f} {basis:.2f} Td (".encode("latin-1")
                + _esc(text) + b") Tj ET\n")

    def _text(self, x: float, top: float, text: str, groesse: float, fett: bool, farbe) -> None:
        self._s.inhalt.append(self._text_bytes(x, top, text, groesse, fett, farbe))

    def _rechteck(self, x0, top, x1, bottom, farbe) -> None:
        self._op(f"{self._fb(farbe)} {x0:.2f} {SEITE_H - bottom:.2f} {x1 - x0:.2f} {bottom - top:.2f} re f")

    def _linie_op(self, x0, x1, top, farbe, staerke=0.8) -> str:
        y = SEITE_H - top
        return f"{self._fb(farbe, True)} {staerke} w {x0:.2f} {y:.2f} m {x1:.2f} {y:.2f} l S"

    def _linie(self, x0, x1, top, farbe, staerke=0.8) -> None:
        self._op(self._linie_op(x0, x1, top, farbe, staerke))

    def _vlinie(self, x, top, bottom, farbe, staerke=0.5) -> None:
        self._op(f"{self._fb(farbe, True)} {staerke} w {x:.2f} {SEITE_H - top:.2f} m {x:.2f} {SEITE_H - bottom:.2f} l S")

    def _farben(self, farben) -> tuple:
        if farben is None:
            return self.farbe_neu, self.farbe_weg
        neu, weg = farben
        return (hex_farbe(neu) if isinstance(neu, str) else neu,
                hex_farbe(weg) if isinstance(weg, str) else weg)

    def _wort(self, x: float, top: float, segmente, groesse: float, fett: bool,
              farben=None, grund=(0, 0, 0)) -> float:
        """Zeichnet ein Wort (evtl. mehrere Stil-Segmente). Gibt die Endposition zurueck."""
        v = self.variante
        fn, fw = self._farben(farben)
        for text, stil in segmente:
            w = breite(text, groesse, fett)
            farbe = grund
            strich = None
            unter = None
            if stil == "n":
                if v == "hintergrund":
                    self._rechteck(x, top - 1, x + w, top + groesse + 1, (0.72, 0.86, 1.0))
                elif v in ("farbe", "strich"):
                    farbe = fn
                elif v == "annot":
                    self._s.annots.append(dict(art="Highlight", c=(0, 0, 1), box=(x, top, x + w, top + groesse)))
            elif stil == "g":
                if v == "farbe":
                    farbe, strich = fw, fw
                elif v == "hintergrund":
                    self._rechteck(x, top - 1, x + w, top + groesse + 1, (1.0, 0.8, 0.8))
                    strich = (0, 0, 0)
                elif v == "strich":
                    strich = (0, 0, 0)
                elif v == "annot":
                    self._s.annots.append(dict(art="StrikeOut", c=(1, 0, 0), box=(x, top, x + w, top + groesse)))
            elif stil in _MISCH:
                schluessel, linie = _MISCH[stil]
                farbe = {"n": fn, "g": fw, "link": LINK_BLAU, "gruen": GRUEN, None: grund}[schluessel]
                if linie == "strich":
                    strich = farbe
                elif linie == "unter":
                    unter = farbe
                if stil == "l":
                    self._s.links.append(dict(box=(x, top, x + w, top + groesse), url=self.link_url))
                    self.stats["links"] += 1
            self._text(x, top, text, groesse, fett, farbe)
            if strich:
                self._linie(x - 0.5, x + w + 0.5, top + groesse * 0.55, strich)
            if unter:
                self._linie(x, x + w, top + groesse * 0.93, unter)
            x += w
        return x

    # ---- Silbentrennung -------------------------------------------------
    def _trenne(self, wort, x: float, max_x: float, groesse: float, fett: bool, modus: str):
        """Teilt ein Wort so, dass der linke Teil samt Trennstrich noch in die Zeile passt.

        Gibt (links, rechts) als Segmentlisten zurueck oder None. Der Trennstrich bekommt
        den Stil des letzten linken Segments (auch bei farbigen Woertern).
        """
        text = "".join(t for t, _ in wort)
        if len(text) < 6 or not all(c.isalpha() for c in text):
            return None
        if modus == "gemischt":
            strich = "-" if self.stats["trennungen"] % 2 == 0 else "­"
        else:
            strich = "-" if modus == "ascii" else "­"
        platz = max_x - x
        passend = [i for i in range(2, len(text) - 1)
                   if text[i - 1] in _VOKALE and text[i] not in _VOKALE
                   and breite(text[:i] + strich, groesse, fett) <= platz]
        # bevorzugt vor einem einzelnen Konsonanten zwischen Vokalen (Um-ge-bung), sonst nach einem Vokal
        kandidaten = [i for i in passend if i + 1 < len(text) and text[i + 1] in _VOKALE] or passend
        if not kandidaten:
            return None
        i = max(kandidaten)
        links: list[tuple[str, str]] = []
        rechts: list[tuple[str, str]] = []
        rest = i
        for seg, stil in wort:
            if rest >= len(seg):
                links.append((seg, stil))
                rest -= len(seg)
            elif rest > 0:
                links.append((seg[:rest], stil))
                rechts.append((seg[rest:], stil))
                rest = 0
            else:
                rechts.append((seg, stil))
        links[-1] = (links[-1][0] + strich, links[-1][1])
        self.stats["trennungen"] += 1
        self.getrennt.append((text, tuple(sorted({s for _, s in wort})), strich))
        return links, rechts

    # ---- Bloecke --------------------------------------------------------
    def block(self, markup: str, fett: bool = False, groesse: float = 11.0,
              zeilenabstand: float = 15.0, abstand_danach: float = 9.0, spalte_x: float = RAND_L, *,
              spalte_bis: float | None = None, farben=None, trennung: str | None = None,
              marker: str | None = None, marker_stil: str = "k", einzug: float = 14.0,
              umbruch_nach_zeile: int | None = None) -> None:
        """Schreibt einen Absatz.

        spalte_bis           rechter Rand (z. B. fuer Spalten)
        farben               (neu, weg) als Hex-Strings, nur fuer diesen Absatz
        trennung             Silbentrennung nur fuer diesen Absatz
        marker/einzug        Aufzaehlungszeichen ("•", "–", "1.", "a)") mit haengendem Einzug
        umbruch_nach_zeile   erzwingt nach dieser Zeile einen Seitenumbruch mitten im Absatz
        """
        self._bereit()
        woerter = parse_markup(markup)
        max_x = SEITE_B - RAND_R if spalte_bis is None else spalte_bis
        text_x = spalte_x + (einzug if marker else 0.0)
        x = text_x
        platz = breite(" ", groesse, fett)
        modus = trennung or self.trennung
        if self.y + zeilenabstand > self._unten():
            self._neue_seite()
        start_seite = len(self.seiten)
        if marker:
            self._wort(spalte_x, self.y, [(marker, marker_stil)], groesse, fett, farben)
        zeilen = 0
        for wort in woerter:
            text = "".join(t for t, _ in wort)
            w = breite(text, groesse, fett)
            if x > text_x and x + w > max_x:
                teil = self._trenne(wort, x, max_x, groesse, fett, modus) if modus else None
                if teil:
                    x = self._wort(x, self.y, teil[0], groesse, fett, farben)
                    wort = teil[1]
                    text = "".join(t for t, _ in wort)
                    w = breite(text, groesse, fett)
                zeilen += 1
                x = text_x
                if umbruch_nach_zeile and zeilen == umbruch_nach_zeile:
                    self._neue_seite()
                    if teil:
                        self.stats["trennung_seitenende"] += 1
                else:
                    self.y += zeilenabstand
                    if self.y + zeilenabstand > self._unten():
                        self._neue_seite()
                        if teil:
                            self.stats["trennung_seitenende"] += 1
            x = self._wort(x, self.y, wort, groesse, fett, farben)
            x += platz
        self.y += zeilenabstand + abstand_danach
        self.letzter_block = (start_seite, len(self.seiten))      # Seite des Anfangs, Seite des Endes

    def ueberschrift(self, markup: str, groesse: float = 14, zeilenabstand: float = 18) -> None:
        self.y += 6
        self.block(markup, fett=True, groesse=groesse, zeilenabstand=zeilenabstand, abstand_danach=8)

    def _zeilen_schaetzen(self, markup: str, x0: float, x1: float, groesse: float = 11.0) -> int:
        x, n = x0, 1
        for wort in parse_markup(markup):
            w = breite("".join(t for t, _ in wort), groesse)
            if x > x0 and x + w > x1:
                n += 1
                x = x0
            x += w + breite(" ", groesse)
        return n

    def zwei_spalten(self, links: list[str], rechts: list[str], abstand: float = 18.0, **kw) -> None:
        """Setzt zwei Textspalten nebeneinander (rechte Spalte endet am rechten Rand)."""
        self._bereit()
        spalte = (SEITE_B - RAND_L - RAND_R - abstand) / 2
        l0, l1 = RAND_L, RAND_L + spalte
        r0, r1 = l1 + abstand, SEITE_B - RAND_R
        hoehe = max(sum(self._zeilen_schaetzen(m, l0, l1) * 15 + 9 for m in links),
                    sum(self._zeilen_schaetzen(m, r0, r1) * 15 + 9 for m in rechts)) + 6
        if self.y + hoehe > self._unten():
            self._neue_seite()
        y0 = self.y
        for m in links:
            self.block(m, spalte_x=l0, spalte_bis=l1, **kw)
        yl = self.y
        self.y = y0
        for m in rechts:
            self.block(m, spalte_x=r0, spalte_bis=r1, **kw)
        self.y = max(self.y, yl)

    # ---- Tabelle ---------------------------------------------------------
    def tabelle(self, zeilen: list[list[str]], breiten: list[float] | None = None, *,
                kopf_farbe: str | None = "4472C4", band_farbe: str | None = "B4C6E7",
                groesse: float = 10.0, zeilenhoehe: float = 18.0, farben=None) -> None:
        """Tabelle mit Linien und Zellentext (Markup erlaubt). Erste Zeile = Kopf.

        kopf_farbe: Hintergrund der Kopfzeile (Text dann weiss + fett), None = ohne.
        band_farbe: Hintergrund jeder zweiten Datenzeile (Baenderung), None = ohne.
        """
        self._bereit()
        n_sp = len(zeilen[0])
        gesamt = SEITE_B - RAND_L - RAND_R
        breiten = breiten or [gesamt / n_sp] * n_sp
        hoehe = zeilenhoehe * len(zeilen)
        if self.y + hoehe > self._unten():
            self._neue_seite()
        top0 = self.y
        schwarz = (0, 0, 0)
        for i, zeile in enumerate(zeilen):
            top = top0 + i * zeilenhoehe
            kopf = i == 0
            if kopf and kopf_farbe:
                self._rechteck(RAND_L, top, RAND_L + sum(breiten), top + zeilenhoehe, hex_farbe(kopf_farbe))
            elif not kopf and band_farbe and i % 2 == 1:
                self._rechteck(RAND_L, top, RAND_L + sum(breiten), top + zeilenhoehe, hex_farbe(band_farbe))
            x = RAND_L
            for j, zelle in enumerate(zeile):
                grund = (1, 1, 1) if (kopf and kopf_farbe) else schwarz
                xx = x + 4
                for wort in parse_markup(zelle):
                    xx = self._wort(xx, top + (zeilenhoehe - groesse) / 2, wort, groesse, kopf, farben, grund)
                    xx += breite(" ", groesse, kopf)
                x += breiten[j]
            self._linie(RAND_L, RAND_L + sum(breiten), top, schwarz, 0.5)
        self._linie(RAND_L, RAND_L + sum(breiten), top0 + hoehe, schwarz, 0.5)
        x = RAND_L
        for b in breiten:
            self._vlinie(x, top0, top0 + hoehe, schwarz)
            x += b
        self._vlinie(x, top0, top0 + hoehe, schwarz)
        self.y = top0 + hoehe + 12

    # ---- Bild, Scan --------------------------------------------------------
    def _bild_registrieren(self, img) -> int:
        img = img.convert("RGB")
        self._bilder.append((img.size[0], img.size[1], zlib.compress(img.tobytes(), 1)))
        return len(self._bilder) - 1

    def bild(self, img, breite_pt: float = 200.0, x: float = RAND_L) -> None:
        """Setzt ein Pillow-Bild als Image-XObject in den Textfluss."""
        self._bereit()
        hoehe = breite_pt * img.size[1] / img.size[0]
        if self.y + hoehe > self._unten():
            self._neue_seite()
        idx = self._bild_registrieren(img)
        self._s.bilder.append(idx)
        self._op(f"q {breite_pt:.2f} 0 0 {hoehe:.2f} {x:.2f} {SEITE_H - (self.y + hoehe):.2f} cm /Im{idx + 1} Do Q")
        self.y += hoehe + 10

    def scan_seite(self, img) -> None:
        """Eigene Seite, die nur aus einem Bild besteht (kein Text, keine Kopf-/Fusszeile)."""
        self.seitenumbruch()
        idx = self._bild_registrieren(img)
        s = self._s
        s.bilder.append(idx)
        s.nur_bild = True
        s.inhalt.append(f"q {SEITE_B} 0 0 {SEITE_H} 0 0 cm /Im{idx + 1} Do Q\n".encode("latin-1"))
        self._seite_voll = True

    # ---- Fussnoten ------------------------------------------------------------
    def fussnote(self, text: str, marke: str = "") -> bool:
        """Haengt eine Fussnote (8 pt) unten an die aktuelle Seite. False, wenn kein Platz mehr ist."""
        self._bereit()
        s = self._s
        gesamt = SEITE_B - RAND_L - RAND_R
        zeilen: list[str] = []
        aktuell = ""
        for wort in (f"{marke} {text}" if marke else text).split():
            probe = f"{aktuell} {wort}".strip()
            if breite(probe, 8.0) > gesamt and aktuell:
                zeilen.append(aktuell)
                aktuell = wort
            else:
                aktuell = probe
        if aktuell:
            zeilen.append(aktuell)
        zusatz = 10.0 * len(zeilen) + (8.0 if not s.fussnoten else 0.0)
        if self.y > SEITE_H - RAND_U - s.fuss_hoehe - zusatz:
            return False
        s.fussnoten.append(zeilen)
        s.fuss_hoehe += zusatz
        return True

    # ---- Ausgabe -------------------------------------------------------
    def _zeige_kopf_fuss(self, nr: int) -> bool:
        kf = self.kopf_fuss
        if kf is True or kf is False or kf is None:
            return bool(kf)
        return nr in kf

    def bytes(self) -> bytes:
        n = len(self.seiten)
        objekte: list[bytes] = []

        def neu(inhalt: bytes) -> int:
            objekte.append(inhalt)
            return len(objekte)

        katalog = neu(b"")
        seiten_obj = neu(b"")
        f1 = neu(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        f2 = neu(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
        bild_ids = []
        for bw, bh, daten in self._bilder:
            bild_ids.append(neu(b"<< /Type /XObject /Subtype /Image /Width %d /Height %d /ColorSpace /DeviceRGB "
                                b"/BitsPerComponent 8 /Filter /FlateDecode /Length %d >>\nstream\n" % (bw, bh, len(daten))
                                + daten + b"\nendstream"))
        seiten_ids = []
        for nr, s in enumerate(self.seiten, 1):
            extra = b""
            if self._zeige_kopf_fuss(nr) and not s.nur_bild:
                kopf = (f"BT /F1 9 Tf {self._fb(self.kopf_farbe)} {RAND_L} {SEITE_H - 36} Td (".encode("latin-1")
                        + _esc(self.kopf_text) + b") Tj ET\n")
                fuss = (f"BT /F1 9 Tf {self._fb(self.kopf_farbe)} {SEITE_B / 2 - 30} 36 Td (".encode("latin-1")
                        + _esc(self.fuss_vorlage.format(nr=nr, n=n)) + b") Tj ET\n")
                extra = kopf + fuss
            fn = b""
            if s.fussnoten:
                top = SEITE_H - RAND_U - s.fuss_hoehe + 2
                fn += (self._linie_op(RAND_L, RAND_L + 100, top, (0.5, 0.5, 0.5), 0.5) + "\n").encode("latin-1")
                top += 6
                for zeilen in s.fussnoten:
                    for z in zeilen:
                        fn += self._text_bytes(RAND_L, top, z, 8.0, False, (0, 0, 0))
                        top += 10
            strom = (self._fb((0, 0, 0)) + "\n").encode("latin-1") + extra + b"".join(s.inhalt) + fn
            inh = neu(b"<< /Length %d >>\nstream\n" % len(strom) + strom + b"endstream")
            annot_ids = []
            for a in s.annots:
                x0, top, x1, bot = a["box"]
                y1, y0 = SEITE_H - top, SEITE_H - bot
                r, g, b = a["c"]
                d = (f"<< /Type /Annot /Subtype /{a['art']} /Rect [{x0:.2f} {y0:.2f} {x1:.2f} {y1:.2f}] "
                     f"/QuadPoints [{x0:.2f} {y1:.2f} {x1:.2f} {y1:.2f} {x0:.2f} {y0:.2f} {x1:.2f} {y0:.2f}] "
                     f"/C [{r} {g} {b}] /F 4 >>")
                annot_ids.append(neu(d.encode("latin-1")))
            for lk in s.links:
                x0, top, x1, bot = lk["box"]
                y1, y0 = SEITE_H - top, SEITE_H - bot
                d = (f"<< /Type /Annot /Subtype /Link /Rect [{x0:.2f} {y0:.2f} {x1:.2f} {y1:.2f}] /Border [0 0 0] "
                     f"/A << /S /URI /URI ({lk['url']}) >> >>")
                annot_ids.append(neu(d.encode("latin-1")))
            annots = b"/Annots [" + b" ".join(b"%d 0 R" % i for i in annot_ids) + b"]" if annot_ids else b""
            if s.rotate:
                annots += b" /Rotate %d" % s.rotate
            xo = b""
            if s.bilder:
                xo = b" /XObject << " + b" ".join(b"/Im%d %d 0 R" % (i + 1, bild_ids[i]) for i in s.bilder) + b" >>"
            seite = neu(b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %d %d] /Contents %d 0 R "
                        b"/Resources << /Font << /F1 %d 0 R /F2 %d 0 R >>%s >> %s >>"
                        % (seiten_obj, SEITE_B, SEITE_H, inh, f1, f2, xo, annots))
            seiten_ids.append(seite)
        objekte[katalog - 1] = b"<< /Type /Catalog /Pages %d 0 R >>" % seiten_obj
        objekte[seiten_obj - 1] = (b"<< /Type /Pages /Kids [" + b" ".join(b"%d 0 R" % i for i in seiten_ids)
                                   + b"] /Count %d >>" % len(seiten_ids))
        out = bytearray(b"%PDF-1.4\n")
        offsets = []
        for i, o in enumerate(objekte, 1):
            offsets.append(len(out))
            out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
        xref = len(out)
        out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objekte) + 1)
        for off in offsets:
            out += b"%010d 00000 n \n" % off
        out += b"trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objekte) + 1, katalog, xref)
        return bytes(out)

    def speichern(self, pfad) -> None:
        with open(pfad, "wb") as f:
            f.write(self.bytes())


def baue_pdf(pfad, elemente: list[tuple[str, str]], variante: str = "farbe") -> None:
    """elemente: [('h', markup), ('p', markup), ...]."""
    d = PdfDokument(variante=variante)
    for art, markup in elemente:
        if art == "h":
            d.ueberschrift(markup)
        else:
            d.block(markup)
    d.speichern(pfad)
