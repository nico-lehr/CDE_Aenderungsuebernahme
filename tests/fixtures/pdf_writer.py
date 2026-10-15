"""Winziger PDF-Schreiber NUR fuer Tests (erfundene Beispieldokumente, Regel 16).

Erzeugt PDFs mit Standard-Schrift Helvetica, farbigem Text, Hintergrund-
Markierungen, Durchstreichungen und PDF-Kommentaren (Annotationen). Keine
zusaetzliche Bibliothek noetig -- die PDF-Struktur wird von Hand geschrieben.

Markup fuer Absaetze:
    {+neuer Text+}   hinzugekommen (blau)
    {-alter Text-}   entfaellt (rot + durchgestrichen)
    Rest             unveraendert
Beispiel: "Die Kapazitaet betraegt {-100 kWh-} {+120 kWh+} je Rack."
"""
from __future__ import annotations

import re

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
}


def breite(text: str, groesse: float, fett: bool = False) -> float:
    faktor = 1.05 if fett else 1.0
    return sum(_BREITEN.get(c, 556) for c in text) * groesse / 1000.0 * faktor


def _esc(text: str) -> bytes:
    return (text.encode("cp1252", errors="replace")
            .replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)"))


def parse_markup(markup: str) -> list[list[tuple[str, str]]]:
    """Markup -> Liste von Woertern; ein Wort ist eine Liste von (Text, Stil), Stil: k|n|g."""
    teile: list[tuple[str, str]] = []
    pos = 0
    for m in re.finditer(r"\{\+(.*?)\+\}|\{-(.*?)-\}", markup, flags=re.S):
        if m.start() > pos:
            teile.append((markup[pos:m.start()], "k"))
        if m.group(1) is not None:
            teile.append((m.group(1), "n"))
        else:
            teile.append((m.group(2), "g"))
        pos = m.end()
    if pos < len(markup):
        teile.append((markup[pos:], "k"))
    woerter: list[list[tuple[str, str]]] = []
    aktuell: list[tuple[str, str]] = []
    for text, stil in teile:
        for stueck in re.split(r"(\s+)", text):
            if not stueck:
                continue
            if stueck.isspace():
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
    s = markup
    if neu:
        s = re.sub(r"\{-.*?-\}", "", s, flags=re.S)
        s = re.sub(r"\{\+(.*?)\+\}", r"\1", s, flags=re.S)
    else:
        s = re.sub(r"\{\+.*?\+\}", "", s, flags=re.S)
        s = re.sub(r"\{-(.*?)-\}", r"\1", s, flags=re.S)
    return re.sub(r"\s+", " ", s).strip()


class _Seite:
    def __init__(self) -> None:
        self.inhalt: list[bytes] = []
        self.annots: list[dict] = []


class PdfDokument:
    """variante: farbe | hintergrund | strich | annot."""

    def __init__(self, variante: str = "farbe", kopf_fuss: bool = True) -> None:
        self.variante = variante
        self.kopf_fuss = kopf_fuss
        self.seiten: list[_Seite] = []
        self.y = RAND_O
        self._neue_seite()

    # ---- Zeichnen ------------------------------------------------------
    def _neue_seite(self) -> None:
        self.seiten.append(_Seite())
        self.y = RAND_O

    @property
    def _s(self) -> _Seite:
        return self.seiten[-1]

    def _op(self, s: str) -> None:
        self._s.inhalt.append(s.encode("latin-1") + b"\n")

    def _text(self, x: float, top: float, text: str, groesse: float, fett: bool, farbe) -> None:
        basis = SEITE_H - (top + groesse * 0.8)
        r, g, b = farbe
        self._s.inhalt.append(
            f"BT /{'F2' if fett else 'F1'} {groesse} Tf {r} {g} {b} rg {x:.2f} {basis:.2f} Td (".encode("latin-1")
            + _esc(text) + b") Tj ET\n")

    def _rechteck(self, x0, top, x1, bottom, farbe) -> None:
        r, g, b = farbe
        self._op(f"{r} {g} {b} rg {x0:.2f} {SEITE_H - bottom:.2f} {x1 - x0:.2f} {bottom - top:.2f} re f")

    def _linie(self, x0, x1, top, farbe, staerke=0.8) -> None:
        r, g, b = farbe
        y = SEITE_H - top
        self._op(f"{r} {g} {b} RG {staerke} w {x0:.2f} {y:.2f} m {x1:.2f} {y:.2f} l S")

    def _wort(self, x: float, top: float, segmente, groesse: float, fett: bool) -> float:
        """Zeichnet ein Wort (evtl. mehrere Stil-Segmente). Gibt die Endposition zurueck."""
        v = self.variante
        for text, stil in segmente:
            w = breite(text, groesse, fett)
            farbe = (0, 0, 0)
            strich = None
            if stil == "n":
                if v == "hintergrund":
                    self._rechteck(x, top - 1, x + w, top + groesse + 1, (0.72, 0.86, 1.0))
                elif v in ("farbe", "strich"):
                    farbe = (0, 0, 1)
                elif v == "annot":
                    self._s.annots.append(dict(art="Highlight", c=(0, 0, 1), box=(x, top, x + w, top + groesse)))
            elif stil == "g":
                if v == "farbe":
                    farbe, strich = (1, 0, 0), (1, 0, 0)
                elif v == "hintergrund":
                    self._rechteck(x, top - 1, x + w, top + groesse + 1, (1.0, 0.8, 0.8))
                    strich = (0, 0, 0)
                elif v == "strich":
                    strich = (0, 0, 0)
                elif v == "annot":
                    self._s.annots.append(dict(art="StrikeOut", c=(1, 0, 0), box=(x, top, x + w, top + groesse)))
            self._text(x, top, text, groesse, fett, farbe)
            if strich:
                self._linie(x - 0.5, x + w + 0.5, top + groesse * 0.55, strich)
            x += w
        return x

    def block(self, markup: str, fett: bool = False, groesse: float = 11.0,
              zeilenabstand: float = 15.0, abstand_danach: float = 9.0, spalte_x: float = RAND_L) -> None:
        woerter = parse_markup(markup)
        max_x = SEITE_B - RAND_R
        x = spalte_x
        platz = breite(" ", groesse, fett)
        if self.y + zeilenabstand > SEITE_H - RAND_U:
            self._neue_seite()
        for wort in woerter:
            text = "".join(t for t, _ in wort)
            w = breite(text, groesse, fett)
            if x > spalte_x and x + w > max_x:
                self.y += zeilenabstand
                x = spalte_x
                if self.y + zeilenabstand > SEITE_H - RAND_U:
                    self._neue_seite()
            x = self._wort(x, self.y, wort, groesse, fett)
            x += platz
        self.y += zeilenabstand + abstand_danach

    def ueberschrift(self, markup: str) -> None:
        self.y += 6
        self.block(markup, fett=True, groesse=14, zeilenabstand=18, abstand_danach=8)

    # ---- Ausgabe -------------------------------------------------------
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
        seiten_ids = []
        for nr, s in enumerate(self.seiten, 1):
            extra = b""
            if self.kopf_fuss:
                kopf = f"BT /F1 9 Tf 0.4 0.4 0.4 rg {RAND_L} {SEITE_H - 36} Td (Beispieldokument - erfunden) Tj ET\n"
                fuss = (f"BT /F1 9 Tf 0.4 0.4 0.4 rg {SEITE_B / 2 - 30} 36 Td (Seite {nr} von {n}) Tj ET\n")
                extra = kopf.encode("latin-1") + fuss.encode("latin-1")
            strom = b"0 0 0 rg\n" + extra + b"".join(s.inhalt)
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
            annots = b"/Annots [" + b" ".join(b"%d 0 R" % i for i in annot_ids) + b"]" if annot_ids else b""
            seite = neu(b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %d %d] /Contents %d 0 R "
                        b"/Resources << /Font << /F1 %d 0 R /F2 %d 0 R >> >> %s >>"
                        % (seiten_obj, SEITE_B, SEITE_H, inh, f1, f2, annots))
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
