"""Zeichenmodul fuer Word-PDFs mit Aenderungsverfolgung ("Markup"-Ansicht) -- NUR fuer erfundene Testdaten (Regel 16).

Generator v3 (`word_markup.py`) beschreibt WAS auf einer Seite steht; dieses Modul zeichnet es so, wie Word ein Dokument
mit Aenderungsverfolgung als PDF ausgibt (Formen vom Orchestrator an echten Beispielen gemessen, Inhalte erfunden):

  * Seite 612 x 792 pt, Fliesstext Times (Base-14), Sprechblasen Helvetica (statt Tahoma)
  * neu       = Text in Autorfarbe + gefuelltes Rechteck gleicher Farbe knapp unter der Grundlinie (~90 % der Zeichenhoehe, 0,36 pt)
  * entfaellt = Text in DERSELBEN Farbe + gefuelltes Rechteck auf ~55 % der Zeichenhoehe (0,46 pt)
  * Listennummer geaendert: gruen (#498205), alte Nummer gestrichen + neue unterstrichen direkt dahinter ("c.a.")
  * gelbe Hinterlegung hinter Platzhaltern, schwarze Aenderungsbalken links (x = 27,3)
  * graue Markup-Leiste rechts (Grau 0,949) mit Sprechblasen (abgerundete Rechtecke aus Bezier-Kurven) und gestrichelten
    Verbindungslinien (dash [0.14 0.14], 0,14 pt) vom Ankerpunkt im Text bis x = 413, dann schraeg zur Sprechblase
  * keine PDF-Annotationen; Schriften nicht eingebettet (Base-14, WinAnsi), Inhalte komprimiert (FlateDecode)

Zeichenbreiten kommen aus den Standard-Metriken von pdfminer (dieselben, mit denen pdfplumber spaeter misst), damit
Zeilenumbruch, Unter- und Durchstreichung genau unter den Zeichen liegen. Dieses Modul importiert nichts aus src/.
"""
from __future__ import annotations

import zlib
from dataclasses import dataclass, field

from pdfminer.fontmetrics import FONT_METRICS

SEITE_B, SEITE_H = 612.0, 792.0
SCHRIFTEN = {"F1": "Times-Roman", "F2": "Times-Bold", "F3": "Helvetica", "F4": "Helvetica-Bold"}
_KENNUNG = {name: fid for fid, name in SCHRIFTEN.items()}

# Farben (RGB 0..1)
WORD_ROT = (0.82, 0.204, 0.22)            # #d13438
WORD_BLAU = (0.0, 0.4706, 0.8314)         # #0078d4 (zweiter Autor)
WORD_GRUEN = (0.2863, 0.5098, 0.0196)     # #498205 (geaenderte Listennummer)
SCHWARZ = (0.0, 0.0, 0.0)
WEISS = (1.0, 1.0, 1.0)
GELB = (1.0, 1.0, 0.0)
LEISTE_GRAU = (0.949, 0.949, 0.949)       # #f2f2f2
AUTORFARBE = {"rot": WORD_ROT, "blau": WORD_BLAU, "gruen": WORD_GRUEN}

# Lage der Linien (Anteil der Zeichenhoehe von oben; Zeichenkasten wie pdfminer: Oberkante = Grundlinie - (1 + Unterlaenge))
UNTER_ANTEIL, UNTER_DICKE = 0.90, 0.36
STRICH_ANTEIL, STRICH_DICKE = 0.55, 0.46
GELB_HOEHE = 9.6

# Markup-Leiste, Sprechblasen, Verbindungslinien, Aenderungsbalken
LEISTE_X0, LEISTE_X1, LEISTE_OBEN, LEISTE_UNTEN = 415.0, 612.0, 95.0, 696.0
BLASE_X0, BLASE_X1 = 433.0, 610.0
BLASE_SCHRIFT = 6.1
BLASE_ZEILE = 11.0
LINIE_BIS_X = 413.0
LINIE_STAERKE = 0.14
BALKEN_X, BALKEN_B = 27.3, 0.53


def breite(text: str, schrift: str, groesse: float) -> float:
    """Breite eines Textes in pt (Standard-Metrik; unbekannte Zeichen wie pdfminer mit 500/1000)."""
    w = FONT_METRICS[schrift][1]
    return sum(w.get(c, 500) for c in text) * groesse / 1000.0


def unterlaenge(schrift: str) -> float:
    """Unterlaenge (positiv, Anteil der Schriftgroesse) wie in pdfminer: der Zeichenkasten reicht so weit unter die Grundlinie."""
    return -float(FONT_METRICS[schrift][0]["Descent"]) / 1000.0


def zeichen_oben(grundlinie: float, schrift: str, groesse: float) -> float:
    """Oberkante des Zeichenkastens (Seitenkoordinate von oben), so wie pdfplumber sie ausgibt."""
    return grundlinie - (1.0 - unterlaenge(schrift)) * groesse


def _esc(text: str) -> bytes:
    return (text.encode("cp1252", errors="replace")
            .replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)"))


def _rg(f: tuple[float, float, float]) -> str:
    return f"{f[0]:.4g} {f[1]:.4g} {f[2]:.4g}"


@dataclass
class Seite:
    """Inhalt einer Seite in Ebenen (Zeichenreihenfolge: grund -> text -> linien -> rand)."""
    grund: list[bytes] = field(default_factory=list)    # Leiste, Hinterlegungen, Tabellenflaechen
    text: list[bytes] = field(default_factory=list)
    linien: list[bytes] = field(default_factory=list)   # Unter-/Durchstreichung, Balken, Tabellenrahmen
    rand: list[bytes] = field(default_factory=list)     # Sprechblasen (Form, dann Text) und Verbindungslinien
    blasen: list[tuple[float, float]] = field(default_factory=list)   # belegte Hoehen der Sprechblasen (oben, unten)


class WordPdf:
    """Sehr kleiner PDF-Schreiber mit den Grundformen eines Word-Exports (Text, Rechtecke, Kurven, gestrichelte Linien)."""

    def __init__(self) -> None:
        self.seiten: list[Seite] = []

    # ---- Seiten -----------------------------------------------------------------------
    def neue_seite(self) -> int:
        self.seiten.append(Seite())
        return len(self.seiten)

    def _s(self, nr: int) -> Seite:
        return self.seiten[nr - 1]

    @staticmethod
    def _y(top: float) -> float:
        return SEITE_H - top

    # ---- Grundformen ---------------------------------------------------------------------
    def _op(self, nr: int, ebene: str, op: str | bytes) -> None:
        getattr(self._s(nr), ebene).append(op if isinstance(op, bytes) else op.encode("latin-1") + b"\n")

    def text(self, nr: int, x: float, grundlinie: float, text: str, schrift: str = "Times-Roman", groesse: float = 8.4,
             farbe: tuple[float, float, float] = SCHWARZ, ebene: str = "text") -> float:
        """Setzt Text mit der Grundlinie `grundlinie` (von oben). Gibt die Endposition x zurueck."""
        if text:
            self._op(nr, ebene, f"BT /{_KENNUNG[schrift]} {groesse:.4g} Tf {_rg(farbe)} rg {x:.3f} {self._y(grundlinie):.3f} "
                                f"Td (".encode("latin-1") + _esc(text) + b") Tj ET\n")
        return x + breite(text, schrift, groesse)

    def rechteck(self, nr: int, x0: float, oben: float, x1: float, unten: float, farbe: tuple[float, float, float],
                 ebene: str = "linien") -> None:
        """Gefuelltes Rechteck (so zeichnet Word Unter-/Durchstreichung, Balken, Hinterlegung und Tabellenrahmen)."""
        self._op(nr, ebene, f"{_rg(farbe)} rg {x0:.3f} {self._y(unten):.3f} {x1 - x0:.3f} {unten - oben:.3f} re f")

    def rahmen(self, nr: int, x0: float, oben: float, x1: float, unten: float, staerke: float = 0.5,
               farbe: tuple[float, float, float] = SCHWARZ) -> None:
        """Rahmen aus vier duennen gefuellten Rechtecken (wie Word-Tabellenlinien)."""
        h = staerke / 2
        self.rechteck(nr, x0 - h, oben - h, x1 + h, oben + h, farbe)
        self.rechteck(nr, x0 - h, unten - h, x1 + h, unten + h, farbe)
        self.rechteck(nr, x0 - h, oben - h, x0 + h, unten + h, farbe)
        self.rechteck(nr, x1 - h, oben - h, x1 + h, unten + h, farbe)

    def unterstreichung(self, nr: int, x0: float, x1: float, grundlinie: float, schrift: str, groesse: float,
                        farbe: tuple[float, float, float]) -> None:
        mitte = zeichen_oben(grundlinie, schrift, groesse) + UNTER_ANTEIL * groesse
        self.rechteck(nr, x0, mitte - UNTER_DICKE / 2, x1, mitte + UNTER_DICKE / 2, farbe)

    def durchstreichung(self, nr: int, x0: float, x1: float, grundlinie: float, schrift: str, groesse: float,
                        farbe: tuple[float, float, float]) -> None:
        mitte = zeichen_oben(grundlinie, schrift, groesse) + STRICH_ANTEIL * groesse
        self.rechteck(nr, x0, mitte - STRICH_DICKE / 2, x1, mitte + STRICH_DICKE / 2, farbe)

    def hinterlegung(self, nr: int, x0: float, x1: float, grundlinie: float, schrift: str, groesse: float) -> None:
        """Gelbe Hinterlegung (Platzhalter): Rechteck mit Hoehe GELB_HOEHE um den Zeichenkasten."""
        oben = zeichen_oben(grundlinie, schrift, groesse)
        mitte = oben + groesse / 2
        self.rechteck(nr, x0, mitte - GELB_HOEHE / 2, x1, mitte + GELB_HOEHE / 2, GELB, ebene="grund")

    def balken(self, nr: int, oben: float, unten: float) -> None:
        """Aenderungsbalken am linken Rand."""
        self.rechteck(nr, BALKEN_X, oben, BALKEN_X + BALKEN_B, unten, SCHWARZ)

    def leiste(self, nr: int) -> None:
        """Graue Markup-Leiste rechts."""
        self.rechteck(nr, LEISTE_X0, LEISTE_OBEN, LEISTE_X1, LEISTE_UNTEN, LEISTE_GRAU, ebene="grund")

    # ---- Sprechblasen ------------------------------------------------------------------------
    def _rund_rechteck(self, x0: float, oben: float, x1: float, unten: float, r: float = 2.5) -> str:
        """Pfad eines abgerundeten Rechtecks aus Bezier-Kurven (PDF-Koordinaten)."""
        k = 0.5523 * r
        ya, yb = self._y(oben), self._y(unten)          # ya oben (gross), yb unten (klein)
        p = [f"{x0 + r:.3f} {ya:.3f} m", f"{x1 - r:.3f} {ya:.3f} l",
             f"{x1 - r + k:.3f} {ya:.3f} {x1:.3f} {ya - r + k:.3f} {x1:.3f} {ya - r:.3f} c",
             f"{x1:.3f} {yb + r:.3f} l",
             f"{x1:.3f} {yb + r - k:.3f} {x1 - r + k:.3f} {yb:.3f} {x1 - r:.3f} {yb:.3f} c",
             f"{x0 + r:.3f} {yb:.3f} l",
             f"{x0 + r - k:.3f} {yb:.3f} {x0:.3f} {yb + r - k:.3f} {x0:.3f} {yb + r:.3f} c",
             f"{x0:.3f} {ya - r:.3f} l",
             f"{x0:.3f} {ya - r + k:.3f} {x0 + r - k:.3f} {ya:.3f} {x0 + r:.3f} {ya:.3f} c", "h"]
        return " ".join(p)

    def _punkt(self, x: float, grundlinie: float, farbe: tuple[float, float, float]) -> str:
        """Kleiner gefuellter Ankerpunkt (Ellipse ~2 x 3 pt aus vier Kurven)."""
        rx, ry = 1.0, 1.5
        cx, cy = x, self._y(grundlinie) + ry * 0.4
        kx, ky = 0.5523 * rx, 0.5523 * ry
        return (f"{_rg(farbe)} rg {cx + rx:.3f} {cy:.3f} m "
                f"{cx + rx:.3f} {cy + ky:.3f} {cx + kx:.3f} {cy + ry:.3f} {cx:.3f} {cy + ry:.3f} c "
                f"{cx - kx:.3f} {cy + ry:.3f} {cx - rx:.3f} {cy + ky:.3f} {cx - rx:.3f} {cy:.3f} c "
                f"{cx - rx:.3f} {cy - ky:.3f} {cx - kx:.3f} {cy - ry:.3f} {cx:.3f} {cy - ry:.3f} c "
                f"{cx + kx:.3f} {cy - ry:.3f} {cx + rx:.3f} {cy - ky:.3f} {cx + rx:.3f} {cy:.3f} c f")

    @staticmethod
    def blasen_zeilen(teile: list[tuple[str, bool]]) -> list[list[tuple[str, bool]]]:
        """Bricht den Sprechblasentext [(Text, fett)] in Zeilen der Blasenbreite um."""
        max_b = BLASE_X1 - BLASE_X0 - 6.0
        zeilen: list[list[tuple[str, bool]]] = [[]]
        x = 0.0
        for text, fett in teile:
            schrift = "Helvetica-Bold" if fett else "Helvetica"
            for i, wort in enumerate(text.split(" ")):
                stueck = (" " if (i > 0 or (x > 0 and zeilen[-1])) else "") + wort
                w = breite(stueck, schrift, BLASE_SCHRIFT)
                if x > 0 and x + w > max_b:
                    zeilen.append([])
                    x = 0.0
                    stueck = wort
                    w = breite(stueck, schrift, BLASE_SCHRIFT)
                zeilen[-1].append((stueck, fett))
                x += w
        return [z for z in zeilen if z]

    def sprechblase(self, nr: int, anker_x: float, anker_grundlinie: float, teile: list[tuple[str, bool]],
                    farbe: tuple[float, float, float] = WORD_ROT) -> tuple[float, float] | None:
        """Sprechblase in der Leiste mit gestrichelter Verbindungslinie zum Ankerpunkt. Liegt sie unter einer schon gesetzten
        Blase, rutscht sie nach unten; passt sie nicht mehr in die Leiste, entfaellt sie (None)."""
        s = self._s(nr)
        zeilen = self.blasen_zeilen(teile)
        hoehe = BLASE_ZEILE * len(zeilen)
        oben = max(LEISTE_OBEN + 2.0, anker_grundlinie - hoehe / 2)
        for a, b in sorted(s.blasen):
            if oben < b + 2.0 and oben + hoehe > a - 2.0:
                oben = b + 2.0
        if oben + hoehe > LEISTE_UNTEN - 2.0:
            return None
        unten = oben + hoehe
        s.blasen.append((oben, unten))
        self._op(nr, "rand", f"{_rg(WEISS)} rg {_rg(farbe)} RG 0.5 w " + self._rund_rechteck(BLASE_X0, oben, BLASE_X1, unten) + " B")
        for i, zeile in enumerate(zeilen):
            x = BLASE_X0 + 3.0
            gl = oben + BLASE_ZEILE * i + 7.6
            for text, fett in zeile:
                schrift = "Helvetica-Bold" if fett else "Helvetica"
                x = self.text(nr, x, gl, text, schrift, BLASE_SCHRIFT, SCHWARZ, ebene="rand")
        mitte = (oben + unten) / 2
        y_a, y_m = self._y(anker_grundlinie), self._y(mitte)
        self._op(nr, "rand", f"q [{LINIE_STAERKE} {LINIE_STAERKE}] 0 d {LINIE_STAERKE} w {_rg(farbe)} RG "
                             f"{anker_x:.3f} {y_a:.3f} m {LINIE_BIS_X:.3f} {y_a:.3f} l {BLASE_X0:.3f} {y_m:.3f} l S Q")
        self._op(nr, "rand", self._punkt(anker_x, anker_grundlinie, farbe))
        return oben, unten

    # ---- Ausgabe --------------------------------------------------------------------------------
    def bytes(self) -> bytes:
        objekte: list[bytes] = []

        def neu(inhalt: bytes) -> int:
            objekte.append(inhalt)
            return len(objekte)

        katalog = neu(b"")
        seiten_obj = neu(b"")
        fonts = {fid: neu(b"<< /Type /Font /Subtype /Type1 /BaseFont /%s /Encoding /WinAnsiEncoding >>" % name.encode())
                 for fid, name in SCHRIFTEN.items()}
        font_res = b" ".join(b"/%s %d 0 R" % (fid.encode(), oid) for fid, oid in fonts.items())
        kinder = []
        for s in self.seiten:
            strom = b"".join(s.grund) + b"".join(s.text) + b"".join(s.linien) + b"".join(s.rand)
            gepackt = zlib.compress(strom, 6)
            inh = neu(b"<< /Length %d /Filter /FlateDecode >>\nstream\n" % len(gepackt) + gepackt + b"\nendstream")
            kinder.append(neu(b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %d %d] /Contents %d 0 R "
                              b"/Resources << /Font << %s >> >> >>" % (seiten_obj, SEITE_B, SEITE_H, inh, font_res)))
        objekte[katalog - 1] = b"<< /Type /Catalog /Pages %d 0 R >>" % seiten_obj
        objekte[seiten_obj - 1] = (b"<< /Type /Pages /Kids [" + b" ".join(b"%d 0 R" % i for i in kinder)
                                   + b"] /Count %d >>" % len(kinder))
        out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
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
