"""Bauhilfe fuer Word-artige Test-PDFs: Aenderungsverfolgung in der Ansicht "Markup" (ADR 0006). Nur erfundene Daten (Regel 16).

Nachgebaut sind nur die FORMEN, die ein deutsches Word beim PDF-Export zeichnet (vom Orchestrator gemessen):
  * Seite 612 x 792 pt, Fliesstext Times 8,4 pt, Zeilenabstand 10 pt, Einzugsebenen x0 = 55, 82, 109, 137, 164, 191, 218
    (Listennummer auf Ebene n, Text ab Ebene n+1 = haengender Einzug), rechter Textrand ~395 pt.
  * Neu = Text Word-Rot (#d13438) + Unterstreichung als gefuelltes Rechteck gleicher Farbe (0,36 pt, ~90 % der Zeichenhoehe);
    entfaellt = Text in DERSELBEN Farbe + Durchstreichung als gefuelltes Rechteck (0,46 pt, ~55 %). Zweiter Autor: Blau.
    Geaenderte Listennummern gruen (#498205): alte Nummer gestrichen + neue unterstrichen direkt dahinter ("c.a.").
  * Graue Markup-Leiste (#f2f2f2, x 415-612, y 95-696) mit Sprechblasen (abgerundete Rechtecke aus Bezier-Kurven, Kontur
    Word-Rot, weiss gefuellt, ~11 pt hoch, Text 6,1 pt) und gestrichelten Verbindungslinien (dash [0.14 0.14], 0,14 pt,
    waagrecht auf der Grundlinie bis x 413, dann schraeg bis x 433) mit kleinem gefuelltem Ankerpunkt (~2 x 3 pt).
  * Gelbe Hinterlegung (9,6 pt hoch) hinter Platzhaltern, schwarzer Aenderungsbalken bei x 27,3 (0,53 pt breit).
  * Kopfzeile (y ~124-145) und Fusszeile (y ~630-662), auf jeder Seite gleich bis auf die Seitenzahl, mit eigenen Aenderungen.

Eine Zeile ist eine Liste von Stuecken (text, art) oder (text, art, farbe). Zwischen Stuecken steht KEIN Leerzeichen.
  art "k" schwarz | "n" neu (farbig + unterstrichen) | "d" entfaellt (farbig + durchgestrichen) | "f" farbig ohne Linie |
      "u" schwarz unterstrichen | "nd" farbig, unter- UND durchgestrichen (Widerspruch) | "nu" farbig, schwarz unterstrichen
  farbe: RGB (0..1), Standard WORT_ROT.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from pdfminer.fontmetrics import FONT_METRICS

SEITE_B, SEITE_H = 612.0, 792.0
WORT_ROT = (0.82, 0.204, 0.22)        # #d13438
WORT_GRUEN = (0.286, 0.51, 0.02)      # #498205 (Listennummern)
WORT_BLAU = (0.0, 0.47, 0.831)        # #0078d4 (zweiter Autor)
GELB = (1.0, 1.0, 0.0)
GRAU_LEISTE = 0.949
GROESSE = 8.4
ZEILE = 10.0
EBENEN = (55.0, 82.0, 109.0, 137.0, 164.0, 191.0, 218.0)
RECHTS = 395.0
LEISTE = (415.0, 95.0, SEITE_B, 696.0)
BLASE_X0, BLASE_X1 = 433.0, 610.0
TEXT_OBEN = 165.0                      # erste Textzeile unter der Kopfzeile
TEXT_UNTEN = 615.0                     # letzte Textzeile ueber der Fusszeile

_SCHRIFT = {"F1": "Times-Roman", "F2": "Times-Bold", "F3": "Helvetica", "F4": "Helvetica-Bold",
            "F5": "Times-Roman"}       # F5: Times-Metrik mit eigener Kodierung fuer Zeichen ausserhalb von WinAnsi
# Zeichen, die Word setzt, die WinAnsi aber nicht kennt: Code in F5, Glyphenname, Breite (1/1000 em)
_SONDER = {"‑": (127, "uni2011", 333),          # geschuetzter Bindestrich
           " ": (129, "uni202F", 200)}          # schmales geschuetztes Leerzeichen


def breite(text: str, groesse: float = GROESSE, schrift: str = "F1") -> float:
    """Breite wie pdfminer sie aus den Standard-Schriftmetriken rechnet (damit die Stuecke genau aneinander passen)."""
    w = FONT_METRICS[_SCHRIFT[schrift]][1]
    return sum(w.get(c, _SONDER[c][2] if c in _SONDER else 500) for c in text) * groesse / 1000.0


def _unterlaenge(schrift: str) -> float:
    return -FONT_METRICS[_SCHRIFT[schrift]][0]["Descent"] / 1000.0


def _esc(text: str, schrift: str = "F1") -> str:
    if schrift == "F5":
        roh = b"".join(bytes([_SONDER[c][0]]) if c in _SONDER else c.encode("cp1252") for c in text)
    else:
        roh = text.encode("cp1252")
    roh = roh.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)")
    return roh.decode("latin-1")


def _f5_schrift(deskriptor: int) -> bytes:
    """Font-Dictionary fuer F5: Times-Breiten fuer WinAnsi + Sonderzeichen ueber /Differences (pdfminer liest dann
    /Widths und die Unterlaenge aus dem /FontDescriptor -- gleiche Kaesten wie bei Times-Roman)."""
    w = FONT_METRICS["Times-Roman"][1]
    breiten = []
    for code in range(32, 256):
        sonder = next((b for c, (k, _, b) in _SONDER.items() if k == code), None)
        if sonder is not None:
            breiten.append(sonder)
            continue
        try:
            breiten.append(int(w.get(bytes([code]).decode("cp1252"), 0)))
        except UnicodeDecodeError:
            breiten.append(0)
    diff = " ".join(f"{k} /{n}" for k, n, _ in _SONDER.values())
    return (b"<< /Type /Font /Subtype /Type1 /BaseFont /TimesWord /FirstChar 32 /LastChar 255 /Widths [%s] "
            b"/Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [%s] >> /FontDescriptor %d 0 R >>"
            % (" ".join(map(str, breiten)).encode(), diff.encode(), deskriptor))


def _rgb(farbe: tuple[float, float, float]) -> str:
    return " ".join(f"{v:.4g}" for v in farbe)


@dataclass
class WordDok:
    """Mehrseitiges Word-artiges PDF. Jede Seite: Hintergrund (Leiste, Flaechen) + Vordergrund (Text, Linien)."""
    seiten: list[dict[str, list[str]]] = field(default_factory=list)
    mit_leiste: bool = True
    mit_kopf_fuss: bool = False
    top: float = TEXT_OBEN

    def __post_init__(self) -> None:
        if not self.seiten:
            self.neue_seite()

    # ------------------------------------------------------------------ Seiten
    @property
    def nr(self) -> int:
        return len(self.seiten)

    def neue_seite(self, leiste: bool | None = None) -> None:
        self.seiten.append({"hinten": [], "vorn": []})
        self.top = TEXT_OBEN
        if self.mit_leiste if leiste is None else leiste:
            x0, t, x1, b = LEISTE
            self._hinten(f"q {GRAU_LEISTE} g {x0:.2f} {SEITE_H - b:.2f} {x1 - x0:.2f} {b - t:.2f} re f Q\n")
        if self.mit_kopf_fuss:
            self.kopf_fuss()

    def _hinten(self, befehl: str) -> None:
        self.seiten[-1]["hinten"].append(befehl)

    def _vorn(self, befehl: str) -> None:
        self.seiten[-1]["vorn"].append(befehl)

    # ------------------------------------------------------------------ Grundformen
    def text(self, x: float, top: float, text: str, groesse: float = GROESSE, farbe: tuple[float, float, float] | None = None,
             schrift: str = "F1", hoch: float = 0.0, basis_groesse: float | None = None) -> float:
        """Text an (x, Oberkante des Zeichenkastens); gibt die Breite zurueck. hoch: so viele pt ueber der Grundlinie der
        Zeile (Schriftgroesse `basis_groesse`), z. B. fuer eine hochgestellte Ziffer."""
        if schrift == "F1" and any(c in _SONDER for c in text):
            schrift = "F5"
        basis = SEITE_H - (top + (basis_groesse or groesse) * (1 - _unterlaenge(schrift))) + hoch
        fb = f"{_rgb(farbe)} rg" if farbe else "0 g"
        self._vorn(f"BT /{schrift} {groesse} Tf {fb} {x:.3f} {basis:.3f} Td ({_esc(text, schrift)}) Tj ET\n")
        return breite(text, groesse, schrift)

    def rechteck(self, x0: float, top: float, x1: float, bot: float, farbe: tuple[float, float, float] | float,
                 hinten: bool = False) -> None:
        fb = f"{farbe} g" if isinstance(farbe, (int, float)) else f"{_rgb(farbe)} rg"
        befehl = f"q {fb} {x0:.3f} {SEITE_H - bot:.3f} {x1 - x0:.3f} {bot - top:.3f} re f Q\n"
        (self._hinten if hinten else self._vorn)(befehl)

    def unterstreichung(self, x0: float, x1: float, top: float, farbe, groesse: float = GROESSE) -> None:
        y = top + 0.9 * groesse
        self.rechteck(x0, y - 0.18, x1, y + 0.18, farbe)

    def durchstreichung(self, x0: float, x1: float, top: float, farbe, groesse: float = GROESSE) -> None:
        y = top + 0.55 * groesse
        self.rechteck(x0, y - 0.23, x1, y + 0.23, farbe)

    def gelb(self, x0: float, x1: float, top: float, groesse: float = GROESSE) -> None:
        """Gelbe Hinterlegung (9,6 pt hoch) hinter dem Text."""
        self.rechteck(x0, top - 0.6, x1, top + 9.0, GELB, hinten=True)

    def kurvenflaeche(self, x0: float, top: float, x1: float, bot: float, farbe: tuple[float, float, float]) -> None:
        """Gefuellte Flaeche als Bezier-Pfad (kein Rechteck-Operator) hinter dem Text -- fuer den Leser keine Markierung."""
        y0, y1 = SEITE_H - bot, SEITE_H - top
        xm = (x0 + x1) / 2
        self._hinten(f"q {_rgb(farbe)} rg {x0:.2f} {y0:.2f} m {x0:.2f} {y1:.2f} {xm:.2f} {y1:.2f} {xm:.2f} {y1:.2f} c "
                     f"{x1:.2f} {y1:.2f} {x1:.2f} {y0:.2f} {x1:.2f} {y0:.2f} c h f Q\n")

    def aenderungsbalken(self, top: float, bot: float) -> None:
        self.rechteck(27.3, top, 27.83, bot, 0.0)

    # ------------------------------------------------------------------ Zeilen und Absaetze
    def zeile(self, x: float, top: float, stuecke: list[tuple], groesse: float = GROESSE,
              fett: bool = False) -> list[tuple[float, float]]:
        """Schreibt die Stuecke hintereinander; gibt je Stueck (x0, x1) zurueck. Ein Stueck darf als viertes Element
        Optionen haben: {"groesse": 5.28, "hoch": 0.6} = kleinere Schrift, so viele pt ueber der Grundlinie (hochgestellt)."""
        out = []
        for st in stuecke:
            text, art = st[0], st[1]
            farbe = (st[2] if len(st) > 2 and st[2] is not None else WORT_ROT)
            opt = st[3] if len(st) > 3 else {}
            schrift = "F2" if fett else "F1"
            textfarbe = None if art in ("k", "u") else farbe
            w = self.text(x, top, text, float(opt.get("groesse", groesse)), textfarbe, schrift,
                          hoch=float(opt.get("hoch", 0.0)), basis_groesse=groesse)
            if art in ("n", "nd"):
                self.unterstreichung(x, x + w, top, farbe, groesse)
            if art in ("u", "nu"):
                self.unterstreichung(x, x + w, top, 0.0, groesse)
            if art in ("d", "nd"):
                self.durchstreichung(x, x + w, top, farbe, groesse)
            out.append((x, x + w))
            x += w
        return out

    def absatz(self, zeilen: list[list[tuple]], ebene: int = 1, abstand: float = 4.0) -> list[float]:
        """Absatz ab Einzugsebene `ebene` (eine Zeile je Eintrag); gibt die Oberkanten der Zeilen zurueck."""
        tops = []
        for z in zeilen:
            self.zeile(EBENEN[ebene], self.top, z)
            tops.append(self.top)
            self.top += ZEILE
        self.top += abstand
        return tops

    def listenpunkt(self, ebene: int, label: list[tuple], zeilen: list[list[tuple]], abstand: float = 4.0) -> list[float]:
        """Listennummer auf Ebene `ebene`, Text ab Ebene `ebene + 1` (haengender Einzug)."""
        self.zeile(EBENEN[ebene], self.top, label)
        return self.absatz(zeilen, ebene + 1, abstand)

    def ueberschrift(self, nummer: list[tuple], titel: list[tuple], abstand: float = 6.0) -> float:
        """Section-Zeile "1.6<Tab>TITEL": Nummer auf Ebene 0, Titel auf Ebene 1."""
        top = self.top
        self.zeile(EBENEN[0], top, nummer)
        self.zeile(EBENEN[1], top, titel)
        self.top += ZEILE + abstand
        return top

    def kapitel(self, stuecke: list[tuple], abstand: float = 6.0) -> float:
        """Kapitelzeile "PART 1 TITEL" auf Ebene 0."""
        top = self.top
        self.zeile(EBENEN[0], top, stuecke)
        self.top += ZEILE + abstand
        return top

    # ------------------------------------------------------------------ Markup-Leiste
    def sprechblase(self, top: float, kopf: str, text: str, anker: tuple[float, float] | None = None,
                    farbe: tuple[float, float, float] = WORT_ROT, hoehe: float = 11.0) -> None:
        """Sprechblase in der Leiste (Oberkante `top`), fetter Kopf ("hat formatiert:") + Text; mit `anker` = (x, Oberkante
        der Textzeile) eine gestrichelte Verbindungslinie vom Ankerpunkt auf der Grundlinie in die Blase."""
        x0, x1, t, b = BLASE_X0, BLASE_X1, top, top + hoehe
        y0, y1 = SEITE_H - b, SEITE_H - t
        r = 2.0
        pfad = (f"{x0 + r:.2f} {y0:.2f} m {x1 - r:.2f} {y0:.2f} l {x1:.2f} {y0:.2f} {x1:.2f} {y0:.2f} {x1:.2f} {y0 + r:.2f} c "
                f"{x1:.2f} {y1 - r:.2f} l {x1:.2f} {y1:.2f} {x1:.2f} {y1:.2f} {x1 - r:.2f} {y1:.2f} c "
                f"{x0 + r:.2f} {y1:.2f} l {x0:.2f} {y1:.2f} {x0:.2f} {y1:.2f} {x0:.2f} {y1 - r:.2f} c "
                f"{x0:.2f} {y0 + r:.2f} l {x0:.2f} {y0:.2f} {x0:.2f} {y0:.2f} {x0 + r:.2f} {y0:.2f} c h")
        self._vorn(f"q 1 g {_rgb(farbe)} RG 0.5 w {pfad} B Q\n")
        x = x0 + 3.0
        x += self.text(x, t + 2.4, kopf, 6.1, None, "F4")
        self.text(x + 1.7, t + 2.4, text, 6.1, None, "F3")
        if anker is not None:
            ax, atop = anker
            self.verbindungslinie(ax, atop, t + hoehe / 2, farbe=farbe)
            self.ankerpunkt(ax, atop, farbe)

    @staticmethod
    def _grundlinie(atop: float) -> float:
        """Grundlinie (PDF-Koordinaten) einer Textzeile mit Oberkante `atop`."""
        return SEITE_H - (atop + GROESSE * (1 - _unterlaenge("F1")))

    def verbindungslinie(self, ax: float, atop: float, ziel_top: float | None = None, bis_x: float = 413.0,
                         farbe: tuple[float, float, float] = WORT_ROT) -> None:
        """Gestrichelte Verbindungslinie: waagrecht auf der Grundlinie der Textzeile (Oberkante `atop`) von x `ax` bis
        `bis_x`, dann (mit `ziel_top`) schraeg bis x 433 auf die Hoehe `ziel_top` in der Leiste."""
        ay = self._grundlinie(atop)
        pfad = f"{ax:.2f} {ay:.2f} m {bis_x:.2f} {ay:.2f} l"
        if ziel_top is not None:
            pfad += f" 433 {SEITE_H - ziel_top:.2f} l"
        self._vorn(f"q {_rgb(farbe)} RG 0.14 w [0.14 0.14] 0 d {pfad} S Q\n")

    def ankerpunkt(self, ax: float, atop: float, farbe: tuple[float, float, float] = WORT_ROT, flach: bool = False) -> None:
        """Ankerpunkt auf der Grundlinie der Textzeile (Oberkante `atop`): kleine gefuellte Ellipse ~2 x 3 pt aus vier Boegen.
        flach=True: nur zwei Boegen mit Endpunkten auf der Grundlinie (pdfplumber meldet dann einen Kasten der Hoehe 0)."""
        ay = self._grundlinie(atop)
        if flach:
            self._vorn(f"q {_rgb(farbe)} rg {ax - 1:.2f} {ay:.2f} m {ax - 1:.2f} {ay + 2:.2f} {ax + 1:.2f} {ay + 2:.2f} "
                       f"{ax + 1:.2f} {ay:.2f} c {ax + 1:.2f} {ay - 1:.2f} {ax - 1:.2f} {ay - 1:.2f} {ax - 1:.2f} {ay:.2f} c f Q\n")
            return
        rx, ry, k = 1.0, 1.5, 0.55
        self._vorn(f"q {_rgb(farbe)} rg {ax - rx:.2f} {ay:.2f} m "
                   f"{ax - rx:.2f} {ay + k * ry:.2f} {ax - k * rx:.2f} {ay + ry:.2f} {ax:.2f} {ay + ry:.2f} c "
                   f"{ax + k * rx:.2f} {ay + ry:.2f} {ax + rx:.2f} {ay + k * ry:.2f} {ax + rx:.2f} {ay:.2f} c "
                   f"{ax + rx:.2f} {ay - k * ry:.2f} {ax + k * rx:.2f} {ay - ry:.2f} {ax:.2f} {ay - ry:.2f} c "
                   f"{ax - k * rx:.2f} {ay - ry:.2f} {ax - rx:.2f} {ay - k * ry:.2f} {ax - rx:.2f} {ay:.2f} c f Q\n")

    def randtext(self, top: float, text: str, farbe: tuple[float, float, float] = WORT_ROT) -> None:
        """Text in Word-Rot (8,4 pt Times) frei in der Leiste (ohne Sprechblasen-Umriss)."""
        self.text(BLASE_X0 + 2, top, text, GROESSE, farbe)

    # ------------------------------------------------------------------ Kopf-/Fusszeile
    def kopf_fuss(self) -> None:
        """Kopfzeile (zwei Zeilen) und Fusszeile mit Seitenzahl; beide mit eigener Aenderung (Platzhalter gestrichen +
        Wert unterstrichen), auf jeder Seite gleich bis auf die Seitenzahl."""
        self.zeile(EBENEN[0], 124.0, [("Projekt Zebrafink ", "k"), ("[Name]", "d"), ("Ostflügel", "n")])
        self.zeile(EBENEN[0], 135.0, [("Abschnitt Pumpenhaus Revision ", "k"), ("[0]", "d"), ("2", "n")])
        self.zeile(EBENEN[0], 640.0, [("Pumpenhaus Leitfaden ", "k"), ("[Datum]", "d"), ("Mai", "n"),
                                      (f" Seite {self.nr}", "k")])

    # ------------------------------------------------------------------ Ausgabe
    def bytes(self) -> bytes:
        objekte: list[bytes] = [b"", b""]

        def neu(inhalt: bytes) -> int:
            objekte.append(inhalt)
            return len(objekte)

        fonts = {k: neu(b"<< /Type /Font /Subtype /Type1 /BaseFont /%s /Encoding /WinAnsiEncoding >>" % v.encode())
                 for k, v in _SCHRIFT.items() if k != "F5"}
        d = FONT_METRICS["Times-Roman"][0]
        deskriptor = neu(b"<< /Type /FontDescriptor /FontName /TimesWord /Flags 34 /FontBBox [-168 -218 1000 898] "
                         b"/ItalicAngle 0 /Ascent %d /Descent %d /CapHeight 662 /StemV 84 >>"
                         % (int(d["Ascent"]), int(d["Descent"])))
        fonts["F5"] = neu(_f5_schrift(deskriptor))
        font_res = " ".join(f"/{k} {n} 0 R" for k, n in fonts.items()).encode()
        ids = []
        for s in self.seiten:
            strom = "".join(s["hinten"] + s["vorn"]).encode("latin-1")
            inh = neu(b"<< /Length %d >>\nstream\n" % len(strom) + strom + b"\nendstream")
            ids.append(neu(b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %d %d] /Contents %d 0 R "
                           b"/Resources << /Font << %s >> >> >>" % (int(SEITE_B), int(SEITE_H), inh, font_res)))
        objekte[0] = b"<< /Type /Catalog /Pages 2 0 R >>"
        objekte[1] = b"<< /Type /Pages /Kids [" + b" ".join(b"%d 0 R" % i for i in ids) + b"] /Count %d >>" % len(ids)
        out = bytearray(b"%PDF-1.4\n")
        offs = []
        for i, o in enumerate(objekte, 1):
            offs.append(len(out))
            out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
        xref = len(out)
        out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objekte) + 1)
        for o in offs:
            out += b"%010d 00000 n \n" % o
        out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objekte) + 1, xref)
        return bytes(out)

    def speichern(self, pfad: Path) -> Path:
        pfad = Path(pfad)
        pfad.parent.mkdir(parents=True, exist_ok=True)
        pfad.write_bytes(self.bytes())
        return pfad


# ---------------------------------------------------------------- erfundener Fliesstext
# Eigenes Vokabular: die Testwoerter (Pumpe, Ventil, Becken, Speicher, Zebrafink, Kranich ...) kommen hier NICHT vor,
# damit ein Test sein markiertes Wort eindeutig findet.
FUELL = ("Der Kessel heizt das Wasser im Keller und haelt die Temperatur im Kreislauf ueber den ganzen Tag stabil",
         "Die Leitung fuehrt durch den Schacht bis zum Verteiler und wird dort mit einer Schelle befestigt",
         "Jede Kontrolle wird im Betriebsbuch eingetragen und vom Pruefer mit Datum und Kuerzel bestaetigt",
         "Das Gehaeuse steht auf einem Sockel aus Beton und ist gegen Feuchte und Frost gut geschuetzt")


def voll(stuecke: list[tuple], ebene: int = 1, start: int = 0) -> list[tuple]:
    """Die Stuecke + schwarze Fuellwoerter, bis die Zeile fast den rechten Rand erreicht (wie eine volle Zeile im Absatz)."""
    rest = RECHTS - EBENEN[ebene] - sum(breite(st[0]) for st in stuecke)
    zusatz = ""
    for w in " ".join(FUELL[(start + i) % len(FUELL)] for i in range(4)).split():
        if breite(zusatz + " " + w) <= rest:      # passende Woerter nehmen, bis die Zeile wirklich voll ist
            zusatz += " " + w
        if rest - breite(zusatz) < 6.0:
            break
    return list(stuecke) + ([(zusatz, "k")] if zusatz else [])


def fuellzeilen(n: int, start: int = 0, ebene: int = 1) -> list[list[tuple]]:
    """n Zeilen schwarzer Fliesstext ab Einzugsebene `ebene`, jede Zeile fast bis zum rechten Rand (die letzte auch)."""
    breite_max = RECHTS - EBENEN[ebene]
    woerter = " ".join(FUELL[(start + i) % len(FUELL)] for i in range(n * 3)).split()
    zeilen: list[list[tuple]] = []
    akt = ""
    for w in woerter:
        probe = (akt + " " + w).strip()
        if breite(probe) > breite_max:
            zeilen.append([(akt, "k")])
            akt = w
            if len(zeilen) == n:
                break
        else:
            akt = probe
    while len(zeilen) < n:
        zeilen.append([(akt or "Ende", "k")])
    return zeilen
