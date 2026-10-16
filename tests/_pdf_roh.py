"""Roh-PDF-Bauer fuer die SICHT-Tests (nur erfundene Daten, Regel 16).

Erweitert tests/redteam/_bau_r4.RohDok um Dinge, die der Textleser (pdfplumber) nicht oder anders sieht:
  * Seiten-Eintraege (z. B. /Annots mit einem Link-Kommentar)
  * Bilder (XObject, einfarbig)
  * Schattierungen (`sh` in einem Rechteck-Ausschnitt) -- pdfplumber sieht sie gar nicht, pdfium zeichnet sie
Alle Objekte werden ueber "{OBJ<n>}" (n-tes Zusatzobjekt) referenziert, wie bei RohDok.
"""
from __future__ import annotations

from dataclasses import dataclass

from tests.fixtures.pdf_writer import SEITE_B, SEITE_H
from tests.redteam._bau_r4 import RohDok


@dataclass
class SichtDok(RohDok):
    seiten_eintraege: str = ""          # zusaetzliche Eintraege im Seiten-Dictionary der ERSTEN Seite (z. B. "/Annots [...]")

    def objekt(self, inhalt: bytes) -> str:
        """Fuegt ein Zusatzobjekt an; gibt den Platzhalter fuer Verweise zurueck ("{OBJ<n>} 0 R")."""
        self.objekte.append(inhalt)
        return "{OBJ%d} 0 R" % (len(self.objekte) - 1)

    def bild(self, x0: float, top: float, x1: float, bot: float, rgb: tuple[int, int, int], name: str = "Im1") -> None:
        """Einfarbiges Bild (2x2 Pixel, gestreckt) an die Stelle setzen."""
        daten = bytes(rgb) * 4
        ref = self.objekt(b"<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceRGB "
                          b"/BitsPerComponent 8 /Length %d >>\nstream\n" % len(daten) + daten + b"\nendstream")
        self._xobjekt(name, ref)
        self.inhalt.append(f"q {x1 - x0:.2f} 0 0 {bot - top:.2f} {x0:.2f} {SEITE_H - bot:.2f} cm /{name} Do Q\n")

    def schattierung(self, x0: float, top: float, x1: float, bot: float, rgb: tuple[float, float, float],
                     name: str = "Sh1", vorn: bool = False) -> None:
        """Einfarbige Schattierung in einem Rechteck-Ausschnitt (fuer pdfplumber unsichtbar)."""
        c = " ".join(f"{v:g}" for v in rgb)
        self.ressourcen += (f" /Shading << /{name} << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 1 0] "
                            f"/Function << /FunctionType 2 /Domain [0 1] /C0 [{c}] /C1 [{c}] /N 1 >> /Extend [true true] >> >>")
        befehl = f"q {x0:.2f} {SEITE_H - bot:.2f} {x1 - x0:.2f} {bot - top:.2f} re W n /{name} sh Q\n"
        if vorn:
            self.inhalt.append(befehl)
        else:
            self.inhalt.insert(0, befehl)            # hinter dem Text

    def link(self, x0: float, top: float, x1: float, bot: float) -> None:
        ref = self.objekt(b"<< /Type /Annot /Subtype /Link /Rect [%.2f %.2f %.2f %.2f] /Border [0 0 0] "
                          b"/A << /S /URI /URI (https://beispiel.invalid/) >> >>"
                          % (x0, SEITE_H - bot, x1, SEITE_H - top))
        self.seiten_eintraege += f" /Annots [{ref}]"

    def _xobjekt(self, name: str, ref: str) -> None:
        self.ressourcen += f" /XObject << /{name} {ref} >>"

    def bytes(self) -> bytes:   # noqa: C901 - wie RohDok.bytes, plus Seiten-Eintraege
        objekte: list[bytes] = [b"", b""]

        def neu(inhalt: bytes) -> int:
            objekte.append(inhalt)
            return len(objekte)

        f1 = neu(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        f2 = neu(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
        zusatz = [neu(b"") for _ in self.objekte]

        def ersetze(text: str) -> str:
            for i, nr in enumerate(zusatz):
                text = text.replace("{OBJ%d}" % i, str(nr))
            return text

        for i, nr in enumerate(zusatz):
            objekte[nr - 1] = ersetze(self.objekte[i].decode("latin-1")).encode("latin-1")
        res = ersetze(self.ressourcen)
        extra = ersetze(self.seiten_eintraege)
        ids = []
        for k, s in enumerate(self.seiten):
            strom = "".join(s).encode("latin-1")
            inh = neu(b"<< /Length %d >>\nstream\n" % len(strom) + strom + b"\nendstream")
            ids.append(neu(b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %d %d] /Contents %d 0 R "
                           b"/Resources << /Font << /F1 %d 0 R /F2 %d 0 R >> %s >> %s >>"
                           % (SEITE_B, SEITE_H, inh, f1, f2, res.encode("latin-1"),
                              (extra if k == 0 else "").encode("latin-1"))))
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
