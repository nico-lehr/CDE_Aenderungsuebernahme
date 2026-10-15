"""Hilfsfunktionen fuer die Red-Team-Tests des PDF-Lesers (nur erfundene Daten, Regel 16).

Alles lokal in tests/redteam/; es wird nichts in src/ oder tests/fixtures/ veraendert.
`tests/redteam/_bau.py` wird nicht angefasst (nur importiert).

Zwei Wege, ein Test-PDF zu bauen:
  * `PdfDokument` aus tests/fixtures/pdf_writer.py (Markup {+neu+} / {-weg-}, Spalten, Tabellen, Fussnoten ...)
  * `roh_pdf(...)` hier: ein PDF direkt aus Inhaltsstrom-Befehlen (fuer Farbraeume, unsichtbaren Text,
    /Rotate, CropBox, seltene Annotationen, Schrift-Kodierungen). Damit lassen sich auch Faelle bauen,
    die der Fixture-Schreiber nicht kann.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from src.models import ADD, DEL, KEEP, LUECKEN_ARTEN, MOD, PdfBlock, PdfToken
from src.pdf_reader import PdfAnalyse, analysiere_pdf
from src.protokoll import bilanz
from tests.fixtures.pdf_writer import PdfDokument, SEITE_B, SEITE_H

# ---------------------------------------------------------------- Rohes PDF
@dataclass
class RohSeite:
    """Eine Seite eines Roh-PDFs. `inhalt` = Inhaltsstrom (Text), `annots` = Annotations-Woerterbuecher (Text,
    ohne die aeusseren << >>; /Type /Annot wird ergaenzt), `attribute` = zusaetzliche Seiten-Eintraege
    (z. B. "/Rotate 90 /CropBox [20 30 575 810]"), `ressourcen` = zusaetzliche Eintraege in /Resources."""
    inhalt: str = ""
    annots: list[str] = field(default_factory=list)
    attribute: str = ""
    ressourcen: str = ""
    bild: bool = False                # ein kleines Bild als /Im1 anlegen
    schrift: str = ""                 # Ersatz fuer das Schrift-Objekt F1 (z. B. mit Differences)
    mediabox: str = ""                # Ersatz fuer die MediaBox, z. B. "0 100 595 942" (Ursprung nicht bei 0/0)


def zeile(x: float, top: float, text: str, groesse: float = 11.0, farbe: str = "0 0 0 rg", fett: bool = False,
          modus: int | None = None, kodierung: str = "cp1252") -> str:
    """Eine Textzeile als Inhaltsstrom-Befehle. `top` = Abstand vom oberen Seitenrand (wie pdfplumber)."""
    basis = SEITE_H - (top + groesse * 0.8)
    tr = f"{modus} Tr " if modus is not None else ""
    roh = text.encode(kodierung, errors="replace").replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)")
    return (f"BT /{'F2' if fett else 'F1'} {groesse} Tf {farbe} {tr}{x:.2f} {basis:.2f} Td (".encode("latin-1")
            + roh + b") Tj ET\n").decode("latin-1")


def roh_pdf(seiten: list[RohSeite], breite: float = SEITE_B, hoehe: float = SEITE_H) -> bytes:
    """Baut ein gueltiges PDF (Helvetica / Helvetica-Bold, WinAnsi) aus Roh-Inhalten."""
    objekte: list[bytes] = [b"", b""]          # 1 = Katalog, 2 = Seitenbaum (werden unten gefuellt)

    def neu(inhalt: bytes) -> int:
        objekte.append(inhalt)
        return len(objekte)

    f1 = neu(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    f2 = neu(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
    bild_id = neu(b"<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 "
                  b"/Length 12 >>\nstream\n" + bytes([0, 0, 255, 255, 0, 0, 0, 255, 0, 200, 200, 200]) + b"\nendstream")
    ids = []
    for s in seiten:
        font1 = f1
        if s.schrift:
            font1 = neu(s.schrift.encode("latin-1"))
        strom = s.inhalt.encode("latin-1")
        inh = neu(b"<< /Length %d >>\nstream\n" % len(strom) + strom + b"\nendstream")
        annot_ids = [neu(("<< /Type /Annot " + a + " >>").encode("latin-1")) for a in s.annots]
        annots = b"/Annots [" + b" ".join(b"%d 0 R" % i for i in annot_ids) + b"] " if annot_ids else b""
        xo = b" /XObject << /Im1 %d 0 R >>" % bild_id if s.bild else b""
        box = s.mediabox or f"0 0 {breite:g} {hoehe:g}"
        ids.append(neu(b"<< /Type /Page /Parent 2 0 R /MediaBox [%s] /Contents %d 0 R "
                       b"/Resources << /Font << /F1 %d 0 R /F2 %d 0 R >>%s %s >> %s %s >>"
                       % (box.encode("latin-1"), inh, font1, f2, xo, s.ressourcen.encode("latin-1"), annots,
                          s.attribute.encode("latin-1"))))
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


# ---------------------------------------------------------------- Dokument-Fabrik
def dok(**kw: Any) -> PdfDokument:
    """PdfDokument OHNE Kopf-/Fusszeile (sonst taucht die erfundene Kopfzeile als Block auf)."""
    kw.setdefault("kopf_fuss", False)
    return PdfDokument(**kw)


def annot(art: str, x0: float, top: float, x1: float, bot: float, farbe: str = "0 0 1") -> str:
    """Annotation (Text fuer RohSeite.annots) auf einer normalen, nicht gedrehten Seite. Koordinaten wie pdfplumber (top)."""
    y1, y0 = SEITE_H - top, SEITE_H - bot
    return (f"/Subtype /{art} /Rect [{x0} {y0} {x1} {y1}] /QuadPoints [{x0} {y1} {x1} {y1} {x0} {y0} {x1} {y0}] "
            f"/C [{farbe}] /F 4")


def zeile_gedreht90(x: float, top: float, text: str, groesse: float = 11.0, farbe: str = "0 0 0 rg") -> str:
    """Textzeile fuer eine Seite mit /Rotate 90, die im Betrachter AUFRECHT erscheint (Querformat-Seite).
    (x, top) = Lage in der angezeigten Seite, wie pdfplumber sie meldet."""
    return (f"BT /F1 {groesse} Tf {farbe} 0 1 -1 0 {top + groesse * 0.8:.2f} {x:.2f} Tm ("
            + text.replace("(", "\\(").replace(")", "\\)") + ") Tj ET\n")


def annot_gedreht90(art: str, x0: float, top0: float, x1: float, top1: float, farbe: str = "0 0 1") -> str:
    """Annotation fuer eine /Rotate-90-Seite; Koordinaten der ANGEZEIGTEN Seite werden in den Seitenraum umgerechnet."""
    ux0, ux1, uy0, uy1 = top0, top1, x0, x1
    return (f"/Subtype /{art} /Rect [{ux0} {uy0} {ux1} {uy1}] /QuadPoints [{ux0} {uy1} {ux0} {uy0} {ux1} {uy1} {ux1} {uy0}] "
            f"/C [{farbe}] /F 4")


# ---------------------------------------------------------------- Lesen
def lese(ordner: Path, dok_: PdfDokument | bytes, name: str = "t.pdf") -> PdfAnalyse:
    """PDF (PdfDokument oder rohe Bytes) in den Ordner schreiben und mit dem echten PDF-Leser lesen."""
    ordner.mkdir(parents=True, exist_ok=True)
    pfad = ordner / name
    if isinstance(dok_, (bytes, bytearray)):
        pfad.write_bytes(bytes(dok_))
    else:
        dok_.speichern(pfad)
    return analysiere_pdf(pfad)


# ---------------------------------------------------------------- Auswerten
def alle_tokens(a: PdfAnalyse) -> list[PdfToken]:
    return [t for b in a.blocks for t in b.tokens]


def woerter(a: PdfAnalyse, status: str | None = None) -> list[PdfToken]:
    return [t for t in alle_tokens(a) if status is None or t.status == status]


def formen(b: PdfBlock) -> set[str]:
    """Alle Wortformen (alt und neu) eines Blocks."""
    return {x for t in b.tokens for x in (t.old, t.new) if x}


def text_von(b: PdfBlock, neu: bool = True) -> str:
    return " ".join((t.new if neu else t.old) for t in b.tokens if (t.new if neu else t.old))


def bloecke_mit(a: PdfAnalyse, *woerter_: str) -> list[PdfBlock]:
    """Bloecke, in denen alle diese Woerter (als ganze Wortform, alt oder neu) vorkommen."""
    return [b for b in a.blocks if all(w in formen(b) for w in woerter_)]


def token(a: PdfAnalyse, wort: str) -> PdfToken | None:
    """Erstes Wort, dessen alte oder neue Form genau `wort` ist."""
    return next((t for t in alle_tokens(a) if wort in (t.old, t.new)), None)


def tokens_mit(a: PdfAnalyse, teil: str) -> list[PdfToken]:
    """Alle Woerter, deren alte oder neue Form `teil` enthaelt."""
    return [t for t in alle_tokens(a) if teil in t.old or teil in t.new]


def hinweis_bei(a: PdfAnalyse, ts: list[PdfToken]) -> bool:
    """True, wenn an einem dieser Woerter ein Hinweis haengt oder der Block, in dem sie stehen, Hinweise hat."""
    ids = {id(t) for t in ts}
    for b in a.blocks:
        if any(id(t) in ids for t in b.tokens) and (b.hinweise or any(t.hinweis for t in b.tokens if id(t) in ids)):
            return True
    return False


def kurz(t: PdfToken) -> str:
    return f"old={t.old!r} new={t.new!r} status={t.status}" + (f" hinweis={t.hinweis!r}" if t.hinweis else "")


_ALLGEMEIN = "keine Änderungen erkannt"


def meldungen(a: PdfAnalyse, mit_warnungen: bool = False) -> list[str]:
    """ALLES, womit der Leser auf Unsicheres hinweist: Luecken, Hinweise der Bloecke und Woerter
    (optional auch Warnungen; die ALLGEMEINE Warnung 'keine Aenderungen erkannt' zaehlt nie)."""
    out = [f"luecke:{l['art']}" for l in a.luecken]
    for b in a.blocks:
        out += [f"block:{h}" for h in b.hinweise]
        out += [f"wort:{t.hinweis}" for t in b.tokens if t.hinweis]
    if mit_warnungen:
        out += [f"warnung:{w}" for w in a.warnungen if _ALLGEMEIN not in w]
    return out


def luecken_arten(a: PdfAnalyse) -> set[str]:
    for l in a.luecken:
        if l["art"] not in LUECKEN_ARTEN:
            raise RuntimeError(f"Testaufbau fehlerhaft: unbekannte Lueckenart {l['art']!r}")
    return {l["art"] for l in a.luecken}


def lueckenwoerter(a: PdfAnalyse, art: str) -> int:
    return sum(int(l.get("woerter", 0)) for l in a.luecken if l["art"] == art)


def vollstaendig(a: PdfAnalyse) -> bool:
    """Was die Bilanz zum Lauf sagen wuerde (nur anhand der Luecken des Lesers)."""
    return bool(bilanz([], a.markierte_woerter, a.ignoriert, a.luecken)["vollstaendig"])


def geaendert(t: PdfToken) -> bool:
    return t.status in (ADD, DEL, MOD)


def kopf_fuss_zone(a: PdfAnalyse, anteil: float = 0.07) -> list[PdfBlock]:
    """Bloecke, die ganz im oberen oder unteren Seitenrand (Kopf-/Fusszeilen-Zone) liegen."""
    out = []
    for b in a.blocks:
        _, h = a.seiten_masse[b.page]
        if b.bbox[3] < anteil * h or b.bbox[1] > (1 - anteil) * h:
            out.append(b)
    return out
