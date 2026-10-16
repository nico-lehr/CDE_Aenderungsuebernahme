"""Zweiter PDF-Erzeuger NUR fuer Tests: LibreOffice (erfundene Beispieldokumente, Regel 16).

Dasselbe logische Modell wie beim eigenen PDF-Schreiber (pdf_writer.py) wird als Flat-ODT (.fodt, von Hand geschriebenes
XML, keine neuen Pakete) geschrieben und mit

    soffice --headless -env:UserInstallation=file:///<tmp>/lo_profil_<pid>_... --convert-to pdf --outdir <tmp> doc.fodt

in ein PDF verwandelt. Jeder Aufruf hat ein eigenes LibreOffice-Profil (Parallellaeufe gehen) und ein Zeitlimit.

So sieht das PDF aus wie aus einem echten Textprogramm (Word-aehnlich):
  * eingebettete Schriften (Carlito, Liberation Sans, Liberation Serif), Blocksatz
  * weiche Trennzeichen (U+00AD) in langen Woertern -> Silbentrennung am Zeilenende (sichtbarer Strich nur bei Trennung)
  * echte Aufzaehlungen/Nummerierungen (Listen-Stile), Fussnoten, Kopf-/Fusszeile mit Seitenzahl-Feld
  * Absaetze ueber Seitenumbrueche (natuerlicher Umbruch), Tabellen, Bilder, Zwei-Spalten-Bereich, Scan-Seite (eigene Seitenvorlage)
  * Markierung NUR per Zeichenformat (wie von Hand eingefaerbt): blau = neu (teils zusaetzlich unterstrichen),
    rot + durchgestrichen = entfaellt; der Strich laeuft ueber ganze Zeilen (auch ueber die Leerzeichen)
  * Aenderungsbalken am Rand (senkrechte Linie links neben geaenderten Absaetzen), je Paar zufaellig an/aus

Die Wahrheit (vorher/nachher, erwartet.json) kommt immer aus dem logischen Modell, nie aus diesem PDF.
Dieses Modul importiert nichts aus src/ und nichts aus generator.py (kein Kreis). Fehlt `soffice`, meldet
`verfuegbar()` False -- Aufrufer ueberspringen dann sauber (Tests: skipif), es gibt keinen Absturz beim Import.
"""
from __future__ import annotations

import base64
import io
import os
import random
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

from .pdf_writer import GRUEN, LINK_BLAU, _teile, muster_bild, scan_bild

ZEITLIMIT_S = 180
SCHRIFTEN = ("Carlito", "Liberation Sans", "Liberation Serif")
SHY = "­"
_FN = "\x01"                       # Platzhalter fuer die Fussnoten-Marke im Markup (nur intern)
_VOKALE = set("aeiouäöüAEIOUÄÖÜ")


class LibreOfficeFehlt(RuntimeError):
    """soffice ist nicht installiert (Erzeuger 'libreoffice' nicht verfuegbar)."""


# ====================================================================== soffice finden
def soffice_pfad() -> str | None:
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    return None


def verfuegbar() -> bool:
    return soffice_pfad() is not None


_VERSION: list[str | None] = []


def lo_version() -> str | None:
    """'LibreOffice 24.2.7.2 ...' -> '24.2.7.2' (None, wenn soffice fehlt oder nicht antwortet)."""
    if _VERSION:
        return _VERSION[0]
    exe = soffice_pfad()
    ver = None
    if exe:
        try:
            r = subprocess.run([exe, "--headless", "--version"], capture_output=True, text=True, timeout=60)
            m = re.search(r"(\d+(?:\.\d+)+)", r.stdout or "")
            ver = m.group(1) if m else (r.stdout or "").strip()[:40] or None
        except (OSError, subprocess.SubprocessError):
            ver = None
    _VERSION.append(ver)
    return ver


# ====================================================================== Hilfen
def _hex(farbe) -> str:
    """(r, g, b) 0..1 oder 'RRGGBB' -> '#rrggbb'."""
    if isinstance(farbe, str):
        return "#" + farbe.lstrip("#").lower()
    return "#" + "".join(f"{int(round(v * 255)):02x}" for v in farbe)


def _shy_wort(wort: str, rng: random.Random) -> str:
    """Setzt weiche Trennzeichen in ein langes Wort: vor einem einzelnen Konsonanten zwischen zwei Vokalen (Ü-ber-wa-chung),
    Abstand >= 3 Zeichen. Keine echte Silbenlehre -- es reicht, dass die Trennstellen plausibel aussehen."""
    if len(wort) < 11 or not wort.isalpha():
        return wort
    stellen = [i for i in range(3, len(wort) - 3)
               if wort[i - 1] in _VOKALE and wort[i] not in _VOKALE and wort[i + 1] in _VOKALE]
    gewaehlt: list[int] = []
    for i in stellen:
        if not gewaehlt or i - gewaehlt[-1] >= 3:
            if rng.random() < 0.8:
                gewaehlt.append(i)
    for i in reversed(gewaehlt):
        wort = wort[:i] + SHY + wort[i:]
    return wort


def _mit_shy(text: str, rng: random.Random) -> str:
    return re.sub(r"[A-Za-zÄÖÜäöüß]{11,}", lambda m: _shy_wort(m.group(0), rng), text)


class _Stile:
    """Sammelt automatische Zeichen-/Absatzstile (Name je Eigenschaften-Satz)."""

    def __init__(self) -> None:
        self.text: dict[tuple, str] = {}
        self.para: dict[tuple, str] = {}

    def t(self, **props: str) -> str:
        key = tuple(sorted(props.items()))
        if key not in self.text:
            self.text[key] = f"T{len(self.text) + 1}"
        return self.text[key]

    def p(self, parent: str, master: str | None = None, **props: str) -> str:
        key = (parent, master, tuple(sorted(props.items())))
        if key not in self.para:
            self.para[key] = f"P{len(self.para) + 1}"
        return self.para[key]

    def xml(self) -> str:
        out = []
        for key, name in self.text.items():
            attrs = " ".join(f'{k}="{escape(v)}"' for k, v in key)
            out.append(f'<style:style style:name="{name}" style:family="text"><style:text-properties {attrs}/></style:style>')
        for (parent, master, props), name in self.para.items():
            mp = f' style:master-page-name="{master}"' if master else ""
            ppar = {k[2:]: v for k, v in props if k.startswith("p:")}
            tpar = {k[2:]: v for k, v in props if k.startswith("t:")}
            inner = ""
            if ppar:
                inner += "<style:paragraph-properties " + " ".join(f'{k}="{escape(v)}"' for k, v in ppar.items()) + "/>"
            if tpar:
                inner += "<style:text-properties " + " ".join(f'{k}="{escape(v)}"' for k, v in tpar.items()) + "/>"
            out.append(f'<style:style style:name="{name}" style:family="paragraph" style:parent-style-name="{parent}"{mp}>'
                       f"{inner}</style:style>")
        return "\n".join(out)


_STRICH = {"style:text-line-through-style": "solid", "style:text-line-through-type": "single",
           "style:text-line-through-width": "auto"}
_UNTER = {"style:text-underline-style": "solid", "style:text-underline-width": "auto",
          "style:text-underline-color": "font-color"}


# ====================================================================== Dokument
class LoDokument:
    """Baut den Inhalt (office:text) und die Stile fuer EIN Paar."""

    def __init__(self, ctx: Any, rng: random.Random) -> None:
        self.ctx = ctx
        self.rng = rng
        kw = dict(ctx.pdfkw)
        self.variante = kw.get("variante", "farbe")
        if self.variante == "annot":           # Kommentar-Markierungen gibt es per Zeichenformat nicht -> wie farbig markiert
            self.variante = "farbe"
        self.fn = _hex(kw["farbe_neu"]) if kw.get("farbe_neu") else "#0000ff"
        self.fw = _hex(kw["farbe_weg"]) if kw.get("farbe_weg") else "#ff0000"
        self.kopf_fuss = kw.get("kopf_fuss", True)
        self.kopf_text = kw.get("kopf_text", "Beispieldokument - erfunden")
        self.kopf_farbe = _hex(kw.get("kopf_farbe", (0.4, 0.4, 0.4)))
        self.fuss_vorlage = kw.get("fuss_vorlage", "Seite {nr} von {n}")
        # je Paar zufaellig (eigener Zufallsstrom, das Modell bleibt unberuehrt)
        self.schrift = rng.choice(SCHRIFTEN)
        self.kopf_schrift = rng.choice(("Liberation Sans", self.schrift))
        self.unter_neu = rng.random() < 0.4              # blau (neu) zusaetzlich unterstrichen
        self.balken = rng.random() < 0.5                 # Aenderungsbalken am Rand
        self.stile = _Stile()
        self.koerper: list[str] = []
        self.fussnoten = 0
        self.bilder = 0
        self.tabellen = 0
        self.listen = 0
        self.nach_scan = False
        self.erste = True
        self.info = {"schrift": self.schrift, "neu_unterstrichen": self.unter_neu, "aenderungsbalken": self.balken,
                     "variante": self.variante}

    # ---- Zeichenformat je Markierungs-Stil ----------------------------------------------
    def _tstil(self, stil: str, farben) -> str | None:
        fn, fw = self.fn, self.fw
        if farben:
            fn, fw = _hex(farben[0]), _hex(farben[1])
        v = self.variante
        p: dict[str, str] = {}
        if stil == "k":
            return None
        if stil == "n":
            if v == "hintergrund":
                p["fo:background-color"] = "#b8dbff"
            else:
                p["fo:color"] = fn
                if self.unter_neu:
                    p.update(_UNTER)
        elif stil == "g":
            if v == "farbe":
                p.update(_STRICH, **{"fo:color": fw})
            elif v == "hintergrund":
                p.update(_STRICH, **{"fo:background-color": "#ffcccc"})
            else:
                p.update(_STRICH)
        elif stil == "bu":
            p.update(_UNTER, **{"fo:color": fn})
        elif stil == "bs":
            p.update(_STRICH, **{"fo:color": fn})
        elif stil == "ru":
            p.update(_UNTER, **{"fo:color": fw})
        elif stil == "r":
            p["fo:color"] = fw
        elif stil == "s":
            p.update(_STRICH)
        elif stil == "l":
            p.update(_UNTER, **{"fo:color": _hex(LINK_BLAU)})
        elif stil == "gn":
            p["fo:color"] = _hex(GRUEN)
        else:
            return None
        return self.stile.t(**p)

    def _text(self, text: str, note: str = "") -> str:
        """Text -> ODF-Inline-XML (Leerzeichenfolgen, Tabs, Zeilenumbrueche, Fussnoten-Platzhalter, weiche Trennzeichen)."""
        text = _mit_shy(text, self.rng)
        out: list[str] = []
        for stueck in re.split(r"( {2,}|\t|\n|\x01)", text):
            if not stueck:
                continue
            if stueck == "\t":
                out.append("<text:tab/>")
            elif stueck == "\n":
                out.append("<text:line-break/>")
            elif stueck == _FN:
                out.append(note)
            elif stueck.startswith("  "):
                out.append(f'<text:s text:c="{len(stueck)}"/>')
            else:
                out.append(escape(stueck))
        return "".join(out)

    def _laeufe(self, markup: str, farben=None, note: str = "") -> str:
        out: list[str] = []
        for text, stil in _teile(markup):
            inhalt = self._text(text, note)
            name = self._tstil(stil, farben)
            if name:
                inhalt = f'<text:span text:style-name="{name}">{inhalt}</text:span>'
            if stil == "l":
                inhalt = (f'<text:a xlink:type="simple" xlink:href="https://beispiel.example/dokument">{inhalt}</text:a>')
            out.append(inhalt)
        return "".join(out)

    # ---- Absatz-Stile -------------------------------------------------------------------------
    def _master(self) -> str | None:
        """Seitenvorlage fuer den naechsten Absatz: am Anfang 'Erste' (falls Kopf nur auf Seite 1), nach einer Scan-Seite
        'Standard' (erzwingt den Seitenwechsel)."""
        if self.erste:
            self.erste = False
            return "Erste"
        if self.nach_scan:
            self.nach_scan = False
            return "Standard"
        return None

    def _pstil(self, parent: str, geaendert: bool = False, ganz: str | None = None, farben=None) -> str:
        props: dict[str, str] = {}
        if self.balken and geaendert and parent != "Liste":
            # Aenderungsbalken: senkrechte Linie 0,4 cm links neben dem Absatz, der Text bleibt an seiner Stelle
            props.update({"p:fo:border-left": "0.75pt solid #000000", "p:fo:padding-left": "0.4cm",
                          "p:fo:margin-left": "-0.43cm", "p:style:join-border": "false"})
        if ganz == "neu":                         # Listenzeichen eines ganz neuen Punkts: blau (wie das Absatzende)
            if self.variante != "hintergrund":
                props["t:fo:color"] = _hex(farben[0]) if farben else self.fn
        elif ganz == "weg":
            props.update({"t:" + k: v for k, v in _STRICH.items()})
            if self.variante == "farbe":
                props["t:fo:color"] = _hex(farben[1]) if farben else self.fw
        master = self._master()
        if not props and master is None:
            return parent
        return self.stile.p(parent, master, **props)

    # ---- Bausteine -----------------------------------------------------------------------------
    def ueberschrift(self, markup: str, ebene: int) -> None:
        name = self._pstil(f"H{min(ebene, 3)}")
        self.koerper.append(f'<text:h text:style-name="{name}" text:outline-level="{min(ebene, 3)}">'
                            f"{self._laeufe(markup)}</text:h>")

    def absatz(self, a: Any) -> str:
        markup = a.markup
        note = ""
        if getattr(a, "fussnote", None) and markup[:1].isalpha():
            m = re.match(r"(\S+)(.*)", markup, re.S)
            if m:
                self.fussnoten += 1
                n = self.fussnoten
                markup = m.group(1) + _FN + m.group(2)
                note = (f'<text:note text:id="ftn{n}" text:note-class="footnote"><text:note-citation>{n}</text:note-citation>'
                        f'<text:note-body><text:p text:style-name="Fussnote">{escape(a.fussnote)}</text:p>'
                        f"</text:note-body></text:note>")
        ganz = _ganz(markup)
        name = self._pstil("Liste" if a.marker else "Text", bool(a.typ), ganz if a.marker else None, a.farben)
        return f'<text:p text:style-name="{name}">{self._laeufe(markup, a.farben, note)}</text:p>'

    def liste(self, items: list[Any]) -> None:
        art = _listen_art(items[0].marker)
        self.listen += 1
        start = _listen_start(items[0].marker)
        teile = []
        for i, a in enumerate(items):
            sv = f' text:start-value="{start}"' if (i == 0 and art in ("num", "alpha") and start != 1) else ""
            teile.append(f"<text:list-item{sv}>{self.absatz(a)}</text:list-item>")
        self.koerper.append(f'<text:list text:style-name="L_{_LISTEN_NAME[art]}">' + "".join(teile) + "</text:list>")

    def zwei_spalten(self, absaetze: list[Any], nr: int) -> None:
        inhalt = "".join(self.absatz(a) for a in absaetze)
        self.koerper.append(f'<text:section text:style-name="Sect2" text:name="Spalten{nr}">{inhalt}</text:section>')

    def tabelle(self, zeilen: list[list[str]], stil: tuple, farben=None) -> None:
        self.tabellen += 1
        n = self.tabellen
        kopf, band = stil
        name = f"Tab{n}"
        sp = len(zeilen[0])
        z = [f'<table:table table:name="{name}" table:style-name="Tabelle">',
             f'<table:table-column table:number-columns-repeated="{sp}"/>']
        for i, zeile in enumerate(zeilen):
            if i == 0 and kopf:
                zs = self._zellstil(kopf)
                ts = self.stile.t(**{"fo:color": "#ffffff", "fo:font-weight": "bold"})
            elif i > 0 and band and i % 2 == 1:
                zs, ts = self._zellstil(band), None
            else:
                zs, ts = self._zellstil(None), None
            z.append("<table:table-row>")
            for zelle in zeile:
                inhalt = self._laeufe(zelle, farben)
                if ts:
                    inhalt = f'<text:span text:style-name="{ts}">{inhalt}</text:span>'
                z.append(f'<table:table-cell table:style-name="{zs}" office:value-type="string">'
                         f'<text:p text:style-name="Tabellentext">{inhalt}</text:p></table:table-cell>')
            z.append("</table:table-row>")
        z.append("</table:table>")
        master = self._master()
        if master:                                  # Seitenvorlage braucht einen Absatz
            self.koerper.append(f'<text:p text:style-name="{self.stile.p("Text", master)}"/>')
        self.koerper.append("".join(z))

    @staticmethod
    def _zellstil(hinter: str | None) -> str:
        return {"4472C4": "ZelleK", "1F4E9F": "ZelleD", "B4C6E7": "ZelleB"}.get((hinter or "").upper(), "ZelleW")

    def bild(self, img, breite_cm: float = 7.0) -> None:
        self.bilder += 1
        hoehe = breite_cm * img.size[1] / img.size[0]
        self.koerper.append(f'<text:p text:style-name="{self._pstil("Text")}">{_frame(img, f"Bild{self.bilder}", breite_cm, hoehe)}'
                            f"</text:p>")

    def scan(self, img) -> None:
        self.bilder += 1
        self.erste = False
        self.koerper.append(f'<text:p text:style-name="{self.stile.p("Scanseite", "Scan")}">'
                            f'{_frame(img, f"Scan{self.bilder}", 20.9, 29.5)}</text:p>')
        self.nach_scan = True

    # ---- Ganzes Dokument ----------------------------------------------------------------------
    def fodt(self) -> str:
        return _VORLAGE.format(
            schrift=self.schrift, kopf_schrift=self.kopf_schrift, auto=self.stile.xml(),
            master=self._master_xml(), koerper="\n".join(self.koerper), kopf_farbe=self.kopf_farbe)

    def _kopf_fuss_xml(self) -> str:
        teile = re.split(r"(\{nr\}|\{n\})", self.fuss_vorlage)
        fuss = "".join('<text:page-number text:select-page="current">1</text:page-number>' if t == "{nr}" else
                       ("<text:page-count>1</text:page-count>" if t == "{n}" else escape(t)) for t in teile if t)
        return (f'<style:header><text:p text:style-name="Kopf">{escape(self.kopf_text)}</text:p></style:header>'
                f'<style:footer><text:p text:style-name="Fuss">{fuss}</text:p></style:footer>')

    def _master_xml(self) -> str:
        kf = self.kopf_fuss
        alle = kf is True
        erste = alle or (isinstance(kf, (set, frozenset, list, tuple)) and 1 in kf)
        kopf = self._kopf_fuss_xml()
        return (f'<style:master-page style:name="Standard" style:page-layout-name="pm1">{kopf if alle else ""}</style:master-page>'
                f'<style:master-page style:name="Erste" style:page-layout-name="pm1" style:next-style-name="Standard">'
                f'{kopf if erste else ""}</style:master-page>'
                f'<style:master-page style:name="Scan" style:page-layout-name="pm2" style:next-style-name="Standard"/>')


_LISTEN_NAME = {"•": "punkt", "–": "strich", "num": "zahl", "alpha": "buchstabe"}


def _listen_art(marker: str) -> str:
    if re.fullmatch(r"\d+\.", marker):
        return "num"
    if re.fullmatch(r"[a-z]\)", marker):
        return "alpha"
    return marker if marker in ("•", "–") else "•"


def _listen_start(marker: str) -> int:
    if re.fullmatch(r"\d+\.", marker):
        return int(marker[:-1])
    if re.fullmatch(r"[a-z]\)", marker):
        return ord(marker[0]) - 96
    return 1


def _ganz(markup: str) -> str | None:
    m = markup.strip()
    if m.startswith("{+") and m.endswith("+}") and m.count("{+") == 1:
        return "neu"
    if m.startswith("{-") and m.endswith("-}") and m.count("{-") == 1:
        return "weg"
    return None


def _frame(img, name: str, breite_cm: float, hoehe_cm: float) -> str:
    puffer = io.BytesIO()
    img.convert("RGB").save(puffer, format="PNG")
    daten = base64.b64encode(puffer.getvalue()).decode("ascii")
    return (f'<draw:frame draw:name="{name}" text:anchor-type="as-char" svg:width="{breite_cm:.2f}cm" '
            f'svg:height="{hoehe_cm:.2f}cm" draw:z-index="0"><draw:image><office:binary-data>{daten}'
            f"</office:binary-data></draw:image></draw:frame>")


def _rgb255(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]


# ====================================================================== Modell -> Dokument
def baue_dokument(ctx: Any, rng: random.Random) -> LoDokument:
    """Geht das Modell (ctx.sektionen) in derselben Reihenfolge durch wie generator.rendere_pdf."""
    d = LoDokument(ctx, rng)
    spalten = 0
    for s in ctx.sektionen:
        if s.scan:
            elemente = [("h", s.kopf_markup)] + [("p", ((a.marker + " ") if a.marker else "") + a.markup)
                                                 for a in s.absaetze if a.art == "text"]
            d.scan(scan_bild(elemente, farbe_neu=_rgb255(d.fn), farbe_weg=_rgb255(d.fw), seed=ctx.seed + s.id))
            continue
        d.ueberschrift(s.kopf_markup, s.ebene)
        if s.zwei_spalten:
            spalten += 1
            d.zwei_spalten([a for a in s.absaetze if a.art == "text"], spalten)
            continue
        i = 0
        abs_ = s.absaetze
        while i < len(abs_):
            a = abs_[i]
            if a.art == "tabelle":
                d.tabelle(a.tabelle or [], a.tabelle_stil, a.farben)
            elif a.art == "bild":
                d.bild(muster_bild(240, 140, a.bild_seed))
            elif a.marker:
                art = _listen_art(a.marker)
                j = i
                while j < len(abs_) and abs_[j].art == "text" and abs_[j].marker and _listen_art(abs_[j].marker) == art:
                    j += 1
                d.liste(abs_[i:j])
                i = j
                continue
            else:
                d.koerper.append(d.absatz(a))
            i += 1
    return d


def konvertiere(fodt: str, ziel_pdf: Path, zeitlimit: int = ZEITLIMIT_S) -> None:
    """Schreibt das .fodt in einen Temp-Ordner und laesst LibreOffice daraus ein PDF machen (eigenes Profil je Aufruf)."""
    exe = soffice_pfad()
    if exe is None:
        raise LibreOfficeFehlt("soffice (LibreOffice) ist nicht installiert")
    with tempfile.TemporaryDirectory(prefix=f"lo_{os.getpid()}_") as tmp_name:
        tmp = Path(tmp_name)
        profil = tmp / f"lo_profil_{os.getpid()}"
        quelle = tmp / "dokument.fodt"
        quelle.write_text(fodt, encoding="utf-8")
        cmd = [exe, "--headless", "--norestore", "--nologo", "--nolockcheck",
               f"-env:UserInstallation={profil.as_uri()}", "--convert-to", "pdf:writer_pdf_Export",
               "--outdir", str(tmp), str(quelle)]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=zeitlimit,
                               env={**os.environ, "HOME": str(tmp)})
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(f"LibreOffice hat das Zeitlimit von {zeitlimit} s ueberschritten") from exc
        pdf = tmp / "dokument.pdf"
        if r.returncode != 0 or not pdf.is_file():
            raise RuntimeError(f"LibreOffice konnte kein PDF erzeugen (Code {r.returncode})")
        shutil.copyfile(pdf, ziel_pdf)


def seitenzahl(pdf: Path) -> int:
    """Nur die Seitenzahl (kein Lesen des Inhalts)."""
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(str(pdf))
    try:
        return len(doc)
    finally:
        doc.close()


def rendere(ctx: Any, ziel_pdf: Path) -> dict[str, Any]:
    """Modell -> PDF ueber LibreOffice. Gibt Angaben fuer erwartet.json zurueck (nur Zahlen/Kennungen)."""
    rng = random.Random(ctx.seed * 7_919 + 104_729)        # eigener Strom: das Modell bleibt je Seed gleich
    d = baue_dokument(ctx, rng)
    konvertiere(d.fodt(), ziel_pdf)
    info = dict(d.info)
    info.update(seiten=seitenzahl(ziel_pdf), fussnoten=d.fussnoten, listen=d.listen, tabellen=d.tabellen, bilder=d.bilder,
                libreoffice_version=lo_version() or "")
    return info


_NS = ('xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
       'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
       'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
       'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
       'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" '
       'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
       'xmlns:xlink="http://www.w3.org/1999/xlink" '
       'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" '
       'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"')

_LISTEN_XML = (
    '<text:list-style style:name="L_punkt"><text:list-level-style-bullet text:level="1" text:bullet-char="•">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.63cm" '
    'fo:text-indent="-0.5cm" fo:margin-left="0.63cm"/></style:list-level-properties></text:list-level-style-bullet></text:list-style>'
    '<text:list-style style:name="L_strich"><text:list-level-style-bullet text:level="1" text:bullet-char="–">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.63cm" '
    'fo:text-indent="-0.5cm" fo:margin-left="0.63cm"/></style:list-level-properties></text:list-level-style-bullet></text:list-style>'
    '<text:list-style style:name="L_zahl"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix=".">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.63cm" '
    'fo:text-indent="-0.5cm" fo:margin-left="0.63cm"/></style:list-level-properties></text:list-level-style-number></text:list-style>'
    '<text:list-style style:name="L_buchstabe"><text:list-level-style-number text:level="1" style:num-format="a" style:num-suffix=")">'
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">'
    '<style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.63cm" '
    'fo:text-indent="-0.5cm" fo:margin-left="0.63cm"/></style:list-level-properties></text:list-level-style-number></text:list-style>'
)

_VORLAGE = ('<?xml version="1.0" encoding="UTF-8"?>\n'
            f'<office:document {_NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">\n'
            '<office:font-face-decls>'
            '<style:font-face style:name="Carlito" svg:font-family="Carlito" style:font-family-generic="swiss" style:font-pitch="variable"/>'
            '<style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" style:font-family-generic="swiss" style:font-pitch="variable"/>'
            '<style:font-face style:name="Liberation Serif" svg:font-family="&apos;Liberation Serif&apos;" style:font-family-generic="roman" style:font-pitch="variable"/>'
            '</office:font-face-decls>\n'
            '<office:styles>'
            '<style:default-style style:family="paragraph"><style:paragraph-properties fo:hyphenation-ladder-count="no-limit"/>'
            '<style:text-properties style:font-name="{schrift}" fo:font-size="11pt" fo:language="de" fo:country="DE" '
            'fo:hyphenate="false"/></style:default-style>'
            '<style:style style:name="Standard" style:family="paragraph"/>'
            '<style:style style:name="Text" style:family="paragraph" style:parent-style-name="Standard">'
            '<style:paragraph-properties fo:text-align="justify" fo:margin-top="0cm" fo:margin-bottom="0.25cm" '
            'fo:orphans="2" fo:widows="2"/></style:style>'
            '<style:style style:name="Liste" style:family="paragraph" style:parent-style-name="Text">'
            '<style:paragraph-properties fo:margin-bottom="0.15cm"/></style:style>'
            '<style:style style:name="Tabellentext" style:family="paragraph" style:parent-style-name="Standard">'
            '<style:text-properties fo:font-size="10pt"/></style:style>'
            '<style:style style:name="Scanseite" style:family="paragraph" style:parent-style-name="Standard">'
            '<style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:line-height="100%"/></style:style>'
            '<style:style style:name="Fussnote" style:family="paragraph" style:parent-style-name="Standard">'
            '<style:text-properties fo:font-size="8pt"/></style:style>'
            '<style:style style:name="Kopf" style:family="paragraph" style:parent-style-name="Standard">'
            '<style:text-properties fo:font-size="9pt" fo:color="{kopf_farbe}"/></style:style>'
            '<style:style style:name="Fuss" style:family="paragraph" style:parent-style-name="Standard">'
            '<style:paragraph-properties fo:text-align="center"/><style:text-properties fo:font-size="9pt" fo:color="{kopf_farbe}"/></style:style>'
            '<style:style style:name="H1" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="1">'
            '<style:paragraph-properties fo:margin-top="0.4cm" fo:margin-bottom="0.2cm" fo:keep-with-next="always"/>'
            '<style:text-properties style:font-name="{kopf_schrift}" fo:font-size="14pt" fo:font-weight="bold"/></style:style>'
            '<style:style style:name="H2" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="2">'
            '<style:paragraph-properties fo:margin-top="0.3cm" fo:margin-bottom="0.15cm" fo:keep-with-next="always"/>'
            '<style:text-properties style:font-name="{kopf_schrift}" fo:font-size="12.5pt" fo:font-weight="bold"/></style:style>'
            '<style:style style:name="H3" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="3">'
            '<style:paragraph-properties fo:margin-top="0.25cm" fo:margin-bottom="0.15cm" fo:keep-with-next="always"/>'
            '<style:text-properties style:font-name="{kopf_schrift}" fo:font-size="12pt" fo:font-weight="bold"/></style:style>'
            '<style:style style:name="Tabelle" style:family="table"><style:table-properties style:width="16cm" table:align="left"/></style:style>'
            + _LISTEN_XML +
            '<text:notes-configuration text:note-class="footnote" style:num-format="1" text:start-value="0" '
            'text:footnotes-position="page" text:start-numbering-at="document"/>'
            '</office:styles>\n'
            '<office:automatic-styles>'
            '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" '
            'fo:margin-top="1.2cm" fo:margin-bottom="1.2cm" fo:margin-left="2.2cm" fo:margin-right="2.2cm"/>'
            '<style:header-style><style:header-footer-properties fo:min-height="0.6cm" fo:margin-bottom="0.5cm"/></style:header-style>'
            '<style:footer-style><style:header-footer-properties fo:min-height="0.6cm" fo:margin-top="0.5cm"/></style:footer-style>'
            '</style:page-layout>'
            '<style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" '
            'fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm"/></style:page-layout>'
            '<style:style style:name="Sect2" style:family="section"><style:section-properties text:dont-balance-text-columns="false">'
            '<style:columns fo:column-count="2" fo:column-gap="0.6cm"/></style:section-properties></style:style>'
            '<style:style style:name="ZelleK" style:family="table-cell"><style:table-cell-properties fo:background-color="#4472c4" '
            'fo:padding="0.08cm" fo:border="0.5pt solid #000000"/></style:style>'
            '<style:style style:name="ZelleD" style:family="table-cell"><style:table-cell-properties fo:background-color="#1f4e9f" '
            'fo:padding="0.08cm" fo:border="0.5pt solid #000000"/></style:style>'
            '<style:style style:name="ZelleB" style:family="table-cell"><style:table-cell-properties fo:background-color="#b4c6e7" '
            'fo:padding="0.08cm" fo:border="0.5pt solid #000000"/></style:style>'
            '<style:style style:name="ZelleW" style:family="table-cell"><style:table-cell-properties '
            'fo:padding="0.08cm" fo:border="0.5pt solid #000000"/></style:style>'
            '{auto}'
            '</office:automatic-styles>\n'
            '<office:master-styles>{master}</office:master-styles>\n'
            '<office:body><office:text>\n{koerper}\n</office:text></office:body>\n'
            '</office:document>\n')
