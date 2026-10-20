"""Liest eine Word-Datei (.docx) mit Aenderungsverfolgung und liefert dieselbe Datenklasse wie der PDF-Leser (`PdfAnalyse`).

Planer und Applier bleiben unveraendert: `analysiere_docx(pfad)` gibt Bloecke mit Woertern (`PdfToken`: old/new/status,
Zeichen-Laeufe `teile`, `trenner`/`trenner_art`) zurueck, genau wie `pdf_reader.analysiere_pdf` bei einem Word-PDF.

Festlegungen (Auftrag DOCX, ADR 0006):
  * legende_modus = "docx", page = 0, bbox = (0, 0, 0, 0), seiten = 0, seiten_masse leer, luecken nur fuer Unbekanntes
    (das XML hat keine Lese-Luecke), markup_leiste = False; Kommentare stehen in `randnotizen` (Art "kommentar", nur Info).
  * Inhalte: w:ins (neu), w:del / w:delText (entfaellt), w:moveTo (neu), w:moveFrom (entfaellt), Formatwechsel (rPrChange,
    pPrChange ...) sind nur Info. Absatzmarken-Aenderungen (verbundene / getrennte Absaetze) sperren den Absatz (Hinweis-Code
    `absatzgrenze_unklar`) und stehen als Luecke; ein ganz neuer bzw. ganz gestrichener Absatz ist dagegen ein normaler Block.
  * Tabellen: eine Zelle = ein Block mit Hinweis `tabelle` an den geaenderten Woertern (der Planer uebernimmt sie nie allein).
  * Kopf- und Fusszeilen werden getrennt gelesen und wie beim PDF in `ignoriert` gefuehrt (Grund "kopf_fuss"); der Planer
    liest die Spec aus diesen Texten (und der Titelzeile).
  * Textfelder, Fussnoten, Endnoten, Formeln und alles Unbekannte erscheinen als Warnung; haben sie Aenderungen, auch als
    Luecke (Bilanz "unvollstaendig"). Nichts wird still weggelassen.
  * Listennummern werden aus numbering.xml selbst berechnet, in zwei Staenden: Original (ohne w:ins, mit w:del) und Endstand
    (mit w:ins, ohne w:del). Unterscheiden sie sich, ist die Nummer ein Label-Wort mit Zeichen-Laeufen (alt "c.", neu "a.",
    wie "c.a." im PDF). Section-Nummern "1.1 TITEL" und "PART 1 TITEL" folgen den Regeln des PDF-Lesers (farben.yaml layout).
  * Buchstabenlisten: Word zaehlt a..z, dann aa, bb, cc ... (config docx.yaml listen.buchstaben_folge: "alphabet" = aa, ab ...).

Nur Python-Standardbibliothek (zipfile, xml.etree.ElementTree). Keine Dokumentinhalte in Warnungen, Luecken oder Fehlern (Regel 16).
"""
from __future__ import annotations

import re
import zipfile
import zlib
from collections import Counter
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET

from . import config
from .fehler import NichtGefunden, NutzerFehler
from .models import ADD, DEL, KEEP, MOD, PdfBlock, PdfToken, luecke
from .pdf_reader import PdfAnalyse, _block_hinweise, _gross_anteil, _kapitel, _laeufe, _re_cfg, _section_von
from .textutil import norm, section_nummer

_OLE_KOPF = bytes.fromhex("D0CF11E0A1B11AE1")
_NETTO = "x"                 # eingefuegt und danach gestrichen: weder im Original noch im Endstand
_GESCHUETZT = ("\xa0", " ", " ")


# ---------------------------------------------------------------- kleine XML-Helfer (Namensraum egal: Strict und Transitional)
def _ln(e: ET.Element) -> str:
    t = e.tag
    return t.rsplit("}", 1)[-1] if isinstance(t, str) else ""


def _attr(e: ET.Element, name: str, default: str | None = None) -> str | None:
    for k, v in e.attrib.items():
        if k.rsplit("}", 1)[-1] == name:
            return v
    return default


def _kind(e: ET.Element, name: str) -> ET.Element | None:
    for k in e:
        if _ln(k) == name:
            return k
    return None


def _an(e: ET.Element | None, name: str = "val", standard: bool = True) -> bool:
    """Ein- und Aus-Schalter wie w:b (kein val = an, val 0/false/off = aus)."""
    if e is None:
        return False
    v = _attr(e, name)
    if v is None:
        return standard
    return v.strip().lower() not in ("0", "false", "off", "none")


def _ganz(text: str | None, standard: int | None = None) -> int | None:
    try:
        return int(str(text).strip())
    except (TypeError, ValueError):
        return standard


def _fehler(meldung: str, art: str = "docx_kaputt") -> NutzerFehler:
    return NutzerFehler(meldung, 400, art)


def _dcfg(cfg: dict | None) -> dict:
    return cfg["docx"] if cfg and "docx" in cfg else config.lade("docx")


def _lay(cfg: dict | None) -> dict:
    if cfg and "layout" in cfg:
        return cfg["layout"]
    return config.lade("farben")["layout"]


# ---------------------------------------------------------------- Paket oeffnen
@dataclass
class _Paket:
    zf: zipfile.ZipFile
    haupt: str
    namen: set[str]
    dcfg: dict

    def lies(self, name: str) -> bytes | None:
        if name not in self.namen:
            return None
        info = self.zf.getinfo(name)
        if info.file_size > float(self.dcfg.get("max_teil_mb", 64)) * 1024 * 1024:
            raise _fehler("Die Word-Datei enthält einen unerwartet großen Teil und wird aus Sicherheitsgründen nicht gelesen.",
                          "docx_zu_gross")
        try:
            return self.zf.read(name)
        except (zipfile.BadZipFile, OSError, RuntimeError, zlib.error):
            raise _fehler("Die Word-Datei ist beschädigt und kann nicht gelesen werden.")

    def xml(self, name: str) -> ET.Element | None:
        roh = self.lies(name)
        if roh is None:
            return None
        return _parse(roh)


def _parse(roh: bytes) -> ET.Element:
    kopf = roh[:4096].lower()
    if b"<!doctype" in kopf or b"<!entity" in roh[:65536].lower():
        raise _fehler("Die Word-Datei enthält unerwartete Strukturangaben (DOCTYPE) und wird aus Sicherheitsgründen nicht gelesen.",
                      "docx_unsicher")
    try:
        return ET.fromstring(roh)
    except ET.ParseError:
        raise _fehler("Der Text der Word-Datei ist beschädigt und kann nicht gelesen werden.")


def _oeffne(pfad: str | Path, dcfg: dict) -> _Paket:
    p = Path(pfad)
    if not p.is_file():
        raise NichtGefunden("Die Word-Datei wurde nicht gefunden.")
    try:
        with open(p, "rb") as f:
            kopf = f.read(8)
            if kopf == _OLE_KOPF:
                f.seek(0)
                probe = f.read(8 * 1024 * 1024)
                if "EncryptedPackage".encode("utf-16-le") in probe or "EncryptionInfo".encode("utf-16-le") in probe:
                    raise _fehler("Die Word-Datei ist mit einem Passwort geschützt. Bitte in Word das Passwort entfernen "
                                  "(Datei, Informationen, Dokument schützen) und die Datei erneut hochladen.", "docx_passwort")
                raise _fehler("Das ist eine alte Word-Datei (.doc, Word 97–2003). Bitte in Word öffnen und über „Speichern unter“ "
                              "als „Word-Dokument (*.docx)“ speichern.", "docx_alt")
    except OSError:
        raise _fehler("Die Word-Datei konnte nicht geöffnet werden.")
    if not zipfile.is_zipfile(p):
        raise _fehler("Die Datei ist keine gültige Word-Datei (.docx) oder sie ist beschädigt.")
    try:
        zf = zipfile.ZipFile(p)
    except (zipfile.BadZipFile, OSError):
        raise _fehler("Die Datei ist keine gültige Word-Datei (.docx) oder sie ist beschädigt.")
    namen = set(zf.namelist())
    if len(namen) > int(dcfg.get("max_teile", 2000)):
        raise _fehler("Die Word-Datei enthält unerwartet viele Teile und wird nicht gelesen.", "docx_zu_gross")
    paket = _Paket(zf, str(dcfg.get("haupt_teil", "word/document.xml")), namen, dcfg)
    if paket.haupt not in namen:
        ziel = _haupt_aus_beziehungen(paket)
        if ziel and ziel in namen:
            paket.haupt = ziel
    if paket.haupt not in namen:
        zf.close()
        raise _fehler("Die Datei enthält keinen Word-Text (word/document.xml fehlt). Ist es wirklich eine .docx-Datei?")
    return paket


def _haupt_aus_beziehungen(paket: _Paket) -> str:
    try:
        wurzel = paket.xml("_rels/.rels")
    except NutzerFehler:
        return ""
    if wurzel is None:
        return ""
    for r in wurzel:
        if (_attr(r, "Type") or "").endswith("/officeDocument"):
            ziel = (_attr(r, "Target") or "").lstrip("/")
            return ziel
    return ""


# ---------------------------------------------------------------- Formatvorlagen und Nummerierung
@dataclass
class _Stil:
    sid: str
    name: str = ""
    basiert: str = ""
    num_id: int | None = None
    ilvl: int | None = None
    gliederung: int | None = None          # outlineLvl (0 = Ebene 1)
    fett: bool | None = None
    gross: bool | None = None


class _Stile:
    def __init__(self, wurzel: ET.Element | None) -> None:
        self.stile: dict[str, _Stil] = {}
        if wurzel is None:
            return
        for s in wurzel:
            if _ln(s) != "style":
                continue
            sid = _attr(s, "styleId") or ""
            st = _Stil(sid)
            n = _kind(s, "name")
            st.name = (_attr(n, "val") or "") if n is not None else ""
            b = _kind(s, "basedOn")
            st.basiert = (_attr(b, "val") or "") if b is not None else ""
            ppr = _kind(s, "pPr")
            if ppr is not None:
                npr = _kind(ppr, "numPr")
                if npr is not None:
                    st.num_id = _ganz(_attr(_kind(npr, "numId"), "val") if _kind(npr, "numId") is not None else None)
                    st.ilvl = _ganz(_attr(_kind(npr, "ilvl"), "val") if _kind(npr, "ilvl") is not None else None)
                ol = _kind(ppr, "outlineLvl")
                if ol is not None:
                    st.gliederung = _ganz(_attr(ol, "val"))
            rpr = _kind(s, "rPr")
            if rpr is not None:
                if _kind(rpr, "b") is not None:
                    st.fett = _an(_kind(rpr, "b"))
                if _kind(rpr, "caps") is not None:
                    st.gross = _an(_kind(rpr, "caps"))
            self.stile[sid] = st

    def kette(self, sid: str) -> list[_Stil]:
        out: list[_Stil] = []
        sehen: set[str] = set()
        while sid and sid in self.stile and sid not in sehen:
            sehen.add(sid)
            out.append(self.stile[sid])
            sid = self.stile[sid].basiert
        return out

    def eigenschaft(self, sid: str, feld: str) -> Any:
        for st in self.kette(sid):
            v = getattr(st, feld)
            if v is not None:
                return v
        return None


@dataclass
class _Lvl:
    start: int = 1
    fmt: str = "decimal"
    text: str = ""
    restart: int | None = None
    suff: str = "tab"
    lgl: bool = False
    pstil: str = ""


def _lies_lvl(e: ET.Element) -> _Lvl:
    v = _Lvl()
    for k in e:
        n = _ln(k)
        if n == "start":
            v.start = int(_ganz(_attr(k, "val"), 1) or 0)
        elif n == "numFmt":
            v.fmt = _attr(k, "val") or "decimal"
        elif n == "lvlText":
            v.text = _attr(k, "val") or ""
        elif n == "lvlRestart":
            v.restart = _ganz(_attr(k, "val"))
        elif n == "suff":
            v.suff = _attr(k, "val") or "tab"
        elif n == "isLgl":
            v.lgl = _an(k)
        elif n == "pStyle":
            v.pstil = _attr(k, "val") or ""
    return v


@dataclass
class _Num:
    abstrakt: int
    start_ueberschreiben: dict[int, int] = field(default_factory=dict)
    lvl_ueberschreiben: dict[int, _Lvl] = field(default_factory=dict)


class _Nummern:
    """numbering.xml: abstractNum (Ebenen) und num (Instanz mit Ueberschreibungen)."""

    def __init__(self, wurzel: ET.Element | None) -> None:
        self.abstrakt: dict[int, dict[int, _Lvl]] = {}
        self.nums: dict[int, _Num] = {}
        if wurzel is None:
            return
        for e in wurzel:
            n = _ln(e)
            if n == "abstractNum":
                aid = _ganz(_attr(e, "abstractNumId"))
                if aid is None:
                    continue
                self.abstrakt[aid] = {(_ganz(_attr(l, "ilvl"), 0) or 0): _lies_lvl(l) for l in e if _ln(l) == "lvl"}
            elif n == "num":
                nid = _ganz(_attr(e, "numId"))
                a = _kind(e, "abstractNumId")
                aid = _ganz(_attr(a, "val")) if a is not None else None
                if nid is None or aid is None:
                    continue
                num = _Num(aid)
                for o in e:
                    if _ln(o) != "lvlOverride":
                        continue
                    il = _ganz(_attr(o, "ilvl"), 0) or 0
                    so = _kind(o, "startOverride")
                    if so is not None and _ganz(_attr(so, "val")) is not None:
                        num.start_ueberschreiben[il] = int(_ganz(_attr(so, "val")) or 0)
                    lv = _kind(o, "lvl")
                    if lv is not None:
                        num.lvl_ueberschreiben[il] = _lies_lvl(lv)
                self.nums[nid] = num

    def lvl(self, num_id: int, ilvl: int) -> _Lvl | None:
        n = self.nums.get(num_id)
        if n is None:
            return None
        if ilvl in n.lvl_ueberschreiben:
            return n.lvl_ueberschreiben[ilvl]
        return self.abstrakt.get(n.abstrakt, {}).get(ilvl)

    def start(self, num_id: int, ilvl: int) -> int:
        n = self.nums.get(num_id)
        if n is not None and ilvl in n.start_ueberschreiben:
            return n.start_ueberschreiben[ilvl]
        l = self.lvl(num_id, ilvl)
        return l.start if l else 1

    def ebene_von_stil(self, num_id: int, sid: str) -> int | None:
        n = self.nums.get(num_id)
        if n is None:
            return None
        for il, l in self.abstrakt.get(n.abstrakt, {}).items():
            if l.pstil == sid:
                return il
        return None


_ROEM = (("m", 1000), ("cm", 900), ("d", 500), ("cd", 400), ("c", 100), ("xc", 90), ("l", 50), ("xl", 40), ("x", 10),
         ("ix", 9), ("v", 5), ("iv", 4), ("i", 1))


def _roemisch(n: int) -> str:
    out = ""
    for s, w in _ROEM:
        while n >= w:
            out += s
            n -= w
    return out


def _buchstaben(n: int, folge: str) -> str:
    if n < 1:
        return str(n)
    if folge == "alphabet":
        s, m = "", n
        while m > 0:
            m, r = divmod(m - 1, 26)
            s = chr(97 + r) + s
        return s
    return chr(97 + (n - 1) % 26) * ((n - 1) // 26 + 1)


_FORMATE = ("decimal", "decimalZero", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", "bullet", "none")


def _zahl(n: int, fmt: str, folge: str) -> str:
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "lowerLetter":
        return _buchstaben(n, folge)
    if fmt == "upperLetter":
        return _buchstaben(n, folge).upper()
    if fmt == "lowerRoman":
        return _roemisch(n) if n > 0 else str(n)
    if fmt == "upperRoman":
        return _roemisch(n).upper() if n > 0 else str(n)
    return str(n)


@dataclass
class _Marke:
    text: str
    suff: str          # "\t" | " " | ""
    fmt: str


class _Zaehler:
    """Laufende Nummern eines Standes (Original oder Endstand)."""

    def __init__(self, nummern: _Nummern, lcfg: dict) -> None:
        self.n = nummern
        self.folge = str(lcfg.get("buchstaben_folge", "wort"))
        self.punkt = str(lcfg.get("aufzaehlungszeichen", "•"))
        self.maxe = int(lcfg.get("max_ebene", 8))
        self.zaehl: dict[int, list[int | None]] = {}
        self.gestartet: set[int] = set()
        self.unbekannt: Counter = Counter()

    def weiter(self, num_id: int, ilvl: int) -> _Marke | None:
        n = self.n.nums.get(num_id)
        if n is None or num_id == 0:
            return None
        ilvl = max(0, min(ilvl, self.maxe))
        lv = self.n.lvl(num_id, ilvl)
        if lv is None:
            return None
        z = self.zaehl.setdefault(n.abstrakt, [None] * (self.maxe + 1))
        if num_id not in self.gestartet:
            self.gestartet.add(num_id)
            for k, wert in n.start_ueberschreiben.items():
                if k < ilvl:
                    z[k] = wert
                    for j in range(k + 1, len(z)):
                        z[j] = None
                elif k == ilvl:
                    z[k] = None
        z[ilvl] = self.n.start(num_id, ilvl) if z[ilvl] is None else z[ilvl] + 1          # type: ignore[operator]
        for j in range(ilvl + 1, len(z)):
            r = (self.n.lvl(num_id, j) or _Lvl()).restart
            if r is None or (r != 0 and ilvl <= r - 1):
                z[j] = None
        if lv.fmt not in _FORMATE:
            self.unbekannt[lv.fmt if re.fullmatch(r"[A-Za-z]{1,30}", lv.fmt or "") else "andere"] += 1
            return None
        suff = {"tab": "\t", "space": " ", "nothing": ""}.get(lv.suff, "\t")
        if lv.fmt == "none":
            return None
        if lv.fmt == "bullet":
            return _Marke(self.punkt, suff, "bullet")

        def ersetze(m: re.Match) -> str:
            k = int(m.group(1)) - 1
            lk = self.n.lvl(num_id, k)
            wert = z[k] if (k < len(z) and z[k] is not None) else self.n.start(num_id, k)
            fmt = "decimal" if (lv.lgl and k != ilvl) else (lk.fmt if lk else "decimal")
            if fmt not in _FORMATE:
                self.unbekannt[fmt if re.fullmatch(r"[A-Za-z]{1,30}", fmt or "") else "andere"] += 1
                fmt = "decimal"
            return _zahl(int(wert), fmt, self.folge)                                          # type: ignore[arg-type]

        text = re.sub(r"%([1-9])", ersetze, lv.text)
        if not text.strip():
            return None
        return _Marke(text, suff, lv.fmt)


# ---------------------------------------------------------------- Statistik (nur Zahlen und Formen)
@dataclass
class DocxZahlen:
    """Zahlen und Formen einer Word-Datei (kein Text, Regel 16)."""
    zaehler: Counter = field(default_factory=Counter)
    listenarten: Counter = field(default_factory=Counter)
    listenebenen: Counter = field(default_factory=Counter)
    unbekannt: Counter = field(default_factory=Counter)
    unbekannte_zahlformate: Counter = field(default_factory=Counter)
    autoren: set[str] = field(default_factory=set)


# ---------------------------------------------------------------- Absaetze
@dataclass
class _Abs:
    zeichen: list[tuple[str, str]] = field(default_factory=list)       # (Zeichen, Status KEEP|ADD|DEL)
    marke: str = ""                      # "" | "ins" | "del" (Absatzmarke)
    stil: str = ""
    num_neu: tuple[int, int] | None = None
    num_alt: tuple[int, int] | None = None
    gliederung: int | None = None
    fett_zeichen: int = 0
    gesamt_zeichen: int = 0
    behaelter: int = 0                   # 0 = Dokumenttext, sonst Nummer der Tabellenzelle
    in_tabelle: bool = False
    zeilen_status: str = ""              # "" | ADD | DEL (ganze Tabellenzeile eingefuegt / gestrichen)
    ex_alt: bool = True
    ex_neu: bool = True
    absatzgrenze: str = ""               # "" | "getrennt" | "verbunden"
    marke_alt: _Marke | None = None
    marke_neu: _Marke | None = None
    tokens: list[PdfToken] = field(default_factory=list)

    @property
    def hat_alt(self) -> bool:
        return any((not c.isspace()) and s in (KEEP, DEL) for c, s in self.zeichen)

    @property
    def hat_neu(self) -> bool:
        return any((not c.isspace()) and s in (KEEP, ADD) for c, s in self.zeichen)


@dataclass
class _Kontext:
    behaelter: int = 0
    in_tabelle: bool = False
    zeilen_status: str = ""


class _Leser:
    def __init__(self, paket: _Paket, cfg: dict | None, teilen: "_Leser | None" = None) -> None:
        self.paket = paket
        self.z = DocxZahlen()
        if teilen is not None:                      # Kopf-/Fusszeilen: Formatvorlagen und Nummerierung des Dokuments mitbenutzen
            self.dcfg, self.lay = teilen.dcfg, teilen.lay
            self.stile, self.nummern, self._stil_re = teilen.stile, teilen.nummern, teilen._stil_re
        else:
            self.dcfg = _dcfg(cfg)
            self.lay = _lay(cfg)
            self.stile = _Stile(paket.xml(str(self.dcfg.get("styles_teil", "word/styles.xml"))))
            self.nummern = _Nummern(paket.xml(str(self.dcfg.get("numbering_teil", "word/numbering.xml"))))
            self._stil_re = re.compile(str(self.dcfg.get("ueberschrift_stil_muster", r"^$")), re.I)
        self.absaetze: list[_Abs] = []
        self.zellen = 0
        self.luecken: list[dict[str, Any]] = []
        self.warnungen: list[str] = []

    # ------------------------------------------------------------ Textinhalt eines Laufs
    def _lauf_text(self, r: ET.Element, status: str, a: _Abs) -> None:
        """Ein w:r: sichtbare Zeichen anhaengen."""
        rpr = _kind(r, "rPr")
        fett = gross = None
        versteckt = False
        if rpr is not None:
            if _kind(rpr, "rPrChange") is not None:
                self.z.zaehler["format_lauf"] += 1
            if _kind(rpr, "b") is not None:
                fett = _an(_kind(rpr, "b"))
            if _kind(rpr, "caps") is not None:
                gross = _an(_kind(rpr, "caps"))
            versteckt = _an(_kind(rpr, "vanish"))
            rs = _kind(rpr, "rStyle")
            if rs is not None:
                sid = _attr(rs, "val") or ""
                fett = self.stile.eigenschaft(sid, "fett") if fett is None else fett
                gross = self.stile.eigenschaft(sid, "gross") if gross is None else gross
        if fett is None:
            fett = self.stile.eigenschaft(a.stil, "fett") if a.stil else None
        if gross is None:
            gross = self.stile.eigenschaft(a.stil, "gross") if a.stil else None
        text = ""
        for k in r:
            n = _ln(k)
            if n in ("t", "delText"):
                text += k.text or ""
            elif n == "tab" or n == "ptab":
                text += "\t"
            elif n in ("br", "cr"):
                text += "\n"
            elif n == "noBreakHyphen":
                text += "‑"
            elif n == "softHyphen":
                continue
            elif n in ("rPr", "lastRenderedPageBreak", "fldChar", "instrText", "delInstrText", "commentReference",
                       "annotationRef", "footnoteRef", "endnoteRef", "separator", "continuationSeparator", "pgNum"):
                continue
            elif n in ("footnoteReference", "endnoteReference"):
                self.z.zaehler["fussnotenmarken"] += 1
            elif n in ("drawing", "pict", "object", "AlternateContent", "movie", "control"):
                self._grafik(k)
            elif n == "sym":
                self.z.unbekannt["sym"] += 1
                self.z.zaehler["sonderzeichen"] += 1
            else:
                self.z.unbekannt[_name_ok(n)] += 1
        if not text:
            return
        if versteckt:
            self.z.zaehler["versteckter_text_woerter"] += len(text.split())
            return
        if gross:
            text = text.upper()
        for c in text:
            a.zeichen.append((c, status))
            if not c.isspace():
                a.gesamt_zeichen += 1
                if fett:
                    a.fett_zeichen += 1

    def _grafik(self, e: ET.Element) -> None:
        """Zeichnung / Textfeld / Objekt: Text darin wird nicht gelesen; Aenderungen darin sind eine Luecke."""
        self.z.zaehler["grafiken"] += 1
        felder = _textfelder(e)
        if felder:
            self.z.zaehler["textfelder"] += len(felder)
            for f in felder:
                geaendert = _geaenderte_woerter(f)
                if geaendert:
                    self.z.zaehler["textfeld_woerter_geaendert"] += geaendert
                self.z.zaehler["textfeld_woerter"] += sum(len((t.text or "").split()) for t in f.iter() if _ln(t) in ("t", "delText"))

    def _laeufe(self, elem: ET.Element, status: str, a: _Abs, ctx: _Kontext) -> None:
        for k in elem:
            n = _ln(k)
            if n == "r":
                if status != _NETTO:
                    self._lauf_text(k, status, a)
                else:
                    self.z.zaehler["netto_null"] += 1
            elif n in ("ins", "moveTo", "del", "moveFrom"):
                neu = ADD if n in ("ins", "moveTo") else DEL
                self.z.zaehler[{"ins": "ins", "moveTo": "move_to", "del": "del", "moveFrom": "move_from"}[n]] += 1
                autor = _attr(k, "author")
                if autor:
                    self.z.autoren.add(autor)
                if status == ADD and neu == DEL:
                    st = _NETTO
                elif status in (DEL, _NETTO):
                    st = status
                else:
                    st = neu
                self._laeufe(k, st, a, ctx)
            elif n in ("hyperlink", "smartTag", "fldSimple", "customXml", "bdo", "dir", "sdtContent", "sdt"):
                if n == "sdt":
                    inh = _kind(k, "sdtContent")
                    if inh is not None:
                        self._laeufe(inh, status, a, ctx)
                else:
                    self._laeufe(k, status, a, ctx)
            elif n in ("pPr", "proofErr", "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd",
                       "permStart", "permEnd", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart",
                       "moveToRangeEnd", "rPrChange", "numberingChange", "cellMerge", "noProof"):
                continue
            else:
                self.z.unbekannt[_name_ok(n)] += 1
                if any(_ln(t) in ("t", "delText") and (t.text or "").strip() for t in k.iter()):
                    self.z.zaehler["unbekannt_mit_text"] += 1

    def _absatz(self, p: ET.Element, ctx: _Kontext) -> _Abs:
        a = _Abs(behaelter=ctx.behaelter, in_tabelle=ctx.in_tabelle, zeilen_status=ctx.zeilen_status)
        ppr = _kind(p, "pPr")
        npr_neu = None
        wechsel = None
        if ppr is not None:
            ps = _kind(ppr, "pStyle")
            a.stil = (_attr(ps, "val") or "") if ps is not None else ""
            npr_neu = _kind(ppr, "numPr")
            ol = _kind(ppr, "outlineLvl")
            if ol is not None:
                a.gliederung = _ganz(_attr(ol, "val"))
            rpr = _kind(ppr, "rPr")
            if rpr is not None:
                if _kind(rpr, "ins") is not None:
                    a.marke = "ins"
                    self.z.zaehler["absatzmarke_ins"] += 1
                elif _kind(rpr, "del") is not None:
                    a.marke = "del"
                    self.z.zaehler["absatzmarke_del"] += 1
                if _kind(rpr, "moveFrom") is not None or _kind(rpr, "moveTo") is not None:
                    self.z.zaehler["absatzmarke_move"] += 1
            wechsel = _kind(ppr, "pPrChange")
        if a.gliederung is None and a.stil:
            a.gliederung = self.stile.eigenschaft(a.stil, "gliederung")
        a.num_neu = self._numpr(npr_neu, a.stil)
        if wechsel is not None:                                          # Formatwechsel am Absatz: das Alte steht im pPrChange
            self.z.zaehler["format_absatz"] += 1
            alt_ppr = _kind(wechsel, "pPr")
            alt_stil, npr_alt = a.stil, None
            if alt_ppr is not None:
                ps2 = _kind(alt_ppr, "pStyle")
                if ps2 is not None:
                    alt_stil = _attr(ps2, "val") or ""
                npr_alt = _kind(alt_ppr, "numPr")
            a.num_alt = self._numpr(npr_alt, alt_stil)
        else:
            a.num_alt = a.num_neu
        self._laeufe(p, KEEP, a, ctx)
        if ctx.zeilen_status:                                            # ganze Tabellenzeile eingefuegt / gestrichen
            a.zeichen = [(c, ctx.zeilen_status if s == KEEP else s) for c, s in a.zeichen]
        self.absaetze.append(a)
        return a

    def _numpr(self, npr: ET.Element | None, sid: str) -> tuple[int, int] | None:
        num_id = ilvl = None
        if npr is not None:
            nn = _kind(npr, "numId")
            il = _kind(npr, "ilvl")
            num_id = _ganz(_attr(nn, "val")) if nn is not None else None
            ilvl = _ganz(_attr(il, "val")) if il is not None else None
        if num_id is None:
            num_id = self.stile.eigenschaft(sid, "num_id") if sid else None
        if num_id is None or num_id == 0:
            return None
        if ilvl is None:
            ilvl = self.stile.eigenschaft(sid, "ilvl") if sid else None
        if ilvl is None:
            ilvl = self.nummern.ebene_von_stil(num_id, sid) if sid else None
        return (num_id, ilvl or 0)

    # ------------------------------------------------------------ Koerper und Tabellen
    def _koerper(self, elem: ET.Element, ctx: _Kontext) -> None:
        for k in elem:
            n = _ln(k)
            if n == "p":
                self._absatz(k, ctx)
            elif n == "tbl":
                self._tabelle(k, ctx)
            elif n == "sdt":
                inh = _kind(k, "sdtContent")
                if inh is not None:
                    self._koerper(inh, ctx)
            elif n in ("sdtContent", "customXml", "ins", "del", "moveFrom", "moveTo"):
                self._koerper(k, ctx)
            elif n in ("sectPr", "bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd", "commentRangeStart",
                       "commentRangeEnd", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
                       "tblPr", "tblGrid", "trPr", "tcPr", "tblPrEx", "sdtPr", "sdtEndPr"):
                continue
            else:
                self.z.unbekannt[_name_ok(n)] += 1
                if any(_ln(t) in ("t", "delText") and (t.text or "").strip() for t in k.iter()):
                    self.z.zaehler["unbekannt_mit_text"] += 1

    def _tabelle(self, tbl: ET.Element, ctx: _Kontext) -> None:
        self.z.zaehler["tabellen"] += 1
        pr = _kind(tbl, "tblPr")
        if pr is not None and _kind(pr, "tblPrChange") is not None:
            self.z.zaehler["format_tabelle"] += 1
        for tr in tbl:
            if _ln(tr) != "tr":
                continue
            zs = ""
            trpr = _kind(tr, "trPr")
            if trpr is not None:
                if _kind(trpr, "ins") is not None:
                    zs = ADD
                    self.z.zaehler["zeilen_ins"] += 1
                elif _kind(trpr, "del") is not None:
                    zs = DEL
                    self.z.zaehler["zeilen_del"] += 1
                if _kind(trpr, "trPrChange") is not None:
                    self.z.zaehler["format_tabelle"] += 1
            self.z.zaehler["tabellenzeilen"] += 1
            for tc in tr:
                if _ln(tc) != "tc":
                    continue
                tcpr = _kind(tc, "tcPr")
                if tcpr is not None:
                    for nm in ("cellIns", "cellDel", "cellMerge"):
                        if _kind(tcpr, nm) is not None:
                            self.z.zaehler["zellen_" + nm.lower()] += 1
                    if _kind(tcpr, "tcPrChange") is not None:
                        self.z.zaehler["format_tabelle"] += 1
                self.zellen += 1
                self.z.zaehler["zellen"] += 1
                self._koerper(tc, _Kontext(behaelter=self.zellen, in_tabelle=True, zeilen_status=zs or ctx.zeilen_status))


def _name_ok(n: str) -> str:
    return n if re.fullmatch(r"[A-Za-z][A-Za-z0-9]{0,39}", n or "") else "andere"


def _textfelder(e: ET.Element) -> list[ET.Element]:
    """w:txbxContent in einer Zeichnung; bei mc:AlternateContent zaehlt nur mc:Choice (Fallback ist dieselbe Box nochmal)."""
    out: list[ET.Element] = []

    def gehe(x: ET.Element) -> None:
        for k in x:
            n = _ln(k)
            if n == "Fallback":
                continue
            if n == "txbxContent":
                out.append(k)
                continue
            gehe(k)

    gehe(e)
    return out


def _geaenderte_woerter(e: ET.Element) -> int:
    n = 0
    for k in e.iter():
        if _ln(k) in ("ins", "del", "moveFrom", "moveTo"):
            n += sum(len((t.text or "").split()) for t in k.iter() if _ln(t) in ("t", "delText"))
    return n


# ---------------------------------------------------------------- Woerter (wie der PDF-Leser)
def _trenner_von(ws: list[tuple[str, str]]) -> tuple[str, str]:
    """(Trenner-Text, Status) eines Leerraums zwischen zwei Woertern. Status nur, wenn alle Zeichen ihn teilen."""
    texte = "".join(c for c, _ in ws)
    if any(c in texte for c in _GESCHUETZT):
        t = "\xa0"
    elif "\t" in texte:
        t = "\t"
    else:
        t = " "
    sts = {s for _, s in ws}
    return t, (sts.pop() if len(sts) == 1 else KEEP)


def _teile(g: list[tuple[str, str]]) -> list[tuple[str, str]]:
    if len({s for _, s in g}) <= 1:
        return []
    out: list[tuple[str, str]] = []
    for c, s in g:
        if out and out[-1][0] == s:
            out[-1] = (s, out[-1][1] + c)
        else:
            out.append((s, c))
    return out


def _tokens(zeichen: list[tuple[str, str]]) -> list[PdfToken]:
    """Zeichen mit Status -> Woerter (PdfToken) wie der PDF-Leser: Wort = Folge ohne Leerraum; Zeichen-Laeufe (`teile`) bei
    Zeichen-genauen Aenderungen; markierter Leerraum = `trenner_art`, oder (eingefuegtes Leerzeichen zwischen zwei alten
    Woertern bzw. gestrichenes zwischen zwei neuen) ein zusammengelegtes Wort mit dem Leerraum als Lauf."""
    woerter: list[list[tuple[str, str]]] = []
    seps: list[tuple[str, str]] = []
    akt: list[tuple[str, str]] = []
    ws: list[tuple[str, str]] = []
    for c, s in zeichen:
        if c.isspace():
            if akt:
                woerter.append(akt)
                akt = []
            ws.append((c, s))
            continue
        if ws:
            if woerter:
                seps.append(_trenner_von(ws))
            ws = []
        akt.append((c, s))
    if akt:
        woerter.append(akt)
    seps = seps[:len(woerter) - 1] + [(" ", KEEP)] if woerter else []
    tokens: list[PdfToken] = []
    for g, (sep_text, _) in zip(woerter, seps):
        st = [s for _, s in g]
        alt = "".join(c for c, s in g if s in (KEEP, DEL))
        neu = "".join(c for c, s in g if s in (KEEP, ADD))
        status = KEEP if all(s == KEEP for s in st) else DEL if all(s == DEL for s in st) else ADD if all(s == ADD for s in st) else MOD
        tokens.append(PdfToken(old=alt, new=neu, status=status, page=0, bbox=(0.0, 0.0, 0.0, 0.0), trenner=sep_text, teile=_teile(g)))
    _kleben(tokens, [st for _, st in seps], [t for t, _ in seps])
    return tokens


def _kleben(tokens: list[PdfToken], sep_status: list[str], sep_text: list[str]) -> None:
    """Wie pdf_reader._leerraum_kleben: markierter Leerraum gehoert nur zur alten bzw. neuen Fassung (siehe dort)."""
    i = 0
    while i < len(tokens) - 1:
        s = sep_status[i]
        l, r = tokens[i], tokens[i + 1]
        if s not in (ADD, DEL):
            i += 1
            continue
        if sep_text[i] != "\t" and ((s == ADD and l.old and r.old) or (s == DEL and l.new and r.new)):
            text = "\xa0" if l.trenner == "\xa0" else " "
            teile: list[tuple[str, str]] = []
            for s_, tx in _laeufe(l) + [(s, text)] + _laeufe(r):
                if teile and teile[-1][0] == s_:
                    teile[-1] = (s_, teile[-1][1] + tx)
                elif tx:
                    teile.append((s_, tx))
            neu = PdfToken(old="".join(tx for s_, tx in teile if s_ in (KEEP, DEL)),
                           new="".join(tx for s_, tx in teile if s_ in (KEEP, ADD)), status=MOD, page=0,
                           bbox=(0.0, 0.0, 0.0, 0.0), trenner=r.trenner, teile=teile)
            if getattr(r, "trenner_art", ""):
                neu.trenner_art = r.trenner_art          # type: ignore[attr-defined]
            tokens[i:i + 2] = [neu]
            del sep_status[i], sep_text[i]
            continue
        l.trenner_art = "a" if s == ADD else "d"         # type: ignore[attr-defined]
        i += 1


# ---------------------------------------------------------------- Nummern und Absatzgrenzen je Behaelter
def _existenz(liste: list[_Abs]) -> None:
    """Wo gibt es einen Absatz im Original und im Endstand? Absatzmarken-Aenderungen verbinden bzw. trennen Absaetze:
    Original: Absaetze, die durch eingefuegte Marken verbunden sind, waren EIN Absatz (der erste mit altem Inhalt zaehlt;
    mehrere = ein getrennter Absatz). Endstand: durch gestrichene Marken verbundene Absaetze sind EIN Absatz."""
    n = len(liste)
    i = 0
    while i < n:
        j = i
        while j < n - 1 and liste[j].marke == "ins":
            j += 1
        grp = liste[i:j + 1]
        mit = [a for a in grp if a.hat_alt]
        if mit:
            rep: _Abs | None = mit[0]
            if len(mit) > 1:
                for a in mit:
                    a.absatzgrenze = "getrennt"
        elif any(a.hat_neu for a in grp):
            rep = grp[-1] if (not grp[-1].hat_neu and grp[-1].marke != "ins") else None
        else:
            rep = grp[-1]
        for a in grp:
            a.ex_alt = a is rep
        i = j + 1
    i = 0
    while i < n:
        j = i
        while j < n - 1 and liste[j].marke == "del":
            j += 1
        grp = liste[i:j + 1]
        mit = [a for a in grp if a.hat_neu]
        if mit:
            rep = mit[0]
            if len(mit) > 1:
                for a in mit:
                    a.absatzgrenze = "verbunden"
        elif any(a.hat_alt for a in grp):
            rep = grp[-1] if (not grp[-1].hat_alt and grp[-1].marke != "del") else None
        else:
            rep = grp[-1]
        for a in grp:
            a.ex_neu = a is rep
        i = j + 1


def _existenz_alle(absaetze: list[_Abs]) -> None:
    je: dict[int, list[_Abs]] = {}
    for a in absaetze:
        je.setdefault(a.behaelter, []).append(a)
    for liste in je.values():
        _existenz(liste)


def _marken_zeichen(a: _Abs) -> list[tuple[str, str]]:
    """Zeichen der Listennummer (und des Trennzeichens dahinter) mit Status: gleich = KEEP, sonst alt (DEL) und neu (ADD)."""
    ma, mn = a.marke_alt, a.marke_neu
    if ma is None and mn is None:
        return []
    suff = (mn or ma).suff                                   # type: ignore[union-attr]
    if ma is not None and mn is not None:
        if ma.text == mn.text:
            z = [(c, KEEP) for c in ma.text]
        else:
            z = _marken_diff(ma.text, mn.text)
        sst = KEEP
    elif ma is not None:
        z = [(c, DEL) for c in ma.text]
        sst = DEL
    else:
        z = [(c, ADD) for c in mn.text]                      # type: ignore[union-attr]
        sst = ADD
    return z + [(c, sst) for c in suff]


def _marken_diff(alt: str, neu: str) -> list[tuple[str, str]]:
    """Alte und neue Nummer Wort fuer Wort: gleiche Woerter bleiben, abweichende sind (alt gestrichen, neu eingefuegt)
    direkt hintereinander ("c." + "a." = ein Label-Wort wie im PDF)."""
    wa = re.split(r"(\s+)", alt)
    wn = re.split(r"(\s+)", neu)
    if len(wa) != len(wn):
        return [(c, DEL) for c in alt] + [(c, ADD) for c in neu]
    out: list[tuple[str, str]] = []
    for x, y in zip(wa, wn):
        if x == y:
            out += [(c, KEEP) for c in x]
        else:
            out += [(c, DEL) for c in x] + [(c, ADD) for c in y]
    return out


# ---------------------------------------------------------------- Hauptfunktion
_HINWEIS_ABSATZ = {
    "getrennt": "In der Word-Datei wurde ein Absatz getrennt (eine Absatzmarke wurde eingefügt). In der Excel steht er noch "
                "als ein Absatz. Bitte prüfen.",
    "verbunden": "In der Word-Datei wurden zwei Absätze verbunden (eine Absatzmarke wurde gestrichen). In der Excel stehen "
                 "sie noch getrennt. Bitte prüfen.",
}


def analysiere_docx(pfad: str | Path, cfg: dict | None = None) -> PdfAnalyse:
    """Word-Datei (.docx) lesen -> PdfAnalyse (siehe Modul-Docstring). Wirft NutzerFehler mit einfacher deutscher Meldung bei
    kaputter, alter (.doc) oder passwortgeschuetzter Datei."""
    return _analysiere(pfad, cfg)[0]


def docx_zahlen(pfad: str | Path, cfg: dict | None = None) -> tuple[PdfAnalyse, DocxZahlen]:
    """Wie `analysiere_docx`, dazu die Zahlen und Formen der Datei (fuer den Steckbrief; kein Text)."""
    return _analysiere(pfad, cfg)


def _analysiere(pfad: str | Path, cfg: dict | None) -> tuple[PdfAnalyse, DocxZahlen]:
    dcfg = _dcfg(cfg)
    paket = _oeffne(pfad, dcfg)
    try:
        return _lies(paket, Path(pfad), cfg)
    finally:
        paket.zf.close()


def _lies(paket: _Paket, pfad: Path, cfg: dict | None) -> tuple[PdfAnalyse, DocxZahlen]:
    leser = _Leser(paket, cfg)
    wurzel = paket.xml(paket.haupt)
    if wurzel is None:
        raise _fehler("Die Datei enthält keinen Word-Text (word/document.xml fehlt).")
    body = _kind(wurzel, "body")
    if body is None:
        raise _fehler("Der Text der Word-Datei fehlt oder ist beschädigt (kein Textbereich gefunden).")
    leser._koerper(body, _Kontext())
    lay = leser.lay
    z = leser.z
    ergebnis = PdfAnalyse(pfad=str(pfad), seiten=0, blocks=[], legende_modus="docx",
                          legende_grund="Word-Datei: die Änderungen stehen im Dokument selbst (Änderungsverfolgung).",
                          markup_leiste=False)
    # 1. Existenz im Original / Endstand und Listennummern (je Behaelter getrennt; Tabellenzellen eigene Gruppen)
    _existenz_alle(leser.absaetze)
    za, zn = _Zaehler(leser.nummern, leser.dcfg.get("listen", {})), _Zaehler(leser.nummern, leser.dcfg.get("listen", {}))
    for a in leser.absaetze:
        if a.num_alt is not None and a.ex_alt:
            a.marke_alt = za.weiter(*a.num_alt)
        if a.num_neu is not None and a.ex_neu:
            a.marke_neu = zn.weiter(*a.num_neu)
        for m in (a.marke_alt, a.marke_neu):
            if m is not None:
                z.listenarten[m.fmt] += 1
        if a.marke_neu is not None and a.num_neu is not None:
            z.listenebenen[str(a.num_neu[1] + 1)] += 1
        if a.marke_alt is not None and a.marke_neu is not None and a.marke_alt.text != a.marke_neu.text:
            z.zaehler["nummern_geaendert"] += 1
    z.unbekannte_zahlformate.update(za.unbekannt)
    z.unbekannte_zahlformate.update(zn.unbekannt)
    # 2. Woerter
    for a in leser.absaetze:
        if not a.ex_alt and not a.ex_neu and not a.zeichen:
            continue
        a.tokens = _tokens(_marken_zeichen(a) + a.zeichen)
    # 3. Bloecke (eine Tabellenzelle = ein Block)
    sec_alt: str | None = None
    sec_neu: str | None = None
    idx = 0
    gesehen_zelle: dict[int, PdfBlock] = {}
    block_von_absatz: dict[int, PdfBlock] = {}
    getrennt = verbunden = 0
    ueber_re = leser._stil_re
    max_ebene = int(leser.dcfg.get("ueberschrift_max_ebene", 2))
    max_w = int(leser.dcfg.get("ueberschrift_max_woerter", 25))
    for a in leser.absaetze:
        toks = a.tokens
        if not toks:
            continue
        if a.absatzgrenze:
            for t in toks:
                t.hinweis = ";".join(sorted(set(filter(None, t.hinweis.split(";"))) | {"absatzgrenze_unklar"}))
            if a.absatzgrenze == "getrennt":
                getrennt += 1
            else:
                verbunden += 1
        if a.in_tabelle:
            b = gesehen_zelle.get(a.behaelter)
            if b is None:
                b = PdfBlock(idx=idx, page=0, kind="absatz", tokens=[], bbox=(0.0, 0.0, 0.0, 0.0), section_alt=sec_alt,
                             section_neu=sec_neu)
                gesehen_zelle[a.behaelter] = b
                ergebnis.blocks.append(b)
                idx += 1
            b.tokens.extend(toks)
            block_von_absatz[id(a)] = b
            continue
        kopf = _ist_kopf(a, toks, lay, ueber_re, max_ebene, max_w, leser)
        if kopf:
            n_alt, n_neu = _section_von(toks, lay)
            if n_alt or n_neu:
                sec_alt, sec_neu = n_alt, n_neu
            else:
                sec_alt = sec_neu = None
        blk = PdfBlock(idx=idx, page=0, kind="ueberschrift" if kopf else "absatz", tokens=toks,
                       bbox=(0.0, 0.0, 0.0, 0.0), section_alt=sec_alt, section_neu=sec_neu)
        ergebnis.blocks.append(blk)
        block_von_absatz[id(a)] = blk
        idx += 1
    zellen_ids = {id(b) for b in gesehen_zelle.values()}
    for b in ergebnis.blocks:
        if id(b) in zellen_ids:
            for t in b.tokens:
                if t.status != KEEP:
                    t.hinweis = ";".join(sorted(set(filter(None, t.hinweis.split(";"))) | {"tabelle"}))
        b.hinweise = _block_hinweise(b.tokens)
    for a in leser.absaetze:
        b = block_von_absatz.get(id(a))
        if a.absatzgrenze and b is not None and _HINWEIS_ABSATZ[a.absatzgrenze] not in b.hinweise:
            b.hinweise.append(_HINWEIS_ABSATZ[a.absatzgrenze])
    # 4. Kopf- und Fusszeilen, Kommentare, Fussnoten
    _kopf_fuss(paket, leser, ergebnis)
    _kommentare(paket, leser, ergebnis)
    _fussnoten(paket, leser)
    # 5. Meldungen
    _meldungen(leser, ergebnis, getrennt, verbunden)
    z.zaehler["absaetze"] = sum(1 for a in leser.absaetze if a.tokens)
    z.zaehler["ueberschriften"] = sum(1 for b in ergebnis.blocks if b.kind == "ueberschrift")
    z.zaehler["bloecke"] = len(ergebnis.blocks)
    return ergebnis, z


def _ist_kopf(a: _Abs, toks: list[PdfToken], lay: dict, ueber_re: re.Pattern, max_ebene: int, max_w: int, leser: _Leser) -> bool:
    """Ueberschrift? Formatvorlage (Ebene 1 und 2), sonst dieselben Muster wie im PDF-Leser ("1.1 TITEL", "PART 1 TITEL",
    Nummer + fett)."""
    if len(toks) > max_w or len(toks) < 1:
        return False
    if a.stil:
        st = leser.stile.stile.get(a.stil)
        namen = [a.stil.replace(" ", "").lower()] + ([st.name.replace(" ", "").lower()] if st and st.name else [])
        for nm in namen:
            m = ueber_re.match(nm)
            if m and int(m.group(1)) <= max_ebene:
                return True
    if len(toks) < 2:
        return False
    t0, t1 = toks[0], toks[1]
    for text in {f"{t0.old} {t1.old}".strip(), f"{t0.new} {t1.new}".strip()}:
        if _kapitel(text, lay):
            return True
    rx = _re_cfg(lay, "section_nummer")
    hat_nr = bool(rx) and any(rx.match(x) for x in (t0.old, t0.new) if x)                 # type: ignore[union-attr]
    if hat_nr and (t0.trenner == "\t" or _gross_anteil(toks[1:]) >= lay.get("section_min_gross_anteil", 0.6)):
        return True
    nr = (section_nummer(t0.old) if t0.old else None) or (section_nummer(t0.new) if t0.new else None)
    fett = a.gesamt_zeichen > 0 and a.fett_zeichen / a.gesamt_zeichen >= float(leser.dcfg.get("fett_anteil", 0.8))
    return bool(nr and fett)


def _kopf_fuss(paket: _Paket, leser: _Leser, ergebnis: PdfAnalyse) -> None:
    kopf_re = re.compile(str(leser.dcfg.get("kopf_teile", r"^word/header\d*\.xml$")))
    fuss_re = re.compile(str(leser.dcfg.get("fuss_teile", r"^word/footer\d*\.xml$")))
    teile = [(0, n, "kopf") for n in paket.namen if kopf_re.match(n)] + [(1, n, "fuss") for n in paket.namen if fuss_re.match(n)]
    for _, name, art in sorted(teile):
        wurzel = paket.xml(name)
        if wurzel is None:
            continue
        teil = _Leser(paket, None, teilen=leser)
        teil._koerper(wurzel, _Kontext(behaelter=-1))
        for k, v in teil.z.zaehler.items():
            if k in ("ins", "del", "move_from", "move_to", "format_lauf", "format_absatz", "absatzmarke_ins", "absatzmarke_del"):
                leser.z.zaehler["kopf_fuss_" + k] += v
            elif k in ("textfelder", "textfeld_woerter", "textfeld_woerter_geaendert", "grafiken", "unbekannt_mit_text",
                       "fussnotenmarken", "versteckter_text_woerter", "sonderzeichen", "netto_null"):
                leser.z.zaehler[k] += v
        leser.z.unbekannt.update(teil.z.unbekannt)
        leser.z.autoren |= teil.z.autoren
        leser.z.zaehler["kopf_fuss_teile"] += 1
        leser.z.zaehler["kopf_teile" if art == "kopf" else "fuss_teile"] += 1
        # Existenz: Absatzmarken im Kopf/Fuss wie im Text auswerten (Gruppen), aber ohne Nummerierung
        _existenz_alle(teil.absaetze)
        for a in teil.absaetze:
            if not a.ex_alt and not a.ex_neu and not a.zeichen:
                continue
            toks = _tokens(a.zeichen)
            if not toks:
                continue
            leser.z.zaehler["kopf_fuss_absaetze"] += 1
            ergebnis.ignoriert.append({"seite": 0, "grund": "kopf_fuss", "woerter": sum(1 for t in toks if t.status != KEEP),
                                       "text": " ".join(t.new or t.old for t in toks)})


def _kommentare(paket: _Paket, leser: _Leser, ergebnis: PdfAnalyse) -> None:
    wurzel = paket.xml(str(leser.dcfg.get("kommentar_teil", "word/comments.xml")))
    if wurzel is None:
        return
    maxz = int(leser.dcfg.get("kommentar_max_zeichen", 160))
    for c in wurzel:
        if _ln(c) != "comment":
            continue
        absaetze = []
        for p in c.iter():
            if _ln(p) == "p":
                absaetze.append("".join((t.text or "") for t in p.iter() if _ln(t) == "t"))
        text = " ".join(" ".join(x.split()) for x in absaetze if x.strip())
        ergebnis.randnotizen.append({"seite": 0, "art": "kommentar", "text": text[:maxz], "y": 0.0})
        leser.z.zaehler["kommentare"] += 1


def _fussnoten(paket: _Paket, leser: _Leser) -> None:
    for schluessel, zaehl in (("fussnoten_teil", "fussnoten"), ("endnoten_teil", "endnoten")):
        wurzel = paket.xml(str(leser.dcfg.get(schluessel, "")))
        if wurzel is None:
            continue
        for fn in wurzel:
            if _ln(fn) not in ("footnote", "endnote"):
                continue
            if (_attr(fn, "type") or "normal") in ("separator", "continuationSeparator", "continuationNotice"):
                continue
            text = [t.text for t in fn.iter() if _ln(t) in ("t", "delText") and (t.text or "").strip()]
            if not text:
                continue
            leser.z.zaehler[zaehl] += 1
            leser.z.zaehler[zaehl + "_woerter"] += sum(len(x.split()) for x in text)
            leser.z.zaehler[zaehl + "_woerter_geaendert"] += _geaenderte_woerter(fn)


def _meldungen(leser: _Leser, ergebnis: PdfAnalyse, getrennt: int, verbunden: int) -> None:
    z = leser.z.zaehler
    w = ergebnis.warnungen
    luecken = ergebnis.luecken
    if getrennt or verbunden:
        teile = []
        if getrennt:
            teile.append(f"{getrennt} Absätze getrennt")
        if verbunden:
            teile.append(f"{verbunden} Absätze verbunden")
        satz = (f"Absatzmarken wurden geändert ({', '.join(teile)}). Betroffene Absätze werden nie automatisch eingetragen; "
                f"ob die Excel eine Zeile mehr oder weniger braucht, muss geprüft werden.")
        w.append(satz)
        luecken.append(luecke("sonstiges", 0, "Absätze wurden in Word getrennt oder verbunden. Die Excel-Zeilen müssen von Hand "
                                              "angepasst werden.", 0))
    n_format = z["format_lauf"] + z["format_absatz"] + z["format_tabelle"]
    if n_format:
        w.append(f"{n_format} reine Formatänderungen (Schrift, Absatzformat, Tabellenformat) gefunden. Sie sind keine "
                 f"Textänderung und werden nicht übernommen.")
    if z["kommentare"]:
        w.append(f"Das Dokument enthält {z['kommentare']} Kommentare. Sie gehören nicht zum Text und wurden nicht als Änderung "
                 f"gelesen; sie stehen als Randnotizen in der Analyse.")
    for schluessel, name in (("fussnoten", "Fußnoten"), ("endnoten", "Endnoten")):
        if z[schluessel]:
            w.append(f"Das Dokument enthält {z[schluessel]} {name} ({z[schluessel + '_woerter']} Wörter). Ihr Text wird nicht "
                     f"gelesen.")
            if z[schluessel + "_woerter_geaendert"]:
                luecken.append(luecke("sonstiges", 0, f"In den {name} stehen Änderungen, die nicht gelesen werden.",
                                      z[schluessel + "_woerter_geaendert"]))
    if z["fussnotenmarken"]:
        w.append(f"{z['fussnotenmarken']} Fußnoten-/Endnotenmarken stehen im Text; sie wurden nicht als Wort gelesen.")
    if z["textfelder"]:
        w.append(f"Das Dokument enthält {z['textfelder']} Textfelder (Text in Rahmen oder Formen, {z['textfeld_woerter']} Wörter). "
                 f"Dieser Text wird nicht gelesen.")
        if z["textfeld_woerter_geaendert"]:
            luecken.append(luecke("sonstiges", 0, "In Textfeldern stehen Änderungen, die nicht gelesen werden.",
                                  z["textfeld_woerter_geaendert"]))
    elif z["grafiken"]:
        w.append(f"Das Dokument enthält {z['grafiken']} Bilder oder Zeichnungen. Sie werden nicht ausgewertet.")
    if leser.z.unbekannt:
        n = sum(leser.z.unbekannt.values())
        w.append(f"{n} unbekannte Bestandteile (z. B. Formeln, Sonderzeichen oder eingebettete Objekte) wurden nicht gelesen.")
        if z["unbekannt_mit_text"] or leser.z.unbekannt.get("sym") or leser.z.unbekannt.get("oMath") or \
                leser.z.unbekannt.get("oMathPara"):
            luecken.append(luecke("sonstiges", 0, "Unbekannte Bestandteile mit Text wurden nicht gelesen; hier kann eine "
                                                  "Änderung fehlen.", z["unbekannt_mit_text"]))
    if leser.z.unbekannte_zahlformate:
        w.append("Eine Listennummer hat ein Zahlenformat, das nicht berechnet wird; ihre Nummer fehlt im Text.")
        luecken.append(luecke("sonstiges", 0, "Listennummern in einem unbekannten Zahlenformat wurden nicht berechnet.", 0))
    if z["versteckter_text_woerter"]:
        w.append(f"{z['versteckter_text_woerter']} Wörter sind als versteckter Text formatiert und wurden nicht gelesen.")
    if z["netto_null"]:
        w.append(f"{z['netto_null']} eingefügte und danach wieder gestrichene Textstellen haben keine Wirkung und wurden "
                 f"übergangen.")
    if z["zeilen_ins"] or z["zeilen_del"]:
        w.append(f"In Tabellen wurden {z['zeilen_ins']} Zeilen eingefügt und {z['zeilen_del']} Zeilen gestrichen.")
    if not any(b.hat_aenderung for b in ergebnis.blocks) and not any(i["woerter"] for i in ergebnis.ignoriert):
        w.append("Im Dokument wurden keine Textänderungen erkannt. Ist die Änderungsverfolgung im Dokument vorhanden "
                 "(Änderungen annehmen oder ablehnen löscht sie)?")


# ---------------------------------------------------------------- Kreuzprobe mit dem Word-PDF
def _woerter_ansicht(a: PdfAnalyse, ansicht: str) -> list[str]:
    out: list[str] = []
    for b in a.blocks:
        for t in b.tokens:
            x = t.old if ansicht == "alt" else t.new
            n = norm(x) if x else ""
            if n:
                out.append(n)
    return out


def vergleiche_mit_pdf(docx_analyse: PdfAnalyse, pdf_analyse: PdfAnalyse, cfg: dict | None = None) -> dict[str, Any]:
    """Kreuzprobe als Sicherheitsnetz: Wortfolge des Word-Lesers gegen die des PDF-Lesers (alte und neue Fassung), nur Info.
    Enthaelt Zahlen und Positionen (Position und Laenge in Woertern), nie Text aus dem Dokument (Regel 16)."""
    maxb = int(_dcfg(cfg).get("beispiele_max", 20))
    out: dict[str, Any] = {"markierte_woerter": {"docx": docx_analyse.markierte_woerter, "pdf": pdf_analyse.markierte_woerter},
                           "bloecke": {"docx": len(docx_analyse.blocks), "pdf": len(pdf_analyse.blocks)}}
    gleich = True
    for ansicht in ("alt", "neu"):
        d, p = _woerter_ansicht(docx_analyse, ansicht), _woerter_ansicht(pdf_analyse, ansicht)
        sm = SequenceMatcher(None, d, p, autojunk=False)
        nur_d = nur_p = 0
        beispiele: list[dict[str, Any]] = []
        for op, i1, i2, j1, j2 in sm.get_opcodes():
            if op == "equal":
                continue
            nur_d += i2 - i1
            nur_p += j2 - j1
            if len(beispiele) < maxb:
                beispiele.append({"position_docx": i1, "laenge_docx": i2 - i1, "position_pdf": j1, "laenge_pdf": j2 - j1})
        gleich = gleich and not (nur_d or nur_p)
        out[ansicht] = {"woerter_docx": len(d), "woerter_pdf": len(p), "nur_docx": nur_d, "nur_pdf": nur_p,
                        "aehnlichkeit": round(sm.ratio(), 4), "beispiele": beispiele}
    out["gleich"] = gleich
    return out
