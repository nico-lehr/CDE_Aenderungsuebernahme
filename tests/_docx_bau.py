"""Bauhilfe fuer Test-.docx mit Word-Aenderungsverfolgung (Auftrag DOCX). Nur erfundene Daten (Regel 16).

Die Dateien entstehen zur Laufzeit im tmp_path (*.docx ist gitignoriert). Das Dokument-XML wird von Hand geschrieben, in ein
zip gepackt und hat die Formen, die Word erzeugt: w:ins / w:del (w:delText) / w:moveFrom / w:moveTo, Absatzmarken-Aenderungen
(w:pPr/w:rPr/w:ins bzw. w:del), rPrChange / pPrChange, Tabellen, Kopf-/Fusszeilen, Kommentare, numbering.xml, styles.xml.

Kurzschrift fuer Absatztext ("Markup", wie in tests/test_register.py):
    {+neu+}  eingefuegt        {-alt-}  gestrichen        {>x>}  verschoben nach (moveTo)    {<x<}  verschoben von (moveFrom)
    {f:text} Text mit reiner Formataenderung (rPrChange, keine Textaenderung)
    \\t      Tabulator (w:tab)
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_KOPF = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_AUTOR = 'w:author="Erfunden" w:date="2026-01-01T00:00:00Z"'

_MARKUP = re.compile(r"\{\+(.+?)\+\}|\{-(.+?)-\}|\{>(.+?)>\}|\{<(.+?)<\}|\{f:(.+?)\}|([^{]+)", re.S)


def _esc(t: str) -> str:
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


class DocxBau:
    """Sammelt Teile und schreibt das Paket."""

    def __init__(self) -> None:
        self.koerper: list[str] = []
        self.kopf: list[str] = []
        self.fuss: list[str] = []
        self.kommentare: list[str] = []
        self.fussnoten: list[str] = []
        self.numbering: str = ""
        self.styles: str = ""
        self._id = 100

    # ------------------------------------------------------------------ Laeufe
    def _nr(self) -> int:
        self._id += 1
        return self._id

    @staticmethod
    def lauf_xml(text: str, entfaellt: bool = False, rpr: str = "") -> str:
        tag = "w:delText" if entfaellt else "w:t"
        teile = text.split("\t")
        inner = "<w:tab/>".join(f'<{tag} xml:space="preserve">{_esc(t)}</{tag}>' if t else "" for t in teile)
        return f"<w:r>{('<w:rPr>' + rpr + '</w:rPr>') if rpr else ''}{inner}</w:r>"

    def markup(self, text: str, rpr: str = "") -> str:
        """Kurzschrift -> Laeufe (XML)."""
        out = []
        for m in _MARKUP.finditer(text):
            ins, dele, nach, von, fmt, norm = m.groups()
            if ins is not None:
                out.append(f'<w:ins w:id="{self._nr()}" {_AUTOR}>{self.lauf_xml(ins, False, rpr)}</w:ins>')
            elif dele is not None:
                out.append(f'<w:del w:id="{self._nr()}" {_AUTOR}>{self.lauf_xml(dele, True, rpr)}</w:del>')
            elif nach is not None:
                out.append(f'<w:moveTo w:id="{self._nr()}" {_AUTOR}>{self.lauf_xml(nach, False, rpr)}</w:moveTo>')
            elif von is not None:
                out.append(f'<w:moveFrom w:id="{self._nr()}" {_AUTOR}>{self.lauf_xml(von, True, rpr)}</w:moveFrom>')
            elif fmt is not None:
                wechsel = f'<w:rPrChange w:id="{self._nr()}" {_AUTOR}><w:rPr/></w:rPrChange>'
                out.append(self.lauf_xml(fmt, False, rpr + "<w:b/>" + wechsel))
            elif norm:
                out.append(self.lauf_xml(norm, False, rpr))
        return "".join(out)

    # ------------------------------------------------------------------ Absaetze
    def absatz_xml(self, markup: str = "", stil: str = "", num: tuple[int, int] | None = None, marke: str = "",
                   ppr_extra: str = "", rpr: str = "", roh: str = "") -> str:
        ppr = ""
        if stil:
            ppr += f'<w:pStyle w:val="{stil}"/>'
        if num is not None:
            ppr += f'<w:numPr><w:ilvl w:val="{num[1]}"/><w:numId w:val="{num[0]}"/></w:numPr>'
        ppr += ppr_extra
        if marke in ("ins", "del"):
            ppr += f'<w:rPr><w:{marke} w:id="{self._nr()}" {_AUTOR}/></w:rPr>'
        return f"<w:p>{('<w:pPr>' + ppr + '</w:pPr>') if ppr else ''}{self.markup(markup, rpr)}{roh}</w:p>"

    def absatz(self, markup: str = "", **kw) -> "DocxBau":
        self.koerper.append(self.absatz_xml(markup, **kw))
        return self

    def roh(self, xml: str) -> "DocxBau":
        self.koerper.append(xml)
        return self

    def tabelle(self, zeilen: list[list[str]], zeilen_status: dict[int, str] | None = None) -> "DocxBau":
        zeilen_status = zeilen_status or {}
        rows = []
        for i, z in enumerate(zeilen):
            trpr = ""
            if zeilen_status.get(i) in ("ins", "del"):
                trpr = f'<w:trPr><w:{zeilen_status[i]} w:id="{self._nr()}" {_AUTOR}/></w:trPr>'
            zellen = "".join(f"<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>{self.absatz_xml(t)}</w:tc>" for t in z)
            rows.append(f"<w:tr>{trpr}{zellen}</w:tr>")
        self.koerper.append(f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid/>{"".join(rows)}</w:tbl>')
        return self

    # ------------------------------------------------------------------ Kopf, Fuss, Kommentare, Fussnoten
    def kopfzeile(self, markup: str) -> "DocxBau":
        self.kopf.append(self.absatz_xml(markup))
        return self

    def fusszeile(self, markup: str) -> "DocxBau":
        self.fuss.append(self.absatz_xml(markup))
        return self

    def kommentar(self, text: str) -> "DocxBau":
        i = len(self.kommentare)
        self.kommentare.append(f'<w:comment w:id="{i}" {_AUTOR}><w:p><w:r><w:t>{_esc(text)}</w:t></w:r></w:p></w:comment>')
        return self

    def fussnote(self, markup: str) -> "DocxBau":
        i = len(self.fussnoten) + 1
        self.fussnoten.append(f'<w:footnote w:id="{i}"><w:p>{self.markup(markup)}</w:p></w:footnote>')
        return self

    # ------------------------------------------------------------------ Nummerierung und Formatvorlagen
    @staticmethod
    def lvl(ilvl: int, fmt: str, text: str, start: int = 1, suff: str = "tab", restart: int | None = None,
            pstil: str = "", lgl: bool = False) -> str:
        return (f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>'
                + (f'<w:lvlRestart w:val="{restart}"/>' if restart is not None else "")
                + (f'<w:pStyle w:val="{pstil}"/>' if pstil else "") + ('<w:isLgl/>' if lgl else "")
                + f'<w:suff w:val="{suff}"/><w:lvlText w:val="{text}"/></w:lvl>')

    def nummerierung(self, abstrakte: dict[int, list[str]], nums: dict[int, tuple[int, dict[int, int]]]) -> "DocxBau":
        """abstrakte: abstractNumId -> Ebenen (lvl-XML); nums: numId -> (abstractNumId, {ilvl: startOverride})."""
        teile = [f'<w:abstractNum w:abstractNumId="{a}">{"".join(l)}</w:abstractNum>' for a, l in abstrakte.items()]
        for n, (a, ueb) in nums.items():
            o = "".join(f'<w:lvlOverride w:ilvl="{il}"><w:startOverride w:val="{v}"/></w:lvlOverride>' for il, v in ueb.items())
            teile.append(f'<w:num w:numId="{n}"><w:abstractNumId w:val="{a}"/>{o}</w:num>')
        self.numbering = f'{_KOPF}<w:numbering xmlns:w="{W}">{"".join(teile)}</w:numbering>'
        return self

    def register_nummerierung(self) -> "DocxBau":
        """numId 1 = Gliederung der Spezifikation: A. / 1. / a. / 1) / a)."""
        lv = self.lvl
        return self.nummerierung({0: [lv(0, "upperLetter", "%1."), lv(1, "decimal", "%2."), lv(2, "lowerLetter", "%3."),
                                      lv(3, "decimal", "%4)"), lv(4, "lowerLetter", "%5)")]}, {1: (0, {})})

    def formatvorlagen(self, stile: list[dict]) -> "DocxBau":
        """stile: [{id, name, basiert, num: (numId, ilvl), fett, gross, gliederung}]"""
        xs = []
        for s in stile:
            ppr = ""
            if s.get("num"):
                ppr += f'<w:numPr><w:ilvl w:val="{s["num"][1]}"/><w:numId w:val="{s["num"][0]}"/></w:numPr>'
            if s.get("gliederung") is not None:
                ppr += f'<w:outlineLvl w:val="{s["gliederung"]}"/>'
            rpr = ("<w:b/>" if s.get("fett") else "") + ("<w:caps/>" if s.get("gross") else "")
            xs.append(f'<w:style w:type="paragraph" w:styleId="{s["id"]}"><w:name w:val="{s.get("name", s["id"])}"/>'
                      + (f'<w:basedOn w:val="{s["basiert"]}"/>' if s.get("basiert") else "")
                      + (f"<w:pPr>{ppr}</w:pPr>" if ppr else "") + (f"<w:rPr>{rpr}</w:rPr>" if rpr else "") + "</w:style>")
        self.styles = f'{_KOPF}<w:styles xmlns:w="{W}">{"".join(xs)}</w:styles>'
        return self

    # ------------------------------------------------------------------ Paket
    def dokument_xml(self) -> str:
        return f'{_KOPF}<w:document xmlns:w="{W}"><w:body>{"".join(self.koerper)}<w:sectPr/></w:body></w:document>'

    def speichern(self, pfad: Path, ohne_dokument: bool = False, extra: dict[str, bytes | str] | None = None) -> Path:
        pfad = Path(pfad)
        pfad.parent.mkdir(parents=True, exist_ok=True)
        teile: dict[str, str] = {
            "[Content_Types].xml": f'{_KOPF}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                                   '<Default Extension="xml" ContentType="application/xml"/></Types>',
            "_rels/.rels": f'{_KOPF}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                           'Target="word/document.xml"/></Relationships>'}
        if not ohne_dokument:
            teile["word/document.xml"] = self.dokument_xml()
        if self.styles:
            teile["word/styles.xml"] = self.styles
        if self.numbering:
            teile["word/numbering.xml"] = self.numbering
        if self.kopf:
            teile["word/header1.xml"] = f'{_KOPF}<w:hdr xmlns:w="{W}">{"".join(self.kopf)}</w:hdr>'
        if self.fuss:
            teile["word/footer1.xml"] = f'{_KOPF}<w:ftr xmlns:w="{W}">{"".join(self.fuss)}</w:ftr>'
        if self.kommentare:
            teile["word/comments.xml"] = f'{_KOPF}<w:comments xmlns:w="{W}">{"".join(self.kommentare)}</w:comments>'
        if self.fussnoten:
            sep = '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
            teile["word/footnotes.xml"] = f'{_KOPF}<w:footnotes xmlns:w="{W}">{sep}{"".join(self.fussnoten)}</w:footnotes>'
        teile.update(extra or {})
        with zipfile.ZipFile(pfad, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, inhalt in teile.items():
                zf.writestr(name, inhalt)
        return pfad


# ---------------------------------------------------------------- Register-Dokument (wie tests/test_register.py: art, label, text)
_STUFE = {"zahl.": 1, "klein.": 2, "zahl)": 3, "klein)": 4}


def _label_art(label: str) -> str:
    kern = label[:-1]
    return ("zahl" if kern.isdigit() else "klein") + label[-1]


def _roh_label(label: str) -> str:
    return re.sub(r"\{[-+](.+?)[-+]\}", r"\1", label)


def _ganz(text: str, kennung: str) -> bool:
    return text.startswith("{" + kennung) and text.endswith(kennung + "}") and "{" not in text[2:-2]


def register_docx(pfad: Path, absaetze: list[tuple], auto: bool = True, vorspann: list[str] | None = None,
                  kopf: list[str] | None = None, fuss: list[str] | None = None) -> Path:
    """Dasselbe erfundene Dokument wie `tests.test_register.wort_pdf`, als .docx.

    absaetze: (art, label_markup, text_markup) mit art part | section | sub | sub_titel | liste | liste_titel | absatz.
    auto=True: die Nummern der Unterpunkte (A. 1. a. 1) a)) kommen aus numbering.xml (der Leser rechnet sie selbst, der Label-
    Markup wird NICHT benutzt); auto=False: Nummern stehen als getippter Text mit Tabulator (der Label-Markup gilt).
    vorspann: Absaetze (Markup) vor der Spezifikation (Deckblatt, Titelzeile); kopf / fuss: Zeilen der Kopf- bzw. Fusszeile."""
    b = DocxBau()
    for z in vorspann or []:
        b.absatz(z)
    for z in kopf or []:
        b.kopfzeile(z)
    for z in fuss or []:
        b.fusszeile(z)
    anzahl = 1 + sum(1 for x in absaetze if x[0] == "section")
    if auto:                                   # je Section eine neue Liste (startOverride), wie ein Gliederungswechsel in Word
        lv = b.lvl
        b.nummerierung({0: [lv(0, "upperLetter", "%1."), lv(1, "decimal", "%2."), lv(2, "lowerLetter", "%3."),
                            lv(3, "decimal", "%4)"), lv(4, "lowerLetter", "%5)")]},
                       {n: (0, {} if n == 1 else {0: 1}) for n in range(1, anzahl + 1)})
    num_id = 1
    for art, label, text in absaetze:
        marke = "ins" if _ganz(text, "+") else "del" if _ganz(text, "-") else ""
        if art == "part":
            b.absatz(f"PART {label} " + text, marke=marke)
        elif art == "section":
            num_id += 1
            b.absatz(label + "\t" + text, marke=marke)
        elif art in ("sub", "sub_titel"):
            if auto:
                b.absatz(text, num=(num_id, 0), marke=marke)
            else:
                b.absatz(label + "\t" + text, marke=marke)
        elif art in ("liste", "liste_titel"):
            if auto:
                b.absatz(text, num=(num_id, _STUFE[_label_art(_roh_label(label))]), marke=marke)
            else:
                b.absatz(label + "\t" + text, marke=marke)
        else:
            b.absatz(text, marke=marke)
    return b.speichern(pfad)
