"""Traegt die geplanten Aenderungen in eine KOPIE der Excel ein.

Grundsaetze
  * Das Original wird nie veraendert; das Ergebnis ist eine neue Datei.
  * Nur Inhalt: Texte in Zellen aendern, Zeilen einfuegen, Inhalt leeren. Keine Formatierung ausser
    dem Kopieren des Zeilenformats fuer neue Zeilen.
  * Alle Operationen beziehen sich auf den Originalzustand. Zuerst die Zelltexte (von hinten nach
    vorn), danach Zeilen einfuegen/loeschen (von unten nach oben) -- so bleiben Koordinaten gueltig.
  * Stimmt der erwartete Text an einer Stelle nicht mehr, wird die Aenderung nicht eingetragen
    (Konflikt) und mit Grund gemeldet.
  * Nach dem Speichern wird die Datei neu gelesen und jede Aenderung nachkontrolliert.
"""
from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from . import config
from .excel_model import ExcelModell
from .models import NICHT, NIEDRIG, PRUEFEN, UEBERNOMMEN, Aenderung, Operation


@dataclass
class UebernahmeErgebnis:
    ausgabe: str
    warnungen: list[str] = field(default_factory=list)
    zeilen_verschoben: dict[str, str] = field(default_factory=dict)


def _text(zelle) -> str:
    v = zelle.value
    if v is None:
        return ""
    if isinstance(v, CellRichText):
        return str(v)
    return v if isinstance(v, str) else ("%g" % v if isinstance(v, float) else str(v))


def _konflikt(ae: Aenderung, grund: str) -> None:
    ae.status, ae.konfidenz, ae.ops = NICHT, NIEDRIG, []
    ae.grund = (ae.grund + "; " if ae.grund else "") + grund


def _schneiden(a: Operation, b: Operation) -> bool:
    ins_a, ins_b = a.start == a.ende, b.start == b.ende
    if ins_a and ins_b:
        return False
    if ins_b:
        return a.start < b.start < a.ende
    if ins_a:
        return b.start < a.start < b.ende
    return a.start < b.ende and b.start < a.ende


def _ueberlappen(ops: list[Operation]) -> set[str]:
    """IDs der Aenderungen, deren Operationen sich in einer Zelle ueberschneiden (die spaetere verliert)."""
    verlierer: set[str] = set()
    s = sorted(ops, key=lambda o: (o.start, o.ende, o.seq))
    for i, a in enumerate(s):
        for b in s[:i]:
            if b.aenderung == a.aenderung or a.aenderung in verlierer or b.aenderung in verlierer:
                continue
            if _schneiden(a, b):
                verlierer.add(a.aenderung)
    return verlierer


def _wende_text_an(orig: str, ops: list[Operation]) -> str:
    text = orig
    for o in sorted(ops, key=lambda o: (o.start, o.seq), reverse=True):
        if o.art == "ersetzen":
            text = text[:o.start] + o.text + text[o.ende:]
        else:   # einfuegen
            text = text[:o.start] + o.text + text[o.start:]
    if any(o.absatz_weg for o in ops) and "\n\n" not in orig and not orig.startswith("\n") and not orig.endswith("\n"):
        text = "\n".join(z for z in text.split("\n") if z.strip())
    if not orig.endswith("\n"):
        text = text.rstrip("\n")
    if not orig.startswith("\n"):
        text = text.lstrip("\n")
    if not orig.endswith(" "):
        text = text.rstrip(" ")
    return text


def _verschiebe_hoehen(ws, ab: int, delta: int) -> None:
    hoehen = {r: ws.row_dimensions[r].height for r in list(ws.row_dimensions.keys()) if r >= ab}
    for r in hoehen:
        ws.row_dimensions[r].height = None
    for r, h in hoehen.items():
        ziel = r + delta
        if ziel >= 1:
            ws.row_dimensions[ziel].height = h


def _verschiebe_verbunden(ws, ab: int, delta: int) -> None:
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= ab:
            ws.merged_cells.remove(rng)
            rng.shift(row_shift=delta)
            ws.merged_cells.add(rng)


def uebernehmen(aenderungen: list[Aenderung], quelle: str | Path, ziel: str | Path,
                freigaben: set[str] | None = None, cfg: dict | None = None) -> UebernahmeErgebnis:
    """Trägt alle Aenderungen mit Status 'uebernommen' sowie freigegebene 'pruefen' ein.

    Die Aenderung-Objekte werden aktualisiert (Konflikte -> nicht_uebernommen,
    freigegebene Vorschlaege -> uebernommen/bestaetigt, Nachkontrolle)."""
    cfg = cfg or config.lade("zuordnung")
    freigaben = freigaben or set()
    quelle, ziel = Path(quelle), Path(ziel)
    ergebnis = UebernahmeErgebnis(ausgabe=str(ziel))
    xl = ExcelModell(quelle)                                    # nur zum Lesen (Layout, Sections)
    wb = load_workbook(str(quelle), rich_text=True, keep_vba=(quelle.suffix.lower() == ".xlsm"))

    aktiv: list[Aenderung] = []
    for ae in aenderungen:
        if ae.status == UEBERNOMMEN:
            aktiv.append(ae)
        elif ae.status == PRUEFEN and ae.id in freigaben:
            ae.status, ae.bestaetigt = UEBERNOMMEN, True
            aktiv.append(ae)
    nach_id = {a.id: a for a in aktiv}

    # --- Stimmt der erwartete Text noch? -------------------------------------------------
    for ae in list(aktiv):
        for op in ae.ops:
            if op.blatt not in wb.sheetnames:
                _konflikt(ae, "Tabellenblatt nicht gefunden.")
                break
            if op.art in ("ersetzen", "einfuegen"):
                aktuell = _text(wb[op.blatt].cell(op.zeile, op.spalte))
                if op.ende > len(aktuell) or (op.art == "ersetzen" and aktuell[op.start:op.ende] != op.erwartet):
                    _konflikt(ae, "Die Excel hat sich seit der Analyse geändert (Text an der Stelle stimmt nicht mehr).")
                    break
        if ae.status == NICHT:
            aktiv.remove(ae)
            nach_id.pop(ae.id, None)

    # --- Zelltext-Operationen je Zelle ---------------------------------------------------
    je_zelle: dict[tuple[str, int, int], list[Operation]] = {}
    for ae in aktiv:
        for op in ae.ops:
            if op.art in ("ersetzen", "einfuegen", "zelle_leeren"):
                je_zelle.setdefault((op.blatt, op.zeile, op.spalte), []).append(op)

    for key, ops in je_zelle.items():
        text_ops = [o for o in ops if o.art != "zelle_leeren"]
        if any(o.art == "zelle_leeren" for o in ops):
            for o in text_ops:                                   # Ersetzen in geleerter Zelle: bereits abgedeckt
                if o.art == "ersetzen":
                    o.abgedeckt = True
        else:
            for aid in _ueberlappen(text_ops):
                if aid in nach_id:
                    _konflikt(nach_id[aid], "Konflikt: überschneidet sich mit einer anderen Änderung in derselben Zelle.")
                    nach_id.pop(aid)
            aktiv = [a for a in aktiv if a.id in nach_id]

    geleerte: set[tuple[str, int, int]] = set()
    geaendert: dict[tuple[str, int, int], list[Aenderung]] = {}
    for (blatt, zeile, spalte), ops in je_zelle.items():
        ops = [o for o in ops if o.aenderung in nach_id]
        if not ops:
            continue
        zelle = wb[blatt].cell(zeile, spalte)
        orig = _text(zelle)
        if any(o.art == "zelle_leeren" for o in ops):
            ein = [o for o in ops if o.art == "einfuegen"]
            neu = ""
            if ein:   # Einfuegungen in geleerte Zelle: Reihenfolge nach seq, Position 0
                neu = "".join(o.text for o in sorted(ein, key=lambda o: o.seq)).strip("\n")
            geleerte.add((blatt, zeile, spalte))
        else:
            neu = _wende_text_an(orig, [o for o in ops if not o.abgedeckt])
        zelle.value = neu if neu != "" else None
        geaendert[(blatt, zeile, spalte)] = [nach_id[o.aenderung] for o in ops]

    # --- Markierung der geaenderten Zellen (optional) -------------------------------------
    modus = cfg["markierung"]["geaenderte_zellen"]
    if modus != "keine":
        for (blatt, zeile, spalte), liste in geaendert.items():
            zelle = wb[blatt].cell(zeile, spalte)
            if modus == "gelb":
                zelle.fill = PatternFill("solid", fgColor="FFF2CC")
            elif modus == "kommentar":
                zelle.comment = Comment("\n".join(f"PDF S.{a.seite}: {a.alt[:80]} → {a.neu[:80]}" for a in liste), "CDE-Tool")

    # --- Zeilen einfuegen / loeschen ------------------------------------------------------
    einfuegungen: dict[tuple[str, int], list[Operation]] = {}
    for ae in aktiv:
        for op in ae.ops:
            if op.art == "zeile_einfuegen":
                pos = op.zeile if op.vor else op.zeile + 1
                einfuegungen.setdefault((op.blatt, pos), []).append(op)
    loeschen: set[tuple[str, int]] = set()
    if cfg["geloeschtes"]["inhalt"] == "zeile_entfernen":
        for (blatt, zeile, spalte) in geleerte:
            lay = xl.layouts.get(blatt)
            if not lay:
                continue
            cols = [c for c in (lay.section_spalte, lay.titel_spalte, lay.inhalt_spalte) if c]
            if all(wb[blatt].cell(zeile, c).value in (None, "") for c in cols) and not wb[blatt].merged_cells.ranges:
                loeschen.add((blatt, zeile))
    ereignisse = [("del", b, r) for b, r in loeschen] + [("ins", b, p) for (b, p) in einfuegungen]
    ereignisse.sort(key=lambda e: (e[1], -e[2], 0 if e[0] == "del" else 1))
    abbildung: dict[str, list[tuple[str, int]]] = {}
    for art, blatt, pos in ereignisse:
        ws = wb[blatt]
        if art == "del":
            ws.delete_rows(pos, 1)
            _verschiebe_hoehen(ws, pos + 1, -1)
            _verschiebe_verbunden(ws, pos + 1, -1)
            abbildung.setdefault(blatt, []).append(("del", pos))
        else:
            ops = sorted(einfuegungen[(blatt, pos)], key=lambda o: o.seq)
            n = len(ops)
            ws.insert_rows(pos, n)
            _verschiebe_hoehen(ws, pos, n)
            _verschiebe_verbunden(ws, pos, n)
            abbildung.setdefault(blatt, []).append(("ins", pos, n))        # type: ignore[arg-type]
            for i, op in enumerate(ops):
                neu_zeile = pos + i
                stil = op.stil_von + (n if op.stil_von >= pos else 0)
                for c in range(1, ws.max_column + 1):
                    ws.cell(neu_zeile, c)._style = copy(ws.cell(stil, c)._style)
                ws.row_dimensions[neu_zeile].height = ws.row_dimensions[stil].height
                for spalte, text in op.zellen.items():
                    ws.cell(neu_zeile, spalte).value = text
                if modus == "gelb":
                    for spalte in op.zellen:
                        ws.cell(neu_zeile, spalte).fill = PatternFill("solid", fgColor="FFF2CC")
                elif modus == "kommentar":
                    a = nach_id[op.aenderung]
                    c0 = min(op.zellen)
                    ws.cell(neu_zeile, c0).comment = Comment(f"PDF S.{a.seite}: neu eingefügt", "CDE-Tool")

    # --- Hinweise: Inhalt in gestrichenen Sections ----------------------------------------
    for ae in aktiv:
        if ae.typ == "section_geloescht":
            for op in ae.ops:
                lay = xl.layouts.get(op.blatt)
                if lay and op.zeile in lay.kopf_zeilen and lay.inhalt_spalte:
                    for r in xl.section_zeilen(op.blatt, op.zeile):
                        if (op.blatt, r, lay.inhalt_spalte) not in geleerte and xl.text(op.blatt, r, lay.inhalt_spalte).strip():
                            if not any(o.zeile == r and o.blatt == op.blatt for a in aktiv for o in a.ops):
                                ergebnis.warnungen.append(
                                    f"Section '{ae.alt}' entfällt, aber in {op.blatt}!{get_column_letter(lay.inhalt_spalte)}{r} "
                                    f"steht noch Inhalt (im PDF nicht gestrichen). Bitte prüfen.")
                                break

    ziel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(ziel))

    # --- Nachkontrolle: gespeicherte Datei neu lesen ---------------------------------------
    def neue_zeile(blatt: str, orig: int) -> int | None:
        z = orig
        for e in sorted(abbildung.get(blatt, []), key=lambda t: t[1]):
            if e[0] == "del":
                if e[1] == orig:
                    return None
                if e[1] < orig:
                    z -= 1
            else:
                if e[1] <= orig:
                    z += e[2]       # type: ignore[misc]
        return z

    # Reihenfolge: Ereignisse in Original-Koordinaten verarbeiten (Einfuegung vor Original-Zeile pos)
    kontrolle = load_workbook(str(ziel), rich_text=True, keep_vba=(quelle.suffix.lower() == ".xlsm"))
    for ae in aktiv:
        if ae.id not in nach_id:
            continue
        ok = True
        for op in ae.ops:
            ws = kontrolle[op.blatt]
            if op.art in ("ersetzen", "einfuegen", "zelle_leeren"):
                z = neue_zeile(op.blatt, op.zeile)
                if z is None:
                    continue
                inhalt = " ".join(_text(ws.cell(z, op.spalte)).split())
                if op.art == "zelle_leeren":
                    ok &= inhalt == "" or any(o.art == "einfuegen" for o in je_zelle.get((op.blatt, op.zeile, op.spalte), []))
                elif op.text.strip():
                    ok &= " ".join(op.text.split()) in inhalt
                elif op.erwartet:
                    erw = " ".join(op.erwartet.split())
                    vorher = " ".join(xl.text(op.blatt, op.zeile, op.spalte).split())
                    ok &= inhalt.count(erw) < vorher.count(erw)
        ae.nachkontrolle = "ok" if ok else "abweichung"
    # Neue Zeilen pruefen
    for (blatt, pos), ops in einfuegungen.items():
        ws = kontrolle[blatt]
        vorher_eingefuegt = sum(e[2] for e in abbildung.get(blatt, []) if e[0] == "ins" and e[1] < pos)   # type: ignore[misc]
        geloescht_davor = sum(1 for e in abbildung.get(blatt, []) if e[0] == "del" and e[1] < pos)
        basis = pos + vorher_eingefuegt - geloescht_davor
        for i, op in enumerate(sorted(ops, key=lambda o: o.seq)):
            ok = all(" ".join(str(ws.cell(basis + i, c).value or "").split()) == " ".join(t.split())
                     for c, t in op.zellen.items())
            a = nach_id.get(op.aenderung)
            if a is not None and not ok:
                a.nachkontrolle = "abweichung"
    return ergebnis
