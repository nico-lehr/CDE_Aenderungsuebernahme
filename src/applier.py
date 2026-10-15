"""Traegt die geplanten Aenderungen in eine KOPIE der Excel ein.

Grundsaetze
  * Das Original wird nie veraendert; das Ergebnis ist eine neue Datei.
  * Nur Inhalt: Texte in Zellen aendern, Zeilen einfuegen, Inhalt leeren. Keine Formatierung ausser
    dem Kopieren des Zeilenformats fuer neue Zeilen.
  * Alle Operationen beziehen sich auf den Originalzustand. Zuerst die Zelltexte (von hinten nach
    vorn), danach Zeilen einfuegen/loeschen (von unten nach oben) -- so bleiben Koordinaten gueltig.
  * Stimmt der erwartete Text an einer Stelle nicht mehr, wird die Aenderung nicht eingetragen
    (Konflikt) und mit Grund gemeldet.
  * Nach dem Speichern wird die Datei neu gelesen. Jede Aenderung bekommt eine `wirkung`: "ok" nur, wenn die
    Zielzellen/-zeilen nachweislich so aussehen wie geplant UND die Aenderung sie gegenueber dem Original wirklich
    veraendert hat. Sonst wird sie auf "pruefen" bzw. "nicht_uebernommen" zurueckgestuft (nie "uebernommen" ohne Wirkung).
  * Eine neue Section und ihre Folgeabsaetze (stehen in derselben Excel-Zeile) gelten als Gruppe: Status und Freigabe
    gelten fuer alle gemeinsam.
"""
from __future__ import annotations

from collections import Counter
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.cell.rich_text import CellRichText
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from . import config
from .excel_model import ExcelModell
from .models import (MITTEL, NICHT, NIEDRIG, PRUEFEN, UEBERNOMMEN, WIRKUNG_ABWEICHUNG, WIRKUNG_KEINE, WIRKUNG_OK,
                     Aenderung, Operation)
from .planner import neutext_probleme
from .textutil import norm


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


def _ws(text: str) -> str:
    """Leerraum vereinheitlichen (zum Vergleichen)."""
    return " ".join(text.split())


def _hinweise(ae: Aenderung, *texte: str) -> None:
    for t in texte:
        if t and t not in ae.hinweise:
            ae.hinweise.append(t)


def _konflikt(ae: Aenderung, grund: str, *hinweise: str) -> None:
    ae.status, ae.konfidenz, ae.ops = NICHT, NIEDRIG, []
    ae.grund = (ae.grund + "; " if ae.grund else "") + grund
    _hinweise(ae, *hinweise)


def _vorschlag(ae: Aenderung, grund: str, *hinweise: str) -> None:
    """Zurueck auf 'pruefen': ein Mensch muss entscheiden (die Operationen bleiben als Vorschlag stehen)."""
    ae.status, ae.konfidenz = PRUEFEN, MITTEL
    ae.grund = (ae.grund + "; " if ae.grund else "") + grund
    _hinweise(ae, *hinweise)


def _setze_wert(zelle, text: str) -> None:
    """Text in eine Zelle schreiben. Beginnt er mit '=', bleibt er Text (keine Formel)."""
    zelle.value = text if text != "" else None
    if text.startswith("="):
        zelle.data_type = "s"
        zelle.quotePrefix = True


def _kommentar(zelle, text: str, autor: str) -> None:
    """Kommentar anhaengen; ein vorhandener Kommentar wird nie ueberschrieben."""
    text = ILLEGAL_CHARACTERS_RE.sub("\ufffd", text)          # Steuerzeichen aus dem PDF duerfen das Speichern nicht abbrechen
    alt = zelle.comment
    if alt is not None and (alt.text or "").strip():
        text = f"{alt.text}\n---\n{text}"
        autor = alt.author or autor
    zelle.comment = Comment(text, autor)


def _wortgrenze(text: str, pos: int) -> bool:
    """Liegt `pos` zwischen zwei Woertern (oder am Rand)? Einfuegen mitten in einem Wort ist falsch."""
    return pos <= 0 or pos >= len(text) or text[pos - 1].isspace() or text[pos].isspace()


def _folge_gruppen(aenderungen: list[Aenderung]) -> dict[str, list[Aenderung]]:
    """Neue Section (Kopf-ID) -> ihre Folgeabsaetze. Die Folgeabsaetze stehen in der Zeile der Section (keine eigene Zeile)."""
    koepfe: dict[tuple, Aenderung] = {}
    for a in aenderungen:
        if a.typ == "section_neu":
            for o in a.ops:
                if o.art == "zeile_einfuegen" and not o.abgedeckt and o.zellen:
                    koepfe[(o.blatt, o.zeile, o.vor, o.seq)] = a
    gruppen: dict[str, list[Aenderung]] = {}
    for a in aenderungen:
        if a.typ != "absatz_neu":
            continue
        marker = next((o for o in a.ops if o.art == "zeile_einfuegen" and o.abgedeckt and not o.zellen), None)
        kopf = None
        if marker is not None:
            kopf = koepfe.get((marker.blatt, marker.zeile, marker.vor, marker.seq))
        elif not a.ops and a.ziel:        # aelterer Plan ohne Markierung: ueber die gemeinsame Zielzeile
            gleiche = {k.id: k for k in koepfe.values() if k.ziel == a.ziel}
            if len(gleiche) == 1:
                kopf = next(iter(gleiche.values()))
        if kopf is not None:
            gruppen.setdefault(kopf.id, []).append(a)
    return gruppen


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
    # Von hinten nach vorn. Bei gleichem Start wird erst ersetzt und dann eingefuegt: der eingefuegte Text steht vor dem
    # ersetzten (Reihenfolge unabhaengig von `seq`, sonst wird der Text zerstueckelt).
    for o in sorted(ops, key=lambda o: (o.start, 1 if o.art == "ersetzen" else 0, o.seq), reverse=True):
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


def _zelltext(orig: str, ops: list[Operation]) -> str:
    """Endtext einer Zelle nach allen Operationen dieser Zelle."""
    if any(o.art == "zelle_leeren" for o in ops):
        ein = [o for o in ops if o.art == "einfuegen"]
        # Einfuegungen in geleerte Zelle: Reihenfolge nach seq, Position 0
        return "".join(o.text for o in sorted(ein, key=lambda o: o.seq)).strip("\n") if ein else ""
    return _wende_text_an(orig, [o for o in ops if not o.abgedeckt])


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
    freigegebene Vorschlaege -> uebernommen/bestaetigt, Nachkontrolle und Wirkung; Aenderungen ohne
    nachweisbare Wirkung werden zurueckgestuft)."""
    cfg = cfg or config.lade("zuordnung")
    anw = cfg.get("anwenden") or config.lade("zuordnung")["anwenden"]
    farbe, autor = anw["markierung_farbe"], anw["kommentar_autor"]
    freigaben = freigaben or set()
    quelle, ziel = Path(quelle), Path(ziel)
    ergebnis = UebernahmeErgebnis(ausgabe=str(ziel))
    xl = ExcelModell(quelle)                                    # nur zum Lesen (Layout, Sections)
    wb = load_workbook(str(quelle), rich_text=True, keep_vba=(quelle.suffix.lower() == ".xlsm"))

    # --- Welche Aenderungen werden angewendet? (neue Section + Folgeabsaetze gelten als Gruppe) ---------------
    gruppen = _folge_gruppen(aenderungen)
    folge_zu_kopf = {f.id: k for k, fl in gruppen.items() for f in fl}
    alle_nach_id = {a.id: a for a in aenderungen}

    def freigegeben(a: Aenderung) -> bool:
        return a.status == PRUEFEN and a.id in freigaben

    aktiv_ids: set[str] = set()
    for kopf_id, folgen in gruppen.items():
        kopf = alle_nach_id[kopf_id]
        mitglieder = [kopf] + folgen
        if kopf.status in (UEBERNOMMEN, PRUEFEN) and (kopf.status == UEBERNOMMEN or any(freigegeben(m) for m in mitglieder)):
            for m in mitglieder:
                if m.status == PRUEFEN:
                    m.status, m.bestaetigt = UEBERNOMMEN, True
                    if m.id not in freigaben:
                        _hinweise(m, "Zusammen mit der zugehörigen neuen Section bzw. ihrem Absatz freigegeben, "
                                     "weil beide in einer Excel-Zeile stehen.")
                if m.status == UEBERNOMMEN:
                    aktiv_ids.add(m.id)
        else:
            for m in folgen:
                if m.status == UEBERNOMMEN:         # darf nicht allein "uebernommen" sein
                    if kopf.status == NICHT:
                        _konflikt(m, "Gehört zu einer neuen Section, die nicht eingetragen wird.")
                    else:
                        _vorschlag(m, "Gehört zu einer neuen Section, die noch nicht freigegeben ist.",
                                   "Dieser Absatz wird nur zusammen mit seiner neuen Section eingetragen.")
    for ae in aenderungen:
        if ae.id in folge_zu_kopf or ae.id in gruppen:
            continue
        if ae.status == UEBERNOMMEN:
            aktiv_ids.add(ae.id)
        elif freigegeben(ae):
            ae.status, ae.bestaetigt = UEBERNOMMEN, True
            aktiv_ids.add(ae.id)
    aktiv: list[Aenderung] = [a for a in aenderungen if a.id in aktiv_ids]
    nach_id = {a.id: a for a in aktiv}

    def entferne(ae: Aenderung, art: str, grund: str, *hinweise: str) -> None:
        """Nimmt eine Aenderung (und die Gruppe, zu der sie gehoert) aus der Anwendung."""
        if ae.id not in nach_id:
            return
        (_konflikt if art == NICHT else _vorschlag)(ae, grund, *hinweise)
        nach_id.pop(ae.id)
        kopf_id = folge_zu_kopf.get(ae.id)
        weitere = list(gruppen.get(ae.id, []))
        if kopf_id and kopf_id in nach_id:
            weitere.append(nach_id[kopf_id])
            weitere += [f for f in gruppen[kopf_id] if f.id != ae.id]
        for w in weitere:
            entferne(w, art, "Gehört zu einer neuen Section, deren Eintragung nicht möglich war.")

    # --- Stimmt der erwartete Text noch? Ist der Neutext schreibbar? ------------------------------------------
    for ae in list(aktiv):
        if ae.id not in nach_id:
            continue
        for op in ae.ops:
            if op.blatt not in wb.sheetnames:
                entferne(ae, NICHT, "Tabellenblatt nicht gefunden.")
                break
            if op.art in ("ersetzen", "einfuegen", "zelle_leeren"):
                zelle = wb[op.blatt].cell(op.zeile, op.spalte)
                if zelle.data_type == "f":
                    entferne(ae, NICHT, "In der Zelle steht eine Formel; sie wird nie verändert.")
                    break
                aktuell = _text(zelle)
                if op.art != "zelle_leeren" and (op.ende > len(aktuell) or
                                                 (op.art == "ersetzen" and aktuell[op.start:op.ende] != op.erwartet)):
                    entferne(ae, NICHT, "Die Excel hat sich seit der Analyse geändert (Text an der Stelle stimmt nicht mehr).")
                    break
                if op.art == "zelle_leeren" and op.erwartet:
                    if aktuell != op.erwartet:
                        entferne(ae, NICHT, "Die Excel hat sich seit der Analyse geändert (Text an der Stelle stimmt nicht mehr).")
                        break
                    ungestrichen = Counter(norm(w) or w for w in aktuell.split()) - Counter(norm(w) or w for w in ae.alt.split())
                    if ungestrichen:       # eine Zelle wird nur geleert, wenn ALLE ihre Woerter gestrichen sind
                        entferne(ae, PRUEFEN, "In der Zelle stehen Wörter, die im PDF nicht gestrichen sind. Sie wird nicht geleert.",
                                 "Die Zelle enthält mehr Text als im PDF gestrichen ist. Bitte von Hand entscheiden.")
                        break
        if ae.id not in nach_id:
            continue
        texte = [o.text for o in ae.ops if o.art in ("ersetzen", "einfuegen")] + [t for o in ae.ops for t in o.zellen.values()]
        if ae.id in folge_zu_kopf:
            texte.append(ae.neu)
        p = neutext_probleme(texte, cfg)
        if p["steuerzeichen"]:
            entferne(ae, NICHT, "Der neue Text enthält Steuerzeichen und kann nicht in die Excel geschrieben werden.")
        elif p["verdaechtig"] and not ae.bestaetigt:
            entferne(ae, PRUEFEN, "Der neue Text könnte aus dem PDF verfälscht gelesen worden sein; er wird nicht automatisch eingetragen.",
                     *p["verdaechtig"])
        elif p["formel"]:
            _hinweise(ae, "Der neue Text beginnt mit „=“. Er wird als Text eingetragen, nicht als Formel.")
    aktiv = [a for a in aktiv if a.id in nach_id]

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
                    entferne(nach_id[aid], NICHT, "Konflikt: überschneidet sich mit einer anderen Änderung in derselben Zelle.")
    aktiv = [a for a in aktiv if a.id in nach_id]

    geleerte: set[tuple[str, int, int]] = set()
    zell_info: dict[tuple[str, int, int], dict] = {}
    for (blatt, zeile, spalte), ops in je_zelle.items():
        ops = [o for o in ops if o.aenderung in nach_id]
        if not ops:
            continue
        zelle = wb[blatt].cell(zeile, spalte)
        orig = _text(zelle)
        neu = _zelltext(orig, ops)
        if any(o.art == "zelle_leeren" for o in ops):
            geleerte.add((blatt, zeile, spalte))
        _setze_wert(zelle, neu)
        zell_info[(blatt, zeile, spalte)] = {"orig": orig, "neu": neu, "ops": ops,
                                             "aen": [nach_id[i] for i in dict.fromkeys(o.aenderung for o in ops)]}

    # --- Markierung der geaenderten Zellen (optional) -------------------------------------
    modus = cfg["markierung"]["geaenderte_zellen"]
    if modus != "keine":
        for (blatt, zeile, spalte), info in zell_info.items():
            zelle = wb[blatt].cell(zeile, spalte)
            if modus == "gelb":
                zelle.fill = PatternFill("solid", fgColor=farbe)
            elif modus == "kommentar":
                _kommentar(zelle, "\n".join(f"PDF S.{a.seite}: {a.alt[:80]} → {a.neu[:80]}" for a in info["aen"]), autor)

    # --- Zeilen einfuegen / loeschen ------------------------------------------------------
    einfuegungen: dict[tuple[str, int], list[Operation]] = {}
    for ae in aktiv:
        for op in ae.ops:
            if op.art == "zeile_einfuegen" and not op.abgedeckt:
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
                    _setze_wert(ws.cell(neu_zeile, spalte), text)
                if modus == "gelb":
                    for spalte in op.zellen:
                        ws.cell(neu_zeile, spalte).fill = PatternFill("solid", fgColor=farbe)
                elif modus == "kommentar":
                    a = nach_id[op.aenderung]
                    _kommentar(ws.cell(neu_zeile, min(op.zellen)), f"PDF S.{a.seite}: neu eingefügt", autor)

    ziel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(ziel))

    # --- Nachkontrolle: gespeicherte Datei neu lesen -----------------------------------------
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

    # Je Aenderung sammeln: in wie vielen Zielen die Wirkung stimmt (ok), fehlt (keine) oder abweicht (abw)
    ok: Counter = Counter()
    keine: Counter = Counter()
    abw: Counter = Counter()
    grund_abw: dict[str, str] = {}

    kontrolle = load_workbook(str(ziel), rich_text=True, keep_vba=(quelle.suffix.lower() == ".xlsm"))
    for (blatt, zeile, spalte), info in zell_info.items():
        z = neue_zeile(blatt, zeile)
        tatsaechlich = "" if z is None else _text(kontrolle[blatt].cell(z, spalte))
        zelle_stimmt = _ws(tatsaechlich) == _ws(info["neu"])
        for ae in info["aen"]:
            eigene = [o for o in info["ops"] if o.aenderung == ae.id]
            if not zelle_stimmt:
                abw[ae.id] += 1
                grund_abw[ae.id] = "Der Inhalt der Zelle in der gespeicherten Datei weicht vom Plan ab."
            elif any(o.art == "einfuegen" and not _wortgrenze(info["orig"], o.start) for o in eigene):
                abw[ae.id] += 1
                grund_abw[ae.id] = "Der Text wurde mitten in ein Wort eingefügt (der Plan passt nicht zum Zellinhalt)."
            else:
                if all(o.abgedeckt for o in eigene):          # Ersetzen in einer Zelle, die ohnehin geleert wurde
                    wirkt = _ws(info["neu"]) != _ws(info["orig"])
                else:
                    ohne = _zelltext(info["orig"], [o for o in info["ops"] if o.aenderung != ae.id])
                    wirkt = _ws(ohne) != _ws(info["neu"])
                (ok if wirkt else keine)[ae.id] += 1
    neue_zeilen: dict[str, tuple[str, int]] = {}
    for (blatt, pos), ops in einfuegungen.items():
        ws = kontrolle[blatt]
        vorher_eingefuegt = sum(e[2] for e in abbildung.get(blatt, []) if e[0] == "ins" and e[1] < pos)   # type: ignore[misc]
        geloescht_davor = sum(1 for e in abbildung.get(blatt, []) if e[0] == "del" and e[1] < pos)
        basis = pos + vorher_eingefuegt - geloescht_davor
        for i, op in enumerate(sorted(ops, key=lambda o: o.seq)):
            stimmt = bool(op.zellen) and all(_ws(_text(ws.cell(basis + i, c))) == _ws(t) for c, t in op.zellen.items())
            if stimmt:
                ok[op.aenderung] += 1
                neue_zeilen[op.aenderung] = (blatt, basis + i)
            else:
                abw[op.aenderung] += 1
                grund_abw[op.aenderung] = "Die neue Zeile in der gespeicherten Datei weicht vom Plan ab."
    # Folgeabsaetze einer neuen Section: Wirkung = die Section-Zeile stimmt UND ihr Absatz steht in der Zelle
    for kopf_id, folgen in gruppen.items():
        kopf_op = next((o for o in nach_id[kopf_id].ops if o.art == "zeile_einfuegen" and not o.abgedeckt), None) \
            if kopf_id in nach_id else None
        for f in folgen:
            if f.id not in nach_id or kopf_op is None:
                continue
            spalten = [c for c, t in kopf_op.zellen.items() if _ws(f.neu) in [_ws(x) for x in t.split("\n")]]
            if kopf_id in neue_zeilen and spalten:
                blatt, r = neue_zeilen[kopf_id]
                zeilen = [_ws(x) for x in _text(kontrolle[blatt].cell(r, spalten[0])).split("\n")]
                if _ws(f.neu) in zeilen:
                    ok[f.id] += 1
                else:
                    abw[f.id] += 1
                    grund_abw[f.id] = "Der Absatz steht nicht in der neuen Section-Zeile."
            elif abw[kopf_id]:
                abw[f.id] += 1
                grund_abw[f.id] = grund_abw.get(kopf_id, "Die zugehörige neue Section stimmt nicht.")
            else:
                keine[f.id] += 1

    for ae in aktiv:
        if abw[ae.id] or (keine[ae.id] and ok[ae.id]):
            w = WIRKUNG_ABWEICHUNG
        elif ok[ae.id]:
            w = WIRKUNG_OK
        else:
            w = WIRKUNG_KEINE
        ae.wirkung = w
        if w == WIRKUNG_OK:
            ae.nachkontrolle = "ok"
            continue
        ae.nachkontrolle = "abweichung"
        if w == WIRKUNG_KEINE:
            _konflikt(ae, "Die Änderung hatte keine Wirkung auf die Excel (nichts hat sich geändert).",
                      "Die Excel ist an dieser Stelle unverändert geblieben. Bitte von Hand prüfen.")
        else:
            _vorschlag(ae, "Die Excel wurde an dieser Stelle geändert, aber nicht so wie geplant.",
                       grund_abw.get(ae.id, "Das Ergebnis weicht vom Plan ab."),
                       "Bitte die Stelle in der Ergebnis-Excel von Hand prüfen.")
            ae.konfidenz = NIEDRIG
            ergebnis.warnungen.append(f"Änderung {ae.id} ({ae.ziel}): in der Ergebnis-Excel weicht die Stelle vom Plan ab. Bitte von Hand prüfen.")

    # --- Hinweise: Inhalt in gestrichenen Sections ----------------------------------------
    aktive_ops = [o for a in aktiv if a.status == UEBERNOMMEN for o in a.ops]
    for ae in aktiv:
        if ae.typ == "section_geloescht" and ae.status == UEBERNOMMEN:
            for op in ae.ops:
                lay = xl.layouts.get(op.blatt)
                if lay and op.zeile in lay.kopf_zeilen and lay.inhalt_spalte:
                    for r in xl.section_zeilen(op.blatt, op.zeile):
                        if (op.blatt, r, lay.inhalt_spalte) not in geleerte and xl.text(op.blatt, r, lay.inhalt_spalte).strip():
                            if not any(o.zeile == r and o.blatt == op.blatt for o in aktive_ops):
                                ergebnis.warnungen.append(
                                    f"Section '{ae.alt}' entfällt, aber in {op.blatt}!{get_column_letter(lay.inhalt_spalte)}{r} "
                                    f"steht noch Inhalt (im PDF nicht gestrichen). Bitte prüfen.")
                                break
    return ergebnis
