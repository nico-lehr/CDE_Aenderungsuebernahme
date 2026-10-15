"""Liest ein Aenderungs-PDF und erkennt, was neu ist und was entfaellt.

Signale (alle in config/farben.yaml einstellbar):
  * Textfarbe            (blau = neu, rot = entfaellt)
  * farbige Hinterlegung (gefuellte Rechtecke hinter dem Text)
  * Durchstreichung      (duenne waagerechte Linie in Zeichenmitte)
  * PDF-Kommentare       (Highlight / StrikeOut)

Ergebnis: Woerter (PdfToken) mit Status keep/del/add/mod, zu Bloecken
(Absatz / Ueberschrift) zusammengefasst. Reine Python-Logik, kein LLM, keine UI.
"""
from __future__ import annotations

import colorsys
import re
import statistics
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pdfplumber

from . import config
from .models import ADD, DEL, KEEP, MOD, PdfBlock, PdfToken
from .textutil import norm, section_nummer


@dataclass
class PdfAnalyse:
    pfad: str
    seiten: int
    blocks: list[PdfBlock]
    ignoriert: list[dict[str, Any]] = field(default_factory=list)
    farben: list[dict[str, Any]] = field(default_factory=list)
    warnungen: list[str] = field(default_factory=list)
    seiten_masse: dict[int, tuple[float, float]] = field(default_factory=dict)
    luecken: list[dict[str, Any]] = field(default_factory=list)   # Schema 2: models.luecke(...)

    @property
    def markierte_woerter(self) -> int:
        n = sum(1 for b in self.blocks for t in b.tokens if t.status != KEEP)
        return n + sum(i["woerter"] for i in self.ignoriert)


# ---------------------------------------------------------------- Farben
def _rgb(farbe: Any) -> tuple[float, float, float] | None:
    if farbe is None or isinstance(farbe, str):
        return None
    if isinstance(farbe, (int, float)):
        farbe = (farbe,)
    try:
        c = tuple(float(x) for x in farbe)
    except (TypeError, ValueError):
        return None
    if len(c) == 1:
        return (c[0], c[0], c[0])
    if len(c) == 3:
        return c  # type: ignore[return-value]
    if len(c) == 4:
        cc, m, y, k = c
        return ((1 - cc) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
    return None


def _hex(rgb: tuple[float, float, float]) -> str:
    return "#%02x%02x%02x" % tuple(max(0, min(255, round(v * 255))) for v in rgb)


def klasse(rgb: tuple[float, float, float] | None, cfg: dict, hintergrund: bool = False) -> str | None:
    """'blau' | 'rot' | None."""
    if rgb is None:
        return None
    r, g, b = (max(0.0, min(1.0, v)) for v in rgb)
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    grenzen = cfg["hintergrund"] if hintergrund else cfg["text"]
    if s < grenzen["min_saettigung"] or v < grenzen.get("min_helligkeit", 0.25):
        return None
    grad = h * 360.0
    lo, hi = cfg["farbton"]["blau"]
    if lo <= grad <= hi:
        return "blau"
    if grad <= cfg["farbton"]["rot_bis"] or grad >= cfg["farbton"]["rot_ab"]:
        return "rot"
    return None


def _bedeutung(farbklasse: str | None, cfg: dict) -> str | None:
    if farbklasse is None:
        return None
    wert = cfg["legende"].get(farbklasse)
    return wert if wert in ("neu", "entfaellt") else None


# ---------------------------------------------------------------- Seite
@dataclass
class _Z:
    text: str
    x0: float
    x1: float
    top: float
    bottom: float
    size: float
    font: str
    status: str


@dataclass
class _Zeile:
    tokens: list[PdfToken]
    top: float
    bottom: float
    x0: float
    x1: float
    size: float
    fett: bool
    seite: int


class _Markierer:
    """Sammelt pro Seite alle Markierungs-Signale und beantwortet: Status je Zeichen."""

    def __init__(self, page: Any, cfg: dict) -> None:
        self.cfg = cfg
        self.flaechen: list[tuple[float, float, float, float, str]] = []   # Hinterlegungen
        self.striche: list[tuple[float, float, float]] = []
        self.annot: list[tuple[float, float, float, float, str]] = []
        self.rect_farben: Counter = Counter()
        w, h = float(page.width), float(page.height)
        if cfg["hintergrund"]["aktiv"]:
            for r in page.rects:
                if not r.get("fill"):
                    continue
                rgb = _rgb(r.get("non_stroking_color"))
                if rgb is None:
                    continue
                if r["width"] * r["height"] > cfg["hintergrund"]["max_flaeche_anteil"] * w * h:
                    continue
                if r["height"] <= 1.6:     # duenne Rechtecke sind Linien (siehe unten)
                    continue
                k = klasse(rgb, cfg, hintergrund=True)
                self.rect_farben[(_hex(rgb), k)] += 1
                bed = _bedeutung(k, cfg)
                if bed:
                    self.flaechen.append((r["x0"], r["top"], r["x1"], r["bottom"], bed))
        d = cfg["durchstreichung"]
        if d["aktiv"]:
            bed = _bedeutung("durchgestrichen", cfg)
            if bed:
                for l in page.lines:
                    if (l.get("linewidth", 1) or 0) <= d["max_linienstaerke_pt"] and abs(l["bottom"] - l["top"]) <= 1.2 \
                            and (l["x1"] - l["x0"]) < d["max_breite_anteil"] * w:
                        self.striche.append((l["x0"], l["x1"], (l["top"] + l["bottom"]) / 2))
                for r in page.rects:
                    if r["height"] <= 1.6 and 2 <= r["width"] < d["max_breite_anteil"] * w:
                        self.striche.append((r["x0"], r["x1"], (r["top"] + r["bottom"]) / 2))
        if cfg["annotationen"]["aktiv"]:
            self._annotationen(page, h)

    def _annotationen(self, page: Any, h: float) -> None:
        cfg = self.cfg
        for a in page.annots:
            data = a.get("data") or {}
            sub = data.get("Subtype")
            name = getattr(sub, "name", str(sub))
            if name == "StrikeOut":
                bed = _bedeutung("durchgestrichen", cfg)
            elif name == "Highlight":
                bed = _bedeutung(klasse(_rgb(data.get("C")), cfg, hintergrund=True), cfg)
            else:
                continue
            if not bed:
                continue
            quad = data.get("QuadPoints")
            try:
                q = [float(x) for x in quad] if quad else []
            except (TypeError, ValueError):
                q = []
            boxen = []
            if len(q) >= 8:
                for i in range(0, len(q) - 7, 8):
                    xs, ys = q[i:i + 8:2], q[i + 1:i + 8:2]
                    boxen.append((min(xs), h - max(ys), max(xs), h - min(ys)))
            else:
                boxen.append((a["x0"], a["top"], a["x1"], a["bottom"]))
            for x0, top, x1, bot in boxen:
                self.annot.append((x0, top, x1, bot, bed))

    def status(self, c: dict) -> str:
        cfg = self.cfg
        bed: set[str] = set()
        rgb = _rgb(c.get("non_stroking_color"))
        b = _bedeutung(klasse(rgb, cfg), cfg)
        if b:
            bed.add(b)
        cx, cy = (c["x0"] + c["x1"]) / 2, (c["top"] + c["bottom"]) / 2
        for x0, top, x1, bot, b in self.flaechen:
            if x0 - 0.5 <= cx <= x1 + 0.5 and top - 0.5 <= cy <= bot + 0.5:
                bed.add(b)
        for x0, top, x1, bot, b in self.annot:
            if x0 - 1 <= cx <= x1 + 1 and top - 1 <= cy <= bot + 1:
                bed.add(b)
        if self.striche:
            d = cfg["durchstreichung"]
            h = c["bottom"] - c["top"]
            lo, hi = c["top"] + d["band"][0] * h, c["top"] + d["band"][1] * h
            breite = max(c["x1"] - c["x0"], 0.1)
            for x0, x1, y in self.striche:
                if lo <= y <= hi:
                    ueber = min(x1, c["x1"]) - max(x0, c["x0"])
                    if ueber / breite >= d["min_abdeckung"]:
                        bed.add(_bedeutung("durchgestrichen", cfg) or "entfaellt")
                        break
        if "entfaellt" in bed:
            return DEL
        if "neu" in bed:
            return ADD
        return KEEP


def _woerter(zeichen: list[_Z], seite: int, cfg: dict) -> list[PdfToken]:
    """Zeichen einer Zeile -> Woerter (mit Statusanteilen)."""
    lay = cfg["layout"]
    zeichen = sorted(zeichen, key=lambda z: z.x0)
    gruppen: list[list[_Z]] = []
    for z in zeichen:
        if gruppen and z.x0 - gruppen[-1][-1].x1 <= lay["wort_luecke"] * z.size:
            gruppen[-1].append(z)
        else:
            gruppen.append([z])
    tokens = []
    for g in gruppen:
        n = len(g)
        n_del = sum(1 for z in g if z.status == DEL)
        n_add = sum(1 for z in g if z.status == ADD)
        anteil = cfg["durchstreichung"]["wort_anteil"]
        if n_del / n >= anteil:
            st = [DEL] * n
        elif n_add / n >= anteil and n_del == 0:
            st = [ADD] * n
        else:
            st = [z.status for z in g]
        old = "".join(z.text for z, s in zip(g, st) if s in (KEEP, DEL))
        new = "".join(z.text for z, s in zip(g, st) if s in (KEEP, ADD))
        if all(s == KEEP for s in st):
            status = KEEP
        elif all(s == DEL for s in st):
            status = DEL
        elif all(s == ADD for s in st):
            status = ADD
        else:
            status = MOD
        bbox = (min(z.x0 for z in g), min(z.top for z in g), max(z.x1 for z in g), max(z.bottom for z in g))
        tokens.append(PdfToken(old=old, new=new, status=status, page=seite, bbox=bbox))
    return tokens


def _zeilen(zeichen: list[_Z], seite: int, cfg: dict) -> list[_Zeile]:
    tol = cfg["layout"]["zeilen_toleranz"]
    zeichen = sorted(zeichen, key=lambda z: ((z.top + z.bottom) / 2, z.x0))
    cluster: list[list[_Z]] = []
    mitten: list[float] = []
    for z in zeichen:
        mitte = (z.top + z.bottom) / 2
        if cluster and abs(mitte - mitten[-1]) <= tol * z.size:
            cluster[-1].append(z)
            mitten[-1] = statistics.fmean((c.top + c.bottom) / 2 for c in cluster[-1])
        else:
            cluster.append([z])
            mitten.append(mitte)
    zeilen = []
    for cl in cluster:
        toks = _woerter(cl, seite, cfg)
        if not toks:
            continue
        fett_n = sum(1 for z in cl if re.search(r"bold|black|heavy|semibold", z.font, re.I))
        zeilen.append(_Zeile(
            tokens=toks, top=min(z.top for z in cl), bottom=max(z.bottom for z in cl),
            x0=min(z.x0 for z in cl), x1=max(z.x1 for z in cl),
            size=statistics.median(z.size for z in cl), fett=fett_n / len(cl) > 0.6, seite=seite))
    zeilen.sort(key=lambda z: z.top)
    return zeilen


# ---------------------------------------------------------------- Bloecke
_AUFZAEHLUNG = re.compile(r"^([•\-–·▪●○*]|\(?\d{1,2}\)|[a-z]\))$")


def _erste_form(t: PdfToken) -> str:
    return t.new or t.old


def _ist_ueberschrift(z: _Zeile, body: float, cfg: dict) -> bool:
    t = z.tokens[0]
    nr_alt = section_nummer(t.old) if t.old else None
    nr_neu = section_nummer(t.new) if t.new else None
    wenige = len(z.tokens) <= 25
    gross = z.size >= body * cfg["layout"]["ueberschrift_faktor"]
    if (nr_alt or nr_neu) and len(z.tokens) >= 2 and wenige and (z.fett or gross):
        return True
    return bool(z.fett and z.size >= body * 1.15 and wenige)


def _kopf_fuss_schluessel(z: _Zeile) -> str:
    text = " ".join(_erste_form(t) for t in z.tokens)
    return re.sub(r"\d+", "#", text.casefold())


def _zu_bloecken(alle_zeilen: list[list[_Zeile]], seiten_h: dict[int, float], cfg: dict,
                 ergebnis: PdfAnalyse) -> None:
    lay = cfg["layout"]
    # Kopf-/Fusszeilen finden
    kandidaten: dict[str, set[int]] = {}
    for zl in alle_zeilen:
        for z in zl:
            h = seiten_h[z.seite]
            if z.top < lay["kopf_fuss_anteil"] * h or z.bottom > (1 - lay["kopf_fuss_anteil"]) * h:
                kandidaten.setdefault(_kopf_fuss_schluessel(z), set()).add(z.seite)
    n_seiten = len(alle_zeilen)
    muster = re.compile(r"^(seite|page)\s*#", re.I)

    def ist_kopf_fuss(z: _Zeile) -> bool:
        h = seiten_h[z.seite]
        if not (z.top < lay["kopf_fuss_anteil"] * h or z.bottom > (1 - lay["kopf_fuss_anteil"]) * h):
            return False
        key = _kopf_fuss_schluessel(z)
        return bool(muster.match(key)) or (n_seiten >= 2 and len(kandidaten.get(key, ())) >= max(2, n_seiten // 2))

    groessen = Counter()
    for zl in alle_zeilen:
        for z in zl:
            if not ist_kopf_fuss(z):
                groessen[round(z.size * 2) / 2] += sum(len(t.old or t.new) for t in z.tokens)
    body = groessen.most_common(1)[0][0] if groessen else 11.0

    sec_alt: str | None = None
    sec_neu: str | None = None
    idx = 0
    seiten_nutz: list[list[_Zeile]] = []
    abstaende: list[float] = []
    for zl in alle_zeilen:
        nutz = []
        for z in zl:
            if ist_kopf_fuss(z):
                gew = sum(1 for t in z.tokens if t.status != KEEP)
                if gew:
                    ergebnis.ignoriert.append({"seite": z.seite, "grund": "kopf_fuss", "woerter": gew,
                                               "text": " ".join(_erste_form(t) for t in z.tokens)})
                continue
            nutz.append(z)
        seiten_nutz.append(nutz)
        abstaende += [b.top - a.top for a, b in zip(nutz, nutz[1:]) if 0.9 * a.size < b.top - a.top < 3 * a.size]
    median = _zeilenabstand(abstaende, body)
    weit = median >= 1.5 * body          # grosser Zeilenabstand: Absatzende erkennt man an der kurzen Zeile
    rechts_x = sorted(z.x1 for zl in seiten_nutz for z in zl if len(z.tokens) >= 5)
    rechts = rechts_x[int(0.9 * (len(rechts_x) - 1))] if rechts_x else None
    alle_z = [z for zl in seiten_nutz for z in zl]
    if alle_z and ergebnis.seiten_masse:
        links = min(z.x0 for z in alle_z)
        breite = max(b for b, _ in ergebnis.seiten_masse.values())
        schwelle_kurz = links + 0.75 * (breite - 2 * links)   # Zeile endet deutlich vor dem rechten Rand
    else:
        schwelle_kurz = None

    def kurz(z: _Zeile) -> bool:
        return (rechts is not None and z.x1 < rechts - 3 * z.size) or (schwelle_kurz is not None and z.x1 < schwelle_kurz)
    for nutz in seiten_nutz:
        if not nutz:
            continue
        aktuell: list[_Zeile] = []
        art = "absatz"

        def abschliessen() -> None:
            nonlocal idx, sec_alt, sec_neu, aktuell, art
            if not aktuell:
                return
            toks = [t for z in aktuell for t in z.tokens]
            if art == "ueberschrift":
                t0 = toks[0]
                n_alt = section_nummer(t0.old) if t0.old else None
                n_neu = section_nummer(t0.new) if t0.new else None
                if n_alt or n_neu:
                    sec_alt, sec_neu = n_alt, n_neu
                else:   # unnummerierte Ueberschrift: Zuordnung ueber Text, Nummern bleiben
                    sec_alt = sec_neu = None
            else:
                toks = _silben(toks, aktuell, cfg)
            bbox = (min(z.x0 for z in aktuell), min(z.top for z in aktuell),
                    max(z.x1 for z in aktuell), max(z.bottom for z in aktuell))
            ergebnis.blocks.append(PdfBlock(idx=idx, page=aktuell[0].seite, kind=art, tokens=toks, bbox=bbox,
                                            section_alt=sec_alt, section_neu=sec_neu))
            idx += 1
            aktuell = []

        vorher: _Zeile | None = None
        for z in nutz:
            kopf = _ist_ueberschrift(z, body, cfg)
            neu_start = (not aktuell or kopf or art == "ueberschrift"
                         or (vorher is not None and z.top - vorher.top > lay["absatz_faktor"] * median)
                         or (weit and vorher is not None and kurz(vorher))
                         or bool(_AUFZAEHLUNG.match(_erste_form(z.tokens[0]))))
            if neu_start:
                abschliessen()
                art = "ueberschrift" if kopf else "absatz"
            aktuell.append(z)
            vorher = z
        abschliessen()


def _zeilenabstand(abstaende: list[float], body: float) -> float:
    """Typischer Zeilenabstand (dokumentweit): kleinste Gruppe (+-8 %) mit mindestens 2 Treffern."""
    if not abstaende:
        return body * 1.3
    gruppen: list[list[float]] = []
    for a in sorted(abstaende):
        if gruppen and a <= gruppen[-1][0] * 1.08:
            gruppen[-1].append(a)
        else:
            gruppen.append([a])
    stark = [g for g in gruppen if len(g) >= 2] or gruppen
    return statistics.median(stark[0])


def _silben(tokens: list[PdfToken], zeilen: list[_Zeile], cfg: dict) -> list[PdfToken]:
    """Fuegt am Zeilenende getrennte Woerter ('Kapazi-' + 'taet') wieder zusammen."""
    lay = cfg["layout"]
    if not lay["silbentrennung_zusammenfuegen"] or len(zeilen) < 2:
        return tokens
    ausnahmen = {norm(w) for w in lay["keine_silbentrennung_vor"]}
    erste_der_zeilen = {id(z.tokens[0]) for z in zeilen[1:]}
    out: list[PdfToken] = []
    for t in tokens:
        if out and id(t) in erste_der_zeilen:
            v = out[-1]
            if (v.status == KEEP and t.status == KEEP and len(v.old) > 2 and v.old.endswith("-")
                    and t.old[:1].islower() and norm(t.old) not in ausnahmen):
                bbox = (v.bbox[0], min(v.bbox[1], t.bbox[1]), max(v.bbox[2], t.bbox[2]), max(v.bbox[3], t.bbox[3]))
                out[-1] = PdfToken(old=v.old[:-1] + t.old, new=v.new[:-1] + t.new, status=KEEP, page=v.page, bbox=bbox)
                continue
        out.append(t)
    return out


# ---------------------------------------------------------------- Einstieg
def analysiere_pdf(pfad: str | Path, cfg: dict | None = None) -> PdfAnalyse:
    cfg = cfg or config.lade("farben")
    pfad = Path(pfad)
    ergebnis = PdfAnalyse(pfad=str(pfad), seiten=0, blocks=[])
    alle_zeilen: list[list[_Zeile]] = []
    seiten_h: dict[int, float] = {}
    text_farben: Counter = Counter()
    rect_farben: Counter = Counter()
    unbekannt: Counter = Counter()
    ohne_text: list[int] = []
    with pdfplumber.open(str(pfad)) as pdf:
        ergebnis.seiten = len(pdf.pages)
        for nr, page in enumerate(pdf.pages, 1):
            seiten_h[nr] = float(page.height)
            ergebnis.seiten_masse[nr] = (float(page.width), float(page.height))
            chars = [c for c in page.chars if c.get("text") and c["text"].strip()]
            if not chars:
                ohne_text.append(nr)
                alle_zeilen.append([])
                continue
            m = _Markierer(page, cfg)
            rect_farben.update(m.rect_farben)
            zs = []
            for c in chars:
                rgb = _rgb(c.get("non_stroking_color"))
                k = klasse(rgb, cfg)
                if rgb is not None:
                    text_farben[(_hex(rgb), k)] += 1
                    if k is None and colorsys.rgb_to_hsv(*(max(0, min(1, v)) for v in rgb))[1] >= 0.35:
                        unbekannt[_hex(rgb)] += 1
                zs.append(_Z(text=c["text"], x0=c["x0"], x1=c["x1"], top=c["top"], bottom=c["bottom"],
                             size=float(c.get("size") or 10), font=str(c.get("fontname") or ""),
                             status=m.status(c)))
            alle_zeilen.append(_zeilen(zs, nr, cfg))
    _zu_bloecken(alle_zeilen, seiten_h, cfg, ergebnis)
    # Farbuebersicht fuer das Dashboard (zum Kalibrieren)
    for (hx, k), n in text_farben.most_common():
        ergebnis.farben.append({"farbe": hx, "zeichen": n, "klasse": k or "-", "art": "Text",
                                "bedeutung": cfg["legende"].get(k, "-") if k else "keine Markierung"})
    for (hx, k), n in rect_farben.most_common():
        ergebnis.farben.append({"farbe": hx, "zeichen": n, "klasse": k or "-", "art": "Hinterlegung (Rechtecke)",
                                "bedeutung": cfg["legende"].get(k, "-") if k else "keine Markierung"})
    for nr in ohne_text:
        ergebnis.warnungen.append(f"Seite {nr} enthält keinen auslesbaren Text (Scan?). Sie wird nicht ausgewertet.")
    for hx, n in unbekannt.most_common(5):
        ergebnis.warnungen.append(f"Farbe {hx} ({n} Zeichen) ist nicht in der Legende (config/farben.yaml) und wird als unverändert behandelt.")
    if not any(b.hat_aenderung for b in ergebnis.blocks) and not ergebnis.ignoriert:
        ergebnis.warnungen.append("Im PDF wurden keine Änderungen erkannt. Bitte Farben im PDF prüfen (config/farben.yaml).")
    return ergebnis
