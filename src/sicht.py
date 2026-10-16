"""SICHT: ein zweites, unabhaengiges Paar Augen fuer den PDF-Leser (ADR 0004, Absatz-Sperre).

Der Textleser (pdfplumber/pdfminer) kann Farben uebersehen: Text in DeviceN-Farbe mit mehreren Kanaelen, Text nur mit
farbiger Kontur, Markierungen als Kurven-Flaeche statt Rechteck ... Damit so etwas nie still verloren geht, prueft dieses
Modul jede Seite ein zweites Mal mit pdfium (pypdfium2) -- ohne neue Pakete, ohne numpy:

  C1 zweiter Leser   Fuell-/Konturfarbe und Darstellungsmodus je Zeichen aus pdfium. Wo pdfplumber keine verlaessliche
                     Farbe hat (keine, Sonderfarbe, DeviceN mit mehreren Kanaelen, nur Kontur, unsichtbar), gilt pdfiums
                     Farbe. Sagt einer klar blau/rot und der andere klar unmarkiert: Hinweis `farbe_uneinig` am Wort.
  C3 Farb-Inventur   Die Seite wird gerendert; kraeftige Farbpixel, die keine erkannte Markierung, kein farbiger Text,
                     kein Link, kein Bild, keine Gestaltung erklaert, sind "unerklaerte Farbe". Liegt sie im Text:
                     Hinweis `tinte_unerklaert` an den Woertern (der Planer sperrt den Absatz); sonst / bei einem Absatz
                     ohne erkannte Aenderung: Luecke `unerklaerte_markierung`.

Alle Schwellen stehen in config/farben.yaml (Abschnitt `sicht:`). Fehler in diesem Modul brechen das Lesen nie ab;
sie werden als Warnung gemeldet (dann fehlt nur die zweite Pruefung).
"""
from __future__ import annotations

import colorsys
import ctypes
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .models import luecke
from .pdf_reader import _block_hinweise, _bedeutung, _hex, _unbekannt_art, klasse

Box = tuple[float, float, float, float]

# Darstellungsmodi (PDF Tr): 0 fuellen, 1 nur Kontur, 2 beides, 3 unsichtbar, 4-7 wie 0-3 plus Beschneiden
_NUR_KONTUR = (1, 5)
_UNSICHTBAR = (3, 7)
_BEIDES = (2, 6)


def _sat(rgb: tuple[float, float, float] | None) -> tuple[float, float]:
    """(Saettigung, Helligkeit) einer Farbe (0..1)."""
    if rgb is None:
        return 0.0, 0.0
    _, s, v = colorsys.rgb_to_hsv(*(max(0.0, min(1.0, x)) for x in rgb))
    return s, v


def _hex_rgb(hx: str) -> tuple[float, float, float] | None:
    if not hx or not hx.startswith("#") or len(hx) != 7:
        return None
    return tuple(int(hx[i:i + 2], 16) / 255 for i in (1, 3, 5))  # type: ignore[return-value]


def _schneiden(a: Box, b: Box, rand: float = 0.0) -> bool:
    return a[0] - rand < b[2] and b[0] - rand < a[2] and a[1] - rand < b[3] and b[1] - rand < a[3]


def _zurueck(rahmen: Any, box: Box) -> Box:
    """Umkehrung von pdf_reader._Rahmen.box: Kasten in Leserichtung -> Kasten wie pdfplumber ihn meldet (Anzeige)."""
    art = getattr(rahmen, "art", "h")
    x0, t, x1, b = box
    w, h = getattr(rahmen, "w", 0.0), getattr(rahmen, "h", 0.0)
    if art == "ab":
        return (w - b, x0, w - t, x1)
    if art == "auf":
        return (t, h - x1, b, h - x0)
    if art == "um":
        return (w - x1, h - b, w - x0, h - t)
    return box


def _umriss(box: Box, d: float) -> list[Box]:
    """Die vier Randstreifen eines Kastens (Breite d)."""
    x0, t, x1, b = box
    return [(x0 - d, t - d, x1 + d, t + d), (x0 - d, b - d, x1 + d, b + d),
            (x0 - d, t - d, x0 + d, b + d), (x1 - d, t - d, x1 + d, b + d)]


def entscheide(z: Any, fuell: tuple[float, float, float] | None, kontur: tuple[float, float, float] | None,
               modus: int, cfg: dict, uneinig_s: float = 0.15) -> str:
    """C1-Regeln fuer EIN Zeichen (z = pdf_reader._Z, wird angepasst). pdfium liefert Fuell-/Konturfarbe und Modus.

    * pdfplumber hat keine verlaessliche Farbe (keine, Sonderfarbe/DeviceN mit mehreren Kanaelen, nur Kontur) -> pdfiums
      Farbe gilt (Kontur bei Modus 1/5, sonst Fuellung); Rueckgabe 'uebernommen', wenn sich die Bedeutung aendert.
    * pdfplumber haelt das Zeichen fuer unsichtbar, pdfium zeichnet es farbig (blau/rot) -> nur Hinweis (die Zeile bleibt
      unsichtbar, damit sich der Aufbau nicht still aendert).
    * Widerspruch: einer klar blau/rot, der andere klar unmarkiert (Saettigung < uneinig_s); oder Modus 2/6 (Fuellung
      und Kontur) mit Kontur in anderer Farbklasse als die Fuellung -> Hinweis 'farbe_uneinig'; Rueckgabe 'uneinig'.
    """
    if modus in _UNSICHTBAR:
        return ""                                       # pdfium zeichnet das Zeichen nicht: keine Farbe
    p_rgb = kontur if modus in _NUR_KONTUR else fuell
    if p_rgb is None:
        return ""
    k_p = klasse(p_rgb, cfg)
    if z.unsichtbar:
        if k_p is not None:
            z.hinweis.add("farbe_uneinig")
            return "uneinig"
        return ""
    if z.hex in ("", "sonderfarbe") or modus in _NUR_KONTUR:
        vorher = z.farbe_bed
        z.hex = _hex(p_rgb)
        z.weiss = min(p_rgb) >= 0.97
        z.farbklasse = k_p
        z.farbe_bed = _bedeutung(k_p, cfg)
        z.unb = _unbekannt_art(p_rgb, cfg) if k_p is None else ""
        return "uebernommen" if z.farbe_bed != vorher else ""
    k_pl = z.farbklasse
    uneinig = (k_pl is not None and k_p is None and _sat(p_rgb)[0] < uneinig_s) or \
              (k_p is not None and k_pl is None and _sat(_hex_rgb(z.hex))[0] < uneinig_s)
    if not uneinig and modus in _BEIDES and kontur is not None:
        k_k = klasse(kontur, cfg)
        uneinig = k_k is not None and k_k != k_pl
    if uneinig:
        z.hinweis.add("farbe_uneinig")
        return "uneinig"
    return ""


class _Raster:
    """Einfaches Gitter fuer schnelle 'liegt etwas hier?'-Abfragen."""

    def __init__(self, boxen: list[Box], zelle: float = 12.0) -> None:
        self.z = zelle
        self.boxen = boxen
        self.netz: dict[tuple[int, int], list[int]] = {}
        for i, (x0, t, x1, b) in enumerate(boxen):
            for gx in range(int(x0 // zelle), int(x1 // zelle) + 1):
                for gy in range(int(t // zelle), int(b // zelle) + 1):
                    self.netz.setdefault((gx, gy), []).append(i)

    def treffer(self, box: Box, rand: float = 0.0) -> list[int]:
        x0, t, x1, b = box
        z = self.z
        ids: set[int] = set()
        for gx in range(int((x0 - rand) // z), int((x1 + rand) // z) + 1):
            for gy in range(int((t - rand) // z), int((b + rand) // z) + 1):
                ids.update(self.netz.get((gx, gy), ()))
        return [i for i in ids if _schneiden(self.boxen[i], box, rand)]


@dataclass
class _Seite:
    """Was SICHT von pdfplumber ueber eine Seite braucht (beim Lesen gemerkt, die Seite ist danach geschlossen)."""
    nr: int
    ursprung: tuple[float, float]               # Ecke der sichtbaren Seite (CropBox) in pdfplumber-Koordinaten
    drehung: int
    breite: float
    hoehe: float
    erklaert: list[Box] = field(default_factory=list)       # Grafik, die SICHT nicht als unerklaert zaehlt
    verdaechtig: list[Box] = field(default_factory=list)    # gefuellte Pfade auf Text, die der Leser nicht auswertet
    farbtext: list[Box] = field(default_factory=list)       # Zeichen, die pdfplumber farbig sieht (auch quer stehende)
    abgleich: bool = False                                   # C1 konnte die Zeichen zuordnen (Koordinaten passen)


class Sicht:
    """Zweiter Leser (C1) und Farb-Inventur (C3) fuer EIN PDF. Ablauf (aus pdf_reader.analysiere_pdf):
    merke_seite() je Seite waehrend pdfplumber offen ist, zweiter_leser() vor der Statusbewertung, inventur() am Ende."""

    def __init__(self, pfad: str | Path, cfg: dict) -> None:
        self.pfad = str(pfad)
        self.cfg = cfg
        self.sc = cfg.get("sicht", {}) or {}
        self.seiten: dict[int, _Seite] = {}
        self.warnungen: list[str] = []
        self.dauer: dict[str, float] = {"leser": 0.0, "inventur": 0.0}
        self.zaehler: dict[str, int] = {"uneinig": 0, "uebernommen": 0, "tinte_woerter": 0, "tinte_stellen": 0}
        self._pdf: Any = None

    # ------------------------------------------------------------------ allgemein
    def _teil(self, name: str) -> dict:
        return self.sc.get(name, {}) or {}

    def _doc(self) -> Any:
        if self._pdf is None:
            import pypdfium2 as pdfium
            self._pdf = pdfium.PdfDocument(self.pfad)
        return self._pdf

    def schliessen(self) -> None:
        if self._pdf is not None:
            try:
                self._pdf.close()
            except Exception:       # pragma: no cover
                pass
            self._pdf = None

    # ------------------------------------------------------------------ waehrend pdfplumber liest
    def merke_seite(self, page: Any, sd: Any) -> None:
        """Merkt Seitenlage und die Grafik der Seite (Rechtecke, Linien, Kurven) in Anzeige-Koordinaten."""
        try:
            self._merke(page, sd)
        except Exception as e:      # pragma: no cover - darf das Lesen nie stoppen
            self.warnungen.append(f"Seite {sd.nr}: Die zweite Prüfung (Seitenbild) konnte die Seite nicht vorbereiten ({type(e).__name__}).")

    def _merke(self, page: Any, sd: Any) -> None:
        from .pdf_reader import _objektfarbe, _Sonderfarben, _zeichenfarbe
        cb = getattr(page, "cropbox", None) or page.bbox
        s = _Seite(nr=sd.nr, ursprung=(float(cb[0]), float(cb[1])), drehung=int(getattr(page, "rotation", 0) or 0),
                   breite=float(page.width), hoehe=float(page.height))
        self.seiten[sd.nr] = s
        sonder = _Sonderfarben(page)
        inv = self._teil("inventur")
        dicke = float(inv.get("linie_max_dicke_pt", 2.0))
        gross = float(self.cfg["hintergrund"]["max_flaeche_anteil"]) * s.breite * s.hoehe
        text = [(float(c["x0"]), float(c["top"]), float(c["x1"]), float(c["bottom"]))
                for c in page.chars if (c.get("text") or "").strip()]
        for c in page.chars:
            rgb, unklar = _zeichenfarbe(c, sonder)
            sat, hell = _sat(rgb)
            if unklar or (sat >= float(inv.get("min_saettigung", 0.25)) and hell >= float(inv.get("min_helligkeit", 0.16))):
                s.farbtext.append((float(c["x0"]), float(c["top"]), float(c["x1"]), float(c["bottom"])))
        raster = _Raster(text)
        for art, objekte in (("rect", page.rects), ("curve", page.curves), ("line", page.lines)):
            for o in objekte:
                box = (float(o["x0"]), float(o["top"]), float(o["x1"]), float(o["bottom"]))
                bw, bh = box[2] - box[0], box[3] - box[1]
                if art == "line" or min(bw, bh) <= dicke:
                    s.erklaert.append(box)          # Linien: der Leser hat sie gesehen und beurteilt (Strich, Rahmen, Balken)
                    continue
                if not raster.treffer(box):
                    s.erklaert.append(box)          # Grafik neben dem Text (Logo, Kasten, Diagramm)
                    continue
                rgb, unklar = _objektfarbe(o, not o.get("fill"), sonder)
                bunt = unklar or _sat(rgb)[0] >= float(self.cfg["hintergrund"]["min_saettigung"])
                if bw * bh > gross:
                    if bunt:
                        s.erklaert.append(box)      # farbiger Seitenhintergrund (Gestaltung)
                    continue
                if not o.get("fill"):
                    s.erklaert += _umriss(box, dicke)   # nur Umriss (Kasten um den Text): nur der Rand ist erklaert
                    continue
                if art == "rect":
                    if bunt:
                        s.erklaert.append(box)      # farbige Rechtecke auf Text hat der Markierer beurteilt
                    continue                        # pdfplumber sieht eine graue/weisse Flaeche -- sieht sie bunt aus, ist das unerklaert
                s.verdaechtig.append(box)           # gefuellte Kurve auf Text: der Leser wertet sie nicht aus

    # ------------------------------------------------------------------ C1: zweiter Leser
    def zweiter_leser(self, seiten_daten: list[Any]) -> None:
        import time
        t0 = time.perf_counter()
        try:
            if self._teil("leser").get("aktiv", True):
                for sd in seiten_daten:
                    if sd.nr in self.seiten and any(not z.leer for z in sd.zeichen):
                        self._leser_seite(sd)
        except Exception as e:      # pragma: no cover
            self.warnungen.append(f"Die zweite Farbprüfung (pdfium) ist fehlgeschlagen ({type(e).__name__}); "
                                  f"Farben stammen nur aus dem Textleser.")
        self.dauer["leser"] += time.perf_counter() - t0

    def _pdfium_zeichen(self, nr: int, s: _Seite) -> list[tuple[str, float, float, Any, Any, int]]:
        """Zeichen der Seite aus pdfium: (Text, Mitte x, Mitte top (Anzeige wie pdfplumber), Fuellfarbe, Konturfarbe, Modus)."""
        import pypdfium2.raw as R
        page = self._doc()[nr - 1]
        tp = page.get_textpage()
        try:
            n = R.FPDFText_CountChars(tp.raw)
            # Seitenraum -> Anzeige: ueber FPDF_PageToDevice (beruecksichtigt Drehung und CropBox), fein aufgeloest
            k = 64
            breite, hoehe = page.get_size()          # schon gedreht (Anzeige), wie die gerenderte Seite
            sx, sy = max(1, round(breite * k)), max(1, round(hoehe * k))
            dx, dy = ctypes.c_int(), ctypes.c_int()

            def punkt(px: float, py: float) -> tuple[float, float]:
                R.FPDF_PageToDevice(page.raw, 0, 0, sx, sy, 0, ctypes.c_double(px), ctypes.c_double(py),
                                    ctypes.byref(dx), ctypes.byref(dy))
                return dx.value / k, dy.value / k

            o0 = punkt(0.0, 0.0)
            ox = punkt(1000.0, 0.0)
            oy = punkt(0.0, 1000.0)
            a, b = (ox[0] - o0[0]) / 1000, (ox[1] - o0[1]) / 1000
            c, d = (oy[0] - o0[0]) / 1000, (oy[1] - o0[1]) / 1000
            ux, uy = s.ursprung
            l_, r_, b_, t_ = (ctypes.c_double() for _ in range(4))
            fr, fg, fb, fa = (ctypes.c_uint() for _ in range(4))
            modi: dict[int, int] = {}
            out = []
            for i in range(n):
                u = R.FPDFText_GetUnicode(tp.raw, i)
                if u in (0, 0xFFFE, 0xFFFF) or not chr(u).strip() or R.FPDFText_IsGenerated(tp.raw, i) == 1:
                    continue
                if not R.FPDFText_GetCharBox(tp.raw, i, ctypes.byref(l_), ctypes.byref(r_), ctypes.byref(b_), ctypes.byref(t_)):
                    continue
                mx, my = (l_.value + r_.value) / 2, (b_.value + t_.value) / 2
                x = o0[0] + a * mx + c * my + ux
                y = o0[1] + b * mx + d * my + uy
                fuell = kontur = None
                if R.FPDFText_GetFillColor(tp.raw, i, ctypes.byref(fr), ctypes.byref(fg), ctypes.byref(fb), ctypes.byref(fa)):
                    fuell = (fr.value / 255, fg.value / 255, fb.value / 255)
                if R.FPDFText_GetStrokeColor(tp.raw, i, ctypes.byref(fr), ctypes.byref(fg), ctypes.byref(fb), ctypes.byref(fa)):
                    kontur = (fr.value / 255, fg.value / 255, fb.value / 255)
                obj = R.FPDFText_GetTextObject(tp.raw, i)
                modus = 0
                if obj:
                    key = ctypes.cast(obj, ctypes.c_void_p).value or 0
                    if key not in modi:
                        modi[key] = int(R.FPDFTextObj_GetTextRenderMode(obj))
                    modus = modi[key]
                out.append((chr(u), x, y, fuell, kontur, modus))
            return out
        finally:
            tp.close()
            page.close()

    def _leser_seite(self, sd: Any) -> None:
        cfg = self.cfg
        le = self._teil("leser")
        s = self.seiten[sd.nr]
        pz = self._pdfium_zeichen(sd.nr, s)
        tol = float(le.get("box_toleranz_pt", 0.75))
        uneinig_s = float(le.get("uneinig_max_saettigung", 0.15))
        netz: dict[tuple[int, int], list[int]] = {}
        for i, (_, x, y, _, _, _) in enumerate(pz):
            netz.setdefault((int(x // 10), int(y // 10)), []).append(i)
        belegt: set[int] = set()
        paare: list[tuple[Any, tuple]] = []
        sichtbar = [z for z in sd.zeichen if not z.leer]
        for z in sichtbar:
            x0, t, x1, b = _zurueck(sd.rahmen, (z.x0, z.top, z.x1, z.bottom))
            kand = []
            for gx in range(int((x0 - tol) // 10), int((x1 + tol) // 10) + 1):
                for gy in range(int((t - tol) // 10), int((b + tol) // 10) + 1):
                    for i in netz.get((gx, gy), ()):
                        _, x, y, _, _, _ = pz[i]
                        if i not in belegt and x0 - tol <= x <= x1 + tol and t - tol <= y <= b + tol:
                            kand.append(i)
            if not kand:
                continue
            mx, my = (x0 + x1) / 2, (t + b) / 2
            kand.sort(key=lambda i: (pz[i][0] != z.text[:1], (pz[i][1] - mx) ** 2 + (pz[i][2] - my) ** 2))
            belegt.add(kand[0])
            paare.append((z, pz[kand[0]]))
        if not sichtbar or len(paare) < float(le.get("min_abgleich", 0.8)) * len(sichtbar):
            self.warnungen.append(f"Seite {sd.nr}: Die zweite Farbprüfung (pdfium) konnte die Zeichen nicht zuordnen; "
                                  f"Farben stammen auf dieser Seite nur aus dem Textleser.")
            return
        s.abgleich = True
        for z, (_, _, _, fuell, kontur, modus) in paare:
            erg = entscheide(z, fuell, kontur, modus, cfg, uneinig_s)
            if erg:
                self.zaehler[erg] += 1

    # ------------------------------------------------------------------ C3: Farb-Inventur
    def inventur(self, ergebnis: Any, seiten_daten: list[Any], markierer: dict[int, Any]) -> None:
        import time
        t0 = time.perf_counter()
        try:
            if self._teil("inventur").get("aktiv", True):
                for sd in seiten_daten:
                    s = self.seiten.get(sd.nr)
                    m = markierer.get(sd.nr)
                    if s is None or m is None:
                        continue                         # Seite ohne Text: Scan-/Bild-Luecken meldet der Leser selbst
                    if not s.abgleich:
                        continue                         # Koordinaten nicht gesichert (Warnung steht schon)
                    self._inventur_seite(ergebnis, sd, s, m)
        except Exception as e:      # pragma: no cover
            self.warnungen.append(f"Die Farb-Inventur (Seitenbild) ist fehlgeschlagen ({type(e).__name__}).")
        finally:
            self.schliessen()
        self.dauer["inventur"] += time.perf_counter() - t0
        ergebnis.warnungen += [w for w in self.warnungen if w not in ergebnis.warnungen]

    def _erklaert(self, ergebnis: Any, sd: Any, s: _Seite, m: Any) -> list[Box]:
        """Alle Bereiche, deren Farbe erklaert ist (Anzeige-Koordinaten)."""
        inv = self._teil("inventur")
        r = sd.rahmen
        lesen: list[Box] = []        # Kaesten in Leserichtung
        smin, vmin = float(inv.get("min_saettigung", 0.25)), float(inv.get("min_helligkeit", 0.16))
        for z in sd.zeichen:
            if z.leer or z.unsichtbar:
                continue
            if z.farbklasse or z.unb or z.hex == "sonderfarbe":
                lesen.append((z.x0, z.top, z.x1, z.bottom))
                continue
            sat, hell = _sat(_hex_rgb(z.hex))
            if sat >= smin and hell >= vmin:
                lesen.append((z.x0, z.top, z.x1, z.bottom))
        for liste in (m.flaechen, m.unb_flaechen, m.annot, m.unb_annot):
            lesen += [tuple(x[:4]) for x in liste]     # type: ignore[misc]
        lesen += list(m.design) + list(m.links)
        for k in m.andere:
            lesen += list(k["boxen"])
        for x0, x1, y, *_ in list(m._linien_roh) + list(m.striche) + list(m.unterstriche):
            lesen.append((x0, y - 1.5, x1, y + 1.5))
        lesen += list(sd.bilder)
        for _, seite, boxen in ergebnis.regionen:
            if seite == sd.nr:
                lesen += list(boxen)
        anteil = float(self.cfg["layout"]["kopf_fuss_anteil"])
        lesen.append((-1e4, -1e4, 1e4, anteil * sd.hoehe))
        lesen.append((-1e4, (1 - anteil) * sd.hoehe, 1e4, 1e4))
        anzeige = [_zurueck(r, b) for b in lesen]
        return anzeige + s.erklaert + s.farbtext

    def _inventur_seite(self, ergebnis: Any, sd: Any, s: _Seite, m: Any) -> None:
        from PIL import Image, ImageChops, ImageDraw
        inv = self._teil("inventur")
        dpi = float(inv.get("dpi", 100))
        f = dpi / 72.0
        page = self._doc()[sd.nr - 1]
        try:
            bild = page.render(scale=f, draw_annots=True).to_pil().convert("RGB")
        finally:
            page.close()
        _, sat, hell = bild.convert("HSV").split()
        s_min = round(float(inv.get("min_saettigung", 0.25)) * 255)
        v_min = round(float(inv.get("min_helligkeit", 0.16)) * 255)
        maske = ImageChops.multiply(sat.point(lambda x: 255 if x >= s_min else 0),
                                    hell.point(lambda x: 255 if x >= v_min else 0))
        self.zaehler["farb_px"] = self.zaehler.get("farb_px", 0) + maske.histogram()[255]
        if maske.getbbox() is None:
            return
        erkl = Image.new("L", maske.size, 0)
        zeichne = ImageDraw.Draw(erkl)
        rand = float(inv.get("rand_px", 2))
        ux, uy = s.ursprung
        for x0, t, x1, b in self._erklaert(ergebnis, sd, s, m):
            zeichne.rectangle(((x0 - ux) * f - rand, (t - uy) * f - rand, (x1 - ux) * f + rand, (b - uy) * f + rand), fill=255)
        rest = ImageChops.subtract(maske, erkl)
        self.zaehler["rest_px"] = self.zaehler.get("rest_px", 0) + rest.histogram()[255]
        if rest.getbbox() is None:
            return
        zp = max(1, int(inv.get("zelle_px", 6)))
        netz = rest.reduce(zp)
        gw, gh = netz.size
        schwelle = float(inv.get("zelle_min_anteil", 0.08)) * 255
        daten = netz.tobytes()
        an = {i for i, v in enumerate(daten) if v >= schwelle}
        stellen: list[Box] = []
        min_zellen = int(inv.get("min_zellen", 3))
        while an:
            start = an.pop()
            gruppe = [start]
            stapel = [start]
            while stapel:
                i = stapel.pop()
                x, y = i % gw, i // gw
                for dx_ in (-1, 0, 1):
                    for dy_ in (-1, 0, 1):
                        nx, ny = x + dx_, y + dy_
                        j = ny * gw + nx
                        if 0 <= nx < gw and 0 <= ny < gh and j in an:
                            an.discard(j)
                            gruppe.append(j)
                            stapel.append(j)
            if len(gruppe) < min_zellen:
                continue
            gx0 = min(i % gw for i in gruppe) * zp
            gy0 = min(i // gw for i in gruppe) * zp
            gx1 = (max(i % gw for i in gruppe) + 1) * zp
            gy1 = (max(i // gw for i in gruppe) + 1) * zp
            bb = rest.crop((gx0, gy0, gx1, gy1)).getbbox()
            if bb is None:
                continue
            stellen.append(((gx0 + bb[0]) / f + ux, (gy0 + bb[1]) / f + uy, (gx0 + bb[2]) / f + ux, (gy0 + bb[3]) / f + uy))
        if not stellen:
            return
        # Aenderungsbalken: schmale senkrechte Striche neben dem Text
        text_boxen = [_zurueck(sd.rahmen, (z.x0, z.top, z.x1, z.bottom)) for z in sd.zeichen if not z.leer]
        raster = _Raster(text_boxen)
        b_breite = float(inv.get("balken_max_breite_pt", 3.0))
        b_hoehe = float(inv.get("balken_min_hoehe_pt", 8.0))
        echte = [st for st in stellen
                 if not (st[2] - st[0] <= b_breite and st[3] - st[1] >= b_hoehe and not raster.treffer(st))]
        if echte:
            self._melden(ergebnis, sd, [sd.rahmen.box(*st) for st in echte])

    def _melden(self, ergebnis: Any, sd: Any, stellen: list[Box]) -> None:
        """Unerklaerte Farbe: Hinweis an den Woertern im Text, sonst Luecke."""
        bloecke = [b for b in ergebnis.blocks if b.page == sd.nr]
        ohne_text = 0
        ohne_aenderung_woerter = 0
        ohne_aenderung_stellen = 0
        for st in stellen:
            getroffen = [b for b in bloecke if _schneiden(b.bbox, st, 1.0)]
            if not getroffen:
                ohne_text += 1
                continue
            self.zaehler["tinte_stellen"] += 1
            neu_gesetzt = 0
            for b in getroffen:
                toks = [t for t in b.tokens if _schneiden(t.bbox, st, 2.0)] or list(b.tokens)
                for t in toks:
                    codes = set(filter(None, t.hinweis.split(";")))
                    if "tinte_unerklaert" not in codes:
                        t.hinweis = ";".join(sorted(codes | {"tinte_unerklaert"}))
                        neu_gesetzt += 1
                b.hinweise += [h for h in _block_hinweise(b.tokens) if h not in b.hinweise]
                if not b.hat_aenderung:
                    ohne_aenderung_woerter += len(toks)
            self.zaehler["tinte_woerter"] += neu_gesetzt
            if not any(b.hat_aenderung for b in getroffen):
                ohne_aenderung_stellen += 1
        if ohne_aenderung_stellen:
            ergebnis.luecken.append(luecke(
                "unerklaerte_markierung", sd.nr,
                f"Seite {sd.nr}: Im Seitenbild ist im Text Farbe zu sehen, die keiner erkannten Markierung entspricht "
                f"({ohne_aenderung_stellen} Stelle(n)). Der Text dort gilt als unverändert. Bitte im PDF prüfen.",
                ohne_aenderung_woerter))
        if ohne_text:
            ergebnis.luecken.append(luecke(
                "unerklaerte_markierung", sd.nr,
                f"Seite {sd.nr}: Im Seitenbild ist neben dem Text Farbe zu sehen, die keiner erkannten Markierung "
                f"entspricht ({ohne_text} Stelle(n)). Bitte im PDF prüfen.", 0))


def bereit(cfg: dict) -> bool:
    """Ist SICHT eingeschaltet (config farben.yaml sicht.aktiv)?"""
    return bool((cfg.get("sicht") or {}).get("aktiv"))

