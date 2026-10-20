"""Laeufe: ein Lauf = ein PDF + eine Excel, analysiert und (nach Freigabe) uebernommen.

Ordner je Lauf (data/runs/<id>/), append-only (Regel 29):
  eingabe/            Kopien von Aenderungsdokument (Word .docx oder PDF) und Excel (Originale bleiben unberuehrt)
  lauf.json           Namen der Eingaben (wird zuletzt, atomar geschrieben)
  plan.json           Analyse (Aenderungen, Bilanz, Aufbau, Farben, quelle_art, stichprobe), atomar geschrieben
  entscheidungen.jsonl Freigabe-Protokoll (Vertrauensmodus): eine Zeile je Entscheidung, nur Kennungen, nur angehaengt
  analyse_fehler.json Nur wenn die Analyse fehlschlug: einfache Meldung, Lauf und Eingaben bleiben
  archiviert.json     Nur wenn der Lauf archiviert ist (Marker; Wiederherstellen benennt ihn um)
  ergebnis_001/       je Uebernahme ein NEUER Ordner: aktualisierte Excel + Protokoll;
                      ein Ergebnis zaehlt erst, wenn ergebnis.json (zuletzt, atomar) da ist
Nichts wird ueberschrieben oder geloescht. Einzige Ausnahme: halbfertige Ordner, die ein Vorgang
selbst gerade angelegt hat und die bei seinem Abbruch noch keinen Inhalt von Wert haben.
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import shutil
import time
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from . import config
from .applier import uebernehmen
from . import konsens, layouts, vertrauen
from .fehler import NichtGefunden, NutzerFehler
from .gegenprobe import absaetze_aus_liste
from .excel_model import ExcelModell, profil_kontext
from .models import GELESENE_SCHEMAS, PRUEFEN, UEBERNOMMEN
from .pdf_reader import analysiere_pdf
from .planner import plane
from .protokoll import (ENTSCHEIDUNGEN, aenderung_aus_dict, bilanz, entscheidung_zeile, haenge_entscheidungen_an,
                        schreibe_json, schreibe_xlsx)

_ID = re.compile(r"^\d{8}-\d{6}-[0-9a-f]{6}$")
_ERGEBNIS = re.compile(r"ergebnis_(\d{3})")
# Windows kann diese Namen (auch mit Endung) nicht als Datei anlegen
_RESERVIERT = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
_PLAN_PFLICHT = ("aenderungen", "bilanz", "excel_sha256", "pdf_name", "excel_name")
_VERSUCHE = 50                                   # Wiederholungen bei Namenskollisionen

log = logging.getLogger("cde.runs")

MELDUNG_BESCHAEDIGT = "Der Lauf ist beschädigt."
MELDUNG_NEUERE_VERSION = ("Dieser Lauf stammt aus einer neueren Programmversion (Schema-Version {v}). "
                          "Bitte das Programm aktualisieren.")


class LaufFehler(NutzerFehler):
    """Ein Lauf kann nicht verwendet werden. `meldung` ist ein einfacher deutscher Satz fuer den Nutzer.
    Nur NutzerFehler (und damit LaufFehler) duerfen ihren Text an den Client weitergeben (src/fehler.py)."""

    def __init__(self, meldung: str, status: int = 409, art: str = "fehler"):
        super().__init__(meldung, status, art)


def _beschaedigt() -> LaufFehler:
    return LaufFehler(MELDUNG_BESCHAEDIGT, 422, "beschaedigt")


def runs_dir() -> Path:
    d = config.pfad("runs_dir")
    d.mkdir(parents=True, exist_ok=True)
    return d


def lauf_dir(run_id: str) -> Path:
    if not _ID.fullmatch(run_id):
        raise NutzerFehler("Ungültige Lauf-ID.", 404)
    d = runs_dir() / run_id
    if not d.is_dir():
        raise NichtGefunden("Lauf nicht gefunden.")
    return d


def _sha(pfad: Path) -> str:
    h = hashlib.sha256()
    with open(pfad, "rb") as f:
        for stueck in iter(lambda: f.read(1 << 20), b""):
            h.update(stueck)
    return h.hexdigest()


# ---------------------------------------------------------------------------------------------
# Atomares Schreiben und Lesen
# ---------------------------------------------------------------------------------------------
def _ersetze(tmp: Path, ziel: Path) -> None:
    """os.replace mit kurzer Wiederholung (Windows: das Ziel kann kurz von einem Leser gehalten werden)."""
    for versuch in range(5):
        try:
            os.replace(tmp, ziel)
            return
        except PermissionError:
            if versuch == 4:
                raise
            time.sleep(0.05 * (versuch + 1))


def _atomar(ziel: Path, schreibe) -> None:
    """Schreibt erst in eine Temp-Datei im selben Ordner und ersetzt dann in einem Schritt.
    Ein Leser sieht so nie eine halbe Datei."""
    tmp = ziel.with_name(f".{ziel.name}.{uuid.uuid4().hex[:8]}.tmp")
    try:
        schreibe(tmp)
        with open(tmp, "rb+") as f:
            os.fsync(f.fileno())
        _ersetze(tmp, ziel)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise


def _schreibe_json_atomar(ziel: Path, daten: dict[str, Any]) -> None:
    """Wie protokoll.schreibe_json (mit schema_version), aber atomar."""
    _atomar(ziel, lambda tmp: schreibe_json(tmp, daten))


def _schreibe_text_atomar(ziel: Path, text: str) -> None:
    _atomar(ziel, lambda tmp: tmp.write_text(text, encoding="utf-8"))


def _lese_json(pfad: Path) -> dict[str, Any]:
    """Liest eine JSON-Datei eines Laufs; unlesbar oder kein Objekt -> LaufFehler 'beschaedigt'."""
    try:
        daten = json.loads(pfad.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        raise _beschaedigt() from None
    if not isinstance(daten, dict):
        raise _beschaedigt()
    return daten


def pruefe_schema(daten: dict[str, Any]) -> int:
    """Prueft schema_version gegen models.GELESENE_SCHEMAS. Fehlt sie, gilt Schema 1."""
    v = daten.get("schema_version", 1)
    if isinstance(v, bool) or not isinstance(v, int):
        raise _beschaedigt()
    if v in GELESENE_SCHEMAS:
        return v
    if v > max(GELESENE_SCHEMAS):
        raise LaufFehler(MELDUNG_NEUERE_VERSION.format(v=v), 409, "neuere_version")
    raise _beschaedigt()


# ---------------------------------------------------------------------------------------------
# Dateinamen
# ---------------------------------------------------------------------------------------------
def pruefe_dateiname(name: str) -> None:
    """Zu lange Namen (Zeichen oder Bytes) werden abgelehnt, bevor irgendetwas angelegt wird."""
    grenze = int(config.lade("server").get("dateiname_max_laenge", 255))
    if len(name) > grenze or len(name.encode("utf-8", "replace")) > grenze:
        raise NutzerFehler(f"Der Dateiname ist zu lang (höchstens {grenze} Zeichen). Bitte die Datei kürzer benennen.", 400)


def _sicherer_name(name: str, fallback: str) -> str:
    name = Path(name or fallback).name
    name = re.sub(r"[^\w.\- äöüÄÖÜß]", "_", name).strip(" .") or fallback
    if name.split(".")[0].rstrip().upper() in _RESERVIERT:
        name = "_" + name
    return name


# ---------------------------------------------------------------------------------------------
# Das Dokument mit den Aenderungen: Word (.docx, bevorzugt) oder PDF
# ---------------------------------------------------------------------------------------------
AENDERUNGS_ENDUNGEN = (".docx", ".pdf")
MELDUNG_ENDUNG = "Die Änderungsdatei muss ein Word-Dokument (.docx) oder ein PDF sein."
MELDUNG_DOC_ALT = "Alte Word-Dateien (.doc) können nicht gelesen werden. Bitte in Word als .docx speichern."
MELDUNG_KEIN_PDF = "Die Datei ist kein gültiges PDF und konnte nicht gelesen werden. Bitte das Dokument noch einmal als PDF speichern."
MELDUNG_KEIN_DOCX = ("Die Datei ist kein gültiges Word-Dokument (.docx) und konnte nicht gelesen werden. Ist sie beschädigt oder mit "
                     "einem Passwort geschützt? Bitte in Word als .docx speichern.")
MELDUNG_KEIN_DOCX_LESER = "Word-Dokumente (.docx) kann diese Programmversion noch nicht lesen. Bitte als PDF speichern."


def quelle_art_von(name: str) -> str:
    """"docx" oder "pdf" nach der Endung des gespeicherten Dateinamens (alles Aeltere ist ein PDF)."""
    return "docx" if Path(name or "").suffix.lower() == ".docx" else "pdf"


def pruefe_aenderungsendung(name: str) -> str:
    """Endung (".docx" oder ".pdf", klein) des Dokuments mit Aenderungen; sonst NutzerFehler (400) in einfachem Deutsch."""
    endung = Path(name or "").suffix.lower()
    if endung == ".doc":
        raise NutzerFehler(MELDUNG_DOC_ALT, 400)
    if endung not in AENDERUNGS_ENDUNGEN:
        raise NutzerFehler(MELDUNG_ENDUNG, 400)
    return endung


def _ist_docx(pfad: Path) -> bool:
    """Ein Word-Dokument ist ein ZIP mit [Content_Types].xml und word/document.xml (nur das Inhaltsverzeichnis wird gelesen)."""
    try:
        with zipfile.ZipFile(pfad) as z:
            return {"[Content_Types].xml", "word/document.xml"} <= set(z.namelist())
    except (zipfile.BadZipFile, OSError):
        return False


def pruefe_aenderungsdatei(pfad: str | Path, name: str | None = None) -> str:
    """Prueft Endung (nach `name`, sonst nach dem Pfad) UND Inhalt des Dokuments mit Aenderungen. Gibt "pdf" oder "docx" zurueck.
    PDF: beginnt mit %PDF-. Word: ZIP mit [Content_Types].xml und word/document.xml. Sonst NutzerFehler (400)."""
    endung = pruefe_aenderungsendung(name if name else Path(pfad).name)
    if endung == ".pdf":
        try:
            with open(pfad, "rb") as f:
                ok = f.read(5) == b"%PDF-"
        except OSError:
            ok = False
        if not ok:
            raise NutzerFehler(MELDUNG_KEIN_PDF, 400)
        return "pdf"
    if not _ist_docx(Path(pfad)):
        raise NutzerFehler(MELDUNG_KEIN_DOCX, 400)
    return "docx"


def lade_analyse(pfad: str | Path):
    """Liest das Dokument mit den Aenderungen und liefert die Analyse (PdfAnalyse): nach Endung .pdf mit pdf_reader, .docx mit
    docx_reader (Import erst hier). Eine Analyse aus Word hat legende_modus "docx", Seite 0 und keine Bildkoordinaten."""
    p = Path(pfad)
    endung = p.suffix.lower()
    if endung == ".pdf":
        return analysiere_pdf(p)
    if endung == ".docx":
        try:
            from .docx_reader import analysiere_docx
        except ModuleNotFoundError as e:
            if e.name != f"{__package__}.docx_reader":
                raise
            raise NutzerFehler(MELDUNG_KEIN_DOCX_LESER, 400) from None
        return analysiere_docx(p)
    raise NutzerFehler(MELDUNG_ENDUNG, 400)


# ---------------------------------------------------------------------------------------------
# Lauf anlegen / analysieren
# ---------------------------------------------------------------------------------------------
def _reserviere(eltern: Path, name_fn) -> Path:
    """Legt einen neuen Ordner mit mkdir OHNE exist_ok an; bei Kollision wird ein neuer Name gebildet."""
    for _ in range(_VERSUCHE):
        ziel = eltern / name_fn()
        try:
            ziel.mkdir()
            return ziel
        except FileExistsError:
            continue
    raise LaufFehler("Gerade wird zu vieles gleichzeitig angelegt. Bitte gleich noch einmal versuchen.", 409)


def neuer_lauf(pdf_pfad: str | Path, excel_pfad: str | Path, pdf_name: str = "", excel_name: str = "") -> str:
    pdf_roh, excel_roh = pdf_name or Path(pdf_pfad).name, excel_name or Path(excel_pfad).name
    pruefe_dateiname(pdf_roh)
    pruefe_dateiname(excel_roh)
    pdf_name = _sicherer_name(pdf_roh, "aenderungen.docx" if Path(pdf_pfad).suffix.lower() == ".docx" else "aenderungen.pdf")
    excel_name = _sicherer_name(excel_roh, "cde.xlsx")
    d = _reserviere(runs_dir(), lambda: f"{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}")
    try:
        (d / "eingabe").mkdir()
        shutil.copy2(pdf_pfad, d / "eingabe" / pdf_name)
        shutil.copy2(excel_pfad, d / "eingabe" / excel_name)
        _schreibe_text_atomar(d / "lauf.json", json.dumps(
            {"pdf": pdf_name, "excel": excel_name, "quelle_art": quelle_art_von(pdf_name),
             "angelegt": datetime.now().isoformat(timespec="seconds")},
            ensure_ascii=False))
    except BaseException:
        shutil.rmtree(d, ignore_errors=True)         # noch kein Lauf (lauf.json fehlt): nichts von Wert
        raise
    return d.name


def _info(d: Path) -> dict[str, Any]:
    info = _lese_json(d / "lauf.json")
    for k in ("pdf", "excel"):
        if not isinstance(info.get(k), str) or not info[k] or Path(info[k]).name != info[k]:
            raise _beschaedigt()
    info.setdefault("angelegt", "")
    info["quelle_art"] = quelle_art_von(info["pdf"])            # aeltere Laeufe kennen das Feld nicht: es sind PDFs
    return info


_SPEC_ERLAUBT = re.compile(r"^[0-9A-Za-z][0-9A-Za-z _.\-]{0,39}$")


def pruefe_spec(spec: str | None) -> str | None:
    """Spec, die der Nutzer vorgibt (Spezifikationsnummer, z. B. "11 22 33"): nur Ziffern, Buchstaben, Leerzeichen, Punkt,
    Unterstrich und Bindestrich, hoechstens 40 Zeichen. Leer = keine Vorgabe (None). Kein Freitext."""
    if spec is None:
        return None
    if not isinstance(spec, str):
        raise NutzerFehler("Die Spec muss ein Text sein.", 400)
    spec = " ".join(spec.split())
    if not spec:
        return None
    if not _SPEC_ERLAUBT.match(spec):
        raise NutzerFehler("Die Spec darf nur Ziffern, Buchstaben, Leerzeichen, Punkt, Unterstrich und Bindestrich enthalten "
                           "(höchstens 40 Zeichen).", 400)
    return spec


def specs_der_cde(xl: ExcelModell) -> list[dict[str, Any]]:
    """Specs, die die CDE liefert: Name (einheitliche Schreibweise), Zahl der Zeilen, erste und letzte Zeile.
    Leer, wenn die Excel keine Spalte mit Spezifikationsnummern hat."""
    je: dict[str, list[int]] = {}
    for blatt in xl.layouts:
        for sp, zeilen in xl.specs(blatt).items():
            je.setdefault(sp, []).extend(zeilen)
    return [{"name": sp, "zeilen": len(z), "von": min(z), "bis": max(z)} for sp, z in sorted(je.items()) if z]


def analysiere(run_id: str, profil: str | None = None, spec: str | None = None) -> dict[str, Any]:
    """Analysiert einen Lauf. `profil` = Name eines Layout-Profils (src/layouts.py): wird nur angenommen, wenn es
    aktiv ist und zu seinem aktuellen Inhalt eine bestaetigte Vorschau gibt (neue Profile gehen nie direkt live).
    `spec` = Spezifikationsnummer, die der Nutzer vorgibt (sonst wird sie aus dem PDF erkannt; wichtig bei einer CDE mit
    mehreren Specs). Eine unbekannte Spec ist kein Fehler: der Plan warnt und traegt nichts automatisch ein."""
    spec = pruefe_spec(spec)
    d = lauf_dir(run_id)
    if (d / "plan.json").exists():                   # Regel 29: eine vorhandene Analyse wird nie ersetzt
        raise LaufFehler("Dieser Lauf wurde schon analysiert. Eine Analyse wird nicht überschrieben. "
                         "Bitte einen neuen Lauf starten.", 409, "schon_analysiert")
    info = _info(d)
    prof = None
    if profil:
        try:
            prof = layouts.bestaetigtes_profil(profil)
        except layouts.ProfilFehler as e:
            raise LaufFehler(e.meldung, 409, "profil") from None
    pdf, excel = d / "eingabe" / info["pdf"], d / "eingabe" / info["excel"]
    pa = lade_analyse(pdf)
    try:
        xl = ExcelModell(excel, prof)
    except layouts.ProfilFehler as e:                # das Profil passt nicht zur Excel: schon einfaches Deutsch
        raise LaufFehler(e.meldung, 409, "profil") from None
    except ValueError as e:
        # Nur zwei bekannte Faelle haben einen Text in einfachem Deutsch (falsche Endung, beschaedigte Datei);
        # jede andere ValueError bleibt unerwartet (Rohtext nur ins Server-Log).
        if excel.suffix.lower() not in (".xlsx", ".xlsm") or isinstance(e.__cause__, zipfile.BadZipFile):
            raise NutzerFehler(str(e), 400) from e
        raise
    plan = plane(pa, xl, spec=spec)
    konsens.bildprobe_anwenden(plan.aenderungen, pa, pdf=pdf)      # zweiter Pruefweg (nur PDF): Abweichung -> "pruefen"
    bil = bilanz(plan.aenderungen, plan.markierte_woerter, plan.ignoriert, plan.luecken, plan.luecken_ausserhalb)
    if (d / "plan.json").exists():                   # gleichzeitige Analyse desselben Laufs: die erste gewinnt
        raise LaufFehler("Dieser Lauf wurde schon analysiert. Eine Analyse wird nicht überschrieben. "
                         "Bitte einen neuen Lauf starten.", 409, "schon_analysiert")
    daten = {
        "run_id": run_id, "pdf_name": info["pdf"], "excel_name": info["excel"],
        "pdf_sha256": _sha(pdf), "excel_sha256": _sha(excel), "seiten": pa.seiten,
        "bilanz": bil, "layout": plan.layout, "profil": {"name": prof.name, "hash": prof.hash} if prof else None,
        "warnungen": plan.warnungen, "ignoriert": plan.ignoriert,
        "luecken": plan.luecken,
        "luecken_ausserhalb": plan.luecken_ausserhalb,   # Welle 6: nicht gelesen, aber auf Seiten ohne CDE-Inhalt (nur Info)
        "farben": pa.farben, "aenderungen": [a.zu_dict() for a in plan.aenderungen],
        "absaetze": [ab.zu_dict() for ab in plan.absaetze],   # Schema 3: fuer die Gegenprobe nach dem Speichern
        # Schema 4, nur Info fuer "Details fuer Fachleute": erkannte Legende und Zahl der Sprechblasen (kein Text daraus)
        "legende_modus": str(getattr(pa, "legende_modus", "") or ""),
        "sprechblasen": len(getattr(pa, "randnotizen", None) or []),
        # Spec-Filter (Welle 3): Spec des PDFs in der CDE ({erkannt, quelle, zeilen, kandidaten, pdf}; None = keine Spec-Spalte)
        # und alle Specs der CDE (Name, Zeilenzahl, von, bis). Nur Nummern und Zeilen, kein Dokumenttext.
        "spec": plan.spec, "specs": specs_der_cde(xl),
        # Vertrauensmodus: Word oder PDF (pdf_name bleibt der Feldname fuer beide) und die Stichprobe (nur IDs, fest aus der Lauf-ID)
        "quelle_art": info["quelle_art"],
    }
    daten["stichprobe"] = vertrauen.ziehe_stichprobe(run_id, daten["aenderungen"])
    _schreibe_json_atomar(d / "plan.json", daten)
    return daten


def neu_analysieren(run_id: str, spec: str | None) -> tuple[str, dict[str, Any]]:
    """Neuer Lauf mit denselben Eingabedateien, aber einer anderen Spec (Regel 29: der bisherige Lauf bleibt unberuehrt,
    nichts wird ueberschrieben). Gibt (neue Lauf-ID, Plan) zurueck. Scheitert die Analyse, bleibt der neue Lauf mit seinen
    Eingaben stehen; die ID steht dann in `LaufFehler.run_id`/`NeuAnalyseFehler`."""
    spec = pruefe_spec(spec)
    d = lauf_dir(run_id)
    info = _info(d)
    quelle_pdf, quelle_excel = d / "eingabe" / info["pdf"], d / "eingabe" / info["excel"]
    if not (quelle_pdf.is_file() and quelle_excel.is_file()):
        raise LaufFehler("Die Dateien dieses Laufs sind nicht mehr vorhanden. Bitte PDF und Excel neu ablegen.", 409)
    profil = None
    try:
        p = _lese_json(d / "plan.json").get("profil")
        profil = str(p["name"]) if isinstance(p, dict) and p.get("name") else None
    except (LaufFehler, OSError, KeyError, TypeError):
        profil = None
    neu = neuer_lauf(quelle_pdf, quelle_excel, info["pdf"], info["excel"])
    try:
        return neu, analysiere(neu, profil, spec)
    except Exception as e:
        e.run_id = neu                                # type: ignore[attr-defined]   # der Aufrufer markiert den Fehler im neuen Lauf
        raise


def markiere_analyse_fehler(run_id: str, meldung: str) -> None:
    """Der Lauf bleibt mit seinen Eingaben liegen; hier steht nur die einfache Meldung, kein Rohtext."""
    d = lauf_dir(run_id)
    if (d / "analyse_fehler.json").exists():
        return
    _schreibe_text_atomar(d / "analyse_fehler.json", json.dumps(
        {"meldung": meldung, "zeit": datetime.now().isoformat(timespec="seconds")}, ensure_ascii=False))


def _analyse_meldung(d: Path) -> str | None:
    p = d / "analyse_fehler.json"
    if not p.exists():
        return None
    try:
        m = json.loads(p.read_text(encoding="utf-8")).get("meldung")
    except (OSError, ValueError, AttributeError):
        m = None
    return m if isinstance(m, str) and m else "Die Analyse ist fehlgeschlagen."


def _lese_plan(d: Path) -> dict[str, Any]:
    """plan.json lesen und pruefen (Schema, Pflichtfelder). Keine Neuanalyse."""
    p = d / "plan.json"
    if not p.exists():
        meldung = _analyse_meldung(d)
        if meldung:
            if not meldung.startswith("Die Analyse"):
                meldung = f"Die Analyse dieses Laufs ist fehlgeschlagen. {meldung}"
            raise LaufFehler(meldung, 409, "fehlgeschlagen")
        raise LaufFehler("Dieser Lauf hat noch keine Analyse.", 409, "angelegt")
    daten = _lese_json(p)
    pruefe_schema(daten)
    if not all(k in daten for k in _PLAN_PFLICHT) or not isinstance(daten["aenderungen"], list) \
            or not isinstance(daten["bilanz"], dict):
        raise _beschaedigt()
    return daten


def lade_plan(run_id: str) -> dict[str, Any]:
    d = lauf_dir(run_id)
    _info(d)                                         # kaputte lauf.json -> "beschaedigt", wie in der Liste
    daten = _lese_plan(d)                            # fehlt plan.json, gibt es KEINE Neuanalyse mehr
    daten.setdefault("spec", None)                   # Plaene vor dem Spec-Filter kennen weder Spec noch Spec-Liste
    daten.setdefault("specs", [])
    daten.setdefault("quelle_art", quelle_art_von(daten.get("pdf_name") if isinstance(daten.get("pdf_name"), str) else ""))
    st = daten.get("stichprobe")                     # Plaene vor dem Vertrauensmodus haben keine Stichprobe
    daten["stichprobe"] = [x for x in st if isinstance(x, str)] if isinstance(st, list) else []
    daten["ergebnisse"] = _ergebnisse_in(d)
    daten["archiviert"] = (d / "archiviert.json").exists()
    return daten


# ---------------------------------------------------------------------------------------------
# Ergebnisse
# ---------------------------------------------------------------------------------------------
def _ergebnis_lesen(e: Path) -> dict[str, Any] | None:
    """Gueltiges Ergebnis oder None (kein/kaputtes ergebnis.json, unbekanntes Schema)."""
    j = e / "ergebnis.json"
    if not j.is_file():
        return None
    try:
        daten = _lese_json(j)
        pruefe_schema(daten)
        sp = daten.get("stichprobe") if isinstance(daten.get("stichprobe"), dict) else {}
        return {"ordner": e.name, "bilanz": daten["bilanz"], "excel": daten["ergebnis_name"],
                "erstellt": daten["erstellt"], "freigaben": daten.get("freigaben", []),
                "abwahl": daten.get("abwahl", []),         # aeltere Ergebnisse kennen keine Abwahl
                "abgelehnt": daten.get("abgelehnt", []),   # ... und keine Ablehnung, keine Stichprobe (Vertrauensmodus)
                "stichprobe": {"passt": _ids_aus(sp.get("passt")), "falsch": _ids_aus(sp.get("falsch"))}}
    except (LaufFehler, KeyError):
        return None


def _ergebnisse_in(d: Path) -> list[dict[str, Any]]:
    out = []
    for e in sorted(d.glob("ergebnis_*")):
        if e.is_dir() and _ERGEBNIS.fullmatch(e.name) and (eintrag := _ergebnis_lesen(e)):
            out.append(eintrag)
    return out


def ergebnisse(run_id: str) -> list[dict[str, Any]]:
    return _ergebnisse_in(lauf_dir(run_id))


def _neuer_ergebnis_ordner(d: Path) -> Path:
    """Naechste freie Nummer (hoechste vorhandene + 1, nie 'Anzahl + 1'); mkdir ohne exist_ok."""
    def name() -> str:
        nummern = [int(m.group(1)) for p in d.glob("ergebnis_*") if (m := _ERGEBNIS.fullmatch(p.name))]
        nr = max(nummern, default=0) + 1
        if nr > 999:
            raise LaufFehler("Dieser Lauf hat schon 999 Übernahmen. Bitte einen neuen Lauf starten.", 409)
        return f"ergebnis_{nr:03d}"
    return _reserviere(d, name)


def _profil_des_plans(plan: dict[str, Any]) -> "layouts.Profil | None":
    """Das Profil, mit dem der Lauf analysiert wurde (plan.json 'profil'), aus seiner Bestaetigung; ohne Profil None."""
    p = plan.get("profil")
    if not p:
        return None
    try:
        return layouts.profil_aus_bestaetigung(str(p["name"]), str(p["hash"]))
    except (layouts.ProfilFehler, KeyError, TypeError) as e:
        raise LaufFehler(getattr(e, "meldung", "Das Profil dieses Laufs ist nicht mehr bestätigt. Bitte neu analysieren."), 409, "profil") from None


def kategorie(a: dict[str, Any]) -> str:
    """Wohin eine Aenderung im Dashboard gehoert (reine Einordnung, nichts wird veraendert):
    automatisch (wird eingetragen) | pruefen (Vorschlag, ein Mensch entscheidet) | ausserhalb (gehoert nicht zur CDE,
    z. B. Deckblatt oder Revisionstabelle; kein offener Punkt, ADR 0006) | nicht (mit Grund nicht eingetragen).
    Was eingetragen wird, steht immer unter "automatisch" (auch wenn es ausserhalb laege), damit es abwaehlbar bleibt."""
    status = a.get("status")
    if status == UEBERNOMMEN:
        return "automatisch"
    if a.get("bereich") == "ausserhalb":
        return "ausserhalb"
    if status == PRUEFEN:
        return "pruefen"
    return "nicht"


def zusammenfassung(aenderungen: list[dict[str, Any]]) -> dict[str, int]:
    """Zahlen fuer den einen Satz oben im Dashboard (gesamt = automatisch + pruefen + ausserhalb + nicht)."""
    zahl = {"gesamt": len(aenderungen), "automatisch": 0, "pruefen": 0, "ausserhalb": 0, "nicht": 0}
    for a in aenderungen:
        zahl[kategorie(a)] += 1
    return zahl


def _id_liste(werte: list[str] | None) -> set[str]:
    return {str(x) for x in (werte or [])}


def _ids_aus(werte: Any) -> list[str]:
    return [x for x in werte if isinstance(x, str)] if isinstance(werte, list) else []


def _pruefe_entscheidungen(plan: dict[str, Any], bekannt: dict[str, Any], freig: set[str], abw: set[str], abl: set[str],
                           pausch: set[str], passt: set[str], falsch: set[str]) -> None:
    """Prueft die Entscheidungen des Menschen gegeneinander und gegen den Plan (NutzerFehler 400 mit einfachem Satz)."""
    unbekannt = (freig | abw | abl | pausch | passt | falsch) - set(bekannt)
    if unbekannt:
        raise NutzerFehler(f"Unbekannte Änderungs-IDs: {sorted(unbekannt)}", 400)
    if freig & abw:
        raise NutzerFehler(f"Eine Änderung kann nicht zugleich freigegeben und abgewählt sein: {sorted(freig & abw)}", 400)
    nicht_auto = sorted(x for x in abw if bekannt[x].status != UEBERNOMMEN)
    if nicht_auto:
        raise NutzerFehler(f"Abwählen geht nur bei automatisch eingetragenen Änderungen: {nicht_auto}", 400)
    kein_vorschlag = sorted(x for x in abl if bekannt[x].status != PRUEFEN)
    if kein_vorschlag:
        raise NutzerFehler(f"Ablehnen geht nur bei Vorschlägen: {kein_vorschlag}", 400)
    if abl & freig:
        raise NutzerFehler(f"Eine Änderung kann nicht zugleich freigegeben und abgelehnt sein: {sorted(abl & freig)}", 400)
    if pausch - freig:
        raise NutzerFehler(f"Pauschal übernehmen geht nur bei freigegebenen Vorschlägen: {sorted(pausch - freig)}", 400)
    stichprobe = set(_ids_aus(plan.get("stichprobe")))
    fremd = sorted((passt | falsch) - stichprobe)
    if fremd:
        raise NutzerFehler(f"Das gehört nicht zur Stichprobe dieses Laufs: {fremd}", 400)
    if passt & falsch:
        raise NutzerFehler(f"Eine Änderung der Stichprobe kann nicht zugleich passen und falsch sein: {sorted(passt & falsch)}", 400)
    if falsch - abw:
        raise NutzerFehler(f"Was in der Stichprobe als falsch gilt, muss auch abgewählt sein: {sorted(falsch - abw)}", 400)
    if passt & abw:
        raise NutzerFehler(f"Eine abgewählte Änderung kann in der Stichprobe nicht „passt“ sein: {sorted(passt & abw)}", 400)


def _entscheidungen_des_menschen(run_id: str, ergebnis: str, plan: dict[str, Any], freig: set[str], abw: set[str], abl: set[str],
                                 pausch: set[str], passt: set[str], falsch: set[str]) -> list[dict[str, Any]]:
    """Eine Zeile je Entscheidung fuers Freigabe-Protokoll (nur Kennungen und Pruefcodes, nie Text aus Dokumenten, Regel 16).
    Offene Vorschlaege (weder freigegeben noch abgelehnt) und nicht eingetragene Aenderungen sind keine Entscheidung."""
    stichprobe = set(_ids_aus(plan.get("stichprobe")))
    zeilen = []
    for a in plan["aenderungen"]:
        i, codes = a.get("id"), [c for c in a.get("pruefcodes") or [] if isinstance(c, str)]
        if a.get("status") == UEBERNOMMEN:
            art = "stichprobe" if i in stichprobe else "automatisch"
            if i in abw:
                entscheidung = "abgelehnt" if i in falsch else "abgewaehlt"
            else:
                entscheidung = "bestaetigt" if i in passt else "uebernommen"
        elif kategorie(a) == "pruefen" and i in freig:
            art, entscheidung = ("vorschlag" if i in pausch else "pruefen"), "uebernommen"
        elif kategorie(a) == "pruefen" and i in abl:
            art, entscheidung = "pruefen", "abgelehnt"
        else:
            continue
        zeilen.append(entscheidung_zeile(run_id, ergebnis, i, art, entscheidung, codes))
    return zeilen


def uebernehmen_lauf(run_id: str, freigaben: list[str] | None = None, abwahl: list[str] | None = None, *,
                     abgelehnt: list[str] | None = None, pauschal: list[str] | None = None,
                     stichprobe_passt: list[str] | None = None, stichprobe_falsch: list[str] | None = None) -> dict[str, Any]:
    """Traegt die sicheren Aenderungen und die freigegebenen Vorschlaege in eine NEUE Excel ein (eigener Ergebnis-Ordner).
    `freigaben`: IDs von Vorschlaegen ("pruefen"), die ein Mensch freigibt. `abwahl`: IDs automatischer Aenderungen
    ("uebernommen"), die ein Mensch NICHT eingetragen haben will (gekoppelte Aenderungen fallen mit weg, siehe Applier).
    Unbekannte IDs, eine Abwahl von etwas nicht Automatischem und dieselbe ID in beiden Listen sind Fehler (400).
    Nur fuers Freigabe-Protokoll (Vertrauensmodus, aendern nichts am Eintragen): `abgelehnt` = Vorschlaege, die der Mensch
    ausdruecklich nicht uebernimmt; `pauschal` = freigegebene Vorschlaege, die er nicht einzeln angesehen hat; `stichprobe_passt` /
    `stichprobe_falsch` = Stichproben-Karten mit "Passt" bzw. "Falsch" (Falsch muss auch in `abwahl` stehen).
    Die Entscheidungen werden an entscheidungen.jsonl des Laufs ANGEHAENGT, sobald die Excel erstellt ist."""
    d = lauf_dir(run_id)
    plan = _lese_plan(d)
    info = _info(d)
    quelle = d / "eingabe" / info["excel"]
    if not quelle.is_file() or _sha(quelle) != plan["excel_sha256"]:
        raise LaufFehler("Die gespeicherte Excel passt nicht mehr zur Analyse. Bitte neuen Lauf starten.", 409)
    try:
        aenderungen = [aenderung_aus_dict(a) for a in plan["aenderungen"]]
        absaetze = absaetze_aus_liste(plan.get("absaetze"))       # Schema <= 2: None (alte Nachkontrolle)
    except (KeyError, TypeError, ValueError, AttributeError):
        raise _beschaedigt() from None
    freig, abw = _id_liste(freigaben), _id_liste(abwahl)
    abl, pausch, passt, falsch = _id_liste(abgelehnt), _id_liste(pauschal), _id_liste(stichprobe_passt), _id_liste(stichprobe_falsch)
    _pruefe_entscheidungen(plan, {a.id: a for a in aenderungen}, freig, abw, abl, pausch, passt, falsch)
    ordner = _neuer_ergebnis_ordner(d)
    try:
        stamm = Path(info["excel"]).stem
        ziel = ordner / f"{stamm}_aktualisiert{Path(info['excel']).suffix}"
        with profil_kontext(_profil_des_plans(plan)):          # der Applier liest die Excel selbst: gleicher Aufbau wie bei der Analyse
            res = uebernehmen(aenderungen, quelle, ziel, freig, absaetze=absaetze, abwahl=abw)
        markierte = plan["bilanz"]["woerter_im_pdf"]
        luecken = plan.get("luecken")
        aussen = plan.get("luecken_ausserhalb")
        aussen = aussen if isinstance(aussen, list) else []
        bil = bilanz(aenderungen, markierte, plan.get("ignoriert", []), luecken if isinstance(luecken, list) else [], aussen)
        if not isinstance(luecken, list):
            # Schema-1-Plan: damals wurden Luecken nicht erfasst. "Vollstaendig" ist UNBEKANNT und darf nie behauptet
            # werden (nie gruen); luecken_geprueft=False sagt das ausdruecklich.
            bil.update(vollstaendig=None, luecken=None, luecken_geprueft=False)
        warn = list(plan.get("warnungen", [])) + res.warnungen
        xinfo = {"pdf_name": info["pdf"], "excel_name": info["excel"], "ergebnis_name": ziel.name, "quelle_art": info["quelle_art"]}
        schreibe_xlsx(ordner / "aenderungsprotokoll.xlsx", xinfo, aenderungen, bil, plan.get("layout", []), warn,
                      plan.get("farben", []), plan.get("luecken", []), plan.get("spec"), aussen)
        _schreibe_json_atomar(ordner / "ergebnis.json", {   # zuletzt: erst damit zaehlt das Ergebnis
            "run_id": run_id, "erstellt": datetime.now().isoformat(timespec="seconds"), "ergebnis_name": ziel.name,
            "freigaben": sorted(freig), "abwahl": sorted(abw), "bilanz": bil, "warnungen": warn,
            "abgelehnt": sorted(abl), "pauschal": sorted(pausch),
            "stichprobe": {"passt": sorted(passt), "falsch": sorted(falsch)},
            "eingabe": {"pdf": info["pdf"], "excel": info["excel"], "pdf_sha256": plan.get("pdf_sha256", ""),
                        "excel_sha256": plan["excel_sha256"]},
            "aenderungen": [a.zu_dict() for a in aenderungen]})
    except BaseException:
        shutil.rmtree(ordner, ignore_errors=True)    # eigener, noch unfertiger Ordner (ohne ergebnis.json)
        raise
    try:                                             # Freigabe-Protokoll: erst jetzt, wo das Ergebnis wirklich zaehlt; nur anhaengen
        haenge_entscheidungen_an(d / ENTSCHEIDUNGEN, _entscheidungen_des_menschen(
            run_id, ordner.name, plan, freig, abw, abl, pausch, passt, falsch))
    except OSError:
        log.exception("Freigabe-Protokoll von Lauf %s konnte nicht ergaenzt werden", run_id)
        warn.append("Das Freigabe-Protokoll (entscheidungen.jsonl) konnte nicht ergänzt werden. Die Excel ist trotzdem fertig.")
    return {"ordner": ordner.name, "excel": ziel.name, "bilanz": bil, "warnungen": warn,
            "freigaben": sorted(freig), "abwahl": sorted(abw), "abgelehnt": sorted(abl),
            "stichprobe": {"passt": sorted(passt), "falsch": sorted(falsch)},
            "aenderungen": [a.zu_dict() for a in aenderungen]}


# ---------------------------------------------------------------------------------------------
# Liste, Archiv, Dateien
# ---------------------------------------------------------------------------------------------
def _eintrag(d: Path) -> dict[str, Any]:
    """Ein Listeneintrag; ein kaputter Lauf wird als 'beschaedigt' angezeigt und stoert die Liste nicht."""
    eintrag: dict[str, Any] = {"run_id": d.name, "pdf": "", "excel": "", "angelegt": "", "ergebnisse": 0,
                               "archiviert": (d / "archiviert.json").exists()}
    try:
        info = _info(d)
    except LaufFehler as e:
        eintrag.update(status="beschaedigt", meldung=e.meldung)
        return eintrag
    eintrag.update(pdf=info["pdf"], excel=info["excel"], angelegt=info["angelegt"],
                   ergebnisse=len(_ergebnisse_in(d)))
    try:
        plan = _lese_plan(d)
    except LaufFehler as e:
        eintrag["status"] = e.art if e.art in ("fehlgeschlagen", "angelegt", "neuere_version") else "beschaedigt"
        if e.art != "angelegt":
            eintrag["meldung"] = e.meldung
        return eintrag
    eintrag["bilanz"] = plan["bilanz"]
    sp = plan.get("spec")
    eintrag["spec"] = sp.get("erkannt") if isinstance(sp, dict) and len(sp.get("kandidaten") or []) > 1 else None
    eintrag["status"] = "uebernommen" if eintrag["ergebnisse"] else "analysiert"
    return eintrag


def liste(mit_archiv: bool = False) -> list[dict[str, Any]]:
    out = []
    for d in sorted(runs_dir().iterdir(), reverse=True):
        if not (d.is_dir() and _ID.fullmatch(d.name) and (d / "lauf.json").exists()):
            continue
        if not mit_archiv and (d / "archiviert.json").exists():
            continue
        out.append(_eintrag(d))
    return out


def archiviere(run_id: str) -> None:
    """Regel 29: nichts wird geloescht. Der Lauf bekommt nur einen Marker und fehlt in der Standardliste."""
    d = lauf_dir(run_id)
    if (d / "archiviert.json").exists():
        return
    _schreibe_text_atomar(d / "archiviert.json", json.dumps(
        {"archiviert": datetime.now().isoformat(timespec="seconds")}, ensure_ascii=False))


def stelle_wieder_her(run_id: str) -> None:
    """Der Marker wird nicht geloescht, sondern mit Zeitstempel umbenannt (der Verlauf bleibt erhalten)."""
    d = lauf_dir(run_id)
    marker = d / "archiviert.json"
    if not marker.exists():
        return
    stempel = datetime.now().strftime("%Y%m%d-%H%M%S")
    for i in range(_VERSUCHE):
        ziel = d / f"archiviert_{stempel}{'' if i == 0 else f'-{i}'}.json.alt"
        if not ziel.exists():
            try:
                os.rename(marker, ziel)
            except FileNotFoundError:
                pass                                  # jemand war schneller: schon wiederhergestellt
            return


def datei(run_id: str, ordner: str, name: str) -> Path:
    d = lauf_dir(run_id)
    if not re.fullmatch(r"ergebnis_\d{3}|eingabe", ordner):
        raise NutzerFehler("Ungültiger Ordner.", 404)
    p = (d / ordner / Path(name).name).resolve()
    if d.resolve() not in p.parents or not p.is_file():
        raise NichtGefunden("Datei nicht gefunden.")
    return p
