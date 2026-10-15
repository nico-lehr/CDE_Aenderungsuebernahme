"""Laeufe: ein Lauf = ein PDF + eine Excel, analysiert und (nach Freigabe) uebernommen.

Ordner je Lauf (data/runs/<id>/), append-only (Regel 29):
  eingabe/            Kopien von PDF und Excel (Originale bleiben unberuehrt)
  lauf.json           Namen der Eingaben (wird zuletzt, atomar geschrieben)
  plan.json           Analyse (Aenderungen, Bilanz, Aufbau, Farben), atomar geschrieben
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
import os
import re
import shutil
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from . import config
from .applier import uebernehmen
from .excel_model import ExcelModell
from .models import GELESENE_SCHEMAS
from .pdf_reader import analysiere_pdf
from .planner import plane
from .protokoll import aenderung_aus_dict, bilanz, schreibe_json, schreibe_xlsx

_ID = re.compile(r"^\d{8}-\d{6}-[0-9a-f]{6}$")
_ERGEBNIS = re.compile(r"ergebnis_(\d{3})")
# Windows kann diese Namen (auch mit Endung) nicht als Datei anlegen
_RESERVIERT = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
_PLAN_PFLICHT = ("aenderungen", "bilanz", "excel_sha256", "pdf_name", "excel_name")
_VERSUCHE = 50                                   # Wiederholungen bei Namenskollisionen

MELDUNG_BESCHAEDIGT = "Der Lauf ist beschädigt."
MELDUNG_NEUERE_VERSION = ("Dieser Lauf stammt aus einer neueren Programmversion (Schema-Version {v}). "
                          "Bitte das Programm aktualisieren.")


class LaufFehler(Exception):
    """Ein Lauf kann nicht verwendet werden. `meldung` ist ein einfacher deutscher Satz fuer den Nutzer."""

    def __init__(self, meldung: str, status: int = 409, art: str = "fehler"):
        super().__init__(meldung)
        self.meldung = meldung
        self.status = status
        self.art = art


def _beschaedigt() -> LaufFehler:
    return LaufFehler(MELDUNG_BESCHAEDIGT, 422, "beschaedigt")


def runs_dir() -> Path:
    d = config.pfad("runs_dir")
    d.mkdir(parents=True, exist_ok=True)
    return d


def lauf_dir(run_id: str) -> Path:
    if not _ID.match(run_id):
        raise ValueError("Ungültige Lauf-ID.")
    d = runs_dir() / run_id
    if not d.is_dir():
        raise FileNotFoundError("Lauf nicht gefunden.")
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
        raise ValueError(f"Der Dateiname ist zu lang (höchstens {grenze} Zeichen). Bitte die Datei kürzer benennen.")


def _sicherer_name(name: str, fallback: str) -> str:
    name = Path(name or fallback).name
    name = re.sub(r"[^\w.\- äöüÄÖÜß]", "_", name).strip(" .") or fallback
    if name.split(".")[0].rstrip().upper() in _RESERVIERT:
        name = "_" + name
    return name


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
    pdf_name = _sicherer_name(pdf_roh, "aenderungen.pdf")
    excel_name = _sicherer_name(excel_roh, "cde.xlsx")
    d = _reserviere(runs_dir(), lambda: f"{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}")
    try:
        (d / "eingabe").mkdir()
        shutil.copy2(pdf_pfad, d / "eingabe" / pdf_name)
        shutil.copy2(excel_pfad, d / "eingabe" / excel_name)
        _schreibe_text_atomar(d / "lauf.json", json.dumps(
            {"pdf": pdf_name, "excel": excel_name, "angelegt": datetime.now().isoformat(timespec="seconds")},
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
    return info


def analysiere(run_id: str) -> dict[str, Any]:
    d = lauf_dir(run_id)
    info = _info(d)
    pdf, excel = d / "eingabe" / info["pdf"], d / "eingabe" / info["excel"]
    pa = analysiere_pdf(pdf)
    xl = ExcelModell(excel)
    plan = plane(pa, xl)
    bil = bilanz(plan.aenderungen, plan.markierte_woerter, plan.ignoriert, plan.luecken)
    daten = {
        "run_id": run_id, "pdf_name": info["pdf"], "excel_name": info["excel"],
        "pdf_sha256": _sha(pdf), "excel_sha256": _sha(excel), "seiten": pa.seiten,
        "bilanz": bil, "layout": plan.layout, "warnungen": plan.warnungen, "ignoriert": plan.ignoriert,
        "luecken": plan.luecken,
        "farben": pa.farben, "aenderungen": [a.zu_dict() for a in plan.aenderungen],
    }
    _schreibe_json_atomar(d / "plan.json", daten)
    return daten


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
            raise LaufFehler(f"Die Analyse dieses Laufs ist fehlgeschlagen. {meldung}", 409, "fehlgeschlagen")
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
        return {"ordner": e.name, "bilanz": daten["bilanz"], "excel": daten["ergebnis_name"],
                "erstellt": daten["erstellt"], "freigaben": daten.get("freigaben", [])}
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


def uebernehmen_lauf(run_id: str, freigaben: list[str] | None = None) -> dict[str, Any]:
    d = lauf_dir(run_id)
    plan = _lese_plan(d)
    info = _info(d)
    quelle = d / "eingabe" / info["excel"]
    if not quelle.is_file() or _sha(quelle) != plan["excel_sha256"]:
        raise LaufFehler("Die gespeicherte Excel passt nicht mehr zur Analyse. Bitte neuen Lauf starten.", 409)
    try:
        aenderungen = [aenderung_aus_dict(a) for a in plan["aenderungen"]]
    except (KeyError, TypeError, ValueError, AttributeError):
        raise _beschaedigt() from None
    freig = set(freigaben or [])
    unbekannt = freig - {a.id for a in aenderungen}
    if unbekannt:
        raise ValueError(f"Unbekannte Änderungs-IDs: {sorted(unbekannt)}")
    ordner = _neuer_ergebnis_ordner(d)
    try:
        stamm = Path(info["excel"]).stem
        ziel = ordner / f"{stamm}_aktualisiert{Path(info['excel']).suffix}"
        res = uebernehmen(aenderungen, quelle, ziel, freig)
        markierte = plan["bilanz"]["woerter_im_pdf"]
        bil = bilanz(aenderungen, markierte, plan.get("ignoriert", []), plan.get("luecken", []))
        warn = list(plan.get("warnungen", [])) + res.warnungen
        xinfo = {"pdf_name": info["pdf"], "excel_name": info["excel"], "ergebnis_name": ziel.name}
        schreibe_xlsx(ordner / "aenderungsprotokoll.xlsx", xinfo, aenderungen, bil, plan.get("layout", []), warn,
                      plan.get("farben", []))
        _schreibe_json_atomar(ordner / "ergebnis.json", {   # zuletzt: erst damit zaehlt das Ergebnis
            "run_id": run_id, "erstellt": datetime.now().isoformat(timespec="seconds"), "ergebnis_name": ziel.name,
            "freigaben": sorted(freig), "bilanz": bil, "warnungen": warn,
            "eingabe": {"pdf": info["pdf"], "excel": info["excel"], "pdf_sha256": plan.get("pdf_sha256", ""),
                        "excel_sha256": plan["excel_sha256"]},
            "aenderungen": [a.zu_dict() for a in aenderungen]})
    except BaseException:
        shutil.rmtree(ordner, ignore_errors=True)    # eigener, noch unfertiger Ordner (ohne ergebnis.json)
        raise
    return {"ordner": ordner.name, "excel": ziel.name, "bilanz": bil, "warnungen": warn,
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
    eintrag["status"] = "uebernommen" if eintrag["ergebnisse"] else "analysiert"
    return eintrag


def liste(mit_archiv: bool = False) -> list[dict[str, Any]]:
    out = []
    for d in sorted(runs_dir().iterdir(), reverse=True):
        if not (d.is_dir() and _ID.match(d.name) and (d / "lauf.json").exists()):
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
        raise ValueError("Ungültiger Ordner.")
    p = (d / ordner / Path(name).name).resolve()
    if d.resolve() not in p.parents or not p.is_file():
        raise FileNotFoundError("Datei nicht gefunden.")
    return p
