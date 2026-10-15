"""Laeufe: ein Lauf = ein PDF + eine Excel, analysiert und (nach Freigabe) uebernommen.

Ordner je Lauf (data/runs/<id>/), append-only (Regel 29):
  eingabe/            Kopien von PDF und Excel (Originale bleiben unberuehrt)
  plan.json           Analyse (Aenderungen, Bilanz, Aufbau, Farben)
  ergebnis_001/       je Uebernahme ein NEUER Ordner: aktualisierte Excel + Protokoll
Nichts wird ueberschrieben.
"""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from . import config
from .applier import uebernehmen
from .excel_model import ExcelModell
from .models import PRUEFEN, UEBERNOMMEN
from .pdf_reader import analysiere_pdf
from .planner import plane
from .protokoll import aenderung_aus_dict, bilanz, schreibe_json, schreibe_xlsx

_ID = re.compile(r"^\d{8}-\d{6}-[0-9a-f]{6}$")


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


def _sicherer_name(name: str, fallback: str) -> str:
    name = Path(name or fallback).name
    name = re.sub(r"[^\w.\- äöüÄÖÜß]", "_", name).strip() or fallback
    return name


def neuer_lauf(pdf_pfad: str | Path, excel_pfad: str | Path, pdf_name: str = "", excel_name: str = "") -> str:
    run_id = f"{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}"
    d = runs_dir() / run_id
    (d / "eingabe").mkdir(parents=True)
    pdf_name = _sicherer_name(pdf_name or Path(pdf_pfad).name, "aenderungen.pdf")
    excel_name = _sicherer_name(excel_name or Path(excel_pfad).name, "cde.xlsx")
    shutil.copy2(pdf_pfad, d / "eingabe" / pdf_name)
    shutil.copy2(excel_pfad, d / "eingabe" / excel_name)
    (d / "lauf.json").write_text(json.dumps({"pdf": pdf_name, "excel": excel_name,
                                              "angelegt": datetime.now().isoformat(timespec="seconds")},
                                             ensure_ascii=False), encoding="utf-8")
    return run_id


def _info(d: Path) -> dict[str, Any]:
    return json.loads((d / "lauf.json").read_text(encoding="utf-8"))


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
    schreibe_json(d / "plan.json", daten)
    return daten


def lade_plan(run_id: str) -> dict[str, Any]:
    d = lauf_dir(run_id)
    p = d / "plan.json"
    if not p.exists():
        return analysiere(run_id)
    daten = json.loads(p.read_text(encoding="utf-8"))
    daten["ergebnisse"] = ergebnisse(run_id)
    return daten


def ergebnisse(run_id: str) -> list[dict[str, Any]]:
    d = lauf_dir(run_id)
    out = []
    for e in sorted(d.glob("ergebnis_*")):
        j = e / "ergebnis.json"
        if j.exists():
            daten = json.loads(j.read_text(encoding="utf-8"))
            out.append({"ordner": e.name, "bilanz": daten["bilanz"], "excel": daten["ergebnis_name"],
                        "erstellt": daten["erstellt"], "freigaben": daten.get("freigaben", [])})
    return out


def uebernehmen_lauf(run_id: str, freigaben: list[str] | None = None) -> dict[str, Any]:
    d = lauf_dir(run_id)
    plan = lade_plan(run_id)
    info = _info(d)
    quelle = d / "eingabe" / info["excel"]
    if _sha(quelle) != plan["excel_sha256"]:
        raise RuntimeError("Die gespeicherte Excel passt nicht mehr zur Analyse. Bitte neuen Lauf starten.")
    aenderungen = [aenderung_aus_dict(a) for a in plan["aenderungen"]]
    freig = set(freigaben or [])
    unbekannt = freig - {a.id for a in aenderungen}
    if unbekannt:
        raise ValueError(f"Unbekannte Änderungs-IDs: {sorted(unbekannt)}")
    nr = len(list(d.glob("ergebnis_*"))) + 1
    ordner = d / f"ergebnis_{nr:03d}"
    ordner.mkdir()
    stamm = Path(info["excel"]).stem
    ziel = ordner / f"{stamm}_aktualisiert{Path(info['excel']).suffix}"
    res = uebernehmen(aenderungen, quelle, ziel, freig)
    markierte = plan["bilanz"]["woerter_im_pdf"]
    bil = bilanz(aenderungen, markierte, plan["ignoriert"], plan.get("luecken", []))
    warn = list(plan["warnungen"]) + res.warnungen
    xinfo = {"pdf_name": info["pdf"], "excel_name": info["excel"], "ergebnis_name": ziel.name}
    schreibe_xlsx(ordner / "aenderungsprotokoll.xlsx", xinfo, aenderungen, bil, plan["layout"], warn, plan["farben"])
    schreibe_json(ordner / "ergebnis.json", {
        "run_id": run_id, "erstellt": datetime.now().isoformat(timespec="seconds"), "ergebnis_name": ziel.name,
        "freigaben": sorted(freig), "bilanz": bil, "warnungen": warn,
        "eingabe": {"pdf": info["pdf"], "excel": info["excel"], "pdf_sha256": plan["pdf_sha256"],
                    "excel_sha256": plan["excel_sha256"]},
        "aenderungen": [a.zu_dict() for a in aenderungen]})
    return {"ordner": ordner.name, "excel": ziel.name, "bilanz": bil, "warnungen": warn,
            "aenderungen": [a.zu_dict() for a in aenderungen]}


def liste() -> list[dict[str, Any]]:
    out = []
    for d in sorted(runs_dir().iterdir(), reverse=True):
        if not (d.is_dir() and _ID.match(d.name) and (d / "lauf.json").exists()):
            continue
        info = _info(d)
        eintrag = {"run_id": d.name, "pdf": info["pdf"], "excel": info["excel"], "angelegt": info["angelegt"],
                   "ergebnisse": len(list(d.glob("ergebnis_*")))}
        p = d / "plan.json"
        if p.exists():
            eintrag["bilanz"] = json.loads(p.read_text(encoding="utf-8"))["bilanz"]
        out.append(eintrag)
    return out


def loeschen(run_id: str) -> None:
    shutil.rmtree(lauf_dir(run_id))


def datei(run_id: str, ordner: str, name: str) -> Path:
    d = lauf_dir(run_id)
    if not re.fullmatch(r"ergebnis_\d{3}|eingabe", ordner):
        raise ValueError("Ungültiger Ordner.")
    p = (d / ordner / Path(name).name).resolve()
    if d.resolve() not in p.parents or not p.is_file():
        raise FileNotFoundError("Datei nicht gefunden.")
    return p
