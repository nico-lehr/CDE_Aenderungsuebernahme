"""Vertrauensmodus: Stichprobe, Dokument-Ampel und Auswertung der Freigabe-Protokolle.

Nico lernt dem Tool schrittweise zu vertrauen. Dieses Modul entscheidet NICHTS ueber das Eintragen (Regel 27): Es zieht eine
feste Stichprobe aus den automatisch eingetragenen Aenderungen, fasst den Stand eines Laufs als Ampel zusammen und wertet
nach einiger Zeit aus, bei welchen Pruefcodes Nico fast immer zustimmt. Schwellen stehen in config/vertrauen.yaml (Regel 28).
Es kommen nie Texte aus Dokumenten oder Zellen vor, nur Kennungen und Zahlen (Regel 16).
"""
from __future__ import annotations

import hashlib
import math
import random
from pathlib import Path
from typing import Any, Iterable

from . import config, protokoll
from .models import UEBERNOMMEN

_STANDARD: dict[str, dict[str, Any]] = {
    "vertrauen": {"aktiv": True, "stichprobe_anteil": 0.10, "stichprobe_min": 3},
    "ampel": {"rot_offene_je_100": 40, "offen_code_praefixe": ["konsens_"]},
    "auswertung": {"min_angesehen": 20, "freigabe_ab": 0.98, "ablehnung_max": 0.02},
}


def einstellungen() -> dict[str, dict[str, Any]]:
    """Einstellungen aus config/vertrauen.yaml; fehlende Abschnitte oder Werte bekommen den Standard (das Tool scheitert nicht an einer Kleinigkeit)."""
    try:
        roh = config.lade("vertrauen")
    except FileNotFoundError:
        roh = {}
    roh = roh if isinstance(roh, dict) else {}
    return {k: {**v, **(roh[k] if isinstance(roh.get(k), dict) else {})} for k, v in _STANDARD.items()}


def _zahl(wert: Any, standard: float, tief: float = 0.0, hoch: float = float("inf")) -> float:
    if isinstance(wert, bool) or not isinstance(wert, (int, float)) or wert != wert:
        return standard
    return min(max(float(wert), tief), hoch)


def _kategorie(a: dict[str, Any]) -> str:
    from .runs import kategorie                      # erst beim Aufruf: runs.py importiert dieses Modul
    return kategorie(a)


def _ids(werte: Any) -> list[str]:
    return [x for x in werte if isinstance(x, str)] if isinstance(werte, list) else []


# ---------------------------------------------------------------------------------------------
# Stichprobe
# ---------------------------------------------------------------------------------------------
def stichprobe_groesse(anzahl_auto: int) -> int:
    """Anteil der automatisch eingetragenen Aenderungen, mindestens `stichprobe_min`, hoechstens alle."""
    cfg = einstellungen()["vertrauen"]
    if not cfg.get("aktiv") or anzahl_auto <= 0:
        return 0
    anteil = _zahl(cfg.get("stichprobe_anteil"), 0.10, 0.0, 1.0)
    mindest = int(_zahl(cfg.get("stichprobe_min"), 3, 0.0))
    return min(anzahl_auto, max(mindest, math.ceil(anteil * anzahl_auto - 1e-9)))


def ziehe_stichprobe(run_id: str, aenderungen: list[dict[str, Any]]) -> list[str]:
    """IDs der Stichprobe aus den AUTOMATISCH eingetragenen Aenderungen. Der Zufall ist fest aus der Lauf-ID abgeleitet
    (nicht aus hash(), das je Start anders waere): gleiche Lauf-ID und gleiche Aenderungen ergeben immer dieselbe Stichprobe.
    Das Ergebnis steht in der Reihenfolge des Plans."""
    auto = [a["id"] for a in aenderungen if a.get("status") == UEBERNOMMEN and isinstance(a.get("id"), str)]
    n = stichprobe_groesse(len(auto))
    if not n:
        return []
    rng = random.Random(int.from_bytes(hashlib.sha256(run_id.encode("utf-8")).digest()[:8], "big"))
    gewaehlt = set(rng.sample(sorted(auto), n))
    return [i for i in auto if i in gewaehlt]


# ---------------------------------------------------------------------------------------------
# Dokument-Ampel
# ---------------------------------------------------------------------------------------------
SATZ = {"gruen": "Alles sicher eingetragen.", "gelb": "Bitte ein paar Punkte ansehen.", "rot": "Hier bitte genau hinsehen."}


def _pl(n: int, ein: str, mehr: str) -> str:
    return f"{n} {ein if n == 1 else mehr}"


def _spec_unklar(plan: dict[str, Any]) -> bool:
    sp = plan.get("spec")
    kandidaten = sp.get("kandidaten") if isinstance(sp, dict) else None
    ohne = isinstance(sp, dict) and isinstance(kandidaten, list) and len(kandidaten) > 1 and not sp.get("erkannt")
    code = any("spec_unklar" in (a.get("pruefcodes") or []) for a in plan.get("aenderungen") or [] if isinstance(a, dict))
    return bool(ohne or code)


def _hat_offenen_code(a: dict[str, Any], praefixe: tuple[str, ...]) -> bool:
    return bool(praefixe) and any(isinstance(c, str) and c.startswith(praefixe) for c in a.get("pruefcodes") or [])


def ampel(plan: dict[str, Any], offene_vorschlaege: int | None = None, stichprobe_offen: int | None = None) -> dict[str, Any]:
    """Stand eines Dokuments als Ampel: {"stufe": gruen|gelb|rot, "satz", "gruende": [...], "offen", "offen_je_100", "stichprobe_offen"}.

    `offene_vorschlaege` / `stichprobe_offen`: wie viele Vorschlaege bzw. Stichproben-Karten der Mensch noch NICHT entschieden hat
    (None = alle). So wird die Ampel beim Arbeiten besser. Rot (nicht per Klick loesbar): Bilanz stimmt nicht, Luecken beim Lesen,
    Spec unklar. Rot auch bei mehr als `rot_offene_je_100` offenen Punkten je 100 Aenderungen. Gelb: offene Punkte oder
    Stichprobe nicht angesehen, oder ob alles gelesen wurde ist unbekannt (aelterer Lauf). Gruen sonst.
    Offene Punkte: unentschiedene Vorschlaege, nicht eingetragene Aenderungen (Handarbeit) und automatische Aenderungen,
    deren Pruefcode mit einem Eintrag aus `offen_code_praefixe` beginnt (z. B. konsens_); unbekannte Codes aendern nichts."""
    cfg = einstellungen()["ampel"]
    praefixe = tuple(str(p) for p in cfg.get("offen_code_praefixe") or [] if isinstance(p, str) and p)
    aen = [a for a in plan.get("aenderungen") or [] if isinstance(a, dict)]
    kats = [(a, _kategorie(a)) for a in aen]
    vorschlaege = sum(1 for _, k in kats if k == "pruefen")
    # "schon eingetragen" (die Excel hat die Aenderung schon) macht keine Arbeit und zaehlt nicht als offener Punkt
    nicht = sum(1 for a, k in kats if k == "nicht" and "schon_eingetragen" not in (a.get("pruefcodes") or []))
    markiert = sum(1 for a, k in kats if k == "automatisch" and _hat_offenen_code(a, praefixe))
    offen_v = vorschlaege if offene_vorschlaege is None else int(min(max(offene_vorschlaege, 0), vorschlaege))
    sp_alle = len(_ids(plan.get("stichprobe")))
    sp_offen = sp_alle if stichprobe_offen is None else int(min(max(stichprobe_offen, 0), sp_alle))
    offen = offen_v + nicht + markiert
    je100 = round(offen * 100 / len(aen), 1) if aen else 0.0

    rot: list[str] = []
    bil = plan.get("bilanz") if isinstance(plan.get("bilanz"), dict) else {}
    if bil.get("stimmt") is False:
        rot.append("Die Bilanz stimmt nicht: Es kann etwas verloren gegangen sein.")
    luecken = plan.get("luecken")
    if bil.get("vollstaendig") is False or (isinstance(luecken, list) and luecken):
        rot.append("Das Dokument wurde nicht ganz gelesen: Dort kann eine Änderung fehlen.")
    if _spec_unklar(plan):
        rot.append("Die Spec des Dokuments ist unklar: Bis sie gewählt ist, wird nichts eingetragen.")
    if aen and je100 > _zahl(cfg.get("rot_offene_je_100"), 40.0, 0.0):
        rot.append(f"Viele offene Punkte: {offen} von {_pl(len(aen), 'Änderung', 'Änderungen')}.")

    gelb: list[str] = []
    if offen_v:
        gelb.append(f"{_pl(offen_v, 'Vorschlag ist', 'Vorschläge sind')} noch offen.")
    if nicht:
        gelb.append(f"{_pl(nicht, 'Änderung wird', 'Änderungen werden')} nicht eingetragen und {'braucht' if nicht == 1 else 'brauchen'} Handarbeit.")
    if markiert:
        gelb.append(f"{_pl(markiert, 'automatische Änderung hat', 'automatische Änderungen haben')} einen offenen Hinweis.")
    if sp_offen:
        gelb.append(f"Stichprobe: {sp_offen} von {sp_alle} noch nicht angesehen.")
    if bil.get("vollstaendig") is None and not isinstance(luecken, list):      # aelterer Abgleich: nie "alles sicher" behaupten
        gelb.append("Ob das Dokument ganz gelesen wurde, ist bei diesem älteren Abgleich nicht geprüft.")

    stufe = "rot" if rot else "gelb" if gelb else "gruen"
    return {"stufe": stufe, "satz": SATZ[stufe], "gruende": rot + gelb if rot else gelb, "offen": offen,
            "offen_je_100": je100, "stichprobe_offen": sp_offen}


# ---------------------------------------------------------------------------------------------
# Auswertung: python -m src.cli vertrauen
# ---------------------------------------------------------------------------------------------
OHNE_CODE = "(ohne Prüfcode)"
_FREIGABE = {("pruefen", "uebernommen"), ("stichprobe", "bestaetigt")}          # ein Mensch hat es angesehen und zugestimmt
_ABLEHNUNG = {("pruefen", "abgelehnt"), ("stichprobe", "abgelehnt")}


def _zaehle(z: dict[str, Any], art: str, entscheidung: str) -> None:
    z["anzahl"] += 1
    z["je_art"][art] = z["je_art"].get(art, 0) + 1
    if (art, entscheidung) in _FREIGABE:
        z["frei"] += 1
    elif (art, entscheidung) in _ABLEHNUNG or entscheidung == "abgewaehlt":
        z["abgelehnt"] += 1                         # abgewaehlt: Mensch hat einen automatischen Eintrag aktiv zurueckgenommen


def letzte_entscheidungen(zeilen: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    """Je Lauf zaehlt nur die letzte erstellte Excel (hoechste Ergebnis-Nummer): Wer Entscheidungen aendert und eine weitere
    Excel erstellt, soll nicht doppelt zaehlen. Die frueheren Zeilen bleiben in der Datei erhalten (append-only)."""
    zeilen = list(zeilen)
    neueste: dict[str, str] = {}
    for z in zeilen:
        e = z.get("ergebnis") if isinstance(z.get("ergebnis"), str) else ""
        if e >= neueste.get(z["run_id"], ""):
            neueste[z["run_id"]] = e
    return [z for z in zeilen if (z.get("ergebnis") if isinstance(z.get("ergebnis"), str) else "") == neueste[z["run_id"]]]


def statistik(zeilen: Iterable[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """Je Pruefcode: anzahl (alle Entscheidungen, auch unbesehene), frei (zugestimmt), abgelehnt, je_art. Eine Entscheidung zaehlt
    bei jedem ihrer Codes mit; ohne Code unter OHNE_CODE. Angesehen = frei + abgelehnt."""
    out: dict[str, dict[str, Any]] = {}
    for z in letzte_entscheidungen(zeilen):
        for code in [c for c in z.get("codes") or [] if isinstance(c, str)] or [OHNE_CODE]:
            _zaehle(out.setdefault(code, {"anzahl": 0, "frei": 0, "abgelehnt": 0, "je_art": {}}), z["art"], z["entscheidung"])
    return out


def quoten(z: dict[str, Any]) -> tuple[int, float | None, float | None]:
    """(angesehen, Freigabequote, Ablehnquote) als Anteile 0..1; None, wenn noch nichts angesehen wurde."""
    angesehen = z["frei"] + z["abgelehnt"]
    return (angesehen, z["frei"] / angesehen, z["abgelehnt"] / angesehen) if angesehen else (0, None, None)


def empfehlung(z: dict[str, Any]) -> str:
    """Ein Satz je Pruefcode. Nur ein Hinweis: das Tool stellt nie von selbst auf Automatik um (Regel 27)."""
    cfg = einstellungen()["auswertung"]
    mindest = int(_zahl(cfg.get("min_angesehen"), 20, 1.0))
    ab = _zahl(cfg.get("freigabe_ab"), 0.98, 0.0, 1.0)
    bis = _zahl(cfg.get("ablehnung_max"), 0.02, 0.0, 1.0)
    angesehen, frei, abl = quoten(z)
    if angesehen < mindest:
        return f"noch zu wenig angesehen ({angesehen} von mindestens {mindest})"
    if abl is not None and abl > bis:
        return "noch nicht automatisieren: zu oft abgelehnt"
    if frei is not None and frei >= ab:
        return "sicher genug zum Automatisieren" if z["je_art"].get("pruefen") else "Automatik bewährt sich"
    return "weiter ansehen: noch nicht eindeutig genug"


def prozent(x: float | None) -> str:
    return "–" if x is None else f"{round(x * 100)} %"


def lese_alle(runs_ordner: Path) -> tuple[list[dict[str, Any]], int, int]:
    """Liest alle entscheidungen.jsonl unter data/runs (nur lesen). Gibt (Zeilen, Zahl der Laeufe mit Protokoll, unlesbare Zeilen) zurueck."""
    zeilen: list[dict[str, Any]] = []
    laeufe = kaputt = 0
    for d in sorted(Path(runs_ordner).iterdir()) if Path(runs_ordner).is_dir() else []:
        datei = d / protokoll.ENTSCHEIDUNGEN
        if d.is_dir() and datei.is_file():
            z, k = protokoll.lese_entscheidungen(datei)
            laeufe += 1 if z else 0
            zeilen += z
            kaputt += k
    return zeilen, laeufe, kaputt


def bericht(zeilen: list[dict[str, Any]], laeufe: int, kaputt: int = 0) -> list[str]:
    """Zeilen fuer `python -m src.cli vertrauen`. Enthaelt nur Pruefcodes, Zahlen und feste Saetze (Regel 16)."""
    from . import codes

    if not zeilen:
        return ["Noch keine Entscheidungen gespeichert.",
                "Sobald du im Dashboard eine Excel erstellst, steht in jedem Lauf eine Datei entscheidungen.jsonl; dann erscheint hier die Auswertung."]
    stat = statistik(zeilen)
    z = [f"Vertrauen: Auswertung von {_pl(laeufe, 'Lauf', 'Läufen')} (je Lauf zählt die zuletzt erstellte Excel).",
         "Freigabe = du hast zugestimmt (Vorschlag übernommen oder Stichprobe „Passt“); Ablehnung = „Falsch“, „Nicht übernehmen“ oder abgewählt.",
         "Eine Entscheidung zählt bei jedem ihrer Prüfcodes mit. „Angesehen“ sind nur Entscheidungen, bei denen du selbst geklickt hast.", ""]
    for code in sorted(stat, key=lambda c: (-stat[c]["anzahl"], c)):
        s = stat[code]
        angesehen, frei, abl = quoten(s)
        z.append(f"{code}: Anzahl {s['anzahl']} · angesehen {angesehen} · Freigabe {prozent(frei)} · Ablehnung {prozent(abl)}")
        beschreibung = codes.kurz(code) if code != OHNE_CODE else "Sichere Einträge ohne Auffälligkeit."
        z.append(f"    {beschreibung or 'Code ohne Beschreibung.'}")
        z.append(f"    Empfehlung: {empfehlung(s)}")
    if kaputt:
        z += ["", f"Hinweis: {_pl(kaputt, 'Zeile', 'Zeilen')} in den Protokollen {'war' if kaputt == 1 else 'waren'} nicht lesbar und {'wurde' if kaputt == 1 else 'wurden'} nicht gezählt."]
    z += ["", "Das ist nur eine Empfehlung. Das Tool trägt nichts von selbst anders ein; Schwellen: config/vertrauen.yaml (auswertung)."]
    return z
