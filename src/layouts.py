"""Layout-Profile: der Aufbau einer Excel (Section-/Titel-/Inhaltsspalte, Modell) von Hand festgelegt.

Wozu: Die Erkennung des Aufbaus (src/excel_model.py) ist eine Heuristik. Ist sie unsicher oder falsch, kann ein Mensch
den Aufbau je Blatt in einer YAML-Datei festlegen -- aber ein Profil geht NIE direkt live (ADR 0003):

  1. Profil steht in config/layouts.yaml (Vorlage, im Repo) oder config/layouts.local.yaml (echte Profile, ignoriert).
  2. `profil vorschau --excel X --profil NAME` zeigt den Aufbau (nur Zahlen/Spaltenbuchstaben, keine Zellinhalte)
     und schreibt einen Vorschau-Datensatz mit dem Hash des Profil-Inhalts nach data/profile/.
  3. `profil bestaetigen --vorschau ID` ist die bewusste Handlung; sie schreibt eine Bestaetigung zu genau diesem Hash.
  4. Nur ein aktives Profil mit Bestaetigung zum aktuellen Hash wird im Analyse-Lauf angenommen. Wird die Datei danach
     geaendert, stimmt der Hash nicht mehr und das Profil wird abgelehnt, bis eine neue Vorschau bestaetigt ist.

Alles in data/profile/ ist append-only (Regel 29): nichts wird ueberschrieben oder geloescht. Zuruecknehmen heisst
`aktiv: false` in der Datei. Der Ordner kommt aus config/paths.yaml (`profile_dir`, Regel 28).
Echte Blatt-/Spaltennamen gehoeren nur in config/layouts.local.yaml und nach data/ (Regel 16), nie ins Repo.
"""
from __future__ import annotations

import difflib
import hashlib
import json
import re
import sys
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml
from openpyxl.utils import column_index_from_string, get_column_letter

from . import config
from .excel_model import ExcelModell, profil_kontext

MODELLE = ("zeile_pro_section", "zeile_pro_absatz")
_PROFIL_KEYS = ("name", "aktiv", "beschreibung", "erkennung", "blaetter")
_ERKENNUNG_KEYS = ("blattnamen", "kopfzeilen_begriffe")
_BLATT_KEYS = ("blatt", "ignorieren", "section_spalte", "titel_spalte", "inhalt_spalte", "modell", "kopfzeile")
_ID = re.compile(r"^\d{8}-\d{6}-[0-9a-f]{6}$")
_SPALTE = re.compile(r"^[A-Z]{1,3}$")


class ProfilFehler(Exception):
    """Ein Profil kann nicht verwendet werden. `meldung` ist einfaches Deutsch fuer den Nutzer."""

    def __init__(self, meldung: str):
        super().__init__(meldung)
        self.meldung = meldung


# ---------------------------------------------------------------------------------------------
# Datenobjekte
# ---------------------------------------------------------------------------------------------
@dataclass
class BlattProfil:
    blatt: str
    ignorieren: bool = False
    section_spalte: int | None = None
    titel_spalte: int | None = None
    inhalt_spalte: int | None = None
    modell: str = ""
    kopfzeile: int = 0                 # Zeile mit den Spaltenueberschriften (0 = keine); darueber/darin gibt es keine Section

    def kanon(self) -> dict[str, Any]:
        b = lambda c: get_column_letter(c) if c else None  # noqa: E731
        return {"blatt": self.blatt, "ignorieren": self.ignorieren, "section_spalte": b(self.section_spalte),
                "titel_spalte": b(self.titel_spalte), "inhalt_spalte": b(self.inhalt_spalte),
                "modell": self.modell, "kopfzeile": self.kopfzeile}


@dataclass
class Profil:
    name: str
    aktiv: bool = False
    beschreibung: str = ""
    blattnamen: tuple[str, ...] = ()              # Erkennung: diese Blaetter muessen vorkommen (nur Vorschlag)
    kopfzeilen_begriffe: tuple[str, ...] = ()     # Erkennung: diese Begriffe stehen in der ersten Zeile (nur Vorschlag)
    blaetter: dict[str, BlattProfil] = field(default_factory=dict)
    quelle: str = ""                              # Dateiname, aus dem das Profil kommt
    hash: str = ""

    def kanon(self) -> dict[str, Any]:
        """Der Inhalt, der den Hash bestimmt. `aktiv` und `beschreibung` gehoeren nicht dazu (Ein-/Ausschalten und
        Kommentare aendern nichts am Aufbau), alles andere schon."""
        return {"name": self.name,
                "erkennung": {"blattnamen": sorted(self.blattnamen), "kopfzeilen_begriffe": sorted(self.kopfzeilen_begriffe)},
                "blaetter": [self.blaetter[k].kanon() for k in sorted(self.blaetter)]}


def hash_von(kanon: dict[str, Any]) -> str:
    text = json.dumps(kanon, sort_keys=True, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def kurz(h: str) -> str:
    return h[:12]


# ---------------------------------------------------------------------------------------------
# YAML lesen und pruefen
# ---------------------------------------------------------------------------------------------
def profil_dateien() -> list[Path]:
    """Vorlage (im Repo) und lokale Datei (ignoriert; hier liegen die echten Profile)."""
    return [config.CONFIG_DIR / "layouts.yaml", config.CONFIG_DIR / "layouts.local.yaml"]


def _nahe(wort: str, erlaubt: tuple[str, ...]) -> str:
    treffer = difflib.get_close_matches(wort, erlaubt, n=1, cutoff=0.6)
    return f" Meinten Sie '{treffer[0]}'?" if treffer else ""


def _unbekannte(d: dict, erlaubt: tuple[str, ...], wo: str, probleme: list[str]) -> None:
    for k in d:
        if k not in erlaubt:
            probleme.append(f"{wo}: Der Schlüssel '{k}' ist unbekannt.{_nahe(str(k), erlaubt)} "
                            f"Erlaubt sind: {', '.join(erlaubt)}.")


def _spalte(wert: Any, schluessel: str, wo: str, probleme: list[str]) -> int | None:
    if wert is None:
        return None
    if not isinstance(wert, str):
        probleme.append(f"{wo}: '{schluessel}' muss ein Spaltenbuchstabe in Anführungszeichen sein, z. B. \"C\".")
        return None
    text = wert.strip().upper()
    if not _SPALTE.match(text) or column_index_from_string(text) > 16384:
        probleme.append(f"{wo}: '{schluessel}' ist kein gültiger Spaltenbuchstabe (erlaubt: A bis XFD, z. B. \"C\").")
        return None
    return column_index_from_string(text)


def _liste_text(wert: Any, schluessel: str, wo: str, probleme: list[str]) -> tuple[str, ...]:
    if wert is None:
        return ()
    if not isinstance(wert, list) or not all(isinstance(x, str) and x.strip() for x in wert):
        probleme.append(f"{wo}: '{schluessel}' muss eine Liste mit Texten sein.")
        return ()
    return tuple(x.strip() for x in wert)


def _blatt_pruefen(eintrag: Any, nr: int, wo: str, probleme: list[str]) -> BlattProfil | None:
    wo = f"{wo}, Blatt-Eintrag {nr}"
    if not isinstance(eintrag, dict):
        probleme.append(f"{wo}: Jeder Eintrag unter 'blaetter' braucht einen Namen ('blatt') und Angaben dazu.")
        return None
    n0 = len(probleme)
    _unbekannte(eintrag, _BLATT_KEYS, wo, probleme)
    name = eintrag.get("blatt")
    if not isinstance(name, str) or not name.strip():
        probleme.append(f"{wo}: Es fehlt der Name des Tabellenblatts ('blatt').")
        return None
    wo = f"{wo} ('{name}')"
    ign = eintrag.get("ignorieren", False)
    if not isinstance(ign, bool):
        probleme.append(f"{wo}: 'ignorieren' muss true oder false sein.")
        ign = False
    sek = _spalte(eintrag.get("section_spalte"), "section_spalte", wo, probleme)
    tit = _spalte(eintrag.get("titel_spalte"), "titel_spalte", wo, probleme)
    inh = _spalte(eintrag.get("inhalt_spalte"), "inhalt_spalte", wo, probleme)
    modell = eintrag.get("modell", "")
    kopf = eintrag.get("kopfzeile", 0)
    if not isinstance(kopf, int) or isinstance(kopf, bool) or not 0 <= kopf <= 1048576:
        probleme.append(f"{wo}: 'kopfzeile' muss eine Zeilennummer sein (0 = keine Kopfzeile).")
        kopf = 0
    if ign:
        if any(eintrag.get(k) is not None for k in ("section_spalte", "titel_spalte", "inhalt_spalte", "modell")):
            probleme.append(f"{wo}: Ein ignoriertes Blatt (ignorieren: true) darf keine Spalten oder kein Modell haben.")
    else:
        if inh is None:
            probleme.append(f"{wo}: Es fehlt 'inhalt_spalte' (oder das Blatt mit 'ignorieren: true' ausschließen).")
        if modell not in MODELLE:
            probleme.append(f"{wo}: 'modell' muss {' oder '.join(MODELLE)} sein.")
        if modell == "zeile_pro_section" and sek is None:
            probleme.append(f"{wo}: Das Modell 'zeile_pro_section' braucht eine 'section_spalte'.")
        if tit is not None and sek is None:
            probleme.append(f"{wo}: Eine 'titel_spalte' ergibt nur zusammen mit einer 'section_spalte' Sinn.")
        belegt = [c for c in (sek, tit, inh) if c is not None]
        if len(belegt) != len(set(belegt)):
            probleme.append(f"{wo}: Section-, Titel- und Inhaltsspalte müssen verschiedene Spalten sein.")
    if len(probleme) > n0:
        return None
    return BlattProfil(blatt=name, ignorieren=ign, section_spalte=sek, titel_spalte=tit, inhalt_spalte=inh,
                       modell="" if ign else modell, kopfzeile=kopf)


def profil_aus_dict(d: Any, quelle: str = "", position: int = 1) -> Profil:
    """Prueft einen Profil-Eintrag und baut das Profil. Bei Fehlern: ProfilFehler mit ALLEN Problemen in einfachem Deutsch."""
    probleme: list[str] = []
    name = d.get("name") if isinstance(d, dict) else None
    wo = f"Profil '{name}'" if isinstance(name, str) and name.strip() else f"Profil Nr. {position} in {quelle or 'der Datei'}"
    if not isinstance(d, dict):
        raise ProfilFehler(f"{wo}: Ein Profil muss aus Schlüsseln und Werten bestehen (name, blaetter, ...).")
    _unbekannte(d, _PROFIL_KEYS, wo, probleme)
    if not isinstance(name, str) or not name.strip() or len(name) > 80 or any(ord(c) < 32 for c in name):
        probleme.append(f"{wo}: Es fehlt ein 'name' (Text, höchstens 80 Zeichen).")
        name = ""
    aktiv = d.get("aktiv", False)
    if not isinstance(aktiv, bool):
        probleme.append(f"{wo}: 'aktiv' muss true oder false sein.")
        aktiv = False
    beschr = d.get("beschreibung", "")
    if not isinstance(beschr, str):
        probleme.append(f"{wo}: 'beschreibung' muss ein Text sein.")
        beschr = ""
    blattnamen = kopfbegriffe = ()
    erk = d.get("erkennung")
    if erk is not None:
        if not isinstance(erk, dict):
            probleme.append(f"{wo}: 'erkennung' muss Schlüssel enthalten ({', '.join(_ERKENNUNG_KEYS)}).")
        else:
            _unbekannte(erk, _ERKENNUNG_KEYS, f"{wo}, erkennung", probleme)
            blattnamen = _liste_text(erk.get("blattnamen"), "blattnamen", f"{wo}, erkennung", probleme)
            kopfbegriffe = _liste_text(erk.get("kopfzeilen_begriffe"), "kopfzeilen_begriffe", f"{wo}, erkennung", probleme)
    blaetter: dict[str, BlattProfil] = {}
    roh = d.get("blaetter")
    if not isinstance(roh, list) or not roh:
        probleme.append(f"{wo}: 'blaetter' muss eine Liste mit mindestens einem Blatt sein.")
    else:
        for i, e in enumerate(roh, 1):
            bp = _blatt_pruefen(e, i, wo, probleme)
            if bp is None:
                continue
            if bp.blatt in blaetter:
                probleme.append(f"{wo}: Das Blatt '{bp.blatt}' ist doppelt aufgeführt.")
            else:
                blaetter[bp.blatt] = bp
    if probleme:
        raise ProfilFehler("\n".join(probleme))
    p = Profil(name=name.strip(), aktiv=aktiv, beschreibung=beschr, blattnamen=blattnamen,
               kopfzeilen_begriffe=kopfbegriffe, blaetter=blaetter, quelle=quelle)
    p.hash = hash_von(p.kanon())
    return p


def lade_alle() -> tuple[dict[str, Profil], list[tuple[str | None, str]]]:
    """Liest alle Profil-Dateien. -> (gueltige Profile nach Name, Fehler als (Profilname oder None, Meldung)).
    Doppelte Namen (auch ueber beide Dateien) machen alle betroffenen Profile ungueltig: nichts wird geraten."""
    profile: dict[str, Profil] = {}
    fehler: list[tuple[str | None, str]] = []
    doppelt: set[str] = set()
    for datei in profil_dateien():
        if not datei.is_file():
            continue
        try:
            roh = yaml.safe_load(datei.read_text(encoding="utf-8")) or {}
        except (OSError, UnicodeDecodeError, yaml.YAMLError):
            fehler.append((None, f"Die Datei {datei.name} konnte nicht gelesen werden (kein gültiges YAML)."))
            continue
        liste = roh.get("profile") if isinstance(roh, dict) else None
        if liste is None:
            liste = []
        if not isinstance(roh, dict) or not isinstance(liste, list):
            fehler.append((None, f"Die Datei {datei.name} ist falsch aufgebaut: Es muss 'profile:' mit einer Liste von Profilen geben."))
            continue
        if isinstance(roh, dict):
            for k in roh:
                if k != "profile":
                    fehler.append((None, f"Die Datei {datei.name} enthält den unbekannten Schlüssel '{k}' (erlaubt ist nur 'profile')."))
        for i, eintrag in enumerate(liste, 1):
            try:
                p = profil_aus_dict(eintrag, datei.name, i)
            except ProfilFehler as e:
                n = eintrag.get("name") if isinstance(eintrag, dict) and isinstance(eintrag.get("name"), str) else None
                fehler.append((n, f"{datei.name}: {e.meldung}"))
                continue
            if p.name in profile or p.name in doppelt:
                doppelt.add(p.name)
                profile.pop(p.name, None)
                fehler.append((p.name, f"Der Profilname '{p.name}' kommt mehrfach vor (zuletzt in {datei.name}). "
                                       "Namen müssen eindeutig sein, deshalb wird keines dieser Profile benutzt."))
            else:
                profile[p.name] = p
    return profile, fehler


# ---------------------------------------------------------------------------------------------
# Speicher: data/profile/ (append-only)
# ---------------------------------------------------------------------------------------------
def profile_dir() -> Path:
    d = config.pfad("profile_dir")
    d.mkdir(parents=True, exist_ok=True)
    return d


def _neue_id() -> str:
    return f"{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}"


def _schreibe_neu(pfad: Path, daten: dict[str, Any]) -> None:
    """Legt eine NEUE Datei an ('x': eine vorhandene wird nie ueberschrieben, Regel 29)."""
    with open(pfad, "x", encoding="utf-8") as f:
        json.dump(daten, f, ensure_ascii=False, indent=1)
        f.flush()


def _inhalt_dict(p: Profil) -> dict[str, Any]:
    return {"kanon": p.kanon(), "aktiv": p.aktiv}


def _profil_aus_kanon(kanon: dict[str, Any]) -> Profil:
    """Baut ein Profil aus dem gespeicherten Inhalt (Bestaetigung) wieder auf, mit derselben Pruefung wie aus der YAML."""
    d = {"name": kanon["name"], "aktiv": True, "erkennung": kanon["erkennung"],
         "blaetter": [{k: v for k, v in b.items() if v is not None and v != "" and not (k == "ignorieren" and v is False)
                       and not (k == "kopfzeile" and v == 0)} for b in kanon["blaetter"]]}
    p = profil_aus_dict(d, "Bestätigung")
    return p


def bestaetigungen() -> dict[str, dict[str, Any]]:
    """Alle gueltigen Bestaetigungen: Hash -> Datensatz. Unlesbare oder veraenderte Dateien zaehlen NICHT (sicher)."""
    out: dict[str, dict[str, Any]] = {}
    for f in sorted(profile_dir().glob("bestaetigt_*.json")):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
            h = d["profil"]["hash"]
            if hash_von(d["profil"]["kanon"]) != h:
                continue                                  # Inhalt passt nicht zum Hash: nicht vertrauen
            out.setdefault(h, d)
        except (OSError, ValueError, KeyError, TypeError):
            continue
    return out


def status(p: Profil) -> str:
    """'bestaetigt' (Hash passt zu einer Bestaetigung) | 'geaendert' (Name bestaetigt, aber anderer Inhalt) | 'unbestaetigt'."""
    alle = bestaetigungen()
    if p.hash in alle:
        return "bestaetigt"
    if any(d["profil"].get("name") == p.name for d in alle.values()):
        return "geaendert"
    return "unbestaetigt"


def finde(name: str) -> Profil:
    profile, fehler = lade_alle()
    if name in profile:
        return profile[name]
    eigene = [m for n, m in fehler if n == name]
    if eigene:
        raise ProfilFehler(f"Das Profil '{name}' enthält Fehler und kann nicht benutzt werden:\n" + "\n".join(eigene))
    raise ProfilFehler(f"Es gibt kein Profil mit dem Namen '{name}'. Vorhandene Profile zeigt: python -m src.cli profil liste")


def bestaetigtes_profil(name: str) -> Profil:
    """Das Profil, das im Analyse-Lauf benutzt werden darf -- oder ProfilFehler (einfaches Deutsch).
    Bedingungen: steht in einer Datei, ist aktiv, und zu GENAU diesem Inhalt (Hash) gibt es eine bestaetigte Vorschau."""
    p = finde(name)
    if not p.aktiv:
        raise ProfilFehler(f"Das Profil '{name}' ist nicht aktiv. Es wird nur benutzt, wenn in der Datei 'aktiv: true' steht.")
    st = status(p)
    if st == "geaendert":
        raise ProfilFehler(f"Das Profil '{name}' wurde nach der Bestätigung geändert und darf deshalb nicht benutzt werden. "
                           "Bitte eine neue Vorschau machen und bestätigen (profil vorschau, profil bestaetigen).")
    if st != "bestaetigt":
        raise ProfilFehler(f"Das Profil '{name}' ist noch nicht bestätigt. Neue Profile gehen nie direkt live: "
                           "Bitte erst eine Vorschau ansehen (profil vorschau) und sie dann bestätigen (profil bestaetigen).")
    return p


def profil_aus_bestaetigung(name: str, hash_: str) -> Profil:
    """Das bestaetigte Profil zu (Name, Hash) -- unabhaengig davon, was inzwischen in der YAML-Datei steht.
    Gebraucht, wenn ein analysierter Lauf spaeter uebernommen wird: dann muss derselbe Aufbau gelten wie bei der Analyse."""
    d = bestaetigungen().get(hash_)
    if d is None or d["profil"].get("name") != name:
        raise ProfilFehler(f"Zum Profil '{name}' (Prüfsumme {kurz(hash_)}) gibt es keine Bestätigung mehr. "
                           "Bitte den Lauf mit einem bestätigten Profil neu analysieren.")
    p = _profil_aus_kanon(d["profil"]["kanon"])
    if p.hash != hash_:
        raise ProfilFehler(f"Die gespeicherte Bestätigung des Profils '{name}' ist beschädigt. Bitte neu bestätigen.")
    return p


# ---------------------------------------------------------------------------------------------
# Vorschau (keine Zellinhalte, Regel 16)
# ---------------------------------------------------------------------------------------------
def _sha256(pfad: Path) -> str:
    h = hashlib.sha256()
    with open(pfad, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def _ohne_inhalt(text: str, blattnamen: list[str], profilname: str = "") -> str:
    """Macht eine Meldung des Excel-Modells frei von Inhalten: Blattnamen werden zu 'Blatt N', alles andere in
    Anfuehrungszeichen (z. B. Spaltenueberschriften) faellt weg. Der Profilname bleibt (er steht in der eigenen YAML)."""
    marke = "\x00PROFIL\x00"
    if profilname:
        text = text.replace(f"'{profilname}'", marke)
    for i, n in sorted(enumerate(blattnamen, 1), key=lambda t: -len(t[1])):
        text = text.replace(f"Blatt '{n}'", f"Blatt {i}").replace(f"'{n}'", f"Blatt {i}")
    text = re.sub(r"\s*\('[^']*'\)", "", text)
    text = re.sub(r"'[^']*'", "[…]", text)
    return text.replace(marke, f"'{profilname}'")


def _spalten(b: dict[str, Any]) -> dict[str, Any]:
    return {k: b[k] for k in ("section_spalte", "titel_spalte", "inhalt_spalte", "modell", "sicherheit")}


def _kopf_begriffe(xl: ExcelModell, blatt: str | None = None) -> set[str]:
    """Woerter der ersten belegten Zeile (nur zum Vergleichen mit der Erkennung, wird nie ausgegeben)."""
    out: set[str] = set()
    for name in ([blatt] if blatt else list(xl.layouts)):
        zeilen = [r for (b, r, _) in xl.zellen if b == name]
        if not zeilen:
            continue
        erste = min(zeilen)
        for (b, r, c), t in xl.zellen.items():
            if b == name and r == erste:
                out.update(re.findall(r"\w+", t.casefold()))
    return out


def passende_profile(xl: ExcelModell, profile: dict[str, Profil]) -> list[str]:
    """Namen der Profile, deren `erkennung` auf diese Excel passt -- NUR ein Vorschlag, nie stilles Anwenden."""
    namen_klein = {n.casefold(): n for n in xl.layouts}
    treffer = []
    for p in profile.values():
        if not p.blattnamen and not p.kopfzeilen_begriffe:
            continue                                      # ohne Muster wird nie vorgeschlagen
        if any(b.casefold() not in namen_klein for b in p.blattnamen):
            continue
        if p.kopfzeilen_begriffe:
            ziele = [namen_klein[b.casefold()] for b in p.blattnamen] or [None]
            vorhanden = set()
            for z in ziele:
                vorhanden |= _kopf_begriffe(xl, z)
            if any(t.casefold() not in vorhanden for t in p.kopfzeilen_begriffe):
                continue
        treffer.append(p.name)
    return sorted(treffer)


def erstelle_vorschau(excel: str | Path, profil: Profil | None = None) -> dict[str, Any]:
    """Zahlen zum Aufbau je Blatt (Profil bzw. automatische Erkennung). Enthaelt KEINE Zellinhalte und keine Blattnamen."""
    excel = Path(excel)
    with profil_kontext(None):
        auto = ExcelModell(excel)                         # Heuristik, zum Vergleichen und fuer Vorschlaege
        xl = ExcelModell(excel, profil) if profil is not None else auto
    namen = list(xl.layouts)
    kandidaten: dict[str, int] = {}
    for k in xl.kandidaten:
        kandidaten[k.blatt] = kandidaten.get(k.blatt, 0) + 1
    blaetter = []
    for nr, (name, lay) in enumerate(xl.layouts.items(), 1):
        b = lay.beschreibung()
        eintrag: dict[str, Any] = {
            "nr": nr, "quelle": "ignoriert" if lay.ignoriert else ("profil" if lay.profil else "automatisch"),
            "section_spalte": b["section_spalte"], "titel_spalte": b["titel_spalte"], "inhalt_spalte": b["inhalt_spalte"],
            "modell": b["modell"], "sections": b["sections"], "zeilen": b["zeilen"], "kandidaten": kandidaten.get(name, 0),
            "sicherheit": b["sicherheit"], "gruende": [_ohne_inhalt(g, namen) for g in b["sicherheit_gruende"]],
            "versteckt": b["versteckt"], "kopfzeile": None, "heuristik": None}
        if profil is not None and name in profil.blaetter and not lay.ignoriert:
            eintrag["kopfzeile"] = profil.blaetter[name].kopfzeile
            eintrag["heuristik"] = _spalten(auto.layouts[name].beschreibung())
        blaetter.append(eintrag)
    warn = [_ohne_inhalt(w, namen, profil.name if profil else "") for w in xl.warnungen]
    vorschlaege = passende_profile(auto, lade_alle()[0]) if profil is None else []
    return {"excel_sha256": _sha256(excel), "blaetter": blaetter, "warnungen": warn, "vorschlaege": vorschlaege}


def schreibe_vorschau(profil: Profil, daten: dict[str, Any]) -> str:
    """Speichert den Vorschau-Datensatz (append-only) und gibt seine ID zurueck."""
    vid = _neue_id()
    _schreibe_neu(profile_dir() / f"vorschau_{vid}.json", {
        "id": vid, "erstellt": datetime.now().isoformat(timespec="seconds"),
        "profil": {"name": profil.name, "hash": profil.hash, "kanon": profil.kanon()},
        "excel_sha256": daten["excel_sha256"], "blaetter": daten["blaetter"], "warnungen": daten["warnungen"]})
    return vid


def _lese_vorschau(vid: str) -> dict[str, Any]:
    if not _ID.match(vid):
        raise ProfilFehler("Die Vorschau-ID ist ungültig. Sie sieht so aus: 20261015-143000-ab12cd (steht am Ende der Vorschau).")
    pfad = profile_dir() / f"vorschau_{vid}.json"
    try:
        d = json.loads(pfad.read_text(encoding="utf-8"))
        if hash_von(d["profil"]["kanon"]) != d["profil"]["hash"]:
            raise ValueError
        return d
    except FileNotFoundError:
        raise ProfilFehler(f"Die Vorschau {vid} wurde nicht gefunden.") from None
    except (OSError, ValueError, KeyError, TypeError):
        raise ProfilFehler(f"Die Vorschau {vid} ist beschädigt. Bitte eine neue Vorschau machen.") from None


def bestaetige(vorschau_id: str) -> tuple[Profil, bool]:
    """Die bewusste Handlung: macht das Profil, das in der Vorschau gezeigt wurde, benutzbar.
    -> (Profil, neu). Abgelehnt, wenn die Datei seit der Vorschau geaendert wurde oder das Profil nicht aktiv ist."""
    v = _lese_vorschau(vorschau_id)
    name, h = v["profil"]["name"], v["profil"]["hash"]
    p = finde(name)
    if p.hash != h:
        raise ProfilFehler(f"Das Profil '{name}' wurde seit der Vorschau geändert. Bestätigt wird nur, was in der Vorschau zu sehen war. "
                           "Bitte eine neue Vorschau machen.")
    if not p.aktiv:
        raise ProfilFehler(f"Das Profil '{name}' ist nicht aktiv. Bitte in der Datei 'aktiv: true' setzen (der Hash ändert sich dadurch nicht) "
                           "und dann bestätigen.")
    if p.hash in bestaetigungen():
        return p, False
    _schreibe_neu(profile_dir() / f"bestaetigt_{p.hash[:16]}_{_neue_id()}.json", {
        "bestaetigt": datetime.now().isoformat(timespec="seconds"), "vorschau": vorschau_id,
        "profil": {"name": p.name, "hash": p.hash, "kanon": p.kanon()}})
    return p, True


# ---------------------------------------------------------------------------------------------
# Kommandozeile: python -m src.cli profil ...
# ---------------------------------------------------------------------------------------------
_STATUS_TEXT = {"bestaetigt": "bestätigt", "geaendert": "GEÄNDERT seit der Bestätigung (neue Vorschau nötig)",
                "unbestaetigt": "nicht bestätigt"}


def formatiere_vorschau(daten: dict[str, Any], profil: Profil | None) -> list[str]:
    z: list[str] = []
    for b in daten["blaetter"]:
        quelle = {"profil": "durch Profil festgelegt", "ignoriert": "laut Profil ignoriert",
                  "automatisch": "automatisch erkannt"}[b["quelle"]]
        z.append(f"Blatt {b['nr']}: {quelle}")
        if b["quelle"] == "ignoriert":
            continue
        sp = lambda k: b[k] or "-"  # noqa: E731
        z.append(f"   Section-Spalte {sp('section_spalte')} | Titel-Spalte {sp('titel_spalte')} | Inhalt-Spalte {sp('inhalt_spalte')} | "
                 f"Modell {b['modell']}" + (f" | Kopfzeile {b['kopfzeile']}" if b["kopfzeile"] else ""))
        z.append(f"   Sections {b['sections']} | Zeilen {b['zeilen']} | Kandidaten {b['kandidaten']} | Sicherheit {b['sicherheit']}"
                 + (" | Blatt ausgeblendet" if b["versteckt"] else ""))
        for g in b["gruende"]:
            z.append(f"   Zweifel: {g}")
        h = b.get("heuristik")
        if h:
            gleich = all(h[k] == b[k] for k in ("section_spalte", "titel_spalte", "inhalt_spalte", "modell"))
            if gleich:
                z.append("   Die automatische Erkennung käme zum gleichen Ergebnis.")
            else:
                sp2 = lambda k: h[k] or "-"  # noqa: E731
                z.append(f"   Die automatische Erkennung hätte gewählt: Section {sp2('section_spalte')} | Titel {sp2('titel_spalte')} | "
                         f"Inhalt {sp2('inhalt_spalte')} | Modell {h['modell']} | Sicherheit {h['sicherheit']}")
    for w in daten["warnungen"]:
        z.append(f"Hinweis: {w}")
    return z


def _vorschau_befehl(args: list[str]) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="cde profil vorschau", description="Zeigt den Aufbau einer Excel (ohne Zellinhalte).")
    ap.add_argument("--excel", required=True)
    ap.add_argument("--profil", help="Name des Profils; ohne Angabe wird die automatische Erkennung gezeigt")
    a = ap.parse_args(args)
    excel = Path(a.excel)
    if not excel.is_file():
        print("Die Excel-Datei wurde nicht gefunden.")
        return 2
    profil = finde(a.profil) if a.profil else None
    daten = erstelle_vorschau(excel, profil)
    if profil:
        print(f"Vorschau für das Profil '{profil.name}' (Prüfsumme {kurz(profil.hash)}"
              f"{'' if profil.aktiv else ', noch NICHT aktiv'})")
    else:
        print("Vorschau der automatischen Erkennung (kein Profil gewählt)")
    print(f"Excel-Prüfsumme: {kurz(daten['excel_sha256'])} | Blätter: {len(daten['blaetter'])}")
    for zeile in formatiere_vorschau(daten, profil):
        print(zeile)
    if profil is None:
        if daten["vorschlaege"]:
            print("Passende Profile (nur ein Vorschlag, es wird nichts automatisch angewendet): " + ", ".join(daten["vorschlaege"]))
        print("Es wurde nichts gespeichert. Mit --profil NAME entsteht eine Vorschau, die bestätigt werden kann.")
        return 0
    vid = schreibe_vorschau(profil, daten)
    print(f"Vorschau-ID: {vid}")
    print(f"Wenn der Aufbau stimmt: python -m src.cli profil bestaetigen --vorschau {vid}")
    return 0


def _bestaetigen_befehl(args: list[str]) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="cde profil bestaetigen", description="Bestätigt ein Profil anhand einer Vorschau.")
    ap.add_argument("--vorschau", required=True, help="ID der Vorschau (steht am Ende von 'profil vorschau')")
    a = ap.parse_args(args)
    p, neu = bestaetige(a.vorschau)
    if neu:
        print(f"Das Profil '{p.name}' (Prüfsumme {kurz(p.hash)}) ist jetzt bestätigt und darf im Analyse-Lauf benutzt werden.")
        print("Ändert sich der Inhalt des Profils, ist eine neue Vorschau und Bestätigung nötig.")
    else:
        print(f"Das Profil '{p.name}' (Prüfsumme {kurz(p.hash)}) war schon bestätigt. Es wurde nichts geändert.")
    return 0


def _liste_befehl(args: list[str]) -> int:
    profile, fehler = lade_alle()
    if not profile and not fehler:
        print("Es sind keine Profile angelegt (siehe config/layouts.yaml als Vorlage; echte Profile in config/layouts.local.yaml).")
    for p in profile.values():
        print(f"{p.name} | {'aktiv' if p.aktiv else 'nicht aktiv'} | {_STATUS_TEXT[status(p)]} | Prüfsumme {kurz(p.hash)} | "
              f"{len(p.blaetter)} Blatt-Einträge | aus {p.quelle}")
    for _, m in fehler:
        print(f"FEHLER: {m}")
    return 1 if fehler else 0


def _analyse_befehl(args: list[str]) -> int:
    import argparse
    from . import runs
    from .cli import _zeige
    ap = argparse.ArgumentParser(prog="cde profil analyse", description="Analyse mit einem bestätigten Profil (schreibt nichts in die Excel).")
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--excel", required=True)
    ap.add_argument("--profil", required=True)
    a = ap.parse_args(args)
    bestaetigtes_profil(a.profil)                 # vorab pruefen: sonst entsteht ein Lauf, der sofort scheitert
    run_id = runs.neuer_lauf(a.pdf, a.excel)
    daten = runs.analysiere(run_id, a.profil)
    print("Lauf:", run_id, "| Profil:", a.profil)
    _zeige(daten)
    return 0


def main(args: list[str]) -> int:
    befehle = {"vorschau": _vorschau_befehl, "bestaetigen": _bestaetigen_befehl, "liste": _liste_befehl,
               "analyse": _analyse_befehl}
    if not args or args[0] not in befehle:
        print("Verwendung: python -m src.cli profil vorschau|bestaetigen|liste|analyse ...\n"
              "  vorschau    --excel X.xlsx [--profil NAME]   Aufbau anzeigen (ohne Zellinhalte)\n"
              "  bestaetigen --vorschau ID                    Profil nach der Vorschau freigeben\n"
              "  liste                                        alle Profile und ihren Status zeigen\n"
              "  analyse     --pdf P --excel X --profil NAME  Analyse mit einem bestätigten Profil")
        return 2
    try:
        return befehle[args[0]](args[1:])
    except ProfilFehler as e:
        print(e.meldung)
        return 2
    except ValueError as e:                        # z. B. kaputte Excel: einfache Meldung aus dem Excel-Modell
        print(str(e))
        return 2
    except SystemExit as e:                        # argparse beendet selbst
        return 2 if e.code else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
