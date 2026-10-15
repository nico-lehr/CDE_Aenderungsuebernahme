"""Erzeugt erfundene Testpaare (PDF + vorher.xlsx + nachher.xlsx + erwartet.json) fuer das Messwerkzeug.

Beispiele (im Projektordner, mit der .venv):
    python scripts/erzeuge_paare.py --korpus entwicklung --seeds 1-60
    python scripts/erzeuge_paare.py --korpus test --seeds 1001-1030
    python scripts/erzeuge_paare.py --korpus entwicklung --seeds 1-5 --ziel D:\\temp\\paare

Wohin:  <paare_dir aus config/paths.yaml>/<korpus>/synth_<seed>_<profil>/
        --ziel DIR ersetzt paare_dir (die Paare landen dann in DIR/<korpus>/).
Regeln: Alles ist erfunden (Regel 16). Bestehende Paar-Ordner werden nie ueberschrieben, sondern
        uebersprungen (Regel 29). Entwicklung (Seeds < 1000) und Test (Seeds >= 1000) bleiben getrennt (Regel 11).
Profil: gemischt (Standard) = jedes dritte Paar "schwer", die anderen "standard"; oder --profil standard|schwer.
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from collections import Counter
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from tests.fixtures import generator  # noqa: E402


def lese_seeds(text: str) -> list[int]:
    """'1-60' -> [1..60]; '1,5,9' -> [1,5,9]; Mischformen wie '1-3,10' sind erlaubt."""
    seeds: list[int] = []
    for teil in text.split(","):
        teil = teil.strip()
        if not teil:
            continue
        if "-" in teil:
            a, b = teil.split("-", 1)
            von, bis = int(a), int(b)
            if bis < von:
                raise ValueError(f"ungueltiger Bereich: {teil}")
            seeds.extend(range(von, bis + 1))
        else:
            seeds.append(int(teil))
    if not seeds:
        raise ValueError("keine Seeds angegeben")
    return sorted(set(seeds))


def profil_fuer(seed: int, modus: str) -> str:
    if modus == "gemischt":
        return "schwer" if seed % 3 == 0 else "standard"
    return modus


def ziel_ordner(korpus: str, ziel: str | None) -> Path:
    if ziel:
        return Path(ziel) / korpus
    from src import config
    try:
        return config.pfad("paare_dir") / korpus
    except KeyError:
        raise SystemExit("config/paths.yaml hat keinen Eintrag 'paare_dir' -- bitte --ziel angeben.")


def schreibe_korpus_json(ordner: Path, korpus: str, seeds: list[int]) -> None:
    """KORPUS.json: Generator-Version, Seeds, Datum. Seeds frueherer Laeufe (gleiche Version) bleiben erhalten."""
    pfad = ordner / "KORPUS.json"
    alt_seeds: list[int] = []
    erstellt = date.today().isoformat()
    if pfad.exists():
        alt = json.loads(pfad.read_text(encoding="utf-8"))
        if alt.get("generator_version") != generator.GENERATOR_VERSION:
            raise SystemExit(f"{pfad} gehoert zu Generator-Version {alt.get('generator_version')}, "
                             f"dieser Generator ist Version {generator.GENERATOR_VERSION}. "
                             "Zahlen verschiedener Versionen sind nicht vergleichbar -- bitte einen neuen Ordner nehmen.")
        alt_seeds = list(alt.get("seeds", []))
        erstellt = alt.get("erstellt", erstellt)
    daten = {"generator_version": generator.GENERATOR_VERSION, "korpus": korpus,
             "seeds": sorted(set(alt_seeds) | set(seeds)), "erstellt": erstellt}
    tmp = ordner / "KORPUS.json.tmp"
    tmp.write_text(json.dumps(daten, ensure_ascii=False) + "\n", encoding="utf-8")
    tmp.replace(pfad)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Erzeugt erfundene Testpaare mit bekannter Wahrheit.")
    ap.add_argument("--korpus", required=True, choices=["entwicklung", "test"])
    ap.add_argument("--seeds", required=True, help="z. B. 1-60 oder 1001-1030 oder 1,5,9")
    ap.add_argument("--profil", default="gemischt", choices=["gemischt", "standard", "schwer"])
    ap.add_argument("--ziel", default=None, help="Basisordner statt paare_dir aus config/paths.yaml")
    args = ap.parse_args(argv)

    try:
        seeds = lese_seeds(args.seeds)
    except ValueError as e:
        print(f"Fehler: {e}", file=sys.stderr)
        return 2
    grenze = generator.SEED_GRENZE
    if args.korpus == "entwicklung" and any(s >= grenze for s in seeds):
        print(f"Fehler: Entwicklungs-Seeds muessen unter {grenze} liegen (Test-Seeds sind getrennt, Regel 11).", file=sys.stderr)
        return 2
    if args.korpus == "test" and any(s < grenze for s in seeds):
        print(f"Fehler: Test-Seeds muessen ab {grenze} liegen (Entwicklungs-Seeds sind getrennt, Regel 11).", file=sys.stderr)
        return 2

    ordner = ziel_ordner(args.korpus, args.ziel)
    ordner.mkdir(parents=True, exist_ok=True)
    start = time.time()
    neu: list[int] = []
    uebersprungen: list[str] = []
    aenderungen = 0
    zellen = 0
    stoerer: Counter[str] = Counter()
    for seed in seeds:
        profil = profil_fuer(seed, args.profil)
        name = f"synth_{seed:04d}_{profil}"
        if (ordner / name).exists():
            uebersprungen.append(name)
            continue
        paar = generator.erzeuge_paar(seed, profil, ordner, args.korpus)
        e = json.loads((paar / "erwartet.json").read_text(encoding="utf-8"))
        aenderungen += len(e["aenderungen"])
        zellen += len({(a["blatt"], a["zeile"], a["spalte"]) for a in e["aenderungen"]})
        stoerer.update(e["stoerer"])
        neu.append(seed)
    schreibe_korpus_json(ordner, args.korpus, seeds)

    for name in uebersprungen:
        print(f"uebersprungen (existiert schon): {name}")
    dauer = time.time() - start
    print(f"{len(neu)} Paare erzeugt, {len(uebersprungen)} uebersprungen in {dauer:.1f} s -> {ordner}")
    if neu:
        print(f"Aenderungen: {aenderungen}, davon verschiedene Soll-Stellen (Blatt/Zeile/Spalte): {zellen}")
        print("Stoerer: " + ", ".join(f"{n} {c}" for n, c in sorted(stoerer.items())))
    return 0


if __name__ == "__main__":
    sys.exit(main())
