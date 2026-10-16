"""Erzeugt erfundene Testpaare (PDF + vorher.xlsx + nachher.xlsx + erwartet.json) fuer das Messwerkzeug.

Beispiele (im Projektordner, mit der .venv):
    python scripts/erzeuge_paare.py --korpus entwicklung --seeds 1-60
    python scripts/erzeuge_paare.py --korpus entwicklung --seeds 101-160 --erzeuger wechselnd
    python scripts/erzeuge_paare.py --korpus ex_test --seeds 1001-1030
    python scripts/erzeuge_paare.py --korpus entwicklung --seeds 1-5 --ziel D:\\temp\\paare

Wohin:  <paare_dir aus config/paths.yaml>/<korpus>/synth_<seed>_<profil>/
        --ziel DIR ersetzt paare_dir (die Paare landen dann in DIR/<korpus>/).
Korpora (Regel 11, Generator v2):
        entwicklung  Seeds < 1000
        ex_test      nur Seeds 1001-1030 (frueherer Testkorpus, offen als Entwicklung erklaert)
        test         nur Seeds >= 2000 (reserviert fuer die ungesehene Endmessung -- nicht zum Einstellen benutzen)
Erzeuger: writer (eigener PDF-Schreiber, Standard) | libreoffice (Flat-ODT -> soffice) | wechselnd (gerade Seeds
        LibreOffice, ungerade writer). Fehlt soffice, werden LibreOffice-Paare mit Hinweis uebersprungen (kein Absturz).
Regeln: Alles ist erfunden (Regel 16). Bestehende Paar-Ordner werden nie ueberschrieben, sondern
        uebersprungen (Regel 29).
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


def erzeuger_fuer(seed: int, modus: str) -> str:
    if modus == "wechselnd":
        return "libreoffice" if seed % 2 == 0 else "writer"
    return modus


def schreibe_korpus_json(ordner: Path, korpus: str, seeds: list[int], erzeuger: dict[int, str] | None = None) -> None:
    """KORPUS.json: Generator-Version, Seeds, Datum, Erzeuger je Seed und LibreOffice-Version.
    Angaben frueherer Laeufe (gleiche Version) bleiben erhalten."""
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
        alt_erz = {int(k): v for k, v in (alt.get("erzeuger") or {}).items()}
        alt_lo = alt.get("libreoffice_version") or ""
    else:
        alt_erz, alt_lo = {}, ""
    erz = {**alt_erz, **(erzeuger or {})}
    lo_version = alt_lo
    if "libreoffice" in (erzeuger or {}).values():
        from tests.fixtures import pdf_lo
        lo_version = pdf_lo.lo_version() or alt_lo
    alle = sorted(set(alt_seeds) | set(seeds) | set(erz))
    daten = {"generator_version": generator.GENERATOR_VERSION, "korpus": korpus,
             "seeds": alle, "erstellt": erstellt,
             "erzeuger": {str(s): erz.get(s, "writer") for s in alle},
             "je_erzeuger": {e: sum(1 for s in alle if erz.get(s, "writer") == e) for e in generator.ERZEUGER},
             "libreoffice_version": lo_version}
    tmp = ordner / "KORPUS.json.tmp"
    tmp.write_text(json.dumps(daten, ensure_ascii=False) + "\n", encoding="utf-8")
    tmp.replace(pfad)


def pruefe_seeds(korpus: str, seeds: list[int]) -> str | None:
    """Regel 11: Fehlermeldung, wenn Seeds nicht zum Korpus passen, sonst None."""
    if korpus == "entwicklung" and any(s >= generator.SEED_GRENZE for s in seeds):
        return f"Entwicklungs-Seeds muessen unter {generator.SEED_GRENZE} liegen (Regel 11)."
    lo, hi = generator.EX_TEST_SEEDS
    if korpus == "ex_test" and any(not lo <= s <= hi for s in seeds):
        return f"ex_test umfasst nur die Seeds {lo}-{hi} (frueherer Testkorpus, Regel 11)."
    if korpus == "test" and any(s < generator.TEST_SEED_MIN for s in seeds):
        return f"Test-Seeds muessen ab {generator.TEST_SEED_MIN} liegen (reservierter Testkorpus, Regel 11)."
    return None


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Erzeugt erfundene Testpaare mit bekannter Wahrheit.")
    ap.add_argument("--korpus", required=True, choices=list(generator.KORPORA))
    ap.add_argument("--seeds", required=True, help="z. B. 1-60 oder 1001-1030 oder 1,5,9")
    ap.add_argument("--profil", default="gemischt", choices=["gemischt", "standard", "schwer"])
    ap.add_argument("--erzeuger", default="writer", choices=[*generator.ERZEUGER, "wechselnd"],
                    help="PDF-Erzeuger: writer (Standard), libreoffice oder wechselnd (gerade Seeds LibreOffice)")
    ap.add_argument("--ziel", default=None, help="Basisordner statt paare_dir aus config/paths.yaml")
    args = ap.parse_args(argv)

    try:
        seeds = lese_seeds(args.seeds)
    except ValueError as e:
        print(f"Fehler: {e}", file=sys.stderr)
        return 2
    fehler = pruefe_seeds(args.korpus, seeds)
    if fehler:
        print(f"Fehler: {fehler}", file=sys.stderr)
        return 2
    if args.korpus == "test":
        print("Hinweis (Regel 11): Der Korpus 'test' ist fuer die ungesehene Endmessung reserviert.")

    from tests.fixtures import pdf_lo
    lo_da = pdf_lo.verfuegbar()
    ordner = ziel_ordner(args.korpus, args.ziel)
    ordner.mkdir(parents=True, exist_ok=True)
    start = time.time()
    neu: list[int] = []
    uebersprungen: list[str] = []
    ohne_lo: list[int] = []
    vorhanden: dict[int, str] = {}
    aenderungen = 0
    zellen = 0
    stoerer: Counter[str] = Counter()
    je_erzeuger: Counter[str] = Counter()
    for seed in seeds:
        profil = profil_fuer(seed, args.profil)
        name = f"synth_{seed:04d}_{profil}"
        if (ordner / name).exists():
            uebersprungen.append(name)
            try:
                e_alt = json.loads((ordner / name / "erwartet.json").read_text(encoding="utf-8"))
                vorhanden[seed] = str(e_alt.get("erzeuger", "writer"))
            except (OSError, ValueError):
                vorhanden[seed] = "writer"
            continue
        erzeuger = erzeuger_fuer(seed, args.erzeuger)
        if erzeuger == "libreoffice" and not lo_da:
            ohne_lo.append(seed)
            continue
        paar = generator.erzeuge_paar(seed, profil, ordner, args.korpus, erzeuger=erzeuger)
        e = json.loads((paar / "erwartet.json").read_text(encoding="utf-8"))
        aenderungen += len(e["aenderungen"])
        zellen += len({(a["blatt"], a["zeile"], a["spalte"]) for a in e["aenderungen"]})
        stoerer.update(e["stoerer"])
        je_erzeuger[erzeuger] += 1
        neu.append(seed)
        vorhanden[seed] = erzeuger
    schreibe_korpus_json(ordner, args.korpus, sorted(vorhanden), vorhanden)

    for name in uebersprungen:
        print(f"uebersprungen (existiert schon): {name}")
    if ohne_lo:
        print(f"HINWEIS: soffice (LibreOffice) fehlt -- {len(ohne_lo)} LibreOffice-Paar(e) nicht erzeugt (Seeds "
              + ", ".join(map(str, ohne_lo)) + ").")
    dauer = time.time() - start
    print(f"{len(neu)} Paare erzeugt, {len(uebersprungen)} uebersprungen in {dauer:.1f} s -> {ordner}")
    if neu:
        print(f"Aenderungen: {aenderungen}, davon verschiedene Soll-Stellen (Blatt/Zeile/Spalte): {zellen}")
        print("Erzeuger: " + ", ".join(f"{n} {c}" for n, c in sorted(je_erzeuger.items())))
        print("Stoerer: " + ", ".join(f"{n} {c}" for n, c in sorted(stoerer.items())))
    return 0


if __name__ == "__main__":
    sys.exit(main())
