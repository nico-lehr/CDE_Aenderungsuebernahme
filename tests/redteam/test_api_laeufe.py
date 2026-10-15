"""Red-Team K-API (runs / api): Gleichzeitigkeit, Append-only, Fehlermeldungen, Sicherheit."""
import json
import threading
from pathlib import Path

import pytest

from src import config, runs
from src.api.app import app
from tests.fixtures.szenario import erzeuge
from tests.redteam._bau import aufbau, scheitern


@pytest.fixture
def client():
    app.config["TESTING"] = True
    return app.test_client()


@pytest.fixture
def fx(tmp_path):
    return erzeuge(tmp_path / "fx")


def _hochladen(client, d, name_pdf="a.pdf", name_xlsx="cde.xlsx", **kw):
    return client.post("/api/lauf", data={"pdf": (open(d["pdf"], "rb"), name_pdf), "excel": (open(d["alt"], "rb"), name_xlsx)},
                       content_type="multipart/form-data", **kw)


def _lauf_anlegen(d) -> str:
    rid = runs.neuer_lauf(d["pdf"], d["alt"])
    runs.analysiere(rid)
    return rid


def _daten_wurzel() -> Path:
    return config.pfad("runs_dir").parent


# ---------------------------------------------------------------------------------------------
# RT-39 / RT-40: gleichzeitige und abgebrochene Uebernahmen
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
def test_rt39_zwei_gleichzeitige_uebernahmen(fx, monkeypatch):
    rid = _lauf_anlegen(fx)
    schranke = threading.Barrier(2)
    echt = Path.mkdir

    def mkdir(self, *a, **k):
        if self.name.startswith("ergebnis_"):          # beide Threads haben ihre Nummer berechnet, bevor einer anlegt
            try:
                schranke.wait(timeout=2)
            except threading.BrokenBarrierError:
                pass
        return echt(self, *a, **k)
    monkeypatch.setattr(Path, "mkdir", mkdir)
    fehler, ordner = [], []

    def arbeit():
        try:
            ordner.append(runs.uebernehmen_lauf(rid, [])["ordner"])
        except Exception as e:                       # noqa: BLE001
            fehler.append(f"{type(e).__name__}: {e}")
    ts = [threading.Thread(target=arbeit) for _ in range(2)]
    for t in ts:
        t.start()
    for t in ts:
        t.join(30)
    assert not fehler and len(set(ordner)) == 2, f"Fehler: {fehler}; angelegte Ordner: {ordner}"


@pytest.mark.redteam
def test_rt40_abgebrochene_uebernahme_hinterlaesst_keinen_halben_ordner(fx, monkeypatch):
    rid = _lauf_anlegen(fx)

    def kaputt(*a, **k):
        raise RuntimeError("simulierter Abbruch mitten im Uebernehmen")
    monkeypatch.setattr(runs, "uebernehmen", kaputt)
    with pytest.raises(RuntimeError):
        runs.uebernehmen_lauf(rid, [])
    halb = [p.name for p in runs.lauf_dir(rid).glob("ergebnis_*") if not (p / "ergebnis.json").exists()]
    eintrag = next(e for e in runs.liste() if e["run_id"] == rid)
    assert not halb and eintrag["ergebnisse"] == 0, \
        f"Halbfertige Ordner {halb}; Lauf-Liste zaehlt {eintrag['ergebnisse']} Ergebnis(se), obwohl keines fertig ist"


# ---------------------------------------------------------------------------------------------
# RT-41: plan.json atomar schreiben
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
def test_rt41_plan_json_nicht_atomar(fx, monkeypatch):
    rid = _lauf_anlegen(fx)                              # plan.json existiert schon (vollstaendig)
    halb, weiter = threading.Event(), threading.Event()
    echt = Path.write_text

    def write_text(self, daten, *a, **k):
        if self.name == "plan.json":                     # Schreiber haengt mitten in der Datei
            with open(self, "w", encoding="utf-8") as f:
                f.write(daten[: len(daten) // 2])
                f.flush()
                halb.set()
                weiter.wait(10)
                f.write(daten[len(daten) // 2:])
            return len(daten)
        return echt(self, daten, *a, **k)
    monkeypatch.setattr(Path, "write_text", write_text)
    schreiber = threading.Thread(target=runs.analysiere, args=(rid,))
    schreiber.start()
    problem = None
    try:
        if halb.wait(5):                                 # nur lesen, wenn der Schreiber wirklich mitten drin ist
            try:
                runs.lade_plan(rid)
            except Exception as e:                       # noqa: BLE001
                problem = f"{type(e).__name__}: {e}"
    finally:
        weiter.set()
        schreiber.join(30)
    assert problem is None, f"Leser sieht halb geschriebene plan.json -> {problem}"


# ---------------------------------------------------------------------------------------------
# RT-42 / RT-43: Regel 29 (append-only) -- Laeufe duerfen nicht verschwinden
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
def test_rt42_delete_loescht_lauf_nicht_von_der_platte(client, fx):
    rid = _hochladen(client, fx).get_json()["run_id"]
    antwort = client.delete(f"/api/lauf/{rid}")
    uebrig = list(_daten_wurzel().rglob("lauf.json"))
    assert uebrig, f"Nach DELETE (Status {antwort.status_code}) ist nichts mehr auf der Platte (kein Archiv)"


@pytest.mark.redteam
def test_rt43_fehlgeschlagene_analyse_loescht_sich_nicht(client, fx, tmp_path):
    kaputt = tmp_path / "kaputt.pdf"
    kaputt.write_bytes(b"%PDF-1.4 sieht aus wie ein PDF, ist aber keins")
    r = client.post("/api/lauf", data={"pdf": (open(kaputt, "rb"), "a.pdf"), "excel": (open(fx["alt"], "rb"), "cde.xlsx")},
                    content_type="multipart/form-data")
    aufbau(r.status_code >= 400, f"Analyse haette fehlschlagen muessen, Status {r.status_code}")
    uebrig = list(_daten_wurzel().rglob("lauf.json"))
    assert uebrig, f"Nach der fehlgeschlagenen Analyse (Status {r.status_code}) sind die hochgeladenen Dateien weg"


# ---------------------------------------------------------------------------------------------
# RT-44..46: Fehlermeldungen in einfachem Deutsch (JSON), ohne Rohtext von Ausnahmen
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
def test_rt44_zu_grosser_upload_liefert_json(client, fx, monkeypatch):
    monkeypatch.setitem(app.config, "MAX_CONTENT_LENGTH", 1024)
    r = client.post("/api/lauf", data={"pdf": (open(fx["pdf"], "rb"), "a.pdf"), "excel": (open(fx["alt"], "rb"), "cde.xlsx")},
                    content_type="multipart/form-data")
    aufbau(r.status_code == 413, f"Limit sollte greifen (413), Status {r.status_code}")
    assert r.is_json and "fehler" in r.get_json(), f"Antwort ist {r.content_type}: {r.data[:80]!r}"


@pytest.mark.redteam
def test_rt45_analysefehler_zeigt_keinen_rohtext(client, fx, monkeypatch):
    def kaputt(run_id):
        raise RuntimeError("ROHTEXT_XYZ No /Root object! C:\\intern\\pfad\\datei.py line 12")
    monkeypatch.setattr(runs, "analysiere", kaputt)
    r = _hochladen(client, fx)
    aufbau(r.status_code >= 400 and r.is_json, f"Fehlerantwort erwartet, Status {r.status_code}")
    assert "ROHTEXT_XYZ" not in r.get_json()["fehler"], f"Meldung enthaelt Rohtext: {r.get_json()['fehler']!r}"


@pytest.mark.redteam
def test_rt46_uebernahmefehler_zeigt_keinen_rohtext(client, fx, monkeypatch):
    rid = _hochladen(client, fx).get_json()["run_id"]

    def kaputt(*a, **k):
        raise RuntimeError("ROHTEXT_XYZ unexpected keyword argument 'x' in C:\\intern\\pfad\\datei.py")
    monkeypatch.setattr(runs, "uebernehmen", kaputt)
    r = client.post(f"/api/lauf/{rid}/uebernehmen", json={"freigaben": []})
    aufbau(r.status_code >= 400 and r.is_json, f"Fehlerantwort erwartet, Status {r.status_code}")
    assert "ROHTEXT_XYZ" not in r.get_json()["fehler"], f"Meldung enthaelt Rohtext: {r.get_json()['fehler']!r}"


# ---------------------------------------------------------------------------------------------
# RT-47: plan.json aus der Zukunft
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
def test_rt47_plan_json_aus_der_zukunft_gibt_klare_meldung(client, fx):
    rid = _hochladen(client, fx).get_json()["run_id"]
    pj = runs.lauf_dir(rid) / "plan.json"
    j = json.loads(pj.read_text(encoding="utf-8"))
    j["schema_version"] = 99
    for a in j["aenderungen"]:
        a["neues_feld_aus_der_zukunft"] = 1
    pj.write_text(json.dumps(j), encoding="utf-8")
    r = client.post(f"/api/lauf/{rid}/uebernehmen", json={"freigaben": []})
    text = (r.get_json() or {}).get("fehler", "")
    klar = r.status_code in (400, 409, 422) and ("Version" in text or "Schema" in text) and "keyword" not in text
    assert klar, f"Status {r.status_code}, Meldung: {text!r}"


# ---------------------------------------------------------------------------------------------
# RT-48 / RT-49: Host-/Origin-Pruefung (Angriff von fremder Webseite bzw. DNS-Rebinding)
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
def test_rt48_fremder_origin_wird_abgelehnt(client, fx):
    rid = _hochladen(client, fx).get_json()["run_id"]
    r = client.post(f"/api/lauf/{rid}/uebernehmen", json={"freigaben": []}, headers={"Origin": "http://boese.example"})
    assert r.status_code in (400, 403), f"Anfrage von Origin http://boese.example wurde angenommen (Status {r.status_code})"


@pytest.mark.redteam
def test_rt49_fremder_host_wird_abgelehnt(client, fx):
    r = client.get("/api/lauf", base_url="http://boese.example:5100")
    assert r.status_code in (400, 403, 421), f"Anfrage mit Host boese.example wurde angenommen (Status {r.status_code})"


# ---------------------------------------------------------------------------------------------
# RT-52: ein kaputter Lauf darf die Lauf-Liste nicht unbrauchbar machen
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
def test_rt52_ein_kaputter_lauf_zerstoert_nicht_die_liste(client, fx, monkeypatch):
    monkeypatch.setitem(app.config, "PROPAGATE_EXCEPTIONS", False)
    gut, kaputt = _lauf_anlegen(fx), _lauf_anlegen(fx)
    (runs.lauf_dir(kaputt) / "plan.json").write_text('{"schema_version": 1, "bilanz": {"aend', encoding="utf-8")
    r = client.get("/api/lauf")
    ids = [e["run_id"] for e in r.get_json()] if r.is_json else []
    assert r.status_code == 200 and gut in ids, f"Lauf-Liste: Status {r.status_code} ({r.content_type}), sichtbare Laeufe: {ids}"


# ---------------------------------------------------------------------------------------------
# RT-50 / RT-51: Dateinamen beim Upload
# ---------------------------------------------------------------------------------------------
@pytest.mark.redteam
@pytest.mark.parametrize("pdf_name, xlsx_name", [("CON.pdf", "NUL.xlsx"), ("com1.pdf", "Aux.xlsx"), ("LPT1.PDF", "prn.xlsx")])
def test_rt50_windows_reservenamen(fx, pdf_name, xlsx_name):
    rid = runs.neuer_lauf(fx["pdf"], fx["alt"], pdf_name, xlsx_name)
    info = json.loads((runs.lauf_dir(rid) / "lauf.json").read_text(encoding="utf-8"))
    reserviert = {"CON", "PRN", "AUX", "NUL", *{f"COM{i}" for i in range(1, 10)}, *{f"LPT{i}" for i in range(1, 10)}}
    schlecht = [n for n in (info["pdf"], info["excel"]) if Path(n).stem.upper() in reserviert]
    assert not schlecht, f"Gespeicherte Dateinamen sind unter Windows nicht anlegbar: {schlecht}"


@pytest.mark.redteam
def test_rt51_sehr_langer_dateiname(client, fx, monkeypatch):
    monkeypatch.setitem(app.config, "PROPAGATE_EXCEPTIONS", False)      # wie im echten Betrieb: Fehlerseite statt Traceback
    r = _hochladen(client, fx, name_pdf="a" * 300 + ".pdf")
    halb = [p.name for p in runs.runs_dir().iterdir() if p.is_dir() and not (p / "plan.json").exists()]
    assert r.status_code == 400 and r.is_json and not halb, \
        f"Status {r.status_code} ({r.content_type}); Lauf-Ordner ohne Analyse: {halb}"
