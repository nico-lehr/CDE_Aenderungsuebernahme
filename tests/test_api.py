"""API-Tests ueber den Flask-Testclient (kein echter Server noetig)."""
import io
import json
from pathlib import Path

import pytest

from src import config, runs
from src.api.app import app, pruefe_bindung
from src.models import SCHEMA_VERSION
from tests.fixtures.szenario import erzeuge


@pytest.fixture
def client():
    app.config["TESTING"] = True
    return app.test_client()


def _hochladen(client, d):
    return client.post("/api/lauf", data={"pdf": (open(d["pdf"], "rb"), "a.pdf"),
                                           "excel": (open(d["alt"], "rb"), "cde.xlsx")},
                       content_type="multipart/form-data")


def test_health_und_port(client):
    j = client.get("/api/health").get_json()
    assert j["ok"] and j["port"] == 5100


def test_ganzer_ablauf(client, tmp_path):
    d = erzeuge(tmp_path)
    r = _hochladen(client, d)
    assert r.status_code == 200
    plan = r.get_json()
    rid = plan["run_id"]
    assert plan["bilanz"]["stimmt"]
    u = client.post(f"/api/lauf/{rid}/uebernehmen", json={"freigaben": []})
    assert u.status_code == 200
    ordner, name = u.get_json()["ordner"], u.get_json()["excel"]
    assert client.get(f"/api/lauf/{rid}/download/{ordner}/{name}").status_code == 200
    # append-only: zweite Uebernahme legt neuen Ordner an
    u2 = client.post(f"/api/lauf/{rid}/uebernehmen", json={"freigaben": []}).get_json()
    assert u2["ordner"] != ordner
    assert client.delete(f"/api/lauf/{rid}").status_code == 200


def test_falsche_dateien(client):
    r = client.post("/api/lauf", data={"pdf": (io.BytesIO(b"x"), "a.txt"), "excel": (io.BytesIO(b"x"), "b.xlsx")},
                    content_type="multipart/form-data")
    assert r.status_code == 400
    r = client.post("/api/lauf", data={"pdf": (io.BytesIO(b"x"), "a.pdf"), "excel": (io.BytesIO(b"x"), "b.xls")},
                    content_type="multipart/form-data")
    assert r.status_code == 400
    r = client.post("/api/lauf", data={"pdf": (io.BytesIO(b"kein pdf"), "a.pdf"), "excel": (io.BytesIO(b"PK.."), "b.xlsx")},
                    content_type="multipart/form-data")
    assert r.status_code == 400


def test_pfad_tricks_werden_abgelehnt(client):
    assert client.get("/api/lauf/../../etc/passwd").status_code in (400, 404)
    assert client.get("/api/lauf/20260101-000000-abcdef/download/eingabe/..%2f..%2fplan.json").status_code in (400, 404)
    assert client.get("/static/../config/paths.yaml").status_code in (400, 404)


# ---------------------------------------------------------------------------------------------
# K-API: Archiv, Fehlerantworten, Host/Origin, Limits, Schema (nur erfundene Daten)
# ---------------------------------------------------------------------------------------------
def _neuer_lauf(client, tmp_path):
    return _hochladen(client, erzeuge(tmp_path / "fx")).get_json()["run_id"]


def test_delete_archiviert_und_loescht_nichts(client, tmp_path):
    rid = _neuer_lauf(client, tmp_path)
    r = client.delete(f"/api/lauf/{rid}")
    assert r.status_code == 200 and r.get_json()["archiviert"] is True
    assert (runs.lauf_dir(rid) / "lauf.json").is_file() and (runs.lauf_dir(rid) / "eingabe").is_dir()
    assert client.get("/api/lauf").get_json() == []
    mit = client.get("/api/lauf?archiviert=1").get_json()
    assert [e["run_id"] for e in mit] == [rid] and mit[0]["archiviert"] is True
    assert client.get(f"/api/lauf/{rid}").status_code == 200              # direkt aufrufbar bleibt er
    w = client.post(f"/api/lauf/{rid}/wiederherstellen")
    assert w.status_code == 200 and w.get_json()["archiviert"] is False
    assert [e["run_id"] for e in client.get("/api/lauf").get_json()] == [rid]


def test_archiv_unbekannter_lauf_gibt_404_json(client):
    for r in (client.delete("/api/lauf/20260101-000000-abcdef"),
              client.post("/api/lauf/20260101-000000-abcdef/wiederherstellen"),
              client.delete("/api/lauf/unsinn")):
        assert r.status_code == 404 and r.is_json and r.get_json()["fehler"]


@pytest.mark.parametrize("methode, url, code", [
    ("get", "/api/gibt-es-nicht", 404), ("get", "/nirgendwo.js", 404), ("post", "/api/health", 405),
    ("put", "/api/lauf", 405)])
def test_http_fehler_sind_immer_json(client, methode, url, code):
    r = getattr(client, methode)(url)
    assert r.status_code == code and r.is_json
    assert set(r.get_json()) == {"fehler"} and "<" not in r.get_json()["fehler"]
    if code == 405:
        assert "Allow" in r.headers


def test_unerwarteter_fehler_zeigt_keinen_rohtext(client, monkeypatch):
    def kaputt(*a, **k):
        raise RuntimeError("GEHEIMER_ROHTEXT C:\\intern\\x.py")
    monkeypatch.setattr(runs, "liste", kaputt)
    r = client.get("/api/lauf")
    assert r.status_code == 500 and r.is_json and "GEHEIMER_ROHTEXT" not in r.get_data(as_text=True)


def test_upload_limit_gesamt_nennt_mb(client, tmp_path, monkeypatch):
    monkeypatch.setitem(app.config, "MAX_CONTENT_LENGTH", 1024)
    r = _hochladen(client, erzeuge(tmp_path / "fx"))
    assert r.status_code == 413 and r.is_json
    assert f"{config.lade('server')['upload_max_mb']} MB" in r.get_json()["fehler"]


def test_upload_limit_je_datei(client, tmp_path, monkeypatch):
    monkeypatch.setitem(config.lade("server"), "upload_max_mb_datei", 0)
    r = _hochladen(client, erzeuge(tmp_path / "fx"))
    assert r.status_code == 413 and r.is_json and "0 MB je Datei" in r.get_json()["fehler"]
    assert runs.liste() == []


def test_config_hat_die_neuen_schluessel():
    s = config.lade("server")
    assert s["upload_max_mb"] == 200 and s["upload_max_mb_datei"] <= s["upload_max_mb"]
    assert s["dateiname_max_laenge"] == 255 and "127.0.0.1" in s["erlaubte_hosts"]
    assert app.config["MAX_CONTENT_LENGTH"] == 200 * 1024 * 1024


def test_fehlgeschlagene_analyse_bleibt_als_lauf_stehen(client, tmp_path):
    d = erzeuge(tmp_path / "fx")
    kaputt = tmp_path / "kaputt.pdf"
    kaputt.write_bytes(b"%PDF-1.4 kein echtes PDF")
    r = client.post("/api/lauf", data={"pdf": (open(kaputt, "rb"), "a.pdf"), "excel": (open(d["alt"], "rb"), "cde.xlsx")},
                    content_type="multipart/form-data")
    assert r.status_code >= 400 and r.is_json
    rid = r.get_json()["run_id"]
    assert r.get_json()["status"] == "fehlgeschlagen"
    eintrag = next(e for e in client.get("/api/lauf").get_json() if e["run_id"] == rid)
    assert eintrag["status"] == "fehlgeschlagen" and eintrag["meldung"]
    g = client.get(f"/api/lauf/{rid}")                                  # keine Neuanalyse, klare Meldung
    assert g.status_code == 409 and "fehlgeschlagen" in g.get_json()["fehler"]
    assert not (runs.lauf_dir(rid) / "plan.json").exists()
    assert (runs.lauf_dir(rid) / "eingabe" / "cde.xlsx").is_file()


def test_kaputte_plan_json(client, tmp_path):
    gut, kaputt = _neuer_lauf(client, tmp_path), _neuer_lauf(client, tmp_path)
    (runs.lauf_dir(kaputt) / "plan.json").write_text('{"schema_version": 2, "bilanz": {"a', encoding="utf-8")
    liste = client.get("/api/lauf")
    assert liste.status_code == 200
    stati = {e["run_id"]: e["status"] for e in liste.get_json()}
    assert stati[gut] == "analysiert" and stati[kaputt] == "beschaedigt"
    r = client.get(f"/api/lauf/{kaputt}")
    assert r.status_code == 422 and r.get_json()["fehler"] == "Der Lauf ist beschädigt."
    u = client.post(f"/api/lauf/{kaputt}/uebernehmen", json={"freigaben": []})
    assert u.status_code == 422 and u.get_json()["fehler"] == "Der Lauf ist beschädigt."
    assert client.get(f"/api/lauf/{kaputt}/vorschau/x.png").status_code == 422


def test_kaputte_lauf_json(client, tmp_path):
    rid = _neuer_lauf(client, tmp_path)
    (runs.lauf_dir(rid) / "lauf.json").write_text("", encoding="utf-8")
    assert client.get("/api/lauf").get_json()[0]["status"] == "beschaedigt"
    assert client.get(f"/api/lauf/{rid}").status_code == 422


def test_schema_aus_der_zukunft_beim_ansehen_und_uebernehmen(client, tmp_path):
    rid = _neuer_lauf(client, tmp_path)
    pj = runs.lauf_dir(rid) / "plan.json"
    j = json.loads(pj.read_text(encoding="utf-8"))
    j["schema_version"] = SCHEMA_VERSION + 1
    pj.write_text(json.dumps(j), encoding="utf-8")
    for r in (client.get(f"/api/lauf/{rid}"), client.post(f"/api/lauf/{rid}/uebernehmen", json={"freigaben": []})):
        assert r.status_code == 409 and "neueren Programmversion" in r.get_json()["fehler"]
    assert not list(runs.lauf_dir(rid).glob("ergebnis_*"))


def test_schema_1_ohne_version_feld_wird_gelesen(client, tmp_path):
    rid = _neuer_lauf(client, tmp_path)
    pj = runs.lauf_dir(rid) / "plan.json"
    j = json.loads(pj.read_text(encoding="utf-8"))
    j.pop("schema_version")
    pj.write_text(json.dumps(j), encoding="utf-8")
    assert client.get(f"/api/lauf/{rid}").status_code == 200


def test_lauf_antwort_enthaelt_ergebnisse_wie_bisher(client, tmp_path):
    rid = _neuer_lauf(client, tmp_path)
    client.post(f"/api/lauf/{rid}/uebernehmen", json={"freigaben": []})
    erg = client.get(f"/api/lauf/{rid}").get_json()["ergebnisse"]
    assert len(erg) == 1 and {"ordner", "bilanz", "excel", "erstellt", "freigaben"} <= set(erg[0])


@pytest.mark.parametrize("host", ["boese.example", "boese.example:5100", "127.0.0.1.boese.example:5100",
                                  "localhost:9999", "10.0.0.5:5100", "[::2]:5100"])
def test_fremder_host_wird_mit_json_abgelehnt(client, host):
    r = client.get("/api/health", base_url=f"http://{host}")
    assert r.status_code == 403 and r.is_json and r.get_json()["fehler"]


@pytest.mark.parametrize("host", ["127.0.0.1:5100", "localhost:5100", "[::1]:5100", "LOCALHOST:5100"])
def test_eigener_host_wird_angenommen(client, host):
    assert client.get("/api/health", base_url=f"http://{host}").status_code == 200


@pytest.mark.parametrize("origin", ["http://boese.example", "http://127.0.0.1:9999", "null", "https://127.0.0.1:5100",
                                    "http://127.0.0.1", "http://127.0.0.1.boese.example:5100", "Unsinn"])
def test_fremder_origin_wird_abgelehnt(client, origin):
    r = client.get("/api/health", headers={"Origin": origin})
    assert r.status_code == 403 and r.is_json


@pytest.mark.parametrize("origin", ["http://127.0.0.1:5100", "http://localhost:5100", "http://[::1]:5100"])
def test_eigener_origin_wird_angenommen(client, origin):
    assert client.get("/api/health", headers={"Origin": origin}).status_code == 200


def test_host_schutz_gilt_auch_fuer_die_startseite(client):
    assert client.get("/", base_url="http://boese.example:5100").status_code == 403


@pytest.mark.parametrize("host", ["127.0.0.1", "localhost", "::1", "[::1]"])
def test_bindung_loopback_erlaubt(host):
    pruefe_bindung(host)


@pytest.mark.parametrize("host", ["0.0.0.0", "::", "192.168.1.5", "", "example.org"])
def test_bindung_nicht_loopback_ist_startfehler(host):
    with pytest.raises(SystemExit) as e:
        pruefe_bindung(host)
    assert "config/server.yaml" in str(e.value)


def test_reservenamen_und_lange_namen_beim_upload(client, tmp_path):
    d = erzeuge(tmp_path / "fx")
    r = client.post("/api/lauf", data={"pdf": (open(d["pdf"], "rb"), "CON.pdf"), "excel": (open(d["alt"], "rb"), "nul.xlsx")},
                    content_type="multipart/form-data")
    assert r.status_code == 200
    info = json.loads((runs.lauf_dir(r.get_json()["run_id"]) / "lauf.json").read_text(encoding="utf-8"))
    assert info["pdf"] == "_CON.pdf" and info["excel"] == "_nul.xlsx"
    lang = client.post("/api/lauf", data={"pdf": (open(d["pdf"], "rb"), "a.pdf"),
                                           "excel": (open(d["alt"], "rb"), "x" * 300 + ".xlsx")},
                       content_type="multipart/form-data")
    assert lang.status_code == 400 and lang.is_json and "zu lang" in lang.get_json()["fehler"]
    assert len(runs.liste()) == 1


def test_upload_liest_nur_die_ersten_bytes(client, tmp_path, monkeypatch):
    def verboten(self, *a, **k):
        raise AssertionError("ganze Datei gelesen")
    d = erzeuge(tmp_path / "fx")
    monkeypatch.setattr(Path, "read_bytes", verboten)
    assert _hochladen(client, d).status_code == 200


def test_uebernehmen_mit_unsinn_im_body(client, tmp_path):
    rid = _neuer_lauf(client, tmp_path)
    r = client.post(f"/api/lauf/{rid}/uebernehmen", json=[1, 2])
    assert r.status_code == 400 and r.is_json
    r = client.post(f"/api/lauf/{rid}/uebernehmen", json={"freigaben": ["gibt-es-nicht"]})
    assert r.status_code in (400, 404) and r.is_json


def test_upload_mit_unbestaetigtem_profil_wird_abgelehnt(client, tmp_path):
    """Neue Profile gehen nie direkt live: ohne bestaetigte Vorschau gibt es eine einfache Meldung (409), der Lauf bleibt stehen."""
    d = erzeuge(tmp_path / "fx")
    with open(d["pdf"], "rb") as p, open(d["alt"], "rb") as x:
        r = client.post("/api/lauf", data={"pdf": (p, "a.pdf"), "excel": (x, "b.xlsx"), "profil": "gibt-es-nicht"},
                        content_type="multipart/form-data")
    assert r.status_code == 409 and r.is_json
    assert r.get_json()["status"] == "fehlgeschlagen" and "fehler" in r.get_json()
