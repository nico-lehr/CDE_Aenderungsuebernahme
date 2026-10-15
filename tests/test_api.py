"""API-Tests ueber den Flask-Testclient (kein echter Server noetig)."""
import io

import pytest

from src.api.app import app
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
