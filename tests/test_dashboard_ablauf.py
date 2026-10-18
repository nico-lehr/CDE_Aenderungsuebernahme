"""Ablauf des Dashboards mit erfundenen Daten (Regel 16).

1. API: Upload eines erfundenen Paars -> Plan mit Schema-4-Feldern -> Freigabe + Abwahl -> Ergebnis enthält die abgewählte Änderung
   NICHT, die freigegebene schon -> Download von Excel und Protokoll; die Original-Datei bleibt bytegleich.
2. Kommandozeile: `lauf --abwahl ID` reicht die Abwahl durch.
3. Browser (Chromium headless über Playwright für Node, nur wenn vorhanden, sonst übersprungen): Seite lädt ohne Konsolenfehler,
   Karte per Taste „A“ freigeben, „Excel erstellen“ -> Schritt 3.
"""
from __future__ import annotations

import hashlib
import io
import json
import os
import shutil
import subprocess
import threading
from pathlib import Path

import pytest
from openpyxl import load_workbook

from src import cli, config, runs
from src.api.app import app
from tests.fixtures.szenario import erzeuge


def _sha(p: Path) -> str:
    return hashlib.sha256(Path(p).read_bytes()).hexdigest()


@pytest.fixture
def paar(tmp_path):
    """Erfundenes Paar mit genau einem Vorschlag ("Rack{+s+}", teilweise hinterlegt) und sonst sicheren Änderungen."""
    return erzeuge(tmp_path / "fx", "hintergrund")


def _cde_text(daten: bytes) -> str:
    ws = load_workbook(io.BytesIO(daten))["CDE"]
    return "\n".join(str(c.value) for row in ws.iter_rows() for c in row if c.value is not None)


# ---------------------------------------------------------------------------------------------------------------- 1. API
@pytest.mark.parametrize("markierung", ["aus", pytest.param("an", marks=pytest.mark.markierung_an)])
def test_ablauf_freigabe_und_abwahl_bis_zum_download(paar, markierung):
    original_vorher = _sha(paar["alt"])
    client = app.test_client()
    with open(paar["pdf"], "rb") as p, open(paar["alt"], "rb") as x:
        r = client.post("/api/lauf", data={"pdf": (p, "aenderungen.pdf"), "excel": (x, "cde.xlsx")}, content_type="multipart/form-data")
    assert r.status_code == 200, r.get_json()
    plan = r.get_json()
    rid = plan["run_id"]

    # Plan fürs Dashboard: Schema 4 (bereich, stuecke, grund_einfach) ist da, die Einordnung kommt vom Server
    for a in plan["aenderungen"]:
        assert isinstance(a["bereich"], str) and isinstance(a["stuecke"], list) and isinstance(a["grund_einfach"], str)
    z = plan["zusammenfassung"]
    assert z["gesamt"] == len(plan["aenderungen"]) and z["pruefen"] == 1 and z["automatisch"] == z["gesamt"] - 1
    (vorschlag,) = [a for a in plan["aenderungen"] if a["kategorie"] == "pruefen"]
    (auto,) = [a for a in plan["aenderungen"] if a["kategorie"] == "automatisch" and (a["alt"], a["neu"]) == ("zwei", "drei")]

    u = client.post(f"/api/lauf/{rid}/uebernehmen", json={"freigaben": [vorschlag["id"]], "abwahl": [auto["id"]]})
    assert u.status_code == 200, u.get_json()
    res = u.get_json()
    nach_id = {a["id"]: a for a in res["aenderungen"]}
    assert nach_id[vorschlag["id"]]["status"] == "uebernommen" and nach_id[vorschlag["id"]]["bestaetigt"]
    assert nach_id[auto["id"]]["status"] == "nicht_uebernommen" and nach_id[auto["id"]]["kategorie"] == "nicht"
    assert res["bilanz"]["stimmt"] and res["bilanz"]["uebernommen"] == z["automatisch"]       # -1 abgewählt, +1 freigegeben

    # Download: die Excel enthält die freigegebene Änderung, die abgewählte NICHT
    excel = client.get(f"/api/lauf/{rid}/download/{res['ordner']}/{res['excel']}")
    assert excel.status_code == 200
    text = _cde_text(excel.data)
    assert "Die Racks werden" in text
    assert "dauert zwei Wochen" in text and "drei Wochen" not in text
    proto = client.get(f"/api/lauf/{rid}/download/{res['ordner']}/aenderungsprotokoll.xlsx")
    assert proto.status_code == 200
    zeilen = list(load_workbook(io.BytesIO(proto.data))["Änderungen"].iter_rows(values_only=True))
    (zeile,) = [z_ for z_ in zeilen if z_[0] == auto["id"]]
    assert zeile[4] == "nicht übernommen" and "abgewählt" in str(zeile[10])                # nichts geht still verloren

    # Ergebnisliste nennt Freigabe und Abwahl; die Original-Datei (und ihre Kopie im Lauf) bleibt bytegleich
    erg = client.get(f"/api/lauf/{rid}").get_json()["ergebnisse"]
    assert erg[-1]["freigaben"] == [vorschlag["id"]] and erg[-1]["abwahl"] == [auto["id"]]
    assert _sha(paar["alt"]) == original_vorher == _sha(runs.lauf_dir(rid) / "eingabe" / "cde.xlsx") == plan["excel_sha256"]


# ---------------------------------------------------------------------------------------------------------------- 2. CLI
def test_cli_lauf_mit_abwahl(paar, capsys):
    def neuer() -> str:                                 # der eben angelegte Lauf (Laeufe derselben Sekunde sortieren zufaellig)
        (rid,) = {e["run_id"] for e in runs.liste()} - bekannt
        bekannt.add(rid)
        return rid
    bekannt: set[str] = set()
    assert cli.main(["analyse", "--pdf", str(paar["pdf"]), "--excel", str(paar["alt"])]) == 0
    plan = runs.lade_plan(neuer())
    (auto,) = [a["id"] for a in plan["aenderungen"] if (a["alt"], a["neu"]) == ("zwei", "drei")]
    capsys.readouterr()
    assert cli.main(["lauf", "--pdf", str(paar["pdf"]), "--excel", str(paar["alt"]), "--abwahl", f" {auto} ,"]) == 0
    out = capsys.readouterr().out
    assert f"Abgewählt (nicht eingetragen): {auto}" in out
    assert runs.ergebnisse(neuer())[0]["abwahl"] == [auto]
    # unbekannte ID: klare Meldung, kein Ergebnis
    assert cli.main(["lauf", "--pdf", str(paar["pdf"]), "--excel", str(paar["alt"]), "--abwahl", "A9999"]) == 2
    assert "Unbekannte" in capsys.readouterr().out
    assert runs.ergebnisse(neuer()) == []


# ---------------------------------------------------------------------------------------------------------------- 3. Browser
_NODE_SKRIPT = r"""
const { chromium } = require('playwright');
const [basis, pdf, xlsx] = process.argv.slice(2);
(async () => {
  let b;
  try { b = await chromium.launch(); } catch (e) { console.log(JSON.stringify({ kein_browser: String(e).slice(0, 200) })); return; }
  const p = await b.newPage({ viewport: { width: 1280, height: 900 } });
  const fehler = [];
  p.on('console', (m) => { if (m.type() === 'error') fehler.push(m.text()); });
  p.on('pageerror', (e) => fehler.push(String(e)));
  const aus = { fehler };
  await p.goto(basis + '/');
  await p.waitForFunction(() => document.getElementById('pill-text').textContent.includes('Port'));
  aus.schritt1 = await p.isVisible('#schritt-dateien');
  aus.aktuell1 = await p.getAttribute('#schritt-knopf-dateien', 'aria-current');
  await p.setInputFiles('#file-pdf', pdf);
  await p.setInputFiles('#file-xlsx', xlsx);
  await p.click('#btn-abgleichen');
  await p.waitForSelector('#schritt-pruefen:not([hidden]) #karten-bitte .karte', { timeout: 60000 });
  aus.karten = await p.locator('#karten-bitte .karte').count();
  aus.satz_vorher = await p.textContent('#fortschritt-satz');
  await p.keyboard.press('?');
  aus.hilfe_offen = await p.evaluate(() => document.getElementById('hilfe').open);
  await p.keyboard.press('Escape');
  aus.hilfe_zu = await p.evaluate(() => !document.getElementById('hilfe').open);
  await p.keyboard.press('j');                                  // erste Karte wählen
  aus.fokus_karte = await p.evaluate(() => document.activeElement.classList.contains('karte'));
  await p.keyboard.press('a');                                  // übernehmen
  aus.gedrueckt = await p.getAttribute('#karten-bitte .karte .entscheid.ja', 'aria-pressed');
  aus.entscheid = await p.getAttribute('#karten-bitte .karte', 'data-entscheid');
  aus.satz_nachher = await p.textContent('#fortschritt-satz');
  aus.balken = await p.getAttribute('#fortschritt', 'aria-valuenow');
  await p.click('#btn-excel');
  await p.waitForSelector('#schritt-fertig:not([hidden])', { timeout: 60000 });
  aus.fertig_titel = await p.textContent('#titel-fertig');
  aus.fertig_satz = await p.textContent('#fertig-satz');
  aus.aktuell3 = await p.getAttribute('#schritt-knopf-fertig', 'aria-current');
  aus.links = await p.$$eval('#fertig-knoepfe a', (l) => l.map((a) => a.getAttribute('href')));
  aus.fertig_fokus = await p.evaluate(() => document.activeElement.id);
  await b.close();
  console.log(JSON.stringify(aus));
})().catch((e) => { console.log(JSON.stringify({ absturz: String(e) })); process.exit(1); });
"""


def _node() -> str | None:
    return shutil.which("node") or (str(p) if (p := Path("/opt/node22/bin/node")).exists() else None)


def _node_umgebung() -> dict[str, str]:
    env = dict(os.environ)
    pfade = [x for x in env.get("NODE_PATH", "").split(os.pathsep) if x]
    knoten = _node()
    if knoten:                                                  # globale Module neben node (z. B. /opt/node22/lib/node_modules)
        pfade.append(str(Path(knoten).resolve().parent.parent / "lib" / "node_modules"))
    env["NODE_PATH"] = os.pathsep.join(pfade)
    return env


def _playwright_da() -> bool:
    if not _node():
        return False
    r = subprocess.run([_node(), "-e", "require.resolve('playwright')"], capture_output=True, env=_node_umgebung(), timeout=30)
    return r.returncode == 0


@pytest.mark.skipif(not _playwright_da(), reason="node mit Playwright nicht vorhanden")
def test_browser_karte_per_taste_freigeben_und_excel_erstellen(paar, tmp_path, monkeypatch):
    from werkzeug.serving import make_server

    server = make_server("127.0.0.1", 0, app, threaded=True)
    monkeypatch.setitem(config.lade("server"), "port", server.server_port)     # Origin-Schutz: der Test-Port ist der eigene
    faden = threading.Thread(target=server.serve_forever, daemon=True)
    faden.start()
    skript = tmp_path / "ablauf.js"
    skript.write_text(_NODE_SKRIPT, encoding="utf-8")
    try:
        r = subprocess.run([_node(), str(skript), f"http://127.0.0.1:{server.server_port}", str(paar["pdf"]), str(paar["alt"])],
                           capture_output=True, text=True, env=_node_umgebung(), timeout=180)
    finally:
        server.shutdown()
        faden.join(10)
    zeilen = [z for z in r.stdout.splitlines() if z.startswith("{")]
    assert zeilen, f"keine Ausgabe vom Browser-Skript: {r.stdout[-500:]} {r.stderr[-500:]}"
    aus = json.loads(zeilen[-1])
    if "kein_browser" in aus:
        pytest.skip(f"Chromium nicht startbar: {aus['kein_browser']}")
    assert "absturz" not in aus, aus
    assert aus["fehler"] == [], f"Konsolenfehler: {aus['fehler']}"
    assert aus["schritt1"] and aus["aktuell1"] == "step"
    assert aus["karten"] == 1 and "1 bitte prüfen (0 erledigt)" in aus["satz_vorher"]
    assert aus["hilfe_offen"] and aus["hilfe_zu"]
    assert aus["fokus_karte"] and aus["gedrueckt"] == "true" and aus["entscheid"] == "ja"
    assert "(1 erledigt)" in aus["satz_nachher"] and aus["balken"] == "1"
    assert aus["aktuell3"] == "step" and aus["fertig_titel"] == "Fertig." and "Änderungen eingetragen" in aus["fertig_satz"]
    assert aus["fertig_fokus"] == "titel-fertig"
    assert len(aus["links"]) == 2 and all(l.startswith("/api/lauf/") for l in aus["links"])
    # Auf dem Server: die per Taste freigegebene Änderung ist als Freigabe gespeichert
    (lauf,) = runs.liste()
    plan = runs.lade_plan(lauf["run_id"])
    (vorschlag,) = [a["id"] for a in plan["aenderungen"] if a["status"] == "pruefen"]
    assert runs.ergebnisse(lauf["run_id"])[0]["freigaben"] == [vorschlag]
