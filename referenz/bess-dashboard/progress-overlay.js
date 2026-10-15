/**
 * progress-overlay.js
 *
 * Blockierendes Vollbild-Overlay waehrend eines Pipeline-Laufs.
 * Zeigt EINEN Gesamtbalken aus dem kombinierten Fortschritt beider Tracks
 * (siehe extraction-state.js) und laesst sich zu einer Pille unten rechts
 * minimieren, damit man waehrend des Laufs in andere Tabs schauen kann.
 *
 * Es gibt bewusst keinen Abbrechen-Knopf: das Backend kennt keinen
 * Cancel-Endpunkt, die Threads sind nicht unterbrechbar. Esc minimiert
 * deshalb, statt zu schliessen -- alles andere waere eine Luege ueber den
 * tatsaechlichen Prozesszustand.
 */

(() => {
  "use strict";

  const el = {};
  let minimized = false;
  let open = false;
  let errored = false;
  let lastFocused = null;
  // Memo der zuletzt GESCHRIEBENEN Werte. Das Polling feuert alle 500ms, der
  // gerundete Prozentwert aendert sich bei einem langen Lauf aber hoechstens
  // 101 Mal -- ohne Guard schrieb paint() bei jedem Tick blind in DOM und
  // Style. lastPercent startet auf -1, damit das paint(0, …) aus show()
  // garantiert durchgeht und der Balken nicht auf dem Wert des Vorlaufs steht.
  let lastPercent = -1;
  let lastMessage = null;
  let lastDetail = null;
  // Der conic-gradient des Pillen-Rings ist ein Repaint hinter einem
  // backdrop-filter. Solange das Overlay offen ist, liegt die Pille unsichtbar
  // darunter -- dann nur vormerken und beim Minimieren einmal nachziehen.
  let dockStale = false;
  // Track + task_id des zuletzt fehlgeschlagenen Laufs, fuer den Klick auf
  // "Lauf fortsetzen" -- resumeTarget bleibt null, solange kein Fehler
  // resumable ist (siehe extraction-state.js failTrack()).
  let resumeTarget = null;

  function cache() {
    el.overlay = document.getElementById("progress-overlay");
    el.title = document.getElementById("progress-overlay-title");
    el.percent = document.getElementById("progress-overlay-percent");
    el.fill = document.getElementById("progress-overlay-fill");
    el.message = document.getElementById("progress-overlay-message");
    el.detail = document.getElementById("progress-overlay-detail");
    el.problems = document.getElementById("progress-overlay-problems");
    el.minimize = document.getElementById("progress-overlay-minimize");
    el.resume = document.getElementById("progress-overlay-resume");
    el.close = document.getElementById("progress-overlay-close");
    el.dock = document.getElementById("progress-dock");
    el.dockRing = document.getElementById("progress-dock-ring");
    el.dockLabel = document.getElementById("progress-dock-label");
    el.dockPercent = document.getElementById("progress-dock-percent");
  }

  // Befunde der Vorab-Pruefung als Liste. textContent statt innerHTML: die Texte
  // kommen vom Server und enthalten Datei- und Ordnernamen.
  function renderProblems(problems) {
    if (!el.problems) return;
    el.problems.textContent = "";
    const list = Array.isArray(problems) ? problems : [];
    el.problems.hidden = list.length === 0;
    list.forEach((problem) => {
      const item = document.createElement("li");
      const what = document.createElement("strong");
      what.textContent = problem.was;
      item.appendChild(what);
      [problem.warum, problem.tun && ("Was du tun kannst: " + problem.tun)].forEach((text) => {
        if (!text) return;
        const line = document.createElement("span");
        line.textContent = text;
        item.appendChild(line);
      });
      el.problems.appendChild(item);
    });
  }

  // blockPage(true) friert alles ausserhalb des Overlays ein (siehe
  // .is-run-blocked in style.css): hinter dem Vollbild-Scrim kostet jeder
  // Repaint zusaetzlich eine Neuberechnung von dessen backdrop-filter.
  function lockScroll(lock) {
    document.body.style.overflow = lock ? "hidden" : "";
    document.body.classList.toggle("is-run-blocked", lock);
  }

  function resetPaintCache() {
    lastPercent = -1;
    lastMessage = null;
    lastDetail = null;
    dockStale = false;
  }

  function writeDock(percent) {
    el.dockPercent.textContent = `${percent} %`;
    el.dockRing.style.setProperty("--dock-progress", String(percent));
    dockStale = false;
  }

  function paint(percent, message, detail) {
    const nextDetail = detail || "";
    // Ein Update ohne message laesst die bisherige stehen (bisheriges Verhalten).
    const nextMessage = message || lastMessage;

    if (percent === lastPercent && nextMessage === lastMessage && nextDetail === lastDetail) return;

    if (percent !== lastPercent) {
      el.percent.textContent = `${percent} %`;
      el.percent.setAttribute("aria-valuenow", String(percent));
      el.fill.style.setProperty("--progress", String(percent / 100));
      if (minimized) writeDock(percent); else dockStale = true;
      lastPercent = percent;
    }
    if (nextMessage !== lastMessage) {
      el.message.textContent = nextMessage;
      lastMessage = nextMessage;
    }
    if (nextDetail !== lastDetail) {
      el.detail.textContent = nextDetail;
      lastDetail = nextDetail;
    }
  }

  function showOverlay() {
    el.dock.classList.remove("visible");
    el.dock.hidden = true;
    el.overlay.hidden = false;
    minimized = false;
    lockScroll(true);
    (errored ? el.close : el.minimize).focus();
  }

  const API = {
    show(title) {
      cache();
      errored = false;
      open = true;
      resumeTarget = null;
      el.overlay.classList.remove("is-error");
      el.close.hidden = true;
      el.resume.hidden = true;
      el.minimize.hidden = false;
      el.title.textContent = title || "Prüfung läuft…";
      el.dockLabel.textContent = "Prüfung läuft…";
      resetPaintCache();
      paint(0, "Dokumente werden vorbereitet…", "");
      lastFocused = document.activeElement;
      showOverlay();
    },

    update({ percent, message, detail }) {
      if (!open || errored) return;
      paint(
        Math.max(0, Math.min(100, Math.round(percent))),
        message,
        detail
      );
    },

    fail(errorMessage, options = {}) {
      if (!open) API.show("Prüfung fehlgeschlagen");
      errored = true;
      el.overlay.classList.add("is-error");
      el.title.textContent = "Prüfung fehlgeschlagen";
      el.message.textContent = errorMessage || "Unbekannter Fehler";
      el.detail.textContent = "";
      renderProblems(options.problems);
      // Direkte Writes am paint()-Guard vorbei: Memo nachziehen, sonst liesse
      // ein spaeteres paint() mit derselben Message den Fehlertext stehen.
      lastMessage = errorMessage || "Unbekannter Fehler";
      lastDetail = "";
      el.minimize.hidden = true;
      el.close.hidden = false;
      // Nur ein durch Serverneustart/Prozessende unterbrochener Lauf mit
      // gespeicherten Startparametern ist fortsetzbar (siehe
      // POST /api/extraction/<task_id>/resume) -- ein inhaltlicher Fehler
      // (z.B. "Kein einziger Extraktionsaufruf war erfolgreich") ist es
      // nicht, dort bringt ein Neustart mit denselben Daten nichts.
      if (options.resumable && options.taskId) {
        resumeTarget = { type: options.track, taskId: options.taskId };
        el.resume.hidden = false;
      } else {
        resumeTarget = null;
        el.resume.hidden = true;
      }
      // Der Fehler muss gesehen werden -- aus der Pille zurueck ins Overlay.
      if (minimized) showOverlay();
      (resumeTarget ? el.resume : el.close).focus();
    },

    minimize() {
      if (!open || minimized) return;
      minimized = true;
      el.overlay.hidden = true;
      el.dock.hidden = false;
      // Der Ring wurde waehrend des Overlays nur vorgemerkt, jetzt wird er
      // sichtbar -- einmal nachziehen.
      if (dockStale) writeDock(lastPercent);
      lockScroll(false);
      // Ohne den Frame dazwischen springt die Pille ohne Transition herein.
      requestAnimationFrame(() => el.dock.classList.add("visible"));
    },

    restore() {
      if (!open || !minimized) return;
      showOverlay();
    },

    hide() {
      open = false;
      minimized = false;
      errored = false;
      resumeTarget = null;
      el.overlay.hidden = true;
      el.overlay.classList.remove("is-error");
      renderProblems(null);
      el.dock.classList.remove("visible");
      lockScroll(false);
      resetPaintCache();
      setTimeout(() => { el.dock.hidden = true; }, 300);
      if (lastFocused && document.contains(lastFocused)) lastFocused.focus();
    },

    isOpen: () => open,
    isMinimized: () => minimized
  };

  window.ProgressOverlay = API;

  document.addEventListener("DOMContentLoaded", () => {
    cache();

    el.minimize.addEventListener("click", () => API.minimize());
    el.dock.addEventListener("click", () => API.restore());
    el.close.addEventListener("click", () => API.hide());
    el.resume.addEventListener("click", () => {
      if (!resumeTarget) return;
      const { type, taskId } = resumeTarget;
      resumeTarget = null;
      el.resume.hidden = true;
      window.resumeExtraction(type, taskId);
    });

    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape" || !open || minimized) return;
      e.preventDefault();
      // Fehlerzustand ist quittierbar, ein laufender Prozess nur minimierbar.
      if (errored) API.hide(); else API.minimize();
    });

    window.addEventListener("extraction-progress", (e) => {
      API.update(e.detail);
    });

    window.addEventListener("extraction-failed", (e) => {
      const track = e.detail.track === "cde" ? "CDE-Mapping" : "Wallet-Extraktion";
      API.fail(`${track}: ${e.detail.error}`, {
        resumable: e.detail.resumable,
        taskId: e.detail.taskId,
        track: e.detail.track
      });
    });
  });
})();
