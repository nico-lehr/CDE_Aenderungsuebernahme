/* CDE-Änderungsübernahme – Dashboard (reines JavaScript, keine Abhängigkeiten, keine externen Quellen).
   Drei Schritte statt vieler Reiter: 1 Dateien · 2 Prüfen · 3 Fertig. Vorbild: „Überarbeiten“ in Word und Review-Werkzeuge
   für Anforderungen: eine Liste, gleiche Farben wie im Ergebnis, eine Entscheidung je Punkt, Tastatur, Technik versteckt.
   Vertrauensmodus: Dokument-Ampel und Stichprobe („Passt“ / „Falsch“) kommen vom Server; das Dashboard zeigt sie und meldet nur die Klicks.
   Regel 8: reine Darstellung. Das Dashboard zeigt, was die API liefert (Kategorie und Zusammenfassung kommen vom Server);
   es rechnet nichts Fachliches nach. Alle Texte aus Daten werden über textContent gesetzt (nie als HTML).
   Alle Aufrufe gehen an relative /api/-Pfade (keine festen Adressen oder Ports). */
(() => {
  'use strict';

  // ------------------------------------------------------------------ Texte (einfaches Deutsch)
  const INFO = {
    pruefen: { titel: 'Was heißt „Bitte prüfen“?', text: 'Hier ist sich das Tool nicht sicher, zum Beispiel weil die Excel an der Stelle etwas anders aussieht als das PDF. Ohne dein „Übernehmen“ wird davon nichts eingetragen. Die Vorschau zeigt, wie die Zelle danach in der Excel aussieht: hellblau durchgestrichen fällt weg, rot kommt dazu.' },
    stichprobe: { titel: 'Was ist die Stichprobe?', text: 'Aus den Änderungen, die das Tool selbst einträgt, wird eine kleine, feste Auswahl zum Ansehen vorgelegt (mindestens 3, sonst etwa jede zehnte). „Passt“ zählt als Bestätigung, „Falsch“ trägt die Änderung nicht ein und wird vermerkt. So lernst du Schritt für Schritt, wie weit du dem Tool vertrauen kannst. Gespeichert werden nur Kennungen, nie Texte aus deinen Dokumenten.' },
    archiv: { titel: 'Archivieren', text: 'Archivierte Abgleiche bleiben vollständig auf diesem Rechner erhalten (Dateien, Prüfung, erstellte Excel-Dateien). Sie stehen nur nicht mehr unter „Zuletzt“. Mit „Archivierte anzeigen“ findest du sie wieder und kannst sie wiederherstellen.' },
  };
  const LEGENDE = {
    linie: 'Linie entscheidet (Word-Änderungsverfolgung: unterstrichen = neu, durchgestrichen = entfällt)',
    farbe: 'Farbe und Strich (von Hand markiert: durchgestrichen = entfällt, farbig = neu)',
    alt: 'Blau = neu, Rot = entfällt (ältere Testdaten)',
    docx: 'Word-Änderungsverfolgung direkt gelesen (kein PDF, keine Farben)',
  };
  const AMPEL_NAME = { gruen: 'grün', gelb: 'gelb', rot: 'rot' };
  const AMPEL_ZEICHEN = { gruen: '✓', gelb: '!', rot: '✕' };
  const SP_STATUS = { passt: 'Passt', falsch: 'Falsch – wird nicht eingetragen' };
  const TXT_DATEI_HIER = 'Hierher ziehen oder klicken (.docx bevorzugt, auch .pdf)';
  const ART = {
    text_ersetzt: 'Text ersetzt', text_eingefuegt: 'Text eingefügt', text_geloescht: 'Text entfällt',
    absatz_neu: 'Absatz neu', absatz_geloescht: 'Absatz entfällt', section_neu: 'Section neu', section_geloescht: 'Section entfällt',
  };
  const STATUS = { uebernommen: 'eingetragen', pruefen: 'Vorschlag', nicht_uebernommen: 'nicht eingetragen' };
  const WIRKUNG = { ok: 'wie geplant', keine: 'keine Wirkung', abweichung: 'weicht ab' };
  const LUECKE_ART = {
    scan_seite: 'Scan-Seite (kein Text lesbar)', mehrspaltig: 'Text in mehreren Spalten', tabelle: 'Tabelle', bild: 'Bild',
    unbekannte_farbe: 'Unbekannte Farbe', nicht_lesbarer_text: 'Nicht lesbarer Text', unerklaerte_markierung: 'Unerklärte Markierung', sonstiges: 'Sonstiges',
  };
  const LAUF_STATUS = {
    analysiert: { text: 'noch keine Excel erstellt', offen: true },
    uebernommen: { text: 'Excel erstellt', offen: true },
    fehlgeschlagen: { text: 'Abgleich fehlgeschlagen', offen: false, warn: true },
    beschaedigt: { text: 'beschädigt', offen: false, warn: true },
    neuere_version: { text: 'aus einer neueren Programmversion', offen: false, warn: true },
    angelegt: { text: 'noch nicht abgeglichen', offen: false },
  };
  const TXT_SERVER = 'Der Server antwortet nicht. Läuft das Dashboard noch?';
  const TXT_UNERWARTET = 'Im Dashboard ist etwas schiefgelaufen. Bitte die Seite neu laden. Deine Abgleiche sind gespeichert.';
  const AUTO_SEITE = 20;          // so viele automatische Karten zuerst
  const ZULETZT = 5;              // so viele Abgleiche unter „Zuletzt“
  const RAND = 160;               // so viele Zeichen unveränderten Textes um eine Änderung herum (lange Zellen)
  const SCHRITTE = ['dateien', 'pruefen', 'fertig'];

  // ------------------------------------------------------------------ Zustand und Helfer
  const state = {
    schritt: 'dateien', plan: null, ergebnis: null, ergebnisse: [],
    entscheid: new Map(),         // Vorschlag-ID -> 'ja' | 'nein'
    abwahl: new Set(),            // automatische Änderungen, die NICHT eingetragen werden sollen
    stichprobe: new Map(),        // Stichproben-ID -> 'passt' | 'falsch' ('falsch' wählt die Änderung zugleich ab)
    pauschal: new Set(),          // Vorschläge, die über „Alle übernehmen“ freigegeben wurden (nicht einzeln angesehen)
    ampelSchluessel: '', ampelNr: 0, ampelZeit: null, ampelZeichen: '',
    files: { pdf: null, xlsx: null }, zeigeArchiv: false, zuletztAlle: false, suche: '', autoLimit: AUTO_SEITE,
  };
  const $ = (id) => document.getElementById(id);

  function h(tag, attrs = {}, ...kinder) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (v === false || v == null) continue;
      if (k === 'class') el.className = v;
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? '' : v);
    }
    for (const k of kinder.flat(2)) if (k != null && k !== false) el.append(k.nodeType ? k : document.createTextNode(String(k)));
    return el;
  }

  const liste = (x) => (Array.isArray(x) ? x : []);
  const text = (x) => (typeof x === 'string' ? x : x == null ? '' : String(x));
  const zahl = (x) => (typeof x === 'number' && isFinite(x) ? x : 0);
  const pl = (n, ein, mehr) => `${n} ${n === 1 ? ein : mehr}`;
  const kurz = (t, n) => { t = text(t).replace(/\s+/g, ' ').trim(); return t.length > n ? t.slice(0, n - 1) + '…' : t; };
  const hatWort = (x) => (x === 0 || x ? String(x) : '–');
  const sicher = (id) => text(id).replace(/[^\w-]/g, '_');
  const lauf = (id) => encodeURIComponent(id);
  const reduziert = () => matchMedia('(prefers-reduced-motion: reduce)').matches;

  function fmtZeit(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})/.exec(text(iso));
    return m ? `${m[3]}.${m[2]}.${m[1]} ${m[4]}:${m[5]}` : '';
  }

  // Einordnung kommt vom Server (kategorie); der Rückfall gilt nur für sehr alte Antworten ohne dieses Feld.
  function kat(a) {
    if (a.kategorie) return a.kategorie;
    if (a.status === 'uebernommen') return 'automatisch';
    if (a.bereich === 'ausserhalb') return 'ausserhalb';
    return a.status === 'pruefen' ? 'pruefen' : 'nicht';
  }
  const istDocx = () => !!(state.plan && state.plan.quelle_art === 'docx');      // Word: keine Seiten, kein PDF-Ausschnitt
  const aenderungen = () => liste(state.plan && state.plan.aenderungen);
  const vorschlaege = () => aenderungen().filter((a) => kat(a) === 'pruefen');
  const automatische = () => aenderungen().filter((a) => kat(a) === 'automatisch');
  const nichtEingetragen = () => aenderungen().filter((a) => kat(a) === 'ausserhalb' || kat(a) === 'nicht');
  function stichprobeKarten() {                         // die Stichprobe des Servers: automatisch eingetragene Änderungen, in Planreihenfolge
    const ids = new Set(liste(state.plan && state.plan.stichprobe));
    return automatische().filter((a) => ids.has(a.id));
  }
  function zusammenfassung() {
    const z = state.plan && state.plan.zusammenfassung;
    if (z && typeof z === 'object') return z;
    const l = aenderungen();
    return { gesamt: l.length, automatisch: automatische().length, pruefen: vorschlaege().length,
      ausserhalb: l.filter((a) => kat(a) === 'ausserhalb').length, nicht: l.filter((a) => kat(a) === 'nicht').length };
  }

  function infoIcon(key) {
    return h('button', { class: 'info-icon', type: 'button', 'data-info': key, 'aria-haspopup': 'dialog', 'aria-expanded': 'false',
      'aria-label': `Erklärung: ${INFO[key] ? INFO[key].titel : ''}` }, 'i');
  }

  // ------------------------------------------------------------------ Meldungen
  let toastZeit = null;
  function toast(msg) {
    const t = $('toast');
    clearTimeout(toastZeit);
    t.textContent = '';
    setTimeout(() => { t.textContent = msg; }, 30);       // kurze Pause, damit Screenreader die Änderung melden
    toastZeit = setTimeout(() => { t.textContent = ''; }, 6000);
  }
  function toastFehler(msg) {
    $('toast-fehler-box').hidden = false;
    $('toast-fehler').textContent = '';
    setTimeout(() => { $('toast-fehler').textContent = msg; }, 30);
  }
  function fehlerSchliessen() { $('toast-fehler-box').hidden = true; $('toast-fehler').textContent = ''; }
  function meldeFehler(e) {
    if (e instanceof ApiFehler) {
      const extra = e.daten && e.daten.run_id ? ' Der Abgleich bleibt unter „Zuletzt“ gespeichert.' : '';
      toastFehler(e.message + extra);
    } else {
      console.error(e);
      toastFehler(TXT_UNERWARTET);
    }
  }

  // ------------------------------------------------------------------ Warte-Anzeige (Fokus + alles andere gesperrt)
  let busyZuvor = null, busyUhr = null;
  function busy(an, titel = 'Einen Moment bitte', msg = 'Bitte warten …') {
    const b = $('busy');
    if (an) {
      $('busy-titel').textContent = titel;
      $('busy-text').textContent = msg;
      $('busy-zeit').textContent = '';
      const start = Date.now();
      clearInterval(busyUhr);
      busyUhr = setInterval(() => {
        const s = Math.round((Date.now() - start) / 1000);
        $('busy-zeit').textContent = s >= 2 ? `Bisher ${s} Sekunden` : '';
      }, 1000);
      if (b.hidden) {
        busyZuvor = document.activeElement;
        for (const id of ['kopf', 'schritte', 'main', 'fuss']) $(id).inert = true;
        b.hidden = false;
        $('busy-card').focus();
      }
    } else if (!b.hidden) {
      clearInterval(busyUhr);
      b.hidden = true;
      for (const id of ['kopf', 'schritte', 'main', 'fuss']) $(id).inert = false;
      if (busyZuvor && document.contains(busyZuvor) && busyZuvor.offsetParent !== null && typeof busyZuvor.focus === 'function') busyZuvor.focus();
      busyZuvor = null;
    }
  }

  // ------------------------------------------------------------------ API
  class ApiFehler extends Error {
    constructor(msg, daten) { super(msg); this.daten = daten || {}; }
  }
  function fehlerText(status, daten) {
    if (daten && typeof daten.fehler === 'string' && daten.fehler) return daten.fehler;
    if (status === 413) return 'Die Dateien sind zu groß.';
    if (status === 404) return 'Das wurde nicht gefunden.';
    if (status >= 500) return 'Es ist ein unerwarteter Fehler aufgetreten. Deine Daten sind unverändert. Bitte noch einmal versuchen.';
    return 'Die Anfrage konnte nicht verarbeitet werden. Bitte noch einmal versuchen.';
  }
  async function api(pfad, opt = {}) {
    let r;
    try { r = await fetch('/api' + pfad, opt); } catch (_) { throw new ApiFehler(TXT_SERVER); }
    let daten = null;
    try { daten = await r.json(); } catch (_) { /* keine JSON-Antwort */ }
    if (!r.ok) throw new ApiFehler(fehlerText(r.status, daten), daten);
    if (daten === null) throw new ApiFehler('Der Server hat eine unerwartete Antwort geschickt. Bitte noch einmal versuchen.');
    return daten;
  }

  // ------------------------------------------------------------------ Info-Popup
  let popup = null, popupAuf = null;
  function schliessePopup(fokus = true) {
    if (!popup) return;
    popup.remove();
    if (popupAuf) { popupAuf.setAttribute('aria-expanded', 'false'); if (fokus && document.contains(popupAuf)) popupAuf.focus(); }
    popup = null; popupAuf = null;
  }
  function oeffnePopup(btn) {
    const inf = INFO[btn.dataset.info];
    if (!inf) return;
    const war = popupAuf === btn;
    schliessePopup(false);
    if (war) return;
    popup = h('div', { class: 'info-popup', role: 'dialog', 'aria-label': inf.titel },
      h('div', { class: 'info-popup-title' }, inf.titel), h('p', { class: 'info-popup-text' }, inf.text),
      h('button', { class: 'info-popup-close', type: 'button', 'aria-label': 'Erklärung schließen', onclick: () => schliessePopup() }, '×'));
    document.body.append(popup);
    popupAuf = btn;
    btn.setAttribute('aria-expanded', 'true');
    setzePopup();
    popup.querySelector('.info-popup-close').focus({ preventScroll: true });
  }
  function setzePopup() {                              // folgt dem Info-Symbol, auch beim Scrollen oder Verkleinern des Fensters
    if (!popup || !popupAuf) return;
    const r = popupAuf.getBoundingClientRect(), pw = popup.offsetWidth, ph = popup.offsetHeight;
    if (r.bottom < 0 || r.top > innerHeight) { schliessePopup(false); return; }
    popup.style.left = `${Math.max(8, Math.min(r.left, innerWidth - pw - 8))}px`;
    popup.style.top = `${r.bottom + 8 + ph > innerHeight ? Math.max(8, r.top - ph - 8) : r.bottom + 8}px`;
  }

  // ------------------------------------------------------------------ Schritte
  function zeigeSchritt(name, fokus = true) {
    if (!SCHRITTE.includes(name)) return;
    state.schritt = name;
    for (const s of SCHRITTE) {
      $('schritt-' + s).hidden = s !== name;
      const k = $('schritt-knopf-' + s);
      if (s === name) k.setAttribute('aria-current', 'step'); else k.removeAttribute('aria-current');
    }
    $('schritt-knopf-pruefen').disabled = !state.plan;
    $('schritt-knopf-fertig').disabled = !state.ergebnis;
    $('schritt-knopf-dateien').classList.toggle('erledigt', !!state.plan);
    $('schritt-knopf-pruefen').classList.toggle('erledigt', !!state.ergebnis);
    zeichneWarnungen();
    zeichneFachleute();
    if (name === 'dateien') ladeZuletzt();
    if (fokus) {
      window.scrollTo({ top: 0, behavior: 'auto' });
      $('titel-' + name).focus({ preventScroll: true });
    }
  }

  // ------------------------------------------------------------------ Hell / Dunkel
  const istDunkel = () => {
    const t = document.documentElement.getAttribute('data-theme');
    return t ? t === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
  };
  function themeKnopf() { $('theme-toggle').setAttribute('aria-pressed', istDunkel() ? 'true' : 'false'); }
  function setzeTheme(t) {
    document.documentElement.setAttribute('data-theme', t);
    try { localStorage.setItem('cde-theme', t); } catch (_) { /* privater Modus */ }
    themeKnopf();
  }

  // ------------------------------------------------------------------ Schritt 1: Dateien
  function dateiGewaehlt(art, datei) {
    const okPdf = art === 'pdf' && /\.(docx|pdf)$/i.test(datei.name);
    const okXlsx = art === 'xlsx' && /\.(xlsx|xlsm)$/i.test(datei.name);
    if (!okPdf && !okXlsx) {
      toastFehler(art === 'pdf'
        ? (/\.doc$/i.test(datei.name) ? 'Alte Word-Dateien (.doc) kann das Tool nicht lesen. Bitte in Word als .docx speichern.' : 'Bitte ein Word-Dokument (.docx) oder eine PDF-Datei wählen.')
        : 'Bitte eine .xlsx-Datei wählen (alte .xls-Dateien bitte in Excel als .xlsx speichern).');
      return;
    }
    state.files[art] = datei;
    $(art === 'pdf' ? 'dz-pdf' : 'dz-xlsx').classList.add('filled');
    $(art === 'pdf' ? 'dz-pdf-sub' : 'dz-xlsx-sub').textContent = `✓ ${datei.name}`;
    $('btn-abgleichen').disabled = !(state.files.pdf && state.files.xlsx);
  }

  function dropzone(id, inputId, art) {
    const dz = $(id), input = $(inputId);
    input.addEventListener('change', () => input.files[0] && dateiGewaehlt(art, input.files[0]));
    ['dragenter', 'dragover'].forEach((e) => dz.addEventListener(e, (ev) => { ev.preventDefault(); dz.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((e) => dz.addEventListener(e, (ev) => { ev.preventDefault(); dz.classList.remove('over'); }));
    dz.addEventListener('drop', (ev) => { const f = ev.dataTransfer.files[0]; if (f) dateiGewaehlt(art, f); });
    dz.addEventListener('keydown', (ev) => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); input.click(); } });
  }

  function dateienZuruecksetzen() {
    state.files = { pdf: null, xlsx: null };
    for (const [dz, sub, inp, t] of [['dz-pdf', 'dz-pdf-sub', 'file-pdf', TXT_DATEI_HIER], ['dz-xlsx', 'dz-xlsx-sub', 'file-xlsx', 'Hierher ziehen oder klicken (.xlsx)']]) {
      $(dz).classList.remove('filled');
      $(sub).textContent = t;
      $(inp).value = '';
    }
    $('btn-abgleichen').disabled = true;
  }

  async function abgleichen() {
    if (!(state.files.pdf && state.files.xlsx)) return;
    const fd = new FormData();
    fd.append('datei', state.files.pdf);                // Word oder PDF: die Endung entscheidet auf dem Server
    fd.append('excel', state.files.xlsx);
    busy(true, /\.docx$/i.test(state.files.pdf.name) ? 'Word-Dokument wird gelesen …' : 'PDF wird gelesen …',
      'Das dauert etwa 15 Sekunden. In die Excel wird dabei noch nichts geschrieben.');
    let plan;
    try { plan = await api('/lauf', { method: 'POST', body: fd }); } catch (e) { busy(false); meldeFehler(e); ladeZuletzt(); return; }
    busy(false);
    oeffne(plan);
    toast('Abgleich fertig. Es wurde noch nichts in die Excel geschrieben.');
  }

  // Zuletzt: die letzten Abgleiche
  function laufZeile(r) {
    const st = LAUF_STATUS[r.status] || { text: 'unbekannter Zustand', offen: false, warn: true };
    const name = r.pdf && r.excel ? `${r.pdf} → ${r.excel}` : `Abgleich ${text(r.run_id)}`;
    const meta = [fmtZeit(r.angelegt)];
    if (r.status === 'uebernommen') meta.push(pl(zahl(r.ergebnisse), 'Excel erstellt', 'Excel-Dateien erstellt'));
    else meta.push(h('span', { class: st.warn ? 'warn' : '' }, st.text));
    if (r.spec) meta.push(`Spec ${r.spec}`);
    if (r.archiviert) meta.push('archiviert');
    const teile = meta.filter(Boolean).flatMap((m, i) => (i ? [' · ', m] : [m]));
    const menu = h('details', { class: 'menu' });
    const eintrag = r.archiviert
      ? h('button', { class: 'menu-eintrag', type: 'button', onclick: () => { menu.open = false; wiederherstellen(r.run_id); } }, 'Wiederherstellen')
      : h('button', { class: 'menu-eintrag', type: 'button', onclick: () => { menu.open = false; archivieren(r.run_id); } }, 'Archivieren');
    menu.append(h('summary', { class: 'menu-knopf', 'aria-label': `Mehr zu ${name}`, title: 'Mehr' }, h('span', { 'aria-hidden': 'true' }, '⋯')),
      h('div', { class: 'menu-liste' }, eintrag));
    return h('li', { class: 'lauf-zeile' + (r.archiviert ? ' archiviert' : '') },
      h('div', { class: 'lauf-info' }, h('div', { class: 'lauf-titel' }, name), h('div', { class: 'lauf-meta' }, teile),
        r.meldung ? h('div', { class: 'lauf-meta' }, text(r.meldung)) : ''),
      h('div', { class: 'lauf-knoepfe' },
        st.offen ? h('button', { class: 'btn btn-secondary btn-small', type: 'button', 'aria-label': `Öffnen: ${name}`, onclick: () => laden(r.run_id) }, 'Öffnen') : '',
        menu));
  }

  async function ladeZuletzt() {
    const box = $('zuletzt-liste');
    try {
      const l = liste(await api(state.zeigeArchiv ? '/lauf?archiviert=1' : '/lauf'));
      const knopf = $('btn-zuletzt-alle');
      knopf.hidden = l.length <= ZULETZT;
      knopf.textContent = state.zuletztAlle ? 'Weniger anzeigen' : `Alle ${l.length} anzeigen`;
      if (!l.length) {
        box.replaceChildren(h('p', { class: 'leer' }, state.zeigeArchiv ? 'Es gibt noch keine Abgleiche.' : 'Noch keine Abgleiche. Lege oben die zwei Dateien ab.'));
        return;
      }
      box.replaceChildren(h('ul', { class: 'lauf-liste' }, (state.zuletztAlle ? l : l.slice(0, ZULETZT)).map(laufZeile)));
    } catch (e) {
      box.replaceChildren(h('p', { class: 'leer' }, e.message));
      if (!(e instanceof ApiFehler)) meldeFehler(e);
    }
  }

  async function laden(id) {
    busy(true, 'Abgleich wird geöffnet …');
    let plan;
    try { plan = await api('/lauf/' + lauf(id)); } catch (e) { busy(false); meldeFehler(e); return; }
    busy(false);
    oeffne(plan);
  }

  async function archivieren(id) {
    try {
      await api('/lauf/' + lauf(id), { method: 'DELETE' });
      toast('Archiviert. Der Abgleich bleibt auf diesem Rechner erhalten. Unter „Archivierte anzeigen“ kannst du ihn wiederherstellen.');
      if (state.plan && state.plan.run_id === id) state.plan.archiviert = true;
      await ladeZuletzt();
      $('zuletzt-titel').focus();
    } catch (e) { meldeFehler(e); }
  }

  async function wiederherstellen(id) {
    try {
      await api(`/lauf/${lauf(id)}/wiederherstellen`, { method: 'POST' });
      toast('Der Abgleich ist wiederhergestellt.');
      if (state.plan && state.plan.run_id === id) state.plan.archiviert = false;
      await ladeZuletzt();
      $('zuletzt-titel').focus();
    } catch (e) { meldeFehler(e); }
  }

  // ------------------------------------------------------------------ Einen Abgleich öffnen
  function oeffne(plan) {
    state.plan = plan;
    state.ergebnisse = liste(plan.ergebnisse);
    state.entscheid = new Map();
    state.abwahl = new Set();
    state.stichprobe = new Map();
    state.pauschal = new Set();
    state.suche = '';
    state.autoLimit = AUTO_SEITE;
    $('suche-auto').value = '';
    $('klapp-auto').open = false;
    $('klapp-nicht').open = false;
    const letzte = state.ergebnisse[state.ergebnisse.length - 1];
    state.ergebnis = letzte ? { ...letzte, aenderungen: null, frueher: true } : null;
    if (letzte) {                                      // die Entscheidungen der letzten Excel wieder vorbelegen
      const v = new Set(vorschlaege().map((a) => a.id)), au = new Set(automatische().map((a) => a.id));
      for (const id of liste(letzte.freigaben)) if (v.has(id)) state.entscheid.set(id, 'ja');
      for (const id of liste(letzte.abgelehnt)) if (v.has(id)) state.entscheid.set(id, 'nein');
      for (const id of liste(letzte.abwahl)) if (au.has(id)) state.abwahl.add(id);
      const sp = new Set(liste(state.plan.stichprobe)), st = letzte.stichprobe || {};
      for (const id of liste(st.passt)) if (sp.has(id) && au.has(id) && !state.abwahl.has(id)) state.stichprobe.set(id, 'passt');
      for (const id of liste(st.falsch)) if (sp.has(id) && state.abwahl.has(id)) state.stichprobe.set(id, 'falsch');
    }
    zeichnePruefen();
    if (letzte) zeichneFertig();
    zeigeSchritt(letzte ? 'fertig' : 'pruefen');
  }

  function neuerAbgleich() {
    state.plan = null;
    state.ergebnis = null;
    state.ergebnisse = [];
    state.entscheid = new Map();
    state.abwahl = new Set();
    state.stichprobe = new Map();
    state.pauschal = new Set();
    state.ampelSchluessel = '';
    dateienZuruecksetzen();
    $('fachleute').open = false;
    zeigeSchritt('dateien');
  }

  // ------------------------------------------------------------------ Darstellung einer Änderung
  function ortText(a) {
    const s = text(a.section);
    const teile = [s ? (/^\d/.test(s) ? `Section ${s}` : s) : (istDocx() ? 'Im Word-Dokument' : `Seite ${hatWort(a.seite)}`)];
    const z = text(a.ziel);
    let m = /Zeile\s+(\d+)\s*\(neu\)\s*$/.exec(z);
    if (m) teile.push(`neue Zeile ${m[1]}`);
    else if ((m = /![A-Z]{1,3}(\d+)(?::[A-Z]{1,3}\d+)?$/.exec(z))) teile.push(`Zeile ${m[1]}`);
    return teile.join(' · ');
  }

  function grundText(a) {
    const g = text(a.grund_einfach) || text(a.grund);
    if (g) return g;
    if (kat(a) === 'ausserhalb') return 'Gehört nicht zur CDE (zum Beispiel Deckblatt oder Revisionstabelle).';
    if (kat(a) === 'pruefen') return 'Das Tool ist sich an dieser Stelle nicht sicher.';
    return '–';
  }

  const altEl = (t) => h('del', { class: 'st-alt' }, h('span', { class: 'sr-only' }, '[fällt weg: '), t, h('span', { class: 'sr-only' }, ']'));
  const neuEl = (t) => h('ins', { class: 'st-neu' }, h('span', { class: 'sr-only' }, '[kommt dazu: '), t, h('span', { class: 'sr-only' }, ']'));
  const stuecke = (a) => liste(a.stuecke).filter((s) => Array.isArray(s) && s.length >= 2 && typeof s[1] === 'string');
  const ausgelassen = () => h('span', { class: 'ausgelassen', 'aria-label': 'Text ausgelassen' }, '…');

  function gekuerzt(t, erstes, letztes) {
    if (erstes && letztes) return t.length > 2 * RAND ? [t.slice(0, RAND), ausgelassen()] : [t];
    if (erstes) return t.length > RAND ? [ausgelassen(), t.slice(-RAND)] : [t];
    if (letztes) return t.length > RAND ? [t.slice(0, RAND), ausgelassen()] : [t];
    return t.length > 2 * RAND ? [t.slice(0, RAND), ' ', ausgelassen(), ' ', t.slice(-RAND)] : [t];
  }

  function fuelleVorschau(p, st, kuerzen) {
    p.replaceChildren(...st.flatMap(([fmt, t], i) => {
      if (fmt === 'alt') return [altEl(t)];
      if (fmt === 'neu') return [neuEl(t)];
      return kuerzen ? gekuerzt(t, i === 0, i === st.length - 1) : [t];
    }));
  }

  // Ergebnis-Vorschau genau wie in der Excel (stuecke); ohne stuecke wie früher: alt durchgestrichen, neu farbig.
  function vorschau(a) {
    const p = h('p', { class: 'vorschau' });
    const st = stuecke(a);
    if (!st.length) {
      if (!a.alt && !a.neu) { p.classList.add('vorschau-leer'); p.textContent = '(kein Text)'; return p; }
      p.append(...[a.alt ? altEl(text(a.alt)) : '', a.alt && a.neu ? ' ' : '', a.neu ? neuEl(text(a.neu)) : ''].filter(Boolean));
      return p;
    }
    const lang = st.some(([fmt, t]) => fmt === 'normal' && t.length > 2 * RAND);
    fuelleVorschau(p, st, lang);
    if (!lang) return p;
    let ganz = false;
    const knopf = h('button', { class: 'link-btn mehr-text', type: 'button', 'aria-expanded': 'false', onclick: () => {
      ganz = !ganz;
      fuelleVorschau(p, st, !ganz);
      knopf.setAttribute('aria-expanded', ganz ? 'true' : 'false');
      knopf.textContent = ganz ? 'Weniger zeigen' : 'Ganzen Text zeigen';
    } }, 'Ganzen Text zeigen');
    return h('div', {}, p, knopf);
  }

  function bildPdf(a) {
    if (istDocx()) return null;                         // Word hat keine Seiten: nur der Text (Ergebnis wie in der Excel), kein PDF-Ausschnitt
    return h('figure', { class: 'im-pdf' }, h('figcaption', {}, `Im PDF · Seite ${hatWort(a.seite)}`),
      h('img', { loading: 'lazy', alt: `Ausschnitt aus dem PDF, Seite ${hatWort(a.seite)}`,
        src: `/api/lauf/${lauf(state.plan.run_id)}/vorschau/${encodeURIComponent(a.id)}.png`,
        onerror: (ev) => { const fig = ev.target.parentNode; ev.target.remove(); if (fig) fig.append(h('p', { class: 'bild-fehlt' }, 'Der PDF-Ausschnitt konnte nicht geladen werden.')); } }));
  }

  const grundZeile = (a) => h('p', { class: 'grund' }, h('span', { class: 'grund-label' }, 'Grund'), grundText(a));

  // ------------------------------------------------------------------ Spec (eine CDE kann mehrere Specs enthalten)
  // Nico filtert immer nach der Spec des PDFs. Die Auswahl kommt vom Server (plan.spec, plan.specs); ein anderer Spec = neuer Abgleich.
  function specInfo() {
    const p = state.plan || {};
    const sp = p.spec && typeof p.spec === 'object' ? p.spec : null;
    const namen = liste(p.specs).filter((x) => x && typeof x.name === 'string');
    const liste_ = namen.length ? namen : liste(sp && sp.kandidaten).map((n) => ({ name: text(n) }));
    const mehr = liste_.length > 1;                                    // bei nur einer Spec: keine Anzeige, kein Auswahlzwang
    const codeUnklar = aenderungen().some((a) => liste(a.pruefcodes).includes('spec_unklar'));
    const unklar = mehr && (!(sp && sp.erkannt) || codeUnklar);
    return { mehr, unklar, erkannt: sp && sp.erkannt ? text(sp.erkannt) : '', zeilen: sp && Array.isArray(sp.zeilen) ? sp.zeilen : null,
      quelle: sp ? text(sp.quelle) : '', liste: liste_ };
  }
  const specZeilen = (x) => (x.von != null && x.bis != null ? `${pl(zahl(x.zeilen), 'Zeile', 'Zeilen')} (${x.von}–${x.bis})` : '');

  function zeichneSpec() {
    const si = specInfo();
    $('spec-zone').hidden = !si.mehr;
    if (!si.mehr) return;
    const zeile = si.erkannt ? [h('span', {}, 'Spec: '), h('strong', {}, si.erkannt),
      si.zeilen ? ` · Zeilen ${si.zeilen[0]}–${si.zeilen[1]}` : '', si.quelle === 'vorgabe' ? ' · von dir gewählt' : si.quelle === 'pdf' ? ' · im PDF erkannt' : ''] : [];
    $('spec-zeile').hidden = si.unklar;
    $('spec-text').replaceChildren(...zeile);
    const wahl = $('spec-wahl');
    wahl.classList.toggle('unklar', si.unklar);
    if (si.unklar) wahl.hidden = false;
    else if ($('btn-spec-andere').getAttribute('aria-expanded') !== 'true') wahl.hidden = true;
    wahl.setAttribute('role', si.unklar ? 'alert' : 'group');
    if (!si.unklar) wahl.setAttribute('aria-labelledby', 'spec-titel');
    $('spec-titel').textContent = si.unklar ? 'Welche Spec gehört zu diesem PDF?' : 'Andere Spec wählen';
    $('spec-frage').textContent = si.unklar
      ? 'Das Tool konnte nicht sicher erkennen, zu welcher Spec dieses PDF gehört. Bis du eine wählst, wird nichts eingetragen. Es werden nur die Zeilen dieser Spec bearbeitet, die anderen bleiben unverändert.'
      : 'Deine CDE enthält mehrere Specs. Es werden nur die Zeilen der gewählten Spec bearbeitet, die anderen bleiben unverändert.';
    $('spec-liste').replaceChildren(...si.liste.map((x) => h('li', {}, h('button', { class: 'spec-knopf', type: 'button', 'data-spec': x.name,
      'aria-current': x.name === si.erkannt ? 'true' : null, 'aria-label': `Spec ${x.name}${specZeilen(x) ? ', ' + specZeilen(x) : ''}${x.name === si.erkannt ? ', gerade gewählt' : ''}`,
      onclick: () => specWaehlen(x.name) }, h('span', {}, x.name), h('span', { class: 'meta' }, specZeilen(x))))));
  }

  function specAndere() {
    const knopf = $('btn-spec-andere'), offen = knopf.getAttribute('aria-expanded') !== 'true';
    knopf.setAttribute('aria-expanded', offen ? 'true' : 'false');
    $('spec-wahl').hidden = !offen;
    if (offen) $('spec-titel').focus({ preventScroll: false });
  }

  async function specWaehlen(name) {
    if (!state.plan) return;
    const id = state.plan.run_id;
    busy(true, 'Neuer Abgleich …', `Die Dateien werden mit der Spec ${name} neu abgeglichen. Das dauert etwa 15 Sekunden. Der bisherige Abgleich bleibt erhalten.`);
    let plan;
    try {
      plan = await api(`/lauf/${lauf(id)}/neu-analysieren`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ spec: name }) });
    } catch (e) { busy(false); meldeFehler(e); return; }
    busy(false);
    oeffne(plan);
    toast(`Neuer Abgleich mit der Spec ${name}. Der vorige bleibt unter „Zuletzt“.`);
  }

  // ------------------------------------------------------------------ Schritt 2: Prüfen
  function statusText(e) {
    if (e === 'ja') return [h('span', { 'aria-hidden': 'true' }, '✓'), 'Wird übernommen'];
    if (e === 'nein') return [h('span', { 'aria-hidden': 'true' }, '✕'), 'Wird nicht übernommen'];
    return [h('span', { 'aria-hidden': 'true' }, '●'), 'Noch offen'];
  }

  function karteBitte(a) {
    const e = state.entscheid.get(a.id) || '';
    const tid = 'kt-' + sicher(a.id), sid = 'ks-' + sicher(a.id);
    const knopf = (wert, beschriftung) => h('button', { class: `btn btn-secondary entscheid ${wert}`, type: 'button', 'data-wert': wert,
      'aria-pressed': e === wert ? 'true' : 'false', 'aria-describedby': tid, onclick: () => entscheide(a.id, wert, false) }, beschriftung);
    return h('li', {}, h('article', { class: 'karte', tabindex: '0', 'data-id': a.id, 'data-entscheid': e || 'offen', 'aria-labelledby': tid, 'aria-describedby': sid },
      h('div', { class: 'karte-kopf' }, h('h3', { id: tid }, ortText(a)), h('p', { class: 'karte-status', id: sid }, statusText(e))),
      vorschau(a),
      grundZeile(a),
      h('div', { class: 'karte-unten' + (istDocx() ? ' ohne-bild' : '') }, bildPdf(a),
        h('div', { class: 'entscheid-knoepfe', role: 'group', 'aria-label': 'Entscheidung' }, knopf('ja', 'Übernehmen'), knopf('nein', 'Nicht übernehmen')))));
  }

  // ------------------------------------------------------------------ Stichprobe (Vertrauensmodus): automatisch Eingetragenes kurz ansehen
  function karteStichprobe(a) {
    const e = state.stichprobe.get(a.id) || '';
    const tid = 'sk-' + sicher(a.id), sid = 'ss-' + sicher(a.id);
    const knopf = (wert, beschriftung) => h('button', { class: `btn btn-secondary entscheid ${wert}`, type: 'button', 'data-wert': wert,
      'aria-pressed': e === wert ? 'true' : 'false', 'aria-describedby': tid, onclick: () => stichprobeEntscheide(a.id, wert, false) }, beschriftung);
    // Schlank gehalten: der Text wie in der Excel und zwei Knöpfe; PDF-Ausschnitt und Grund liegen eingeklappt dabei
    return h('li', {}, h('article', { class: 'karte karte-sp', tabindex: '0', 'data-id': a.id, 'data-art': 'stichprobe', 'data-entscheid': e || 'offen',
      'aria-labelledby': tid, 'aria-describedby': sid },
      h('div', { class: 'karte-kopf' }, h('h3', { id: tid }, ortText(a)), h('p', { class: 'karte-status', id: sid }, stichprobeStatus(e))),
      vorschau(a),
      h('div', { class: 'sp-unten' },
        h('details', { class: 'mini' }, h('summary', {}, istDocx() ? 'Grund ansehen' : 'Im PDF ansehen'), bildPdf(a), grundZeile(a)),
        h('div', { class: 'entscheid-knoepfe', role: 'group', 'aria-label': 'Stichprobe: Entscheidung' }, knopf('passt', 'Passt'), knopf('falsch', 'Falsch')))));
  }

  function stichprobeStatus(e) {
    if (e === 'passt') return [h('span', { 'aria-hidden': 'true' }, '✓'), SP_STATUS.passt];
    if (e === 'falsch') return [h('span', { 'aria-hidden': 'true' }, '✕'), SP_STATUS.falsch];
    return [h('span', { 'aria-hidden': 'true' }, '●'), 'Noch nicht angesehen'];
  }

  function aktualisiereStichprobeKarte(id) {
    const k = karteZu('karten-stichprobe', id);
    if (!k) return;
    const e = state.stichprobe.get(id) || '';
    k.dataset.entscheid = e || 'offen';
    for (const b of k.querySelectorAll('.entscheid')) b.setAttribute('aria-pressed', b.dataset.wert === e ? 'true' : 'false');
    k.querySelector('.karte-status').replaceChildren(...stichprobeStatus(e));
  }

  function zeichneStichprobeKopf() {
    const l = stichprobeKarten();
    const angesehen = l.filter((a) => state.stichprobe.has(a.id)).length;
    $('stichprobe-stand').textContent = `${angesehen} von ${l.length} angesehen`;
  }

  function zeichneStichprobe() {
    const l = stichprobeKarten();
    $('klapp-stichprobe').hidden = !l.length;
    $('anzahl-stichprobe').textContent = String(l.length);
    $('karten-stichprobe').replaceChildren(...l.map(karteStichprobe));
    $('klapp-stichprobe').open = l.some((a) => !state.stichprobe.has(a.id));       // schon alles angesehen: eingeklappt
    zeichneStichprobeKopf();
  }

  // „Falsch“ wählt die Änderung ab (wie „Trotzdem nicht eintragen“) und wird als Ablehnung vermerkt; ein zweiter Klick hebt es auf.
  function stichprobeEntscheide(id, wert, weiter) {
    if (state.stichprobe.get(id) === wert && !weiter) state.stichprobe.delete(id);
    else state.stichprobe.set(id, wert);
    setzeAbwahl(id, state.stichprobe.get(id) === 'falsch', true);
    aktualisiereStichprobeKarte(id);
    zeichneStichprobeKopf();
    zeichneFortschritt();
    if (weiter) naechsteOffene(id);
  }

  // ------------------------------------------------------------------ Dokument-Ampel (Regeln und Texte kommen vom Server)
  function zeichneAmpel(a) {
    const el = $('ampel');
    if (!a || !AMPEL_NAME[a.stufe]) { el.hidden = true; return; }
    const gruende = liste(a.gruende).map(text).filter(Boolean);
    const zeichen = [a.stufe, text(a.satz), ...gruende].join('|');
    el.hidden = false;
    if (zeichen === state.ampelZeichen) return;          // nur neu zeichnen (und ansagen), wenn sich etwas ändert
    state.ampelZeichen = zeichen;
    el.dataset.stufe = a.stufe;
    $('ampel-symbol').textContent = AMPEL_ZEICHEN[a.stufe];
    $('ampel-stufe').textContent = `Dokument-Ampel ${AMPEL_NAME[a.stufe]}: `;
    $('ampel-satz').textContent = text(a.satz);
    $('ampel-gruende').replaceChildren(...gruende.map((g) => h('li', {}, g)));
  }

  // Stand der Entscheidungen -> Ampel vom Server (die Regeln stehen nur dort). Nur wenn sich die Zahlen ändern, kurz verzögert.
  function ampelZahlen() {
    const bitte = vorschlaege();
    return { offen: bitte.filter((a) => !state.entscheid.has(a.id)).length,
      stichprobeOffen: stichprobeKarten().filter((a) => !state.stichprobe.has(a.id)).length };
  }
  function planeAmpel() {
    if (!state.plan) return;
    const z = ampelZahlen(), schluessel = `${state.plan.run_id}|${z.offen}|${z.stichprobeOffen}`;
    if (schluessel === state.ampelSchluessel) return;
    state.ampelSchluessel = schluessel;
    const nr = ++state.ampelNr, id = state.plan.run_id;
    clearTimeout(state.ampelZeit);
    state.ampelZeit = setTimeout(async () => {
      try {
        const a = await api(`/lauf/${lauf(id)}/ampel?offen=${z.offen}&stichprobe_offen=${z.stichprobeOffen}`);
        if (nr === state.ampelNr && state.plan && state.plan.run_id === id) zeichneAmpel(a);
      } catch (_) { /* die Ampel bleibt, wie sie war */ }
    }, 120);
  }

  function karteZu(listeId, id) { return [...$(listeId).querySelectorAll('.karte')].find((k) => k.dataset.id === id) || null; }

  function aktualisiereKarte(id) {
    const k = karteZu('karten-bitte', id);
    if (!k) return;
    const e = state.entscheid.get(id) || '';
    k.dataset.entscheid = e || 'offen';
    for (const b of k.querySelectorAll('.entscheid')) b.setAttribute('aria-pressed', b.dataset.wert === e ? 'true' : 'false');
    k.querySelector('.karte-status').replaceChildren(...statusText(e));
  }

  function entscheide(id, wert, weiter) {
    if (state.entscheid.get(id) === wert && !weiter) state.entscheid.delete(id);    // erneuter Klick hebt die Entscheidung auf
    else state.entscheid.set(id, wert);
    state.pauschal.delete(id);                                                      // einzeln entschieden
    aktualisiereKarte(id);
    zeichneFortschritt();
    if (weiter) naechsteOffene(id);
  }

  const istOffen = (k) => !(k.dataset.art === 'stichprobe' ? state.stichprobe : state.entscheid).has(k.dataset.id);
  function naechsteOffene(id) {                          // erst die Stichprobe, dann „Bitte prüfen“ (Reihenfolge wie auf der Seite)
    const karten = [...$('schritt-pruefen').querySelectorAll('#karten-stichprobe .karte, #karten-bitte .karte')].filter((k) => k.offsetParent !== null);
    const i = karten.findIndex((k) => k.dataset.id === id);
    const reihe = [...karten.slice(i + 1), ...karten.slice(0, Math.max(i, 0))];
    const ziel = reihe.find(istOffen);
    if (ziel) fokusKarte(ziel);
    else if (karten[i + 1]) fokusKarte(karten[i + 1]);
    else $('btn-excel').focus();
  }

  function fokusKarte(k) {
    k.focus({ preventScroll: true });
    k.scrollIntoView({ block: 'nearest', behavior: reduziert() ? 'auto' : 'smooth' });
  }

  function karteAuto(a) {
    const tid = 'ka-' + sicher(a.id);
    const ab = state.abwahl.has(a.id);
    const partner = a.gruppe ? automatische().filter((x) => x.gruppe === a.gruppe && x.id !== a.id).length : 0;
    const hinweis = h('p', { class: 'gruppe-hinweis', hidden: !(ab && partner) },
      `Hängt mit ${pl(partner, 'weiteren Änderung', 'weiteren Änderungen')} zusammen. ${partner === 1 ? 'Die wird' : 'Die werden'} dann auch nicht eingetragen.`);
    const cb = h('input', { type: 'checkbox', role: 'switch', 'aria-describedby': tid, onchange: (ev) => setzeAbwahl(a.id, ev.target.checked) });
    cb.checked = ab;
    const karte = h('article', { class: 'karte kompakt' + (ab ? ' abgewaehlt' : ''), tabindex: '0', 'data-id': a.id, 'data-partner': String(partner), 'aria-labelledby': tid },
      h('div', { class: 'karte-kopf' }, h('h3', { id: tid }, ortText(a)), h('label', { class: 'schalter' }, cb, 'Trotzdem nicht eintragen')),
      vorschau(a), hinweis,
      h('details', { class: 'mini' }, h('summary', {}, istDocx() ? 'Grund ansehen' : 'Im PDF ansehen'), bildPdf(a), grundZeile(a)));
    return h('li', {}, karte);
  }

  function setzeAbwahl(id, an, vonStichprobe = false) {
    if (an) state.abwahl.add(id); else state.abwahl.delete(id);
    const k = karteZu('karten-auto', id);
    if (k) {
      k.classList.toggle('abgewaehlt', an);
      k.querySelector('input[type=checkbox]').checked = an;
      k.querySelector('.gruppe-hinweis').hidden = !(an && Number(k.dataset.partner) > 0);
    }
    if (!vonStichprobe && karteZu('karten-stichprobe', id)) {      // dieselbe Änderung in der Stichprobe: abwählen = „Falsch“
      if (an) state.stichprobe.set(id, 'falsch');
      else if (state.stichprobe.get(id) === 'falsch') state.stichprobe.delete(id);
      aktualisiereStichprobeKarte(id);
      zeichneStichprobeKopf();
    }
    zeichneFortschritt();
  }

  function autoTreffer() {
    const woerter = state.suche.toLowerCase().split(/\s+/).filter(Boolean);
    if (!woerter.length) return automatische();
    return automatische().filter((a) => {
      const heu = [a.alt, a.neu, a.section, a.ziel, a.grund_einfach, a.grund, ...stuecke(a).map((s) => s[1])].map(text).join('\n').toLowerCase();
      return woerter.every((w) => heu.includes(w));
    });
  }

  function zeichneAuto() {
    const alle = automatische(), l = autoTreffer();
    const sicht = l.slice(0, state.autoLimit);
    $('karten-auto').replaceChildren(...sicht.map(karteAuto));
    const knopf = $('btn-auto-alle');
    knopf.hidden = l.length <= sicht.length;
    knopf.textContent = `Alle ${l.length} anzeigen`;
    $('suche-info').textContent = state.suche ? `${l.length} von ${alle.length} gefunden.` : '';
  }

  function zeichneNicht() {
    $('liste-nicht').replaceChildren(...nichtEingetragen().map((a) => {
      const st = stuecke(a).filter(([fmt]) => fmt !== 'normal').map((s) => s[1]).join(' ');
      return h('li', {}, h('span', { class: 'nicht-ort' }, ortText(a)), h('span', { class: 'nicht-text' }, `„${kurz(st || a.neu || a.alt, 140)}“`),
        h('span', { class: 'nicht-grund' }, `Grund: ${grundText(a)}`));
    }));
  }

  function zeichnePruefen() {
    const p = state.plan;
    if (!p) return;
    $('lauf-name').textContent = `${text(p.pdf_name)} → ${text(p.excel_name)}${p.archiviert ? ' · archiviert' : ''}`;
    const bitte = vorschlaege();
    $('anzahl-bitte').textContent = String(bitte.length);
    $('karten-bitte').replaceChildren(...bitte.map(karteBitte));
    zeichneStichprobe();
    state.ampelZeichen = '';
    zeichneAmpel(p.ampel);
    state.ampelSchluessel = `${p.run_id}|${bitte.length}|${stichprobeKarten().length}`;     // Ausgangslage: noch nichts entschieden
    $('btn-spec-andere').setAttribute('aria-expanded', 'false');
    zeichneSpec();
    $('bitte-leer').hidden = bitte.length > 0;
    $('btn-alle-ja').hidden = bitte.length < 2;
    $('anzahl-auto').textContent = String(automatische().length);
    $('anzahl-nicht').textContent = String(nichtEingetragen().length);
    $('klapp-auto').hidden = !automatische().length;
    $('klapp-nicht').hidden = !nichtEingetragen().length;
    zeichneAuto();
    zeichneNicht();
    zeichneFortschritt();
  }

  function zeichneFortschritt() {
    if (!state.plan) return;
    const z = zusammenfassung();
    const bitte = vorschlaege();
    const erledigt = bitte.filter((a) => state.entscheid.has(a.id)).length;
    const ja = bitte.filter((a) => state.entscheid.get(a.id) === 'ja').length;
    const abw = automatische().filter((a) => state.abwahl.has(a.id)).length;
    const teile = [`${pl(zahl(z.gesamt), 'Änderung', 'Änderungen')} gefunden`,
      `${zahl(z.automatisch)} automatisch eingetragen${abw ? ` (${abw} abgewählt)` : ''}`,
      `${bitte.length} bitte prüfen (${erledigt} erledigt)`];
    const draussen = zahl(z.ausserhalb) + zahl(z.nicht);
    if (draussen) teile.push(`${draussen} nicht in der CDE`);
    const sp = stichprobeKarten(), spOffen = sp.filter((a) => !state.stichprobe.has(a.id)).length;
    if (sp.length) teile.push(`Stichprobe ${sp.length - spOffen} von ${sp.length} angesehen`);
    $('fortschritt-satz').textContent = teile.join(' · ');
    const balken = $('fortschritt');
    balken.hidden = !bitte.length;
    balken.setAttribute('aria-valuemax', String(bitte.length));
    balken.setAttribute('aria-valuenow', String(erledigt));
    balken.setAttribute('aria-valuetext', `${erledigt} von ${bitte.length} Vorschlägen geprüft`);
    $('fortschritt-fuellung').style.width = bitte.length ? `${Math.round((erledigt / bitte.length) * 100)}%` : '0%';

    const offen = bitte.length - erledigt;
    const saetze = [];
    if (specInfo().unklar) {
      $('excel-hinweis').textContent = 'Bitte zuerst die Spec wählen (siehe oben). Bis dahin wird nichts eingetragen.';
      return;
    }
    if (offen) saetze.push(`${offen === 1 ? '1 Vorschlag ist' : `${offen} Vorschläge sind`} noch offen und ${offen === 1 ? 'wird' : 'werden'} nicht eingetragen.`);
    else if (bitte.length) saetze.push('Alle Vorschläge sind entschieden.');
    if (spOffen) saetze.push(`In der Stichprobe ${spOffen === 1 ? 'ist 1 Änderung' : `sind ${spOffen} Änderungen`} noch nicht angesehen (das hält nichts auf).`);
    saetze.push(`Eingetragen ${automatische().length - abw + ja === 1 ? 'wird 1 Änderung' : `werden ${automatische().length - abw + ja} Änderungen`}.`);
    if (state.ergebnisse.length) saetze.push(`Zu diesem Abgleich gibt es schon ${pl(state.ergebnisse.length, 'Excel', 'Excel-Dateien')}; eine neue kommt dazu, die frühere bleibt.`);
    $('excel-hinweis').textContent = saetze.join(' ');
    planeAmpel();
  }

  async function excelErstellen() {
    if (!state.plan) return;
    if (specInfo().unklar) {                           // nichts wird eingetragen, solange keine Spec gewählt ist
      toastFehler('Bitte zuerst wählen, welche Spec zu diesem PDF gehört. Bis dahin wird nichts eingetragen.');
      $('spec-zone').scrollIntoView({ block: 'center', behavior: reduziert() ? 'auto' : 'smooth' });
      $('spec-titel').focus({ preventScroll: true });
      return;
    }
    const freigaben = vorschlaege().filter((a) => state.entscheid.get(a.id) === 'ja').map((a) => a.id);
    const abwahl = automatische().filter((a) => state.abwahl.has(a.id)).map((a) => a.id);
    // Nur fürs Freigabe-Protokoll (ändert nichts am Eintragen): ausdrücklich Abgelehntes, pauschal Freigegebenes, Stichprobe Passt/Falsch
    const abgelehnt = vorschlaege().filter((a) => state.entscheid.get(a.id) === 'nein').map((a) => a.id);
    const pauschal = freigaben.filter((id) => state.pauschal.has(id));
    const stichprobe = { passt: [], falsch: [] };
    for (const a of stichprobeKarten()) { const e = state.stichprobe.get(a.id); if (e) stichprobe[e].push(a.id); }
    if (automatische().length - abwahl.length + freigaben.length <= 0) {
      toastFehler('Es gibt nichts einzutragen. Gib mindestens einen Vorschlag mit „Übernehmen“ frei.');
      return;
    }
    const id = state.plan.run_id;
    busy(true, 'Excel wird erstellt …', 'Die Änderungen werden in eine Kopie deiner Excel eingetragen. Das Original bleibt unverändert.');
    let res;
    try {
      res = await api(`/lauf/${lauf(id)}/uebernehmen`, { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ freigaben, abwahl, abgelehnt, pauschal, stichprobe }) });
    } catch (e) { busy(false); meldeFehler(e); return; }
    state.ergebnis = res;
    try { state.ergebnisse = liste((await api('/lauf/' + lauf(id))).ergebnisse); }
    catch (_) { state.ergebnisse = [...state.ergebnisse, { ordner: res.ordner, excel: res.excel, bilanz: res.bilanz, erstellt: '', freigaben, abwahl }]; }
    busy(false);
    zeichneFertig();
    zeichneFortschritt();
    zeigeSchritt('fertig');
    toast('Die neue Excel ist fertig.');
  }

  // ------------------------------------------------------------------ Schritt 3: Fertig
  const dlPfad = (ordner, name) => `/api/lauf/${lauf(state.plan.run_id)}/download/${encodeURIComponent(ordner)}/${encodeURIComponent(name)}`;

  function zeichneFertig() {
    const e = state.ergebnis;
    if (!e || !state.plan) return;
    const b = e.bilanz || {};
    const abw = zahl(b.nachkontrolle_abweichungen);
    const problem = b.stimmt !== true || abw > 0;
    $('fertig-symbol').textContent = problem ? '!' : '✓';
    $('fertig-symbol').classList.toggle('warn', problem);
    $('titel-fertig').textContent = problem ? 'Fertig – bitte kurz ansehen.' : 'Fertig.';
    const wann = e.frueher && fmtZeit(e.erstellt) ? ` (erstellt am ${fmtZeit(e.erstellt)})` : '';
    $('fertig-satz').textContent = `${pl(zahl(b.uebernommen), 'Änderung', 'Änderungen')} eingetragen.${wann}`;
    const warn = [];
    if (abw) {
      const betroffen = liste(e.aenderungen).filter((a) => a.nachkontrolle === 'abweichung' || a.wirkung === 'keine' || a.wirkung === 'abweichung');
      warn.push(h('div', { class: 'banner gelb', role: 'status' }, h('span', { class: 'banner-symbol', 'aria-hidden': 'true' }, '!'),
        h('div', { class: 'banner-inhalt' },
          h('div', { class: 'banner-titel' }, `Bei ${pl(abw, 'Änderung', 'Änderungen')} ist die Excel nicht genau so geworden wie geplant`),
          h('p', { class: 'banner-text' }, 'Sie sind im Protokoll markiert. Bitte diese Stellen in der neuen Excel ansehen.'),
          betroffen.length ? h('ul', { class: 'banner-liste' }, betroffen.slice(0, 20).map((a) => h('li', {}, `${ortText(a)}: ${grundText(a)}`))) : '')));
    }
    $('fertig-warn').replaceChildren(...warn);
    $('fertig-knoepfe').replaceChildren(
      h('a', { class: 'btn btn-primary', href: dlPfad(e.ordner, e.excel), download: '' }, 'Excel herunterladen'),
      h('a', { class: 'btn btn-secondary', href: dlPfad(e.ordner, 'aenderungsprotokoll.xlsx'), download: '' }, 'Protokoll herunterladen'),
      h('button', { class: 'btn btn-secondary', type: 'button', onclick: neuerAbgleich }, 'Neuer Abgleich'));
  }

  // ------------------------------------------------------------------ Warnungen oben (nur wenn etwas nicht stimmt)
  let warnSchluessel = '';
  function banner(farbe, titel, satz) {
    return h('div', { class: 'banner ' + farbe, role: 'alert' }, h('span', { class: 'banner-symbol', 'aria-hidden': 'true' }, '!'),
      h('div', { class: 'banner-inhalt' }, h('div', { class: 'banner-titel' }, titel), h('p', { class: 'banner-text' }, satz)));
  }
  function zeichneWarnungen() {
    const out = [];
    if (state.plan && state.schritt !== 'dateien') {
      const b = (state.schritt === 'fertig' && state.ergebnis && state.ergebnis.bilanz) || state.plan.bilanz || {};
      if (b.stimmt === false) {
        out.push(['bilanz', banner('rot', 'Die Bilanz stimmt nicht',
          `Die im ${istDocx() ? 'Word-Dokument' : 'PDF'} markierten Wörter passen nicht zu den gefundenen Änderungen. Es kann etwas verloren gegangen sein. Bitte das Ergebnis nicht verwenden und den Fehler melden.`)]);
      }
      const luecken = liste(state.plan.luecken);
      if (b.vollstaendig === false || luecken.length) {
        const seiten = [...new Set(luecken.map((l) => l.seite).filter((s) => typeof s === 'number' && s > 0))].sort((x, y) => x - y);
        if (istDocx()) {
          out.push(['luecken-docx', banner('gelb', 'Word-Dokument unvollständig gelesen',
            'An einigen Stellen konnte das Tool nicht alles lesen. Dort kann eine Änderung fehlen. Bitte diese Stellen im Word-Dokument selbst ansehen.')]);
        } else {
          const wo = seiten.length ? `${seiten.length === 1 ? 'Seite' : 'den Seiten'} ${seiten.join(', ')}` : 'einigen Seiten';
          out.push([`luecken${seiten.join(',')}`, banner('gelb', 'PDF unvollständig gelesen',
            `Auf ${wo} konnte das Tool nicht alles lesen. Dort kann eine Änderung fehlen. Bitte diese Seiten im PDF selbst ansehen.`)]);
        }
      }
    }
    const schluessel = out.map((x) => x[0]).join('|');
    if (schluessel === warnSchluessel) return;          // nur neu zeichnen (und ansagen), wenn sich etwas ändert
    warnSchluessel = schluessel;
    $('warn-box').replaceChildren(...out.map((x) => x[1]));
  }

  // ------------------------------------------------------------------ Details für Fachleute (alles Technische an einer Stelle)
  function dl(paare) {
    return h('dl', { class: 'dl-grid' }, paare.filter(Boolean).flatMap(([k, v]) => [h('dt', {}, k), h('dd', {}, hatWort(v))]));
  }
  function tabelle(titel, kopf, zeilen, breit = false) {
    return h('div', { class: 'table-wrap' }, h('table', { class: 'tabelle' + (breit ? ' breit' : ''), 'aria-label': titel },
      h('thead', {}, h('tr', {}, kopf.map((x) => h('th', { scope: 'col' }, x)))),
      h('tbody', {}, zeilen.map((z) => h('tr', {}, z.map((x) => h('td', {}, x)))))));
  }
  const jaNein = (x) => (x === true ? 'ja' : x === false ? 'nein' : 'unbekannt (älterer Abgleich)');

  function zeichneFachleute() {
    const d = $('fachleute');
    d.hidden = !state.plan;
    if (!state.plan || !d.open) return;
    const p = state.plan, e = state.schritt === 'fertig' ? state.ergebnis : null;
    const b = (e && e.bilanz) || p.bilanz || {};
    const l = liste(e && e.aenderungen ? e.aenderungen : p.aenderungen);
    const luecken = liste(p.luecken);
    const warn = liste(e && e.warnungen ? e.warnungen : p.warnungen);
    const teile = [
      h('h3', {}, 'Abgleich'),
      dl([['Lauf-ID', p.run_id], [istDocx() ? 'Word-Dokument' : 'PDF', p.pdf_name], ['Excel', p.excel_name], istDocx() ? null : ['PDF-Seiten', p.seiten],
        ['Archiviert', p.archiviert ? 'ja' : 'nein'], e ? ['Ergebnis-Ordner', e.ordner] : null,
        liste(p.stichprobe).length ? ['Stichprobe', liste(p.stichprobe).join(', ')] : null]),
      h('h3', {}, e ? 'Wort-Bilanz (nach dem Eintragen)' : 'Wort-Bilanz'),
      dl([['Wörter im PDF markiert', b.woerter_im_pdf], ['davon eingetragen', b.woerter_uebernommen], ['davon Vorschlag', b.woerter_pruefen],
        ['davon nicht eingetragen', b.woerter_nicht_uebernommen], ['bewusst ignoriert (Kopf-/Fußzeile)', b.woerter_ignoriert],
        ['Bilanz stimmt', jaNein(b.stimmt)], ['PDF vollständig gelesen', jaNein(b.vollstaendig)], ['Lücken beim Lesen', b.luecken], ...(b.luecken_ausserhalb ? [['Nicht gelesen, aber außerhalb der CDE (Info)', b.luecken_ausserhalb]] : []),
        ['Abweichungen nach dem Eintragen', b.nachkontrolle_abweichungen]]),
      h('h3', {}, 'Lücken: Stellen, die nicht gelesen werden konnten'),
      luecken.length ? h('ul', { class: 'einfach-liste' }, luecken.map((x) => h('li', {}, `${istDocx() ? 'Im Word-Dokument' : `Seite ${hatWort(x.seite)}`} · ${LUECKE_ART[x.art] || text(x.art) || 'Sonstiges'}`,
        text(x.beschreibung) ? ` – ${text(x.beschreibung)}` : '', zahl(x.woerter) ? ` (${pl(zahl(x.woerter), 'Wort', 'Wörter')})` : ''))) : h('p', { class: 'hint' }, 'Keine.'),
      h('h3', {}, 'Hinweise zum Abgleich'),
      warn.length ? h('ul', { class: 'einfach-liste' }, warn.map((x) => h('li', {}, text(x)))) : h('p', { class: 'hint' }, 'Keine.'),
      h('h3', {}, 'Spec'),
      specInfo().liste.length ? dl([['Gewählte Spec', specInfo().erkannt || 'nicht eindeutig'], ['Quelle', { vorgabe: 'von dir gewählt', pdf: 'im PDF erkannt', einzige: 'einzige Spec der CDE' }[specInfo().quelle] || '–'],
        ['Specs der CDE', specInfo().liste.map((x) => `${x.name}${specZeilen(x) ? ' (' + specZeilen(x) + ')' : ''}`).join(' · ')],
        ['Im PDF genannt', liste(p.spec && p.spec.pdf).join(', ') || 'keine']]) : h('p', { class: 'hint' }, 'Die CDE hat keine Spalte mit Spezifikationsnummern.'),
      h('h3', {}, 'Erkannter Aufbau der Excel'),
      liste(p.layout).length ? tabelle('Erkannter Aufbau der Excel', ['Blatt', 'Section', 'Titel', 'Inhalt', 'Modell', 'Sections', 'Sicherheit'],
        liste(p.layout).map((x) => [text(x.blatt), x.section_spalte || '–', x.titel_spalte || '–', x.inhalt_spalte || '–',
          x.modell === 'zeile_pro_section' ? 'eine Zeile je Section' : 'eine Zeile je Absatz', hatWort(x.sections), text(x.sicherheit) || '–']))
        : h('p', { class: 'hint' }, 'Für diesen Abgleich ist kein Aufbau gespeichert.'),
      h('h3', {}, 'Farben im PDF'),
      dl([['Erkannte Legende', LEGENDE[p.legende_modus] || text(p.legende_modus) || 'nicht gespeichert (älterer Abgleich)'],
        ['Sprechblasen am Rand', typeof p.sprechblasen === 'number' ? p.sprechblasen : 'nicht gespeichert (älterer Abgleich)']]),
      liste(p.farben).length ? tabelle('Farben im PDF', ['Farbe', 'Zeichen', 'Art', 'Gruppe', 'Bedeutung'], liste(p.farben).map((f) => {
        const feld = h('span', { class: 'farbe-feld', 'aria-hidden': 'true' });
        if (/^#[0-9a-f]{3,8}$/i.test(text(f.farbe))) feld.style.backgroundColor = f.farbe;     // Farbe des PDFs, keine Dashboard-Farbe
        return [[feld, text(f.farbe)], hatWort(f.zeichen), text(f.art), text(f.klasse) === '-' ? '–' : text(f.klasse), text(f.bedeutung)];
      })) : h('p', { class: 'hint' }, 'Keine Farben gespeichert.'),
      h('h3', {}, `Alle Änderungen (${l.length})`),
      l.length ? tabelle('Alle Änderungen', ['Nr.', 'Status', 'Art', 'Seite', 'Section', 'Stelle', 'Sicherheit', 'Prüfcodes', 'Hinweise', 'Wirkung', 'Jetzt in der Excel'],
        l.map((a) => [text(a.id), (STATUS[a.status] || text(a.status)) + (a.bereich === 'ausserhalb' ? ' (außerhalb der CDE)' : '') + (a.bestaetigt ? ', freigegeben' : ''),
          ART[a.typ] || text(a.typ), istDocx() ? '–' : hatWort(a.seite), text(a.section) || '–', text(a.ziel) || '–',
          `${text(a.konfidenz) || '–'}${a.score ? ` (${Math.round(zahl(a.score) * 100)} %)` : ''}`,
          liste(a.pruefcodes).join(', ') || '–', liste(a.hinweise).join(' · ') || '–',
          [WIRKUNG[a.wirkung] || text(a.wirkung), a.nachkontrolle === 'abweichung' ? 'Nachkontrolle weicht ab' : ''].filter(Boolean).join(', ') || '–',
          kurz(a.excel_text, 160) || '–']), true) : h('p', { class: 'hint' }, 'Im PDF wurden keine Änderungen gefunden.'),
      h('h3', {}, 'Erstellte Excel-Dateien'),
      state.ergebnisse.length ? h('ul', { class: 'ergebnis-liste' }, state.ergebnisse.slice().reverse().map((x) => {
        const bb = x.bilanz || {};
        return h('li', {}, h('div', {}, h('strong', {}, text(x.ordner)),
          h('div', { class: 'meta' }, [fmtZeit(x.erstellt), `${zahl(bb.uebernommen)} eingetragen`, `${liste(x.freigaben).length} freigegeben`,
            `${liste(x.abwahl).length} abgewählt`, bb.stimmt === true ? '' : 'Bilanz stimmt nicht'].filter(Boolean).join(' · '))),
        h('div', { class: 'btn-reihe' }, h('a', { class: 'btn btn-secondary btn-small', href: dlPfad(x.ordner, x.excel), download: '' }, 'Excel'),
          h('a', { class: 'btn btn-secondary btn-small', href: dlPfad(x.ordner, 'aenderungsprotokoll.xlsx'), download: '' }, 'Protokoll')));
      })) : h('p', { class: 'hint' }, 'Noch keine.'),
    ];
    $('fachleute-inhalt').replaceChildren(...teile);
  }

  // ------------------------------------------------------------------ Tastatur (wie „Überarbeiten“ in Word)
  let hilfeZuvor = null;
  function hilfeAuf() {
    const d = $('hilfe');
    if (d.open) return;
    hilfeZuvor = document.activeElement;
    if (typeof d.showModal === 'function') d.showModal(); else d.setAttribute('open', '');
    $('hilfe-zu').focus();
  }
  function hilfeZu() {
    const d = $('hilfe');
    if (typeof d.close === 'function') d.close(); else d.removeAttribute('open');
  }

  function sichtbareKarten() { return [...$('schritt-pruefen').querySelectorAll('.karte')].filter((k) => k.offsetParent !== null); }

  function springe(karte, richtung) {
    const l = sichtbareKarten();
    if (!l.length) return;
    let i = karte ? l.indexOf(karte) + richtung : (richtung > 0 ? 0 : l.length - 1);
    i = Math.max(0, Math.min(l.length - 1, i));
    fokusKarte(l[i]);
  }

  function tasten(ev) {
    if (ev.defaultPrevented || ev.ctrlKey || ev.metaKey || ev.altKey) return;
    if ($('hilfe').open || !$('busy').hidden) return;
    const ziel = ev.target instanceof Element ? ev.target : null;
    if (ziel && ziel.closest('input, textarea, select, [contenteditable="true"]')) return;
    if (ev.key === '?') { ev.preventDefault(); hilfeAuf(); return; }
    if (state.schritt !== 'pruefen' || !state.plan) return;
    const karte = ziel ? ziel.closest('.karte') : null;
    const taste = ev.key.length === 1 ? ev.key.toLowerCase() : ev.key;
    if (taste === 'j' || (taste === 'ArrowDown' && karte)) { ev.preventDefault(); springe(karte, 1); }
    else if (taste === 'k' || (taste === 'ArrowUp' && karte)) { ev.preventDefault(); springe(karte, -1); }
    else if (taste === 'a' || taste === 'n') {
      ev.preventDefault();
      if (!karte) { springe(null, 1); return; }             // erst eine Karte wählen, nie blind entscheiden
      const id = karte.dataset.id;
      if (karte.dataset.art === 'stichprobe') stichprobeEntscheide(id, taste === 'a' ? 'passt' : 'falsch', true);
      else if (karte.classList.contains('kompakt')) { setzeAbwahl(id, taste === 'n'); springe(karte, 1); }
      else entscheide(id, taste === 'a' ? 'ja' : 'nein', true);
    }
  }

  // ------------------------------------------------------------------ Start
  document.addEventListener('DOMContentLoaded', async () => {
    try { const t = localStorage.getItem('cde-theme'); if (t === 'dark' || t === 'light') document.documentElement.setAttribute('data-theme', t); } catch (_) { /* ignorieren */ }
    themeKnopf();
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', themeKnopf);
    $('theme-toggle').addEventListener('click', () => setzeTheme(istDunkel() ? 'light' : 'dark'));
    for (const b of document.querySelectorAll('.schritt')) b.addEventListener('click', () => zeigeSchritt(b.dataset.schritt));
    dropzone('dz-pdf', 'file-pdf', 'pdf');
    dropzone('dz-xlsx', 'file-xlsx', 'xlsx');
    for (const e of ['dragover', 'drop']) window.addEventListener(e, (ev) => ev.preventDefault());   // eine daneben gezogene Datei darf die Seite nicht verlassen
    $('btn-abgleichen').addEventListener('click', abgleichen);
    $('btn-archiv').addEventListener('click', () => {
      state.zeigeArchiv = !state.zeigeArchiv;
      $('btn-archiv').setAttribute('aria-pressed', state.zeigeArchiv ? 'true' : 'false');
      $('btn-archiv').textContent = state.zeigeArchiv ? 'Archivierte ausblenden' : 'Archivierte anzeigen';
      ladeZuletzt();
    });
    $('btn-zuletzt-alle').addEventListener('click', () => { state.zuletztAlle = !state.zuletztAlle; ladeZuletzt(); });
    $('btn-alle-ja').addEventListener('click', () => {
      for (const a of vorschlaege()) { state.entscheid.set(a.id, 'ja'); state.pauschal.add(a.id); aktualisiereKarte(a.id); }
      zeichneFortschritt();
    });
    $('btn-tasten').addEventListener('click', hilfeAuf);
    $('btn-hilfe').addEventListener('click', hilfeAuf);
    $('hilfe-zu').addEventListener('click', hilfeZu);
    $('hilfe').addEventListener('close', () => {          // Fokus zurück, aber nur, wenn ihn inzwischen niemand woanders hingesetzt hat
      const frei = !document.activeElement || document.activeElement === document.body || $('hilfe').contains(document.activeElement);
      if (frei && hilfeZuvor && document.contains(hilfeZuvor)) hilfeZuvor.focus();
      hilfeZuvor = null;
    });
    let suchZeit = null;
    $('suche-auto').addEventListener('input', (ev) => {
      clearTimeout(suchZeit);
      suchZeit = setTimeout(() => { state.suche = ev.target.value; state.autoLimit = AUTO_SEITE; zeichneAuto(); }, 150);
    });
    $('btn-auto-alle').addEventListener('click', () => { state.autoLimit = Infinity; zeichneAuto(); });
    $('btn-excel').addEventListener('click', excelErstellen);
    $('btn-spec-andere').addEventListener('click', specAndere);
    $('btn-zurueck-pruefen').addEventListener('click', () => zeigeSchritt('pruefen'));
    $('fachleute').addEventListener('toggle', zeichneFachleute);
    $('toast-schliessen').addEventListener('click', fehlerSchliessen);

    document.addEventListener('click', (ev) => {
      const icon = ev.target.closest('.info-icon');
      if (icon) { ev.preventDefault(); ev.stopPropagation(); oeffnePopup(icon); return; }
      if (popup && !popup.contains(ev.target)) schliessePopup(false);
      for (const m of document.querySelectorAll('details.menu[open]')) if (!m.contains(ev.target)) m.open = false;
    });
    document.addEventListener('keydown', (ev) => {
      if (ev.key === 'Escape') {
        const menu = document.querySelector('details.menu[open]');
        if (popup) schliessePopup();
        else if (menu) { menu.open = false; menu.querySelector('summary').focus(); }
        else if (!$('toast-fehler-box').hidden) fehlerSchliessen();
        return;
      }
      tasten(ev);
    });
    window.addEventListener('scroll', setzePopup, true);
    window.addEventListener('resize', setzePopup);
    window.addEventListener('unhandledrejection', (ev) => { ev.preventDefault(); meldeFehler(ev.reason); });
    window.addEventListener('error', (ev) => { console.error(ev.error || ev.message); toastFehler(TXT_UNERWARTET); });

    zeigeSchritt('dateien', false);
    try {
      const g = await api('/health');
      $('port-text').textContent = g.port;
      $('pill-text').textContent = `Lokal · Port ${g.port}`;
    } catch (_) {
      $('pill-text').textContent = 'Server nicht erreichbar';
      $('status-pill').classList.add('bad');
      toastFehler(TXT_SERVER);
    }
  });
})();
