/* CDE-Änderungsübernahme – Dashboard (reines JavaScript, keine Abhängigkeiten, keine externen Quellen).
   Regel 8: reine Darstellung. Das Dashboard zeigt, was die API liefert; es rechnet nichts Fachliches nach.
   Alle Texte aus Daten werden über textContent gesetzt (Fremdtext wird nie als HTML eingesetzt).
   Alle Aufrufe gehen an relative /api/-Pfade (keine festen Adressen oder Ports). */
(() => {
  'use strict';

  // ------------------------------------------------------------------ Beschriftungen (einfaches Deutsch)
  const STATUS = {
    uebernommen: { text: 'Übernommen', farbe: 'green', rang: 2 },
    pruefen: { text: 'Zu prüfen', farbe: 'amber', rang: 0 },
    nicht_uebernommen: { text: 'Nicht übernommen', farbe: 'red', rang: 1 },
  };
  const ART = {
    text_ersetzt: 'Text ersetzt', text_eingefuegt: 'Text eingefügt', text_geloescht: 'Text entfällt',
    absatz_neu: 'Absatz neu', absatz_geloescht: 'Absatz entfällt',
    section_neu: 'Section neu', section_geloescht: 'Section entfällt',
  };
  const KONF = { hoch: { farbe: 'green', rang: 2 }, mittel: { farbe: 'amber', rang: 1 }, niedrig: { farbe: 'red', rang: 0 } };
  const LUECKE_ART = {
    scan_seite: 'Scan-Seite (kein Text lesbar)', mehrspaltig: 'Text in mehreren Spalten', tabelle: 'Tabelle', bild: 'Bild',
    unbekannte_farbe: 'Unbekannte Farbe', nicht_lesbarer_text: 'Nicht lesbarer Text', sonstiges: 'Sonstiges',
  };
  const WIRKUNG = {
    ok: { text: 'Wirkung geprüft', lang: 'Die Excel hat sich nachweislich wie geplant geändert.', farbe: 'green', schlecht: false },
    keine: { text: 'Keine Wirkung', lang: 'Die Änderung hat an der Excel nichts bewirkt. Sie gilt deshalb nicht als übernommen.', farbe: 'red', schlecht: true },
    abweichung: { text: 'Weicht vom Plan ab', lang: 'Die Excel sieht nach dem Eintragen anders aus als geplant.', farbe: 'red', schlecht: true },
  };
  const LAUF_STATUS = {
    analysiert: { text: 'Analysiert, noch nichts übernommen', farbe: 'blue', offen: true },
    uebernommen: { text: 'Übernommen', farbe: 'green', offen: true },
    fehlgeschlagen: { text: 'Analyse fehlgeschlagen', farbe: 'red', offen: false },
    beschaedigt: { text: 'Beschädigt', farbe: 'red', offen: false },
    neuere_version: { text: 'Aus neuerer Programmversion', farbe: 'amber', offen: false },
    angelegt: { text: 'Noch nicht analysiert', farbe: 'grey', offen: false },
  };
  const INFO = {
    bilanz: { titel: 'Was ist die Bilanz?', text: 'Das Tool zählt alle im PDF markierten Wörter. Jedes davon muss übernommen, zu prüfen oder mit Grund nicht übernommen sein. Stimmt die Summe nicht, ist etwas verloren gegangen.' },
    luecken: { titel: 'Was sind Lücken?', text: 'Lücken sind Stellen im PDF, die das Tool nicht sicher lesen konnte, zum Beispiel Scan-Seiten, Tabellen oder Text in mehreren Spalten. Dort kann eine Änderung fehlen. Bitte diese Stellen im PDF selbst ansehen.' },
    sicherheit: { titel: 'Sicherheit', text: 'Wie sicher das Tool bei der Stelle in der Excel ist. Hoch: eindeutig. Mittel oder niedrig: bitte genau ansehen. Unsichere Treffer trägt das Tool nie von selbst ein.' },
    auswahl: { titel: 'Vorschläge freigeben', text: 'Nur Vorschläge mit Häkchen werden beim Erzeugen der Excel eingetragen. Sichere Treffer (Status „Übernommen“) trägt das Tool automatisch ein. Das lässt sich hier noch nicht abwählen. Die Original-Excel bleibt immer unverändert.' },
    vorher_nachher: { titel: 'Vorher / Nachher', text: 'Links steht, was jetzt in der Excel an dieser Stelle steht. Rechts steht, was sich im PDF ändert: durchgestrichen ist der alte Text, blau der neue.' },
    wirkung: { titel: 'Wirkung', text: 'Nach dem Eintragen prüft das Tool, ob die Excel sich wirklich wie geplant geändert hat. „Keine Wirkung“ heißt: Es hat sich nichts geändert. Dann gilt die Änderung nicht als übernommen.' },
    layout: { titel: 'Erkannter Aufbau', text: 'So hat das Tool die Tabellenblätter verstanden: in welcher Spalte die Section-Nummer, der Titel und der Inhalt stehen. Weicht das stark von der Wirklichkeit ab, bitte melden.' },
    archiv: { titel: 'Archivieren', text: 'Archivierte Läufe bleiben vollständig auf diesem Rechner erhalten (Eingabe, Analyse, Ergebnisse). Sie erscheinen nur nicht mehr in der Liste. Mit „Archivierte anzeigen“ findest du sie wieder und kannst sie wiederherstellen.' },
    farben: { titel: 'Farben im PDF', text: 'Zeigt, welche Farben im geöffneten PDF vorkommen und wie sie gedeutet werden. Passt etwas nicht, lassen sich die Schwellen in config/farben.yaml anpassen.' },
  };
  const TXT_SERVER = 'Der Server antwortet nicht. Läuft das Dashboard noch?';
  const TXT_UNERWARTET = 'Im Dashboard ist etwas schiefgelaufen. Bitte die Seite neu laden. Deine Läufe sind gespeichert.';
  const SEITENGROESSE = 200;

  // ------------------------------------------------------------------ Zustand und Helfer
  const neuerFilter = () => ({ suche: '', status: 'alle', warn: false, typ: '', seite: '', section: '', konfidenz: '', sort: 'nr' });
  const state = {
    plan: null, ergebnis: null, ergebnisse: [], f: neuerFilter(), freig: new Set(), offen: new Set(), limit: SEITENGROESSE,
    files: { pdf: null, xlsx: null }, zeigeArchiv: false,
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
    for (const k of kinder.flat()) if (k != null && k !== false) el.append(k.nodeType ? k : document.createTextNode(String(k)));
    return el;
  }

  const liste = (x) => (Array.isArray(x) ? x : []);
  const text = (x) => (typeof x === 'string' ? x : x == null ? '' : String(x));
  const zahl = (x) => (typeof x === 'number' && isFinite(x) ? x : 0);
  const pl = (n, ein, mehr) => `${n} ${n === 1 ? ein : mehr}`;
  const kurz = (t, n) => { t = text(t); return t.length > n ? t.slice(0, n - 1) + '…' : t; };
  const natuerlich = (a, b) => text(a).localeCompare(text(b), 'de', { numeric: true, sensitivity: 'base' });
  const hatWort = (x) => (x === 0 || x ? String(x) : '–');

  function fmtZeit(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})/.exec(text(iso));
    return m ? `${m[3]}.${m[2]}.${m[1]} ${m[4]}:${m[5]}` : '';
  }

  function infoIcon(key) {
    const t = INFO[key] ? INFO[key].titel : '';
    return h('button', { class: 'info-icon', type: 'button', 'data-info': key, 'aria-haspopup': 'dialog', 'aria-expanded': 'false', 'aria-label': `Erklärung: ${t}` }, 'i');
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
    const box = $('toast-fehler-box');
    box.hidden = false;
    $('toast-fehler').textContent = '';
    setTimeout(() => { $('toast-fehler').textContent = msg; }, 30);
  }
  function fehlerSchliessen() { $('toast-fehler-box').hidden = true; $('toast-fehler').textContent = ''; }
  function meldeFehler(e) {
    if (e instanceof ApiFehler) {
      const extra = e.daten && e.daten.run_id ? ' Der Lauf bleibt im Verlauf gespeichert.' : '';
      toastFehler(e.message + extra);
    } else {
      console.error(e);
      toastFehler(TXT_UNERWARTET);
    }
  }

  // ------------------------------------------------------------------ Warte-Anzeige (Fokus + alles andere gesperrt)
  let busyZuvor = null;
  function busy(an, msg = 'Bitte warten …') {
    const b = $('busy');
    if (an) {
      $('busy-text').textContent = msg;
      if (b.hidden) {
        busyZuvor = document.activeElement;
        for (const id of ['kopf', 'seiten-tabs', 'main', 'fuss']) $(id).inert = true;
        b.hidden = false;
        $('busy-card').focus();
      }
    } else if (!b.hidden) {
      b.hidden = true;
      for (const id of ['kopf', 'seiten-tabs', 'main', 'fuss']) $(id).inert = false;
      if (busyZuvor && document.contains(busyZuvor) && typeof busyZuvor.focus === 'function') busyZuvor.focus();
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
  const lauf = (id) => encodeURIComponent(id);

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

  // ------------------------------------------------------------------ Tabs
  const TABS = ['lauf', 'pruefen', 'verlauf', 'farben'];
  function zeigeTab(name, fokus = false) {
    for (const b of document.querySelectorAll('.page-tab')) {
      const aktiv = b.dataset.tab === name;
      b.setAttribute('aria-selected', aktiv ? 'true' : 'false');
      b.tabIndex = aktiv ? 0 : -1;
      if (aktiv && fokus) b.focus();
    }
    for (const s of document.querySelectorAll('.tab')) s.hidden = s.id !== 'tab-' + name;
    if (name === 'verlauf') ladeVerlauf();
  }
  function tabTaste(ev) {
    const i = TABS.indexOf(ev.target.dataset.tab);
    if (i < 0) return;
    let n = null;
    if (ev.key === 'ArrowRight') n = (i + 1) % TABS.length;
    else if (ev.key === 'ArrowLeft') n = (i + TABS.length - 1) % TABS.length;
    else if (ev.key === 'Home') n = 0;
    else if (ev.key === 'End') n = TABS.length - 1;
    if (n === null) return;
    ev.preventDefault();
    zeigeTab(TABS[n], true);
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

  // ------------------------------------------------------------------ Upload
  function dateiGewaehlt(art, datei) {
    const dz = art === 'pdf' ? $('dz-pdf') : $('dz-xlsx');
    const okPdf = art === 'pdf' && /\.pdf$/i.test(datei.name);
    const okXlsx = art === 'xlsx' && /\.(xlsx|xlsm)$/i.test(datei.name);
    if (!okPdf && !okXlsx) {
      toastFehler(art === 'pdf' ? 'Bitte eine PDF-Datei wählen.' : 'Bitte eine .xlsx-Datei wählen (alte .xls-Dateien bitte in Excel als .xlsx speichern).');
      return;
    }
    state.files[art] = datei;
    dz.classList.add('filled');
    $(art === 'pdf' ? 'dz-pdf-sub' : 'dz-xlsx-sub').textContent = `✓ ${datei.name}`;
    $('btn-analyse').disabled = !(state.files.pdf && state.files.xlsx);
  }

  function dropzone(id, inputId, art) {
    const dz = $(id), input = $(inputId);
    input.addEventListener('change', () => input.files[0] && dateiGewaehlt(art, input.files[0]));
    ['dragenter', 'dragover'].forEach((e) => dz.addEventListener(e, (ev) => { ev.preventDefault(); dz.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((e) => dz.addEventListener(e, (ev) => { ev.preventDefault(); dz.classList.remove('over'); }));
    dz.addEventListener('drop', (ev) => { const f = ev.dataTransfer.files[0]; if (f) dateiGewaehlt(art, f); });
    dz.addEventListener('keydown', (ev) => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); input.click(); } });
  }

  async function analysieren() {
    const fd = new FormData();
    fd.append('pdf', state.files.pdf);
    fd.append('excel', state.files.xlsx);
    busy(true, 'Analysiere das PDF und die Excel … (bei vielen Seiten kann das einen Moment dauern)');
    try {
      const plan = await api('/lauf', { method: 'POST', body: fd });
      oeffne(plan);
      toast('Analyse fertig. Es wurde noch nichts in die Excel geschrieben.');
    } catch (e) { meldeFehler(e); } finally { busy(false); }
  }

  function oeffne(plan) {
    state.plan = plan;
    state.ergebnis = null;
    state.ergebnisse = liste(plan.ergebnisse);
    state.f = neuerFilter();
    state.freig = new Set();
    state.offen = new Set();
    state.limit = SEITENGROESSE;
    zeichne();
    zeigeTab('pruefen');
  }

  // ------------------------------------------------------------------ Hilfen zur Darstellung der Änderungen
  const aktuelleAenderungen = () => liste(state.ergebnis ? state.ergebnis.aenderungen : state.plan && state.plan.aenderungen);
  const aktuelleBilanz = () => (state.ergebnis && state.ergebnis.bilanz) || (state.plan && state.plan.bilanz) || {};
  const artText = (typ) => ART[typ] || text(typ).replace(/_/g, ' ');
  const wirkungSchlecht = (a) => !!(WIRKUNG[a.wirkung] && WIRKUNG[a.wirkung].schlecht);
  const nachkontrolleAbweichung = (a) => a.nachkontrolle === 'abweichung';
  const hatProblem = (a) => wirkungSchlecht(a) || nachkontrolleAbweichung(a);
  const hatHinweis = (a) => liste(a.hinweise).length > 0;
  const auffaellig = (a) => hatProblem(a) || hatHinweis(a);

  function kachel(wert, label, farbe = '', istText = false) {
    return h('div', { class: 'stat-card ' + farbe }, h('span', { class: 'stat-value' + (istText ? ' text' : '') }, wert), h('span', { class: 'stat-label' }, label));
  }

  function banner(farbe, symbol, titel, kinder, rolle) {
    return h('div', { class: 'banner ' + farbe, role: rolle || 'region', 'aria-label': titel },
      h('span', { class: 'banner-symbol', 'aria-hidden': 'true' }, symbol),
      h('div', { class: 'banner-inhalt' }, h('div', { class: 'banner-titel' }, titel), kinder));
  }

  // ------------------------------------------------------------------ Hauptdarstellung
  function zeichne() {
    const p = state.plan;
    $('leer-pruefen').hidden = !!p;
    $('pruefen-inhalt').hidden = !p;
    zeichneFarben();
    if (!p) return;
    const b = aktuelleBilanz();
    $('run-titel').textContent = `${text(p.pdf_name)}  →  ${text(p.excel_name)}`;
    $('run-sub').textContent = `Lauf ${text(p.run_id)} · ${pl(zahl(p.seiten), 'PDF-Seite', 'PDF-Seiten')}`;
    $('archiv-hinweis').hidden = !p.archiviert;

    zeichneBanner(b);
    $('stats').replaceChildren(
      kachel(zahl(b.aenderungen), 'Änderungen erkannt'),
      kachel(zahl(b.uebernommen), 'Übernommen', 'green'),
      kachel(zahl(b.pruefen), 'Zu prüfen', 'amber'),
      kachel(zahl(b.nicht_uebernommen), 'Nicht übernommen', 'red'),
      gelesenKachel(b));
    zeichneBilanzDetail(b);

    const warn = liste(state.ergebnis ? state.ergebnis.warnungen : p.warnungen);
    const w = $('warnungen');
    w.hidden = !warn.length;
    w.replaceChildren(h('strong', {}, 'Hinweise zum Lauf'), h('ul', {}, warn.map((x) => h('li', {}, text(x)))));

    zeichneLayout(p);
    zeichneErgebnisse();
    fuelleFilter();
    zeichneChips();
    zeichneTabelle();
    zeichneAuswahl();
  }

  function unvollstaendig(b) { return b.vollstaendig === false || (liste(state.plan && state.plan.luecken).length > 0); }

  function gelesenKachel(b) {
    if (unvollstaendig(b)) return kachel('Unvollständig', 'Gelesen', 'amber', true);
    if (b.vollstaendig === true) return kachel('Vollständig', 'Gelesen', b.stimmt === true && !zahl(b.nachkontrolle_abweichungen) ? 'green' : 'grey', true);
    return kachel('Unbekannt', 'Gelesen', 'grey', true);
  }

  function zeichneBanner(b) {
    const box = [];
    const luecken = liste(state.plan.luecken);
    const abw = zahl(b.nachkontrolle_abweichungen);
    const stimmt = b.stimmt === true;
    const unvoll = unvollstaendig(b);

    if (!stimmt) {
      box.push(banner('rot', '!', 'Die Bilanz stimmt nicht', h('p', { class: 'banner-text' },
        `Die Wörter im PDF (${hatWort(b.woerter_im_pdf)} markiert) und die Summe aus „übernommen“, „zu prüfen“, „nicht übernommen“ und „ignoriert“ passen nicht zusammen. ` +
        'Etwas kann verloren gegangen sein. Bitte dieses Ergebnis nicht verwenden und den Fehler melden.'), 'alert'));
    }
    if (unvoll) {
      const n = luecken.length || zahl(b.luecken);
      const kinder = [h('p', { class: 'banner-text' },
        `Das Tool konnte nicht alles im PDF lesen (${pl(n, 'Lücke', 'Lücken')}). An diesen Stellen kann eine Änderung fehlen. ` +
        'Bitte diese Stellen im PDF selbst prüfen, bevor du dich auf das Ergebnis verlässt.')];
      if (luecken.length) {
        kinder.push(h('ul', { class: 'luecken-liste', 'aria-label': 'Lücken im PDF' }, luecken.map((l) => h('li', {},
          h('span', { class: 'seite' }, l.seite ? `Seite ${l.seite}` : 'Seite unbekannt'), '· ',
          h('span', { class: 'art' }, LUECKE_ART[l.art] || text(l.art).replace(/_/g, ' ') || 'Sonstiges'),
          text(l.beschreibung) ? '– ' + text(l.beschreibung) : '',
          zahl(l.woerter) ? h('span', { class: 'woerter' }, ` (${pl(zahl(l.woerter), 'Wort', 'Wörter')})`) : ''))));
      }
      box.push(banner('gelb', '!', 'Unvollständig gelesen', kinder, 'alert'));
    } else if (b.vollstaendig === undefined || b.vollstaendig === null) {
      box.push(banner('gelb', '?', 'Vollständigkeit nicht bekannt', h('p', { class: 'banner-text' },
        'Dieser Lauf stammt aus einer älteren Programmversion. Ob das PDF vollständig gelesen wurde, ist dort nicht festgehalten.')));
    }
    if (abw) {
      box.push(banner('rot', '!', `${pl(abw, 'Änderung weicht', 'Änderungen weichen')} ab`, [h('p', { class: 'banner-text' },
        'Bei der Nachkontrolle hat die Excel sich nicht wie geplant geändert (keine Wirkung oder Abweichung). Diese Zeilen sind unten markiert.'),
        h('button', { class: 'btn btn-secondary btn-small', type: 'button', onclick: nurProbleme }, 'Nur diese zeigen')], 'alert'));
    }
    if (stimmt && b.vollstaendig === true && !abw) {
      box.push(banner('gruen', '✓', 'Bilanz stimmt', h('p', { class: 'banner-text' },
        `Alle ${hatWort(b.woerter_im_pdf)} markierten Wörter im PDF sind übernommen, zu prüfen oder mit Grund aufgelistet. Das PDF wurde vollständig gelesen. Nichts ging still verloren.`), 'status'));
    }
    $('banner-box').replaceChildren(...box);
  }

  function zeichneBilanzDetail(b) {
    const zeilen = [
      ['Wörter im PDF markiert', b.woerter_im_pdf], ['davon übernommen', b.woerter_uebernommen], ['davon zu prüfen', b.woerter_pruefen],
      ['davon nicht übernommen', b.woerter_nicht_uebernommen], ['bewusst ignoriert', b.woerter_ignoriert],
      ['Lücken beim Lesen', b.luecken],
    ];
    $('bilanz-dl').replaceChildren(...zeilen.flatMap(([k, v]) => [h('dt', {}, k), h('dd', {}, hatWort(v))]));
  }

  function zeichneLayout(p) {
    const box = $('layout-inhalt');
    const zeilen = liste(p.layout);
    box.replaceChildren(
      h('p', { class: 'hint' }, 'So hat das Tool den Aufbau je Tabellenblatt verstanden. Weicht das stark von der Wirklichkeit ab, bitte melden. ', infoIcon('layout')),
      zeilen.length ? h('div', { class: 'table-wrap' }, h('table', { class: 'changes', 'aria-label': 'Erkannter Aufbau der Excel' },
        h('thead', {}, h('tr', {}, ['Blatt', 'Section-Spalte', 'Titel-Spalte', 'Inhalt-Spalte', 'Modell', 'Sections'].map((x) => h('th', { scope: 'col' }, x)))),
        h('tbody', {}, zeilen.map((l) => h('tr', {}, [l.blatt, l.section_spalte || '–', l.titel_spalte || '–', l.inhalt_spalte || '–',
          l.modell === 'zeile_pro_section' ? 'eine Zeile je Section' : 'eine Zeile je Absatz', hatWort(l.sections)].map((x) => h('td', {}, text(x)))))))) :
        h('p', { class: 'hint' }, 'Für diesen Lauf ist kein Aufbau gespeichert.'));
  }

  // ------------------------------------------------------------------ Ergebnisse (neu + frühere)
  function ergebnisNummer(ordner) { const m = /ergebnis_(\d+)/.exec(text(ordner)); return m ? m[1] : text(ordner); }
  function downloadLink(ordner, name, label, primaer) {
    return h('a', { class: 'btn ' + (primaer ? 'btn-primary' : 'btn-secondary'), href: `/api/lauf/${lauf(state.plan.run_id)}/download/${encodeURIComponent(ordner)}/${encodeURIComponent(name)}` }, label);
  }

  function zeichneErgebnisse() {
    const e = state.ergebnis, box = $('ergebnis-box');
    box.hidden = !e;
    if (e) {
      const b = e.bilanz || {};
      const schlecht = b.stimmt !== true || zahl(b.nachkontrolle_abweichungen) > 0;
      box.className = 'box ' + (schlecht ? 'ergebnis-schlecht' : 'ergebnis-neu');
      const hinweisText = schlecht
        ? 'Es wurde ein Ergebnis angelegt, aber bei der Prüfung ist etwas aufgefallen (siehe oben). Bitte vor der Verwendung genau ansehen.'
        : 'Die Original-Excel wurde nicht verändert. Das Protokoll listet jede Änderung mit alt/neu, Quelle im PDF und Stelle in der Excel.';
      box.replaceChildren(h('h3', {}, schlecht ? `Ergebnis ${ergebnisNummer(e.ordner)} erstellt – bitte prüfen` : `Fertig – Ergebnis ${ergebnisNummer(e.ordner)} liegt bereit`),
        h('p', { class: 'hint' }, hinweisText),
        h('div', { class: 'btn-reihe' }, downloadLink(e.ordner, e.excel, 'Aktualisierte Excel herunterladen', true),
          downloadLink(e.ordner, 'aenderungsprotokoll.xlsx', 'Änderungsprotokoll herunterladen', false)));
    }
    const fr = $('fruehere-box');
    const andere = state.ergebnisse.filter((x) => !(e && x.ordner === e.ordner)).slice().reverse();
    fr.hidden = !andere.length;
    if (!andere.length) return;
    const titel = e ? 'Frühere Ergebnisse dieses Laufs' : `Dieser Lauf hat schon ${pl(andere.length, 'Ergebnis', 'Ergebnisse')}`;
    fr.replaceChildren(h('h3', {}, titel),
      e ? '' : h('p', { class: 'hint' }, 'Du kannst die vorhandenen Ergebnisse unten herunterladen. Ein weiteres Ergebnis ist möglich, die früheren bleiben erhalten.'),
      h('ul', { class: 'ergebnis-liste' }, andere.map((x) => {
        const bb = x.bilanz || {};
        const flags = [];
        if (bb.stimmt !== true) flags.push(h('span', { class: 'badge red' }, 'Bilanz stimmt nicht'));
        if (bb.vollstaendig === false) flags.push(h('span', { class: 'badge amber' }, 'Unvollständig gelesen'));
        if (zahl(bb.nachkontrolle_abweichungen)) flags.push(h('span', { class: 'badge red' }, pl(zahl(bb.nachkontrolle_abweichungen), 'Abweichung', 'Abweichungen')));
        return h('li', {},
          h('div', {}, h('strong', {}, `Ergebnis ${ergebnisNummer(x.ordner)}`),
            h('div', { class: 'meta' }, [fmtZeit(x.erstellt), `${pl(liste(x.freigaben).length, 'Vorschlag', 'Vorschläge')} freigegeben`, `${zahl(bb.uebernommen)} übernommen`].filter(Boolean).join(' · ')),
            flags.length ? h('div', {}, flags) : ''),
          h('div', { class: 'btn-reihe' }, downloadLink(x.ordner, x.excel, 'Excel', false), downloadLink(x.ordner, 'aenderungsprotokoll.xlsx', 'Protokoll', false)));
      })));
  }

  // ------------------------------------------------------------------ Filter, Suche, Sortierung
  function fuelleSelect(id, alle, optionen) {
    const sel = $(id), alt = sel.value;
    sel.replaceChildren(h('option', { value: '' }, alle), ...optionen.map(([v, t]) => h('option', { value: v }, t)));
    sel.value = optionen.some(([v]) => v === alt) ? alt : '';
  }
  function fuelleFilter() {
    const l = aktuelleAenderungen();
    const typen = [...new Set(l.map((a) => a.typ))].sort((a, b) => natuerlich(artText(a), artText(b)));
    const seiten = [...new Set(l.map((a) => a.seite))].filter((s) => s != null).sort((a, b) => a - b);
    const sections = [...new Set(l.map((a) => a.section).filter(Boolean))].sort(natuerlich);
    fuelleSelect('f-typ', 'Alle Arten', typen.map((t) => [t, artText(t)]));
    fuelleSelect('f-seite', 'Alle Seiten', seiten.map((s) => [String(s), `Seite ${s}`]));
    fuelleSelect('f-section', 'Alle Sections', [['__ohne', 'Ohne Section'], ...sections.map((s) => [s, s])]);
    fuelleSelect('f-konfidenz', 'Alle', ['hoch', 'mittel', 'niedrig'].filter((k) => l.some((a) => a.konfidenz === k)).map((k) => [k, k]));
    $('f-suche').value = state.f.suche;
    $('f-typ').value = state.f.typ; $('f-seite').value = state.f.seite; $('f-section').value = state.f.section;
    $('f-konfidenz').value = state.f.konfidenz; $('f-sort').value = state.f.sort;
  }

  function zeichneChips() {
    const l = aktuelleAenderungen();
    const n = (s) => l.filter((a) => a.status === s).length;
    const zeilen = [['alle', `Alle (${l.length})`], ['uebernommen', `Übernommen (${n('uebernommen')})`],
      ['pruefen', `Zu prüfen (${n('pruefen')})`], ['nicht_uebernommen', `Nicht übernommen (${n('nicht_uebernommen')})`]];
    const fr = $('filter-row');
    const fokusIdx = [...fr.children].indexOf(document.activeElement);
    const chips = zeilen.map(([k, t]) => h('button', { class: 'chip', type: 'button', 'aria-pressed': state.f.status === k ? 'true' : 'false',
      onclick: () => { state.f.status = k; neuFiltern(); } }, t));
    chips.push(h('button', { class: 'chip', type: 'button', 'aria-pressed': state.f.warn ? 'true' : 'false',
      onclick: () => { state.f.warn = !state.f.warn; neuFiltern(); } }, `Mit Hinweis oder Abweichung (${l.filter(auffaellig).length})`));
    fr.replaceChildren(...chips);
    if (fokusIdx >= 0 && chips[fokusIdx]) chips[fokusIdx].focus();
    const offen = state.ergebnis ? 0 : l.filter((a) => a.status === 'pruefen').length;
    $('pruef-hinweis').hidden = !offen;
    $('pruef-hinweis-titel').textContent = `${pl(offen, 'Vorschlag wartet', 'Vorschläge warten')} auf deine Prüfung`;
  }

  function neuFiltern() { state.limit = SEITENGROESSE; zeichneChips(); zeichneTabelle(); zeichneAuswahl(); }
  function nurProbleme() { state.f = { ...neuerFilter(), warn: true }; fuelleFilter(); neuFiltern(); }

  function gefiltert() {
    const f = state.f;
    const tokens = f.suche.toLowerCase().split(/\s+/).filter(Boolean);
    const l = aktuelleAenderungen().filter((a) => {
      if (f.status !== 'alle' && a.status !== f.status) return false;
      if (f.warn && !auffaellig(a)) return false;
      if (f.typ && a.typ !== f.typ) return false;
      if (f.seite && String(a.seite) !== f.seite) return false;
      if (f.section === '__ohne' ? !!a.section : (f.section && a.section !== f.section)) return false;
      if (f.konfidenz && a.konfidenz !== f.konfidenz) return false;
      if (tokens.length) {
        const heu = [a.alt, a.neu, a.section, a.ziel, a.excel_text].map(text).join('\n').toLowerCase();
        if (!tokens.every((t) => heu.includes(t))) return false;
      }
      return true;
    });
    const rang = { nr: (x, y) => natuerlich(x.id, y.id),
      seite: (x, y) => zahl(x.seite) - zahl(y.seite),
      section: (x, y) => natuerlich(x.section || '', y.section || ''),
      status: (x, y) => ((STATUS[x.status] || {}).rang ?? 9) - ((STATUS[y.status] || {}).rang ?? 9),
      konfidenz: (x, y) => ((KONF[x.konfidenz] || {}).rang ?? 9) - ((KONF[y.konfidenz] || {}).rang ?? 9),
      auffaellig: (x, y) => (hatProblem(y) - hatProblem(x)) || (hatHinweis(y) - hatHinweis(x)) };
    const vgl = rang[f.sort] || rang.nr;
    return l.slice().sort((x, y) => vgl(x, y) || natuerlich(x.id, y.id));
  }

  // ------------------------------------------------------------------ Tabelle
  function warnZeilen(a) {
    const out = [];
    const w = WIRKUNG[a.wirkung];
    if (w && w.schlecht) out.push(h('div', { class: 'warn-zeile rot' }, h('span', { class: 'sym', 'aria-hidden': 'true' }, '⚠'), h('span', {}, `${w.text}: ${w.lang}`)));
    if (nachkontrolleAbweichung(a) && !(w && w.schlecht)) out.push(h('div', { class: 'warn-zeile rot' }, h('span', { class: 'sym', 'aria-hidden': 'true' }, '⚠'), h('span', {}, 'Nachkontrolle: Die Excel weicht von dem ab, was geplant war.')));
    const hin = liste(a.hinweise);
    for (const x of hin.slice(0, 2)) out.push(h('div', { class: 'warn-zeile' }, h('span', { class: 'sym', 'aria-hidden': 'true' }, '⚠'), h('span', {}, h('span', { class: 'sr-only' }, 'Hinweis: '), text(x))));
    if (hin.length > 2) out.push(h('div', { class: 'warn-mehr' }, `+ ${pl(hin.length - 2, 'weiterer Hinweis', 'weitere Hinweise')} (Zeile öffnen)`));
    return out;
  }

  function statusZelle(a) {
    const s = STATUS[a.status] || { text: text(a.status) || 'Unbekannt', farbe: 'grey' };
    const w = WIRKUNG[a.wirkung];
    const teile = [h('span', { class: 'badge ' + s.farbe }, s.text)];
    if (w && w.schlecht) teile.push(h('span', { class: 'badge red' }, w.text));
    else if (nachkontrolleAbweichung(a)) teile.push(h('span', { class: 'badge red' }, 'Abweichung'));
    if (w && !w.schlecht) teile.push(h('span', { class: 'badge green' }, w.text));
    return teile;
  }

  function konfZelle(a) {
    const k = KONF[a.konfidenz];
    return h('span', { class: 'konf' }, h('span', { class: 'dot ' + (k ? k.farbe : 'grey'), 'aria-hidden': 'true' }), text(a.konfidenz) || '–');
  }

  function zeile(a) {
    const kannFreigeben = a.status === 'pruefen' && !state.ergebnis;
    const chk = kannFreigeben ? h('label', { class: 'chk-ziel' }, h('input', { type: 'checkbox', 'aria-label': `Vorschlag ${a.id} übernehmen`,
      onclick: (ev) => ev.stopPropagation(),
      onchange: (ev) => { ev.target.checked ? state.freig.add(a.id) : state.freig.delete(a.id); zeichneAuswahl(); } })) : '';
    if (chk) chk.firstChild.checked = state.freig.has(a.id);
    const offen = state.offen.has(a.id);
    const tr = h('tr', { class: 'row' + (hatProblem(a) ? ' bad' : hatHinweis(a) ? ' warn' : ''), tabindex: '0', 'data-id': a.id,
      'aria-expanded': offen ? 'true' : 'false', 'aria-controls': 'detail-' + a.id,
      onclick: (ev) => { if (!ev.target.closest('label, input, a, button')) schalteDetail(tr, a); },
      onkeydown: (ev) => zeilenTaste(ev, tr, a) },
      h('td', { class: 'col-chk' }, chk), h('td', {}, text(a.id)), h('td', {}, hatWort(a.seite)), h('td', {}, artText(a.typ)), h('td', {}, a.section || '–'),
      h('td', { class: 'zeile-text' },
        a.alt ? h('span', { class: 'alt' }, kurz(a.alt, 70)) : '', a.alt && a.neu ? h('span', { class: 'pfeil', 'aria-label': 'ersetzt durch' }, '→') : '',
        a.neu ? h('span', { class: 'neu' }, kurz(a.neu, 70)) : '', warnZeilen(a)),
      h('td', { class: 'col-opt' }, a.ziel || '–'),
      h('td', { class: 'col-opt' }, konfZelle(a)),
      h('td', {}, statusZelle(a)));
    return tr;
  }

  function schalteDetail(tr, a) {
    const folge = tr.nextElementSibling;
    if (state.offen.has(a.id)) {
      state.offen.delete(a.id);
      if (folge && folge.classList.contains('detail')) folge.remove();
      tr.setAttribute('aria-expanded', 'false');
    } else {
      state.offen.add(a.id);
      tr.after(detail(a));
      tr.setAttribute('aria-expanded', 'true');
    }
  }

  function zeilenTaste(ev, tr, a) {
    if (ev.target !== tr) return;                      // Tasten in der Checkbox gehören der Checkbox
    const rows = [...$('tabelle').querySelectorAll('tr.row')];
    const i = rows.indexOf(tr);
    if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); schalteDetail(tr, a); }
    else if (ev.key === 'ArrowDown' && rows[i + 1]) { ev.preventDefault(); rows[i + 1].focus(); }
    else if (ev.key === 'ArrowUp' && rows[i - 1]) { ev.preventDefault(); rows[i - 1].focus(); }
    else if (ev.key === 'Home') { ev.preventDefault(); rows[0].focus(); }
    else if (ev.key === 'End') { ev.preventDefault(); rows[rows.length - 1].focus(); }
    else if (ev.key === 'Escape' && state.offen.has(a.id)) { ev.preventDefault(); schalteDetail(tr, a); }
  }

  function zeichneTabelle() {
    const tb = $('tabelle');
    const alle = aktuelleAenderungen();
    const l = gefiltert();
    $('treffer-info').textContent = `${l.length} von ${alle.length} Änderungen werden angezeigt.`;
    if (!l.length) {
      tb.replaceChildren(h('tr', {}, h('td', { colspan: 9, class: 'empty' }, alle.length ? 'Keine Einträge zu diesem Filter. Filter zurücksetzen, um alle zu sehen.' : 'In diesem Lauf wurden keine Änderungen erkannt.')));
      $('mehr-zeile').hidden = true;
      return;
    }
    const sicht = l.slice(0, state.limit);
    const zeilen = [];
    for (const a of sicht) {
      const tr = zeile(a);
      zeilen.push(tr);
      if (state.offen.has(a.id)) zeilen.push(detail(a));
    }
    tb.replaceChildren(...zeilen);
    const rest = l.length - sicht.length;
    $('mehr-zeile').hidden = rest <= 0;
    $('btn-mehr').textContent = `Weitere ${Math.min(rest, SEITENGROESSE)} anzeigen (${rest} noch ausgeblendet)`;
  }

  function vnBox(titel, kinder, info) {
    return h('section', { class: 'vn-box' }, h('h4', {}, titel, info ? infoIcon(info) : ''), kinder);
  }

  function detail(a) {
    const bild = h('figure', { class: 'detail-bild' },
      h('img', { loading: 'lazy', alt: `Ausschnitt aus dem PDF, Seite ${hatWort(a.seite)}`, src: `/api/lauf/${lauf(state.plan.run_id)}/vorschau/${encodeURIComponent(a.id)}.png`,
        onerror: (ev) => { const fig = ev.target.parentNode; ev.target.remove(); fig.append(h('p', { class: 'bild-fehlt' }, 'Der PDF-Ausschnitt konnte nicht geladen werden.')); } }),
      h('figcaption', {}, `PDF-Ausschnitt, Seite ${hatWort(a.seite)}`));

    const excel = vnBox('Dort steht jetzt', [
      h('p', { class: 'vn-text' }, a.excel_text ? text(a.excel_text) : '– (an dieser Stelle ist kein Text hinterlegt)'),
      h('p', { class: 'vn-meta' }, `Stelle in der Excel: ${a.ziel || '–'}`)], 'vorher_nachher');
    const pdf = vnBox('Im PDF: alt → neu', [
      a.alt ? h('p', { class: 'vn-text alt' }, text(a.alt)) : h('p', { class: 'vn-text hint' }, 'Alt: – (nichts Altes im PDF)'),
      a.neu ? h('p', { class: 'vn-text neu' }, text(a.neu)) : h('p', { class: 'vn-text hint' }, 'Neu: – (der Text entfällt)'),
      h('p', { class: 'vn-meta' }, `${artText(a.typ)} · Seite ${hatWort(a.seite)}`)]);

    const hin = liste(a.hinweise), w = WIRKUNG[a.wirkung];
    const wichtig = [];
    if (w && w.schlecht) wichtig.push(h('li', {}, h('strong', {}, `${w.text}: `), w.lang));
    if (nachkontrolleAbweichung(a)) wichtig.push(h('li', {}, h('strong', {}, 'Nachkontrolle: '), 'Die Excel weicht von dem ab, was geplant war.'));
    const box = hin.length || wichtig.length ? h('div', { class: 'detail-hinweise' + (hatProblem(a) ? ' rot' : ''), role: 'group', 'aria-label': 'Hinweise zu dieser Änderung' },
      h('strong', {}, '⚠ Bitte beachten'), h('ul', {}, wichtig, hin.map((x) => h('li', {}, text(x))))) : '';

    const konf = KONF[a.konfidenz] ? a.konfidenz : text(a.konfidenz) || '–';
    const kv = h('dl', { class: 'kv' },
      h('dt', {}, 'Sicherheit'), h('dd', {}, `${konf}${a.score ? ` (Übereinstimmung ${Math.round(zahl(a.score) * 100)} %)` : ''}`),
      h('dt', {}, 'Grund'), h('dd', {}, a.grund || '–'),
      h('dt', {}, 'Wirkung in der Excel'), h('dd', {}, w ? `${w.text}. ${w.lang}` : 'Noch nicht angewendet.'),
      h('dt', {}, 'Wörter'), h('dd', {}, hatWort(a.woerter)),
      a.kontext ? [h('dt', {}, 'Umgebung im PDF'), h('dd', {}, kurz(a.kontext, 400))] : '');

    return h('tr', { class: 'detail', id: 'detail-' + a.id }, h('td', { colspan: 9 },
      h('div', { class: 'detail-grid' }, bild, h('div', {}, h('div', { class: 'vn-grid' }, excel, pdf), box, kv))));
  }

  // ------------------------------------------------------------------ Auswahl der Vorschläge
  const vorschlaege = () => liste(state.plan && state.plan.aenderungen).filter((a) => a.status === 'pruefen');

  function zeichneAuswahl() {
    const p = state.plan;
    if (!p) return;
    const fertig = !!state.ergebnis;
    const v = vorschlaege();
    $('auswahl-zeile').hidden = fertig || !v.length;
    $('chk-alle').closest('label').hidden = fertig || !v.length;
    $('auswahl-info').textContent = `${state.freig.size} von ${pl(v.length, 'Vorschlag', 'Vorschlägen')} freigegeben`;
    const sicht = gefiltert().filter((a) => a.status === 'pruefen');
    const gew = sicht.filter((a) => state.freig.has(a.id)).length;
    const c = $('chk-alle');
    c.checked = sicht.length > 0 && gew === sicht.length;
    c.indeterminate = gew > 0 && gew < sicht.length;
    c.disabled = !sicht.length;
    for (const cb of $('tabelle').querySelectorAll('tr.row input[type=checkbox]')) {
      const id = cb.closest('tr').dataset.id;
      cb.checked = state.freig.has(id);
    }
    // Knopf
    const bt = $('btn-uebernehmen');
    const hat = state.ergebnisse.length > 0;
    bt.textContent = (hat ? 'Weiteres Ergebnis erzeugen' : 'Excel erzeugen') + (state.freig.size ? ` (${pl(state.freig.size, 'Vorschlag', 'Vorschläge')} freigegeben)` : '');
    bt.hidden = fertig;
    $('btn-zurueck-auswahl').hidden = !fertig;
  }

  function waehle(ids, an) {
    for (const id of ids) an ? state.freig.add(id) : state.freig.delete(id);
    zeichneAuswahl();
  }

  // ------------------------------------------------------------------ Farben
  const FARBE_OK = /^#[0-9a-f]{3,8}$/i;
  function zeichneFarben() {
    const p = state.plan;
    $('farben-leer').hidden = !!p;
    $('farben-inhalt').hidden = !p;
    if (!p) return;
    const fl = liste(p.farben);
    $('farben-sub').textContent = fl.length ? `Farben im PDF „${text(p.pdf_name)}“ und wie sie gedeutet werden.` : 'Im PDF wurden keine Farben erkannt oder der Lauf hat keine Farbübersicht gespeichert.';
    $('farben-tabelle').replaceChildren(...fl.map((f) => {
      const feld = h('span', { class: 'farbe-feld', 'aria-hidden': 'true' });
      if (FARBE_OK.test(text(f.farbe))) feld.style.backgroundColor = f.farbe;       // Farbe des PDFs, keine Dashboard-Farbe
      return h('tr', {}, h('td', {}, feld, text(f.farbe)), h('td', { class: 'zahl' }, hatWort(f.zeichen)), h('td', {}, text(f.art)),
        h('td', {}, text(f.klasse) === '-' ? '–' : text(f.klasse)), h('td', {}, text(f.bedeutung)));
    }));
  }

  // ------------------------------------------------------------------ Aktionen
  async function uebernehmen() {
    if (!state.plan || state.ergebnis) return;
    const b = state.plan.bilanz || {};
    if (!zahl(b.uebernommen) && !state.freig.size) { toastFehler('Es gibt nichts Sicheres zu übernehmen. Wähle Vorschläge aus oder prüfe die Liste.'); return; }
    if (state.ergebnisse.length && !confirm(`Für diesen Lauf gibt es schon ${pl(state.ergebnisse.length, 'Ergebnis', 'Ergebnisse')}. Soll ein weiteres Ergebnis erzeugt werden? Die früheren bleiben erhalten.`)) return;
    busy(true, 'Trage die Änderungen in eine Kopie der Excel ein …');
    try {
      const id = state.plan.run_id;
      const res = await api(`/lauf/${lauf(id)}/uebernehmen`, { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ freigaben: [...state.freig] }) });
      state.ergebnis = res;
      try { state.ergebnisse = liste((await api('/lauf/' + lauf(id))).ergebnisse); }
      catch (_) { state.ergebnisse = [...state.ergebnisse, { ordner: res.ordner, excel: res.excel, bilanz: res.bilanz, erstellt: '', freigaben: [...state.freig] }]; }
      state.f = neuerFilter();
      state.limit = SEITENGROESSE;
      zeichne();
      busy(false);
      toast('Das Ergebnis ist fertig. Die Excel und das Protokoll stehen zum Download bereit.');
      $('ergebnis-box').scrollIntoView({ block: 'nearest' });
      $('ergebnis-box').focus();
    } catch (e) { meldeFehler(e); } finally { busy(false); }
  }

  function zurueckZurAuswahl() {
    state.ergebnis = null;
    state.f = neuerFilter();
    state.limit = SEITENGROESSE;
    zeichne();
    $('btn-uebernehmen').focus();
  }

  // ------------------------------------------------------------------ Verlauf
  function laufKarte(r) {
    const st = LAUF_STATUS[r.status] || { text: 'Unbekannter Zustand', farbe: 'grey', offen: false };
    const name = r.pdf && r.excel ? `${r.pdf} → ${r.excel}` : `Lauf ${text(r.run_id)}`;
    const b = r.bilanz && typeof r.bilanz === 'object' ? r.bilanz : null;
    const meta = [fmtZeit(r.angelegt)];
    if (b) meta.push(`${pl(zahl(b.aenderungen), 'Änderung', 'Änderungen')}`, `${zahl(b.pruefen)} zu prüfen`, `${zahl(b.nicht_uebernommen)} nicht übernommen`);
    meta.push(pl(zahl(r.ergebnisse), 'Ergebnis', 'Ergebnisse'));
    const badges = [h('span', { class: 'badge ' + st.farbe }, st.text)];
    if (r.archiviert) badges.push(h('span', { class: 'badge grey' }, 'Archiviert'));
    if (b && b.stimmt === false) badges.push(h('span', { class: 'badge red' }, 'Bilanz stimmt nicht'));
    if (b && b.vollstaendig === false) badges.push(h('span', { class: 'badge amber' }, 'Unvollständig gelesen'));
    if (b && zahl(b.nachkontrolle_abweichungen)) badges.push(h('span', { class: 'badge red' }, pl(zahl(b.nachkontrolle_abweichungen), 'Abweichung', 'Abweichungen')));
    const knoepfe = [];
    if (st.offen) knoepfe.push(h('button', { class: 'btn btn-primary btn-small', type: 'button', 'aria-label': `Lauf öffnen: ${name}`, onclick: () => laden(r.run_id) }, 'Öffnen'));
    knoepfe.push(r.archiviert
      ? h('button', { class: 'btn btn-secondary btn-small', type: 'button', 'aria-label': `Lauf wiederherstellen: ${name}`, onclick: () => wiederherstellen(r.run_id) }, 'Wiederherstellen')
      : h('button', { class: 'btn btn-secondary btn-small', type: 'button', 'aria-label': `Lauf archivieren: ${name}`, onclick: () => archivieren(r.run_id) }, 'Archivieren'));
    return h('li', { class: 'run-card' + (r.archiviert ? ' archiviert' : '') },
      h('div', { class: 'run-info' }, h('div', { class: 'run-name' }, name), h('div', { class: 'meta' }, meta.filter(Boolean).join(' · ')),
        h('div', { class: 'badges' }, badges), r.meldung ? h('p', { class: 'run-meldung' }, text(r.meldung)) : ''),
      h('div', { class: 'run-knoepfe' }, knoepfe));
  }

  async function ladeVerlauf() {
    const box = $('verlauf-liste');
    try {
      const l = liste(await api(state.zeigeArchiv ? '/lauf?archiviert=1' : '/lauf'));
      if (!l.length) {
        box.replaceChildren(h('p', { class: 'empty' }, state.zeigeArchiv ? 'Es gibt noch keine Läufe.' : 'Noch keine Läufe. Starte einen neuen Lauf im ersten Tab.'));
        return;
      }
      box.replaceChildren(h('ul', { class: 'run-liste' }, l.map(laufKarte)));
    } catch (e) {
      box.replaceChildren(h('p', { class: 'empty' }, e.message));
      meldeFehler(e);
    }
  }

  async function laden(id) {
    busy(true, 'Lade Lauf …');
    try { oeffne(await api('/lauf/' + lauf(id))); }
    catch (e) { meldeFehler(e); } finally { busy(false); }
  }

  async function archivieren(id) {
    try {
      await api('/lauf/' + lauf(id), { method: 'DELETE' });
      toast('Der Lauf ist archiviert. Er bleibt auf diesem Rechner erhalten. Unter „Archivierte anzeigen“ kannst du ihn wiederherstellen.');
      if (state.plan && state.plan.run_id === id) { state.plan.archiviert = true; $('archiv-hinweis').hidden = false; }
      await ladeVerlauf();
      $('verlauf-titel').focus();
    } catch (e) { meldeFehler(e); }
  }

  async function wiederherstellen(id) {
    try {
      await api(`/lauf/${lauf(id)}/wiederherstellen`, { method: 'POST' });
      toast('Der Lauf ist wiederhergestellt.');
      if (state.plan && state.plan.run_id === id) { state.plan.archiviert = false; $('archiv-hinweis').hidden = true; }
      if ($('tab-verlauf').hidden === false) { await ladeVerlauf(); $('verlauf-titel').focus(); }
    } catch (e) { meldeFehler(e); }
  }

  // ------------------------------------------------------------------ Start
  function verdrahteFilter() {
    let zeit = null;
    $('f-suche').addEventListener('input', (ev) => { clearTimeout(zeit); zeit = setTimeout(() => { state.f.suche = ev.target.value; neuFiltern(); }, 150); });
    const feld = { 'f-typ': 'typ', 'f-seite': 'seite', 'f-section': 'section', 'f-konfidenz': 'konfidenz', 'f-sort': 'sort' };
    for (const [id, key] of Object.entries(feld)) $(id).addEventListener('change', (ev) => { state.f[key] = ev.target.value; neuFiltern(); });
    $('btn-filter-reset').addEventListener('click', () => { const s = state.f.sort; state.f = neuerFilter(); state.f.sort = s; fuelleFilter(); neuFiltern(); });
    $('btn-nur-vorschlaege').addEventListener('click', () => { state.f = { ...neuerFilter(), status: 'pruefen' }; fuelleFilter(); neuFiltern(); });
    $('btn-mehr').addEventListener('click', () => { state.limit += SEITENGROESSE; zeichneTabelle(); });
    $('btn-alle-vorschlaege').addEventListener('click', () => waehle(vorschlaege().map((a) => a.id), true));
    $('btn-sichtbare').addEventListener('click', () => waehle(gefiltert().filter((a) => a.status === 'pruefen').map((a) => a.id), true));
    $('btn-keine').addEventListener('click', () => waehle([...state.freig], false));
    $('chk-alle').addEventListener('change', (ev) => waehle(gefiltert().filter((a) => a.status === 'pruefen').map((a) => a.id), ev.target.checked));
  }

  document.addEventListener('DOMContentLoaded', async () => {
    try { const t = localStorage.getItem('cde-theme'); if (t === 'dark' || t === 'light') document.documentElement.setAttribute('data-theme', t); } catch (_) { /* ignorieren */ }
    themeKnopf();
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', themeKnopf);
    $('theme-toggle').addEventListener('click', () => setzeTheme(istDunkel() ? 'light' : 'dark'));
    for (const b of document.querySelectorAll('.page-tab')) { b.addEventListener('click', () => zeigeTab(b.dataset.tab)); b.addEventListener('keydown', tabTaste); }
    for (const b of document.querySelectorAll('[data-gehe]')) b.addEventListener('click', () => zeigeTab(b.dataset.gehe, true));
    dropzone('dz-pdf', 'file-pdf', 'pdf');
    dropzone('dz-xlsx', 'file-xlsx', 'xlsx');
    for (const e of ['dragover', 'drop']) window.addEventListener(e, (ev) => ev.preventDefault());   // eine daneben gezogene Datei darf die Seite nicht verlassen
    $('btn-analyse').addEventListener('click', analysieren);
    $('btn-uebernehmen').addEventListener('click', uebernehmen);
    $('btn-zurueck-auswahl').addEventListener('click', zurueckZurAuswahl);
    $('btn-lauf-wiederherstellen').addEventListener('click', () => state.plan && wiederherstellen(state.plan.run_id));
    $('toast-schliessen').addEventListener('click', fehlerSchliessen);
    $('chk-archiv').addEventListener('change', (ev) => { state.zeigeArchiv = ev.target.checked; ladeVerlauf(); });
    $('btn-verlauf-neu').addEventListener('click', ladeVerlauf);
    verdrahteFilter();

    document.addEventListener('click', (ev) => {
      const icon = ev.target.closest('.info-icon');
      if (icon) { ev.preventDefault(); ev.stopPropagation(); oeffnePopup(icon); return; }
      if (popup && !popup.contains(ev.target)) schliessePopup(false);
    });
    document.addEventListener('keydown', (ev) => { if (ev.key === 'Escape') { if (popup) schliessePopup(); else if (!$('toast-fehler-box').hidden) fehlerSchliessen(); } });
    window.addEventListener('scroll', setzePopup, true);
    window.addEventListener('resize', setzePopup);
    window.addEventListener('unhandledrejection', (ev) => { ev.preventDefault(); meldeFehler(ev.reason); });
    window.addEventListener('error', (ev) => { console.error(ev.error || ev.message); toastFehler(TXT_UNERWARTET); });

    zeichne();
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
