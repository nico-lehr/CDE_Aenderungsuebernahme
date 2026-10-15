/* CDE-Änderungsübernahme – Dashboard (reines JavaScript, keine Abhängigkeiten).
   Alle Texte aus Daten werden über textContent gesetzt (kein innerHTML mit Fremdtext). */
(() => {
  'use strict';

  const STATUS = {
    uebernommen: { text: 'Übernommen', farbe: 'green' },
    pruefen: { text: 'Zu prüfen', farbe: 'amber' },
    nicht_uebernommen: { text: 'Nicht übernommen', farbe: 'red' },
  };
  const ART = {
    text_ersetzt: 'Text ersetzt', text_eingefuegt: 'Text eingefügt', text_geloescht: 'Text gelöscht',
    absatz_neu: 'Absatz neu', absatz_geloescht: 'Absatz gelöscht',
    section_neu: 'Section neu', section_geloescht: 'Section gelöscht',
  };

  const state = { plan: null, filter: 'alle', freig: new Set(), files: { pdf: null, xlsx: null }, offen: new Set(), ergebnis: null };
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

  function toast(text, fehler = false) {
    const t = $('toast');
    t.textContent = text;
    t.className = 'toast' + (fehler ? ' error' : '');
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => { t.hidden = true; }, fehler ? 8000 : 4000);
  }

  function busy(an, text = 'Bitte warten …') { $('busy-text').textContent = text; $('busy').hidden = !an; }

  async function api(pfad, opt = {}) {
    const r = await fetch('/api' + pfad, opt);
    let daten = null;
    try { daten = await r.json(); } catch (_) { /* keine JSON-Antwort */ }
    if (!r.ok) throw new Error((daten && daten.fehler) || `Fehler ${r.status}`);
    return daten;
  }

  // ------------------------------------------------------------------ Tabs
  function zeigeTab(name) {
    document.querySelectorAll('.page-tab').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
    document.querySelectorAll('.tab').forEach((s) => { s.hidden = s.id !== 'tab-' + name; });
    if (name === 'verlauf') ladeVerlauf();
  }

  // ------------------------------------------------------------------ Theme
  function setzeTheme(t) {
    document.documentElement.setAttribute('data-theme', t);
    try { localStorage.setItem('cde-theme', t); } catch (_) { /* privater Modus */ }
  }

  // ------------------------------------------------------------------ Upload
  function dateiGewaehlt(art, datei) {
    const dz = art === 'pdf' ? $('dz-pdf') : $('dz-xlsx');
    const okPdf = art === 'pdf' && /\.pdf$/i.test(datei.name);
    const okXlsx = art === 'xlsx' && /\.(xlsx|xlsm)$/i.test(datei.name);
    if (!okPdf && !okXlsx) {
      toast(art === 'pdf' ? 'Bitte eine PDF-Datei wählen.' : 'Bitte eine .xlsx-Datei wählen (alte .xls bitte in Excel als .xlsx speichern).', true);
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
    } catch (e) { toast(e.message, true); } finally { busy(false); }
  }

  function oeffne(plan) {
    state.plan = plan;
    state.filter = 'alle';
    state.freig = new Set();
    state.offen = new Set();
    state.ergebnis = null;
    zeichne();
    zeigeTab('pruefen');
  }

  // ------------------------------------------------------------------ Darstellung
  function kachel(wert, label, farbe = '') { return h('div', { class: 'stat-card ' + farbe }, h('span', { class: 'stat-value' }, wert), h('span', { class: 'stat-label' }, label)); }

  function zeichne() {
    const p = state.plan;
    $('leer-pruefen').hidden = !!p;
    $('pruefen-inhalt').hidden = !p;
    if (!p) return;
    $('run-titel').textContent = `${p.pdf_name}  →  ${p.excel_name}`;
    $('run-sub').textContent = `Lauf ${p.run_id} · ${p.seiten} PDF-Seiten`;
    const b = (state.ergebnis && state.ergebnis.bilanz) || p.bilanz;
    const st = $('stats'); st.replaceChildren(
      kachel(b.aenderungen, 'Änderungen erkannt'),
      kachel(b.uebernommen, 'Übernommen', 'green'),
      kachel(b.pruefen, 'Zu prüfen', 'amber'),
      kachel(b.nicht_uebernommen, 'Nicht übernommen', 'red'));
    const bx = $('bilanz-box');
    bx.className = 'bilanz' + (b.stimmt && !b.nachkontrolle_abweichungen ? '' : ' bad');
    bx.textContent = b.stimmt
      ? `Bilanz stimmt: alle ${b.woerter_im_pdf} markierten Wörter im PDF sind übernommen, zu prüfen oder mit Grund aufgelistet. Nichts ging still verloren.` +
        (b.nachkontrolle_abweichungen ? ` ACHTUNG: ${b.nachkontrolle_abweichungen} Änderung(en) weichen bei der Nachkontrolle ab.` : '')
      : `ACHTUNG: Die Bilanz stimmt nicht (${b.woerter_im_pdf} Wörter im PDF markiert). Bitte Ergebnis nicht verwenden und melden.`;
    const w = $('warnungen');
    const warn = (state.ergebnis && state.ergebnis.warnungen) || p.warnungen || [];
    w.hidden = !warn.length;
    w.replaceChildren(h('strong', {}, 'Hinweise'), h('ul', {}, warn.map((x) => h('li', {}, x))));
    zeichneLayout(p);
    zeichneErgebnis();
    zeichneFilter();
    zeichneTabelle();
    zeichneFarben(p);
    $('btn-uebernehmen').disabled = !!state.ergebnis;
    $('btn-alle-vorschlaege').disabled = !!state.ergebnis;
    $('btn-uebernehmen').textContent = state.freig.size ? `Excel erzeugen (${state.freig.size} Vorschläge freigegeben)` : 'Excel erzeugen';
  }

  function zeichneLayout(p) {
    const box = $('layout-inhalt');
    box.replaceChildren(h('p', { class: 'hint' }, 'So hat das Tool den Aufbau je Tabellenblatt verstanden. Weicht das stark von der Realität ab, bitte melden.'),
      h('div', { class: 'table-wrap' }, h('table', { class: 'changes' },
        h('thead', {}, h('tr', {}, ['Blatt', 'Section-Spalte', 'Titel-Spalte', 'Inhalt-Spalte', 'Modell', 'Sections'].map((x) => h('th', {}, x)))),
        h('tbody', {}, p.layout.map((l) => h('tr', {}, [l.blatt, l.section_spalte || '–', l.titel_spalte || '–', l.inhalt_spalte || '–',
          l.modell === 'zeile_pro_section' ? 'eine Zeile je Section' : 'eine Zeile je Absatz', l.sections].map((x) => h('td', {}, x))))))));
  }

  function zeichneErgebnis() {
    const box = $('ergebnis-box'), e = state.ergebnis;
    box.hidden = !e;
    if (!e) return;
    const link = (name, text, primaer) => h('a', { class: 'btn ' + (primaer ? 'btn-primary' : 'btn-secondary'),
      href: `/api/lauf/${state.plan.run_id}/download/${e.ordner}/${encodeURIComponent(name)}` }, text);
    box.replaceChildren(h('h3', {}, 'Fertig – die aktualisierte Excel liegt bereit'),
      h('p', { class: 'hint' }, 'Die Original-Excel wurde nicht verändert. Das Protokoll listet jede Änderung mit alt/neu, Quelle im PDF und Stelle in der Excel.'),
      link(e.excel, '⬇ Aktualisierte Excel', true), link('aenderungsprotokoll.xlsx', '⬇ Änderungsprotokoll', false));
  }

  function zeichneFilter() {
    const b = (state.ergebnis && state.ergebnis.bilanz) || state.plan.bilanz;
    const zeilen = [['alle', `Alle (${b.aenderungen})`], ['uebernommen', `Übernommen (${b.uebernommen})`],
      ['pruefen', `Zu prüfen (${b.pruefen})`], ['nicht_uebernommen', `Nicht übernommen (${b.nicht_uebernommen})`]];
    $('filter-row').replaceChildren(...zeilen.map(([k, t]) => h('button', {
      class: 'chip' + (state.filter === k ? ' active' : ''), type: 'button',
      onclick: () => { state.filter = k; zeichneFilter(); zeichneTabelle(); } }, t)));
  }

  function aktuelleAenderungen() { return (state.ergebnis ? state.ergebnis.aenderungen : state.plan.aenderungen); }

  function zeichneTabelle() {
    const tb = $('tabelle');
    const liste = aktuelleAenderungen().filter((a) => state.filter === 'alle' || a.status === state.filter);
    if (!liste.length) { tb.replaceChildren(h('tr', {}, h('td', { colspan: 8, class: 'empty' }, 'Keine Einträge in dieser Ansicht.'))); return; }
    const zeilen = [];
    for (const a of liste) {
      const s = STATUS[a.status] || { text: a.status, farbe: 'grey' };
      const kannFreigeben = a.status === 'pruefen' && !state.ergebnis;
      const chk = kannFreigeben ? h('input', { type: 'checkbox', 'aria-label': 'Vorschlag übernehmen',
        ...(state.freig.has(a.id) ? { checked: true } : {}),
        onclick: (ev) => { ev.stopPropagation(); ev.target.checked ? state.freig.add(a.id) : state.freig.delete(a.id); zeichneKnopf(); } }) : '';
      zeilen.push(h('tr', { class: 'row', onclick: () => { state.offen.has(a.id) ? state.offen.delete(a.id) : state.offen.add(a.id); zeichneTabelle(); } },
        h('td', { class: 'col-chk' }, chk), h('td', {}, a.id), h('td', {}, a.seite), h('td', {}, ART[a.typ] || a.typ), h('td', {}, a.section || '–'),
        h('td', {}, a.alt ? h('span', { class: 'alt' }, kurz(a.alt, 70)) : '', a.alt && a.neu ? h('span', { class: 'pfeil' }, '→') : '',
          a.neu ? h('span', { class: 'neu' }, kurz(a.neu, 70)) : ''),
        h('td', {}, a.ziel || '–'),
        h('td', {}, h('span', { class: 'badge ' + s.farbe }, s.text), a.nachkontrolle === 'abweichung' ? h('span', { class: 'badge red' }, 'Abweichung') : '')));
      if (state.offen.has(a.id)) zeilen.push(detail(a));
    }
    tb.replaceChildren(...zeilen);
  }

  function zeichneKnopf() { $('btn-uebernehmen').textContent = state.freig.size ? `Excel erzeugen (${state.freig.size} Vorschläge freigegeben)` : 'Excel erzeugen'; }

  function kurz(t, n) { return t.length > n ? t.slice(0, n - 1) + '…' : t; }

  function detail(a) {
    const kv = h('dl', { class: 'kv' },
      h('dt', {}, 'Alt im PDF'), h('dd', {}, a.alt || '–'), h('dt', {}, 'Neu im PDF'), h('dd', {}, a.neu || '–'),
      h('dt', {}, 'Stelle in Excel'), h('dd', {}, a.ziel || '–'), h('dt', {}, 'Dort steht jetzt'), h('dd', {}, a.excel_text || '–'),
      h('dt', {}, 'Sicherheit'), h('dd', {}, `${a.konfidenz}${a.score ? ` (Übereinstimmung ${Math.round(a.score * 100)} %)` : ''}`),
      h('dt', {}, 'Hinweis / Grund'), h('dd', {}, a.grund || '–'));
    const bild = h('img', { loading: 'lazy', alt: `Ausschnitt aus dem PDF, Seite ${a.seite}`, src: `/api/lauf/${state.plan.run_id}/vorschau/${a.id}.png` });
    return h('tr', { class: 'detail' }, h('td', { colspan: 8 }, h('div', { class: 'detail-grid' }, bild, kv)));
  }

  function zeichneFarben(p) {
    $('farben-tabelle').replaceChildren(...(p.farben || []).map((f) => h('tr', {},
      h('td', {}, h('span', { style: `display:inline-block;width:14px;height:14px;border-radius:4px;vertical-align:middle;margin-right:8px;background:${f.farbe};border:1px solid var(--grey-300)` }), f.farbe),
      h('td', {}, f.zeichen), h('td', {}, f.art), h('td', {}, f.klasse), h('td', {}, f.bedeutung))));
  }

  // ------------------------------------------------------------------ Aktionen
  async function uebernehmen() {
    if (!state.plan || state.ergebnis) { toast('Dieser Lauf wurde bereits übernommen. Für eine erneute Übernahme bitte einen neuen Lauf starten.', true); return; }
    const b = state.plan.bilanz;
    if (!b.uebernommen && !state.freig.size) { toast('Es gibt nichts Sicheres zu übernehmen. Wähle Vorschläge aus oder prüfe die Liste.', true); return; }
    busy(true, 'Trage die Änderungen in eine Kopie der Excel ein …');
    try {
      const res = await api(`/lauf/${state.plan.run_id}/uebernehmen`, { method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ freigaben: [...state.freig] }) });
      state.ergebnis = res;
      zeichne();
      toast('Fertig. Die aktualisierte Excel und das Protokoll stehen zum Download bereit.');
    } catch (e) { toast(e.message, true); } finally { busy(false); }
  }

  function alleVorschlaege() {
    if (state.ergebnis) return;
    state.plan.aenderungen.filter((a) => a.status === 'pruefen').forEach((a) => state.freig.add(a.id));
    zeichneTabelle(); zeichneKnopf();
  }

  async function ladeVerlauf() {
    const box = $('verlauf-liste');
    try {
      const liste = await api('/lauf');
      if (!liste.length) { box.replaceChildren(h('p', { class: 'empty' }, 'Noch keine Läufe.')); return; }
      box.replaceChildren(...liste.map((r) => h('div', { class: 'run-card' },
        h('div', {}, h('strong', {}, `${r.pdf} → ${r.excel}`), h('div', { class: 'meta' },
          `${r.angelegt.replace('T', ' ')} · ${r.bilanz ? `${r.bilanz.aenderungen} Änderungen, ${r.bilanz.nicht_uebernommen} nicht übernommen` : 'nicht analysiert'} · ${r.ergebnisse} Ergebnis(se)`)),
        h('div', {}, h('button', { class: 'btn btn-secondary btn-small', type: 'button', onclick: () => laden(r.run_id) }, 'Öffnen'), ' ',
          h('button', { class: 'btn btn-danger btn-small', type: 'button', onclick: () => loeschen(r.run_id) }, 'Löschen')))));
    } catch (e) { toast(e.message, true); }
  }

  async function laden(id) {
    busy(true, 'Lade Lauf …');
    try { const p = await api('/lauf/' + id); oeffne(p); }
    catch (e) { toast(e.message, true); } finally { busy(false); }
  }

  async function loeschen(id) {
    if (!confirm('Diesen Lauf endgültig von diesem Rechner löschen (Eingabe, Analyse und Ergebnisse)?')) return;
    try {
      await api('/lauf/' + id, { method: 'DELETE' });
      if (state.plan && state.plan.run_id === id) { state.plan = null; state.ergebnis = null; zeichne(); }
      ladeVerlauf();
    } catch (e) { toast(e.message, true); }
  }

  // ------------------------------------------------------------------ Start
  document.addEventListener('DOMContentLoaded', async () => {
    try { const t = localStorage.getItem('cde-theme'); if (t) document.documentElement.setAttribute('data-theme', t); } catch (_) { /* ignorieren */ }
    $('theme-toggle').addEventListener('click', () => {
      const dunkel = document.documentElement.getAttribute('data-theme') === 'dark' ||
        (!document.documentElement.getAttribute('data-theme') && matchMedia('(prefers-color-scheme: dark)').matches);
      setzeTheme(dunkel ? 'light' : 'dark');
    });
    document.querySelectorAll('.page-tab').forEach((b) => b.addEventListener('click', () => zeigeTab(b.dataset.tab)));
    dropzone('dz-pdf', 'file-pdf', 'pdf');
    dropzone('dz-xlsx', 'file-xlsx', 'xlsx');
    $('btn-analyse').addEventListener('click', analysieren);
    $('btn-uebernehmen').addEventListener('click', uebernehmen);
    $('btn-alle-vorschlaege').addEventListener('click', alleVorschlaege);
    try { const g = await api('/health'); $('port-text').textContent = g.port; $('pill-text').textContent = `Lokal · Port ${g.port}`; }
    catch (_) { $('pill-text').textContent = 'Server nicht erreichbar'; }
  });
})();
