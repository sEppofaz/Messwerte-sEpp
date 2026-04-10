'use strict';

// ═══════════════════════════════════════════════════════════════
//  Messdaten PWA – app.js
//  Requires: SheetJS (xlsx) loaded before this script
// ═══════════════════════════════════════════════════════════════

// ── Config ────────────────────────────────────────────────────

const DROPBOX_PATH = '/Apps/Claude/Messdaten/Messdaten sEpp-Claude.xlsx';
const APP_KEY      = 's2ggv6zysmzn7fa';
const APP_VERSION  = 'v5';

const TABLE_MAP = {
  'Strom':           'Tabelle3',
  'PV-Werte':        'Tabelle2',
  'Wasser':          'Tabelle1',
  'Heizung Stunden': 'Tabelle5',
  'Maschinen':       'Tabelle4',
};

const TABS = [
  {
    key: 'strom', sheet: 'Strom', label: '⚡ Strom',
    fields: [
      { key: 'zaehler',   label: 'Zählerstand (kWh)',    type: 'decimal', req: true  },
      { key: 'bemerkung', label: 'Bemerkung',             type: 'text',    req: false },
    ],
    headers: ['Datum', 'Zähler', 'Δ kWh', 'kWh/Tag'],
  },
  {
    key: 'pv', sheet: 'PV-Werte', label: '☀️ PV',
    fields: [
      { key: 'zaehler',   label: 'Zähler gesamt (kWh)',   type: 'decimal', req: true  },
      { key: 'pv1',       label: 'PV1 Zähler (optional)', type: 'decimal', req: false },
      { key: 'bemerkung', label: 'Bemerkung',              type: 'text',    req: false },
    ],
    headers: ['Datum', 'Zähler', 'Δ kWh', 'kWh/Tag', 'PV1/2%'],
  },
  {
    key: 'wasser', sheet: 'Wasser', label: '💧 Wasser',
    fields: [
      { key: 'zaehler',   label: 'Wasserstand (m³)',       type: 'decimal', req: true  },
      { key: 'ph',        label: 'pH-Wert',                type: 'decimal', req: false },
      { key: 'haerte',    label: 'Härte (°dH)',            type: 'decimal', req: false },
      { key: 'druck',     label: 'Druck (bar)',             type: 'decimal', req: false },
      { key: 'bemerkung', label: 'Bemerkung',               type: 'text',    req: false },
    ],
    headers: ['Datum', 'm³', 'Δ m³', 'm³/Tag'],
  },
  {
    key: 'heizung', sheet: 'Heizung Stunden', label: '🔥 Heizung',
    fields: [
      { key: 'zaehler',   label: 'Volllast-Stunden',       type: 'decimal', req: true  },
      { key: 'druck',     label: 'Druck (bar)',             type: 'decimal', req: false },
      { key: 'bemerkung', label: 'Bemerkung',               type: 'text',    req: false },
    ],
    headers: ['Datum', 'Std.', 'Δ Std.', 'Std/Tag'],
  },
  {
    key: 'maschinen', sheet: 'Maschinen', label: '🔧 Masch.',
    fields: [
      { key: 'kategorie', label: 'Kategorie',               type: 'text',    req: true  },
      { key: 'thema',     label: 'Thema',                   type: 'text',    req: false },
      { key: 'zaehler',   label: 'Betriebsstd.',            type: 'decimal', req: false },
      { key: 'kosten',    label: 'Kosten [€]',              type: 'decimal', req: false },
    ],
    headers: ['Datum', 'Kategorie', 'Thema', 'Kosten [€]'],
  },
];

// ── Dropbox OAuth2 PKCE ───────────────────────────────────────

const TOKEN_URL = 'https://api.dropboxapi.com/oauth2/token';
const AUTH_URL  = 'https://www.dropbox.com/oauth2/authorize';
const CONTENT   = 'https://content.dropboxapi.com/2';

function b64url(buf) {
  return btoa(String.fromCharCode(...buf))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}

async function pkce() {
  const verifier  = b64url(crypto.getRandomValues(new Uint8Array(32)));
  const digest    = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
  const challenge = b64url(new Uint8Array(digest));
  return { verifier, challenge };
}

function canonicalUrl() {
  const u = new URL(location.href);
  u.search = '';
  u.hash   = '';
  if (u.pathname.endsWith('index.html')) u.pathname = u.pathname.slice(0, -10);
  if (!u.pathname.endsWith('/'))         u.pathname += '/';
  return u.href;
}

async function startAuth() {
  const appKey      = APP_KEY;
  const redirectUri = canonicalUrl();
  const { verifier, challenge } = await pkce();

  sessionStorage.setItem('pkce_verifier',    verifier);
  sessionStorage.setItem('dropbox_app_key',  appKey);
  sessionStorage.setItem('redirect_uri',     redirectUri);

  const p = new URLSearchParams({
    response_type:         'code',
    client_id:             appKey,
    redirect_uri:          redirectUri,
    code_challenge:        challenge,
    code_challenge_method: 'S256',
    token_access_type:     'offline',
  });
  location.href = AUTH_URL + '?' + p;
}

async function handleCallback() {
  const code = new URLSearchParams(location.search).get('code');
  if (!code) return;

  const appKey      = APP_KEY;
  const verifier    = sessionStorage.getItem('pkce_verifier');
  const redirectUri = sessionStorage.getItem('redirect_uri')    || canonicalUrl();

  if (!appKey || !verifier) {
    alert('Auth-Fehler: Sitzungsdaten fehlen. Bitte erneut verbinden.');
    history.replaceState({}, '', location.pathname);
    return;
  }

  try {
    const r = await fetch(TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        code,
        grant_type:    'authorization_code',
        client_id:     appKey,
        redirect_uri:  redirectUri,
        code_verifier: verifier,
      }),
    });
    if (!r.ok) throw new Error(await r.text());
    const d = await r.json();
    localStorage.setItem('dropbox_access_token', d.access_token);
    localStorage.setItem('dropbox_refresh_token', d.refresh_token);
    localStorage.setItem('dropbox_expires',       Date.now() + d.expires_in * 1000);
  } catch (e) {
    alert('Token-Fehler: ' + e.message);
  }
  history.replaceState({}, '', location.pathname);
}

function isConnected() {
  return !!localStorage.getItem('dropbox_refresh_token');
}

function disconnect() {
  ['dropbox_access_token', 'dropbox_refresh_token', 'dropbox_expires']
    .forEach(k => localStorage.removeItem(k));
  init();
}

async function applyUpdate() {
  if ('caches' in window) {
    const keys = await caches.keys();
    await Promise.all(keys.map(k => caches.delete(k)));
  }
  if ('serviceWorker' in navigator) {
    const regs = await navigator.serviceWorker.getRegistrations();
    await Promise.all(regs.map(r => r.unregister()));
  }
  location.reload(true);
}

async function getToken() {
  const exp = +localStorage.getItem('dropbox_expires');
  if (Date.now() < exp - 60_000) return localStorage.getItem('dropbox_access_token');

  const r = await fetch(TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type:    'refresh_token',
      refresh_token: localStorage.getItem('dropbox_refresh_token'),
      client_id:     APP_KEY,
    }),
  });
  if (!r.ok) throw new Error('Token abgelaufen – bitte neu verbinden.');
  const d = await r.json();
  localStorage.setItem('dropbox_access_token', d.access_token);
  localStorage.setItem('dropbox_expires',       Date.now() + d.expires_in * 1000);
  return d.access_token;
}

// ── Dropbox file I/O ─────────────────────────────────────────

async function dbDownload(token) {
  const r = await fetch(CONTENT + '/files/download', {
    method:  'POST',
    headers: {
      Authorization:     'Bearer ' + token,
      'Dropbox-API-Arg': JSON.stringify({ path: DROPBOX_PATH }),
    },
  });
  if (!r.ok) {
    let detail;
    try { detail = (await r.json()).error_summary; } catch { detail = r.status; }
    throw new Error('Dropbox ' + r.status + ': ' + detail);
  }
  return r.arrayBuffer();
}

async function dbUpload(token, data) {
  const r = await fetch(CONTENT + '/files/upload', {
    method:  'POST',
    headers: {
      Authorization:      'Bearer ' + token,
      'Content-Type':     'application/octet-stream',
      'Dropbox-API-Arg':  JSON.stringify({ path: DROPBOX_PATH, mode: 'overwrite', autorename: false }),
    },
    body: data,
  });
  if (!r.ok) throw new Error('Upload fehlgeschlagen: ' + r.status);
}

// ── Excel helpers (SheetJS) ──────────────────────────────────

function xlLastRow(ws) {
  if (!ws['!ref']) return 3;
  const rng = XLSX.utils.decode_range(ws['!ref']);
  for (let r = rng.e.r; r >= 1; r--) {
    const c = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    if (c && c.v !== undefined) return r + 1;   // returns 1-indexed
  }
  return 3;
}

function xlGet(ws, row, col) {   // 1-indexed
  const cell = ws[XLSX.utils.encode_cell({ r: row - 1, c: col - 1 })];
  return cell ? cell.v : undefined;
}

function xlSet(ws, row, col, val) {  // 1-indexed
  const addr = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  if (val === null || val === undefined) { delete ws[addr]; return; }
  if (val instanceof Date) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const days = (val - excelEpoch) / 86400000;
    ws[addr] = { t: 'n', v: days, z: 'dd.mm.yy' };
  } else if (typeof val === 'number') {
    ws[addr] = { t: 'n', v: val };
  } else {
    ws[addr] = { t: 's', v: String(val) };
  }
  _expandRef(ws, row - 1, col - 1);
}

function xlFormula(ws, row, col, f) {  // 1-indexed
  const addr = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  ws[addr] = { t: 'n', f };
  _expandRef(ws, row - 1, col - 1);
}

function _expandRef(ws, r, c) {
  const rng = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  rng.e.r = Math.max(rng.e.r, r);
  rng.e.c = Math.max(rng.e.c, c);
  ws['!ref'] = XLSX.utils.encode_range(rng);
}

function fmtDate(v) {
  if (v == null) return '–';
  if (typeof v === 'number') {
    const offset = workbookIs1904 ? 24107 : 25569;
    const d = new Date(Math.round((v - offset) * 86400) * 1000);
    return d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: '2-digit', timeZone: 'UTC' });
  }
  return String(v);
}

function safeRound(v, d = 2) {
  const n = typeof v === 'number' ? v : parseFloat(v);
  return isNaN(n) ? (v ?? '–') : Math.round(n * 10 ** d) / 10 ** d;
}

// ── Recent rows ──────────────────────────────────────────────

function recentRows(ws, key, count = 10) {
  const last = xlLastRow(ws);

  if (key === 'maschinen') {
    const startRow = Math.max(2, last - count + 1);
    const rows = [], rowNums = [];
    for (let r = startRow; r <= last; r++) {
      rows.push([
        fmtDate(xlGet(ws, r, 1)),
        xlGet(ws, r, 2) ?? '–',
        xlGet(ws, r, 3) ?? '–',
        safeRound(xlGet(ws, r, 5)),
      ]);
      rowNums.push(r);
    }
    const hasMore = startRow > 2;
    return { rows: rows.slice(-count), rowNums: rowNums.slice(-count), hasMore };
  }

  const start = Math.max(3, last - count);
  const raw   = [];
  for (let r = start; r <= last; r++) {
    raw.push({
      dv:  xlGet(ws, r, 1),
      z:   key === 'pv' ? xlGet(ws, r, 4) : xlGet(ws, r, 2),
      pv1: key === 'pv' ? xlGet(ws, r, 6) : null,
      r,
    });
  }

  const rows = [], rowNums = [];
  for (let i = 1; i < raw.length; i++) {
    const { dv, z, pv1, r }               = raw[i];
    const { dv: dvp, z: zp, pv1: pv1p }   = raw[i - 1];
    let delta = null, perDay = null, ratio = null;
    try {
      delta = Math.round((z - zp) * 1000) / 1000;
      if (typeof dv === 'number' && typeof dvp === 'number') {
        const days = dv - dvp;
        if (days > 0) {
          const dp = { strom: 1, pv: 1, wasser: 4, heizung: 2 }[key] ?? 2;
          perDay = Math.round(delta / days * 10 ** dp) / 10 ** dp;
        }
      }
    } catch {}
    if (key === 'pv' && pv1 != null && pv1p != null) {
      try {
        const dPv1 = pv1 - pv1p;
        if (delta > 0) {
          const pct1 = Math.round(dPv1 / delta * 100);
          ratio = pct1 + '/' + (100 - pct1);
        }
      } catch {}
    }
    const row = [fmtDate(dv), z, safeRound(delta), safeRound(perDay)];
    if (key === 'pv') row.push(ratio ?? '–');
    rows.push(row);
    rowNums.push(r);
  }
  const hasMore = start > 3;
  return { rows: rows.slice(-count), rowNums: rowNums.slice(-count), hasMore };
}

// ── Read row fields for editing ──────────────────────────────

function readRowFields(ws, key, rowNum) {
  const r = rowNum;
  const datumSerial = xlGet(ws, r, 1);
  let dateStr = '';
  if (typeof datumSerial === 'number') {
    const offset = workbookIs1904 ? 24107 : 25569;
    const ms = Math.round((datumSerial - offset) * 86400) * 1000;
    dateStr = new Date(ms).toISOString().slice(0, 10);
  }

  const fields = {};
  if (key === 'strom') {
    fields.zaehler   = xlGet(ws, r, 2);
    fields.bemerkung = xlGet(ws, r, 8);
  } else if (key === 'pv') {
    fields.zaehler   = xlGet(ws, r, 4);
    fields.pv1       = xlGet(ws, r, 6);
    fields.bemerkung = xlGet(ws, r, 17);
  } else if (key === 'wasser') {
    fields.zaehler   = xlGet(ws, r, 2);
    fields.ph        = xlGet(ws, r, 7);
    fields.haerte    = xlGet(ws, r, 8);
    fields.druck     = xlGet(ws, r, 9);
    fields.bemerkung = xlGet(ws, r, 10);
  } else if (key === 'heizung') {
    fields.zaehler   = xlGet(ws, r, 2);
    fields.druck     = xlGet(ws, r, 3);
    fields.bemerkung = xlGet(ws, r, 9);
  } else if (key === 'maschinen') {
    fields.kategorie = xlGet(ws, r, 2);
    fields.thema     = xlGet(ws, r, 3);
    fields.zaehler   = xlGet(ws, r, 4);
    fields.kosten    = xlGet(ws, r, 5);
  }

  return { dateStr, fields };
}

// ── Write row ────────────────────────────────────────────────

function writeRow(ws, key, p, n, dateSerial, fields) {
  const sc = (col, v) => xlSet(ws, n, col, v);
  const sf = (col, f) => xlFormula(ws, n, col, f);
  const z  = fields.zaehler;

  const dateAddr = XLSX.utils.encode_cell({ r: n - 1, c: 0 });
  ws[dateAddr] = { t: 'n', v: dateSerial, z: 'DD.MM.YYYY' };
  _expandRef(ws, n - 1, 0);

  if (key === 'strom') {
    sc(2, z);
    sf(3, `C${p}+B${n}-B${p}`);
    sf(4, `C${n}-C${p}`);
    sf(5, `D${n}/(A${n}-A${p})`);
    sf(6, `(C${n}-$C$3)/(A${n}-$A$3)`);
    sf(7, `A${n}-A${p}`);
    sc(8, fields.bemerkung || '');

  } else if (key === 'pv') {
    sf(2, `YEAR(A${n})`);
    sf(3, `D${n}-D${p}`);
    sc(4, z);
    sf(5, `E${p}+C${n}`);
    sc(6, fields.pv1 != null ? fields.pv1 : xlGet(ws, p, 6));
    sf(7, `E${n}-F${n}`);
    sf(8, `K${n}/I${n}`);
    sf(9, `C${n}/DATEDIF(A${p},A${n},"D")`);
    sf(10, `(F${n}-F${p})/DATEDIF(A${p},A${n},"D")`);
    sf(11, `(G${n}-G${p})/DATEDIF(A${p},A${n},"D")`);
    sf(12, `J${n}/21.45`);
    sf(13, `K${n}/29.5`);
    sf(14, `I${n}/(21.45+29.5)`);
    sf(15, `(E${n}-$E$3)/(A${n}-$A$3)`);
    sf(16, `(G${n}-$G$64)/(A${n}-$A$64)`);
    sc(17, fields.bemerkung || '');

  } else if (key === 'wasser') {
    sc(2, z);
    sf(3, `B${n}-B${p}`);
    sf(4, `C${n}/(A${n}-A${p})`);
    sf(5, `(B${n}-$B$3)/(A${n}-$A$3)`);
    sf(6, `A${n}-A${p}`);
    sc(7, fields.ph);
    sc(8, fields.haerte);
    sc(9, fields.druck);
    sc(10, fields.bemerkung || '');

  } else if (key === 'heizung') {
    sc(2, z);
    sc(3, fields.druck);
    sf(4, `B${n}-B${p}`);
    sf(5, `D${n}/(A${n}-A${p})`);
    sf(6, `E${n}/24`);
    sf(7, `(B${n}-$B$4)/(A${n}-$A$4)`);
    sf(8, `A${n}-A${p}`);
    sc(9, fields.bemerkung || '');

  } else if (key === 'maschinen') {
    sc(2, fields.kategorie || '');
    sc(3, fields.thema || '');
    sc(4, z);
    sc(5, fields.kosten);
  }
}

function extendTable(ws, sheetName, n) {
  const tname = TABLE_MAP[sheetName];
  if (!tname || !ws['!tables']) return;
  const t = ws['!tables'].find(t => t.name === tname);
  if (t) {
    const rng = XLSX.utils.decode_range(t.ref);
    rng.e.r = n - 1;
    t.ref = XLSX.utils.encode_range(rng);
  }
}

// ── App state ────────────────────────────────────────────────

let tabIdx         = 0;
let katCategories  = [];
let workbookIs1904 = false;
let editRowNum     = null;   // null = neuer Eintrag, Zahl = Bearbeitungsmodus
let recentCount    = 10;     // Anzahl angezeigter Einträge

// ── Setup screen ─────────────────────────────────────────────

function renderSetup() {
  document.getElementById('root').innerHTML = `
    <div class="setup">
      <div class="setup-icon"><img src="icon.svg" alt="Messdaten"></div>
      <h1>Messdaten</h1>
      <p>Einmalig mit Dropbox verbinden,<br>danach öffnet die App direkt.</p>
      <button class="btn-primary" style="width:100%;max-width:320px" onclick="startAuth()">
        Mit Dropbox verbinden →
      </button>
    </div>
  `;
}

// ── App shell ────────────────────────────────────────────────

function renderApp() {
  document.getElementById('root').innerHTML = `
    <div id="app">
      <div class="app-header">
        <img src="icon.svg" alt="">
        <span class="app-header-title">Messdaten</span>
      </div>
      <div class="tab-bar" id="tab-bar"></div>
      <div class="scroll" id="scroll">
        <div class="card" id="form-card"></div>
        <div class="recent-section" id="recent-section">
          <div class="recent-header">LETZTE EINTRÄGE <span class="recent-hint">Eintrag antippen zum Bearbeiten</span></div>
          <div class="recent-table" id="recent-table">Lade…</div>
        </div>
        <div class="footer-links">
          <button class="link-btn" id="update-btn" style="color:var(--blue)" onclick="applyUpdate()">🔄 Aktualisieren</button>
          <button class="link-btn" onclick="disconnect()">Dropbox trennen</button>
          <span class="app-version">${APP_VERSION}</span>
        </div>
      </div>
    </div>
  `;
  renderTabBar();
  renderForm();
}

function renderTabBar() {
  document.getElementById('tab-bar').innerHTML = TABS.map((t, i) =>
    `<button class="tab-btn${i === tabIdx ? ' active' : ''}" onclick="switchTab(${i})">${t.label}</button>`
  ).join('');
}

function switchTab(i) {
  tabIdx = i;
  editRowNum = null;
  recentCount = 10;
  renderTabBar();
  renderForm();
  loadRecent();
}

// ── Form ─────────────────────────────────────────────────────

function renderForm() {
  const tab    = TABS[tabIdx];
  const today  = new Date().toISOString().slice(0, 10);
  const isEdit = editRowNum !== null;

  let html = `
    <div class="field-group">
      <label>Datum</label>
      <input type="date" id="f-datum" value="${today}"${isEdit ? '' : ` max="${today}"`}>
    </div>
  `;
  for (const f of tab.fields) {
    const reqStar = f.req ? ' <span class="req">*</span>' : '';
    if (f.key === 'kategorie') {
      html += `
        <div class="field-group">
          <label>${f.label}${reqStar}</label>
          <input id="f-kategorie" type="text" autocomplete="off"
                 autocorrect="off" spellcheck="false"
                 oninput="onKatInput(this)">
          <div id="kat-bar" class="suggestion-bar" style="display:none"></div>
        </div>
      `;
    } else {
      const mode = f.type === 'decimal' ? 'decimal' : 'text';
      html += `
        <div class="field-group">
          <label>${f.label}${reqStar}</label>
          <input id="f-${f.key}" type="text" inputmode="${mode}"
                 placeholder="${f.req ? '' : 'optional'}">
        </div>
      `;
    }
  }

  if (isEdit) {
    html += `<div class="edit-banner">✏️ Zeile ${editRowNum} wird bearbeitet</div>`;
    html += `<button class="btn-secondary" onclick="cancelEdit()">Abbrechen</button>`;
  }
  html += `
    <button class="btn-primary" onclick="onSubmit()">${isEdit ? 'Änderung speichern' : 'Eintragen'}</button>
    <div class="status" id="status"></div>
  `;
  document.getElementById('form-card').innerHTML = html;
}

// ── Kategorie autocomplete ────────────────────────────────────

function onKatInput(inp) {
  const prefix = inp.value.trim().toLowerCase();
  const bar    = document.getElementById('kat-bar');
  if (!bar) return;
  if (!prefix || !katCategories.length) { bar.style.display = 'none'; return; }
  const matches = katCategories.filter(c => c.toLowerCase().startsWith(prefix));
  if (!matches.length) { bar.style.display = 'none'; return; }
  bar.innerHTML = matches.map(c => {
    const safe = c.replace(/&/g, '&amp;').replace(/'/g, '&#39;');
    return `<button class="suggestion-chip" onclick="pickKat('${safe}')">${safe}</button>`;
  }).join('');
  bar.style.display = 'flex';
}

function pickKat(cat) {
  const inp = document.getElementById('f-kategorie');
  if (inp) { inp.value = cat; inp.focus(); }
  const bar = document.getElementById('kat-bar');
  if (bar) bar.style.display = 'none';
}

// ── Render recent table (HTML) ────────────────────────────────

function renderRecentTable(tab, rows, rowNums, hasMore) {
  const hdrs       = tab.headers;
  const isMaschinen = tab.key === 'maschinen';

  const colClass = (i) => {
    if (i === 0) return 'col-date';
    if (isMaschinen && i === 1) return 'col-kat';
    if (isMaschinen && i === 2) return 'col-thema';
    if (isMaschinen && i === 3) return 'col-num col-kosten';
    return 'col-num';
  };

  let html = '<table class="rt"><thead><tr>';
  hdrs.forEach((h, i) => {
    html += `<th class="${colClass(i)}">${h}</th>`;
  });
  html += '</tr></thead><tbody>';

  // Neueste zuerst
  for (let i = rows.length - 1; i >= 0; i--) {
    const row    = rows[i];
    const rNum   = rowNums[i];
    const active = rNum === editRowNum;
    html += `<tr class="rt-row${active ? ' rt-editing' : ''}" data-row="${rNum}" onclick="startEdit(${rNum})">`;
    row.forEach((v, ci) => {
      html += `<td class="${colClass(ci)}">${v ?? '–'}</td>`;
    });
    html += '</tr>';
  }

  html += '</tbody></table>';

  if (hasMore) {
    html += `<div class="more-bar"><button class="more-btn" onclick="loadMore()">Mehr anzeigen ↑</button></div>`;
  }

  return html;
}

// ── Load recent ──────────────────────────────────────────────

async function loadRecent() {
  const tbl = document.getElementById('recent-table');
  if (tbl) tbl.textContent = 'Lade…';
  const tab = TABS[tabIdx];
  try {
    const token = await getToken();
    const buf   = await dbDownload(token);
    const wb    = XLSX.read(new Uint8Array(buf), { type: 'array' });
    workbookIs1904 = wb.Workbook?.WBProps?.date1904 ?? false;
    const ws    = wb.Sheets[tab.sheet];
    if (!ws) throw new Error(`Tabellenblatt "${tab.sheet}" nicht gefunden.`);

    const { rows, rowNums, hasMore } = recentRows(ws, tab.key, recentCount);

    if (tab.key === 'maschinen') {
      const last = xlLastRow(ws);
      katCategories = [...new Set(
        Array.from({ length: Math.max(0, last - 1) }, (_, i) => xlGet(ws, i + 2, 2))
          .filter(Boolean).map(String)
      )].sort();
    }

    if (tbl) tbl.innerHTML = renderRecentTable(tab, rows, rowNums, hasMore);

  } catch (e) {
    if (tbl) { tbl.textContent = '❌ ' + e.message; }
  }
}

function loadMore() {
  recentCount += 10;
  loadRecent();
}

// ── Edit existing row ─────────────────────────────────────────

async function startEdit(rowNum) {
  const tab    = TABS[tabIdx];
  const status = document.getElementById('status');
  if (status) { status.textContent = '⏳ Lade…'; status.className = 'status'; }

  try {
    const token = await getToken();
    const buf   = await dbDownload(token);
    const wb    = XLSX.read(new Uint8Array(buf), { type: 'array' });
    workbookIs1904 = wb.Workbook?.WBProps?.date1904 ?? false;
    const ws    = wb.Sheets[tab.sheet];
    if (!ws) throw new Error('Blatt nicht gefunden');

    const { dateStr, fields } = readRowFields(ws, tab.key, rowNum);

    editRowNum = rowNum;
    renderForm();

    const datumEl = document.getElementById('f-datum');
    if (datumEl && dateStr) datumEl.value = dateStr;

    for (const f of tab.fields) {
      const el = document.getElementById('f-' + f.key);
      if (el && fields[f.key] != null) {
        el.value = String(fields[f.key]);
      }
    }

    updateEditHighlight();
    document.getElementById('form-card')?.scrollIntoView({ behavior: 'smooth', block: 'start' });

  } catch (e) {
    if (status) { status.textContent = '❌ ' + e.message; status.className = 'status err'; }
  }
}

function cancelEdit() {
  editRowNum = null;
  renderForm();
  updateEditHighlight();
}

function updateEditHighlight() {
  document.querySelectorAll('.rt-row').forEach(tr => {
    tr.classList.toggle('rt-editing', +tr.dataset.row === editRowNum);
  });
}

// ── Submit ───────────────────────────────────────────────────

async function onSubmit() {
  const tab    = TABS[tabIdx];
  const isEdit = editRowNum !== null;
  const status = document.getElementById('status');
  const setStatus = (msg, ok = true) => {
    if (!status) return;
    status.textContent = msg;
    status.className   = 'status ' + (ok ? 'ok' : 'err');
  };

  for (const f of tab.fields) {
    if (f.req && !document.getElementById('f-' + f.key)?.value.trim()) {
      setStatus('⚠️ ' + f.label + ' erforderlich', false);
      return;
    }
  }

  const dateStr = document.getElementById('f-datum').value;
  const [yr, mo, dy] = dateStr.split('-').map(Number);
  const date = new Date(yr, mo - 1, dy, 12, 0, 0);

  const fields = {};
  for (const f of tab.fields) {
    let v = document.getElementById('f-' + f.key)?.value.trim() ?? '';
    if (f.type === 'decimal' && v) {
      v = v.replace(',', '.');
      const n = parseFloat(v);
      fields[f.key] = isNaN(n) ? null : n;
    } else {
      fields[f.key] = v || null;
    }
  }

  setStatus('⏳ Speichern…');
  try {
    const token = await getToken();
    const buf   = await dbDownload(token);
    const wb    = XLSX.read(new Uint8Array(buf), { type: 'array' });
    workbookIs1904 = wb.Workbook?.WBProps?.date1904 ?? false;
    const ws    = wb.Sheets[tab.sheet];
    if (!ws) throw new Error(`Tabellenblatt "${tab.sheet}" nicht gefunden.`);

    const epoch      = workbookIs1904 ? new Date(Date.UTC(1904, 0, 1)) : new Date(Date.UTC(1899, 11, 30));
    const dateSerial = Math.round((date - epoch) / 86400000);

    let savedRow;
    if (isEdit) {
      const n = editRowNum;
      const p = n - 1;
      writeRow(ws, tab.key, p, n, dateSerial, fields);
      savedRow = n;
    } else {
      const p = xlLastRow(ws);
      const n = p + 1;
      writeRow(ws, tab.key, p, n, dateSerial, fields);
      extendTable(ws, tab.sheet, n);
      savedRow = n;
    }

    const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    await dbUpload(token, new Uint8Array(out));

    const successMsg = isEdit
      ? '✅ Zeile ' + savedRow + ' aktualisiert → Dropbox'
      : '✅ Zeile ' + savedRow + ' gespeichert → Dropbox';

    if (isEdit) {
      editRowNum = null;
      renderForm();
      const s = document.getElementById('status');
      if (s) { s.textContent = successMsg; s.className = 'status ok'; }
    } else {
      for (const f of tab.fields) {
        const el = document.getElementById('f-' + f.key);
        if (el) el.value = '';
      }
      const bar = document.getElementById('kat-bar');
      if (bar) bar.style.display = 'none';
      setStatus(successMsg);
    }

    loadRecent();

  } catch (e) {
    setStatus('❌ ' + e.message.slice(0, 100), false);
  }
}

// ── Init ─────────────────────────────────────────────────────

async function init() {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }

  if (location.search.includes('code=')) {
    await handleCallback();
  }

  if (isConnected()) {
    renderApp();
    loadRecent();
  } else {
    renderSetup();
  }
}

init();
