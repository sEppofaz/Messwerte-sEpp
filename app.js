'use strict';

// ── Config ────────────────────────────────────────────────────
const DROPBOX_XLS_PATH  = '/Apps/Claude/Messdaten/Messdaten sEpp-Claude.xlsx';
const DROPBOX_JSON_PATH = '/Apps/Claude/Messdaten/messdaten.json';
const APP_KEY           = 's2ggv6zysmzn7fa';
const APP_VERSION       = 'v6';

const TABS = [
  {
    key: 'strom', label: '⚡ Strom',
    fields: [
      { key: 'zaehler',   label: 'Zählerstand (kWh)',     type: 'decimal', req: true  },
      { key: 'bemerkung', label: 'Bemerkung',              type: 'text',    req: false },
    ],
    headers: ['Datum', 'Zähler neu', 'Zähler ges.', 'Δ kWh', 'kWh/Tag'],
  },
  {
    key: 'pv', label: '☀️ PV',
    fields: [
      { key: 'zaehler',   label: 'Zähler gesamt (kWh)',   type: 'decimal', req: true  },
      { key: 'pv1',       label: 'PV1 Zähler (optional)', type: 'decimal', req: false },
      { key: 'bemerkung', label: 'Bemerkung',              type: 'text',    req: false },
    ],
    headers: ['Datum', 'Zähler', 'Δ kWh', 'kWh/Tag', 'PV1/2%'],
  },
  {
    key: 'wasser', label: '💧 Wasser',
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
    key: 'heizung', label: '🔥 Heizung',
    fields: [
      { key: 'zaehler',   label: 'Volllast-Stunden',       type: 'decimal', req: true  },
      { key: 'druck',     label: 'Druck (bar)',             type: 'decimal', req: false },
      { key: 'bemerkung', label: 'Bemerkung',               type: 'text',    req: false },
    ],
    headers: ['Datum', 'Std.', 'Δ Std.', 'Std/Tag'],
  },
  {
    key: 'wallbox', label: '🔌 Wallbox',
    fields: [
      { key: 'zaehler',   label: 'Zählerstand (kWh)',      type: 'decimal', req: true  },
      { key: 'bemerkung', label: 'Bemerkung',               type: 'text',    req: false },
    ],
    headers: ['Datum', 'Zähler', 'Δ kWh', 'kWh/Tag'],
  },
  {
    key: 'maschinen', label: '🔧 Masch.',
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
  u.search = ''; u.hash = '';
  if (u.pathname.endsWith('index.html')) u.pathname = u.pathname.slice(0, -10);
  if (!u.pathname.endsWith('/'))         u.pathname += '/';
  return u.href;
}

async function startAuth() {
  const redirectUri = canonicalUrl();
  const { verifier, challenge } = await pkce();
  sessionStorage.setItem('pkce_verifier', verifier);
  sessionStorage.setItem('redirect_uri',  redirectUri);
  const p = new URLSearchParams({
    response_type: 'code', client_id: APP_KEY,
    redirect_uri: redirectUri,
    code_challenge: challenge, code_challenge_method: 'S256',
    token_access_type: 'offline',
  });
  location.href = AUTH_URL + '?' + p;
}

async function handleCallback() {
  const code = new URLSearchParams(location.search).get('code');
  if (!code) return;
  const verifier    = sessionStorage.getItem('pkce_verifier');
  const redirectUri = sessionStorage.getItem('redirect_uri') || canonicalUrl();
  if (!verifier) {
    alert('Auth-Fehler: Sitzungsdaten fehlen. Bitte erneut verbinden.');
    history.replaceState({}, '', location.pathname);
    return;
  }
  try {
    const r = await fetch(TOKEN_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        code, grant_type: 'authorization_code', client_id: APP_KEY,
        redirect_uri: redirectUri, code_verifier: verifier,
      }),
    });
    if (!r.ok) throw new Error(await r.text());
    const d = await r.json();
    localStorage.setItem('dropbox_access_token', d.access_token);
    localStorage.setItem('dropbox_refresh_token', d.refresh_token);
    localStorage.setItem('dropbox_expires',       Date.now() + d.expires_in * 1000);
  } catch (e) { alert('Token-Fehler: ' + e.message); }
  history.replaceState({}, '', location.pathname);
}

function isConnected() { return !!localStorage.getItem('dropbox_refresh_token'); }

function disconnect() {
  ['dropbox_access_token', 'dropbox_refresh_token', 'dropbox_expires']
    .forEach(k => localStorage.removeItem(k));
  _data = null;
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
      grant_type: 'refresh_token',
      refresh_token: localStorage.getItem('dropbox_refresh_token'),
      client_id: APP_KEY,
    }),
  });
  if (!r.ok) throw new Error('Token abgelaufen – bitte neu verbinden.');
  const d = await r.json();
  localStorage.setItem('dropbox_access_token', d.access_token);
  localStorage.setItem('dropbox_expires',       Date.now() + d.expires_in * 1000);
  return d.access_token;
}

// ── Date helpers ─────────────────────────────────────────────

function fmtDateStr(s) {
  if (!s) return '–';
  const [y, m, d] = s.split('-');
  return `${d}.${m}.${y.slice(2)}`;
}

function dateDiffDays(a, b) {
  return (new Date(b) - new Date(a)) / 86400000;
}

function safeRound(v, d = 2) {
  const n = typeof v === 'number' ? v : parseFloat(v);
  if (isNaN(n)) return v ?? '–';
  return d === 0 ? String(Math.round(n)) : n.toFixed(d);
}

function serialToIso(serial, is1904) {
  if (serial == null) return null;
  if (typeof serial === 'string') {
    const m = serial.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    if (/^\d{4}-\d{2}-\d{2}$/.test(serial)) return serial;
    return null;
  }
  if (typeof serial !== 'number') return null;
  const offset = is1904 ? 24107 : 25569;
  const ms = Math.round((serial - offset) * 86400) * 1000;
  return new Date(ms).toISOString().slice(0, 10);
}

// ── Dropbox JSON I/O ─────────────────────────────────────────

async function jsonDownload(token) {
  const r = await fetch(CONTENT + '/files/download', {
    method: 'POST',
    headers: {
      Authorization:     'Bearer ' + token,
      'Dropbox-API-Arg': JSON.stringify({ path: DROPBOX_JSON_PATH }),
    },
  });
  if (r.status === 409) return null; // file not found
  if (!r.ok) {
    let detail;
    try { detail = (await r.json()).error_summary; } catch { detail = r.status; }
    throw new Error('Dropbox ' + r.status + ': ' + detail);
  }
  return r.json();
}

async function jsonUpload(token, data) {
  const r = await fetch(CONTENT + '/files/upload', {
    method: 'POST',
    headers: {
      Authorization:     'Bearer ' + token,
      'Content-Type':    'application/octet-stream',
      'Dropbox-API-Arg': JSON.stringify({ path: DROPBOX_JSON_PATH, mode: 'overwrite', autorename: false }),
    },
    body: new TextEncoder().encode(JSON.stringify(data)),
  });
  if (!r.ok) throw new Error('Upload fehlgeschlagen: ' + r.status);
}

// ── Excel download + helpers (migration only) ─────────────────

async function xlsDownload(token) {
  const r = await fetch(CONTENT + '/files/download', {
    method: 'POST',
    headers: {
      Authorization:     'Bearer ' + token,
      'Dropbox-API-Arg': JSON.stringify({ path: DROPBOX_XLS_PATH }),
    },
  });
  if (!r.ok) throw new Error('Excel-Download fehlgeschlagen: ' + r.status);
  return r.arrayBuffer();
}

function xlGet(ws, row, col) {
  const cell = ws[XLSX.utils.encode_cell({ r: row - 1, c: col - 1 })];
  return cell ? cell.v : undefined;
}

function xlLastRow(ws) {
  if (!ws['!ref']) return 1;
  const rng = XLSX.utils.decode_range(ws['!ref']);
  for (let r = rng.e.r; r >= 1; r--) {
    const c = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    if (c && c.v !== undefined) return r + 1;
  }
  return 1;
}

// ── App state ────────────────────────────────────────────────

let tabIdx        = 0;
let katCategories = [];
let editIdx       = null;
let recentCount   = 10;
let _data         = null;

function emptyData() {
  return { v: 1, strom: [], pv: [], wasser: [], heizung: [], wallbox: [], maschinen: [] };
}

// ── Data load / save ─────────────────────────────────────────

async function loadData() {
  const token = await getToken();
  const json  = await jsonDownload(token);
  _data = json || emptyData();
  if (!_data.wallbox) _data.wallbox = [];
  return _data;
}

async function saveData() {
  const token = await getToken();
  await jsonUpload(token, _data);
}

// ── Strom: recalculate zaehler_ges from index onwards ────────

function recalcStromGes(arr, fromIdx) {
  for (let i = Math.max(1, fromIdx); i < arr.length; i++) {
    const prev = arr[i - 1];
    arr[i].zaehler_ges = (prev.zaehler_ges ?? prev.zaehler) + (arr[i].zaehler - prev.zaehler);
  }
}

// ── Recent rows from JSON ────────────────────────────────────

function recentRowsJson(entries, key, count) {
  if (!entries || entries.length === 0) return { rows: [], idxs: [], hasMore: false };

  if (key === 'maschinen') {
    const start = Math.max(0, entries.length - count);
    const slice = entries.slice(start);
    return {
      rows: slice.map(e => [
        fmtDateStr(e.datum),
        e.kategorie ?? '–',
        e.thema     ?? '–',
        e.kosten != null ? safeRound(e.kosten, 2) : '–',
      ]),
      idxs:    slice.map((_, i) => start + i),
      hasMore: start > 0,
    };
  }

  const start = Math.max(0, entries.length - count - 1);
  const slice = entries.slice(start);
  const dd    = { strom: [0, 0, 1], pv: [0, 0, 1], wasser: [0, 1, 2], heizung: [0, 0, 1], wallbox: [0, 0, 1] }[key] ?? [2, 2, 2];
  const rows  = [], idxs = [];

  for (let i = 1; i < slice.length; i++) {
    const cur  = slice[i];
    const prv  = slice[i - 1];
    let delta = null, perDay = null;

    if (key === 'strom') {
      const curG = cur.zaehler_ges ?? cur.zaehler;
      const prvG = prv.zaehler_ges ?? prv.zaehler;
      delta = Math.round((curG - prvG) * 1000) / 1000;
      const days = dateDiffDays(prv.datum, cur.datum);
      if (days > 0) perDay = Math.round(delta / days * 10) / 10;
      rows.push([
        fmtDateStr(cur.datum),
        safeRound(cur.zaehler, 0),
        safeRound(curG, 0),
        safeRound(delta, 0),
        safeRound(perDay, 1),
      ]);

    } else if (key === 'pv') {
      const dz   = cur.zaehler - prv.zaehler;
      delta      = Math.round(dz * 1000) / 1000;
      const days = dateDiffDays(prv.datum, cur.datum);
      if (days > 0) perDay = Math.round(dz / days * 10) / 10;
      let ratio = null;
      if (cur.pv1 != null && prv.pv1 != null && delta > 0) {
        const pct1 = Math.round((cur.pv1 - prv.pv1) / delta * 100);
        ratio = pct1 + '/' + (100 - pct1);
      }
      rows.push([
        fmtDateStr(cur.datum),
        safeRound(cur.zaehler, 0),
        safeRound(delta, 0),
        safeRound(perDay, 1),
        ratio ?? '–',
      ]);

    } else {
      const dz   = parseFloat(cur.zaehler) - parseFloat(prv.zaehler);
      delta      = Math.round(dz * 1000) / 1000;
      const days = dateDiffDays(prv.datum, cur.datum);
      if (days > 0) perDay = Math.round(dz / days * 10 ** dd[2]) / 10 ** dd[2];
      rows.push([
        fmtDateStr(cur.datum),
        safeRound(cur.zaehler, dd[0]),
        safeRound(delta, dd[1]),
        safeRound(perDay, dd[2]),
      ]);
    }

    idxs.push(start + i);
  }

  return { rows: rows.slice(-count), idxs: idxs.slice(-count), hasMore: start > 0 };
}

// ── Migration from Excel ──────────────────────────────────────

async function migrateFromExcel() {
  const btn = document.getElementById('migrate-btn');
  if (btn) btn.textContent = '⏳ Migriere…';
  try {
    const token = await getToken();
    const buf   = await xlsDownload(token);
    const wb    = XLSX.read(new Uint8Array(buf), { type: 'array', cellStyles: true });
    const is1904 = wb.Workbook?.WBProps?.date1904 ?? false;

    const result = emptyData();

    // Strom (data from row 3, col: 1=datum, 2=zaehler_neu, 3=zaehler_ges, 8=bemerkung)
    const wsStrom = wb.Sheets['Strom'];
    if (wsStrom) {
      const last = xlLastRow(wsStrom);
      for (let r = 3; r <= last; r++) {
        const datum = serialToIso(xlGet(wsStrom, r, 1), is1904);
        if (!datum) continue;
        result.strom.push({
          datum,
          zaehler:     xlGet(wsStrom, r, 2) ?? null,
          zaehler_ges: xlGet(wsStrom, r, 3) ?? null,
          bemerkung:   xlGet(wsStrom, r, 8) || null,
        });
      }
    }

    // PV-Werte (data from row 3, col: 1=datum, 4=zaehler, 6=pv1, 17=bemerkung)
    const wsPv = wb.Sheets['PV-Werte'];
    if (wsPv) {
      const last = xlLastRow(wsPv);
      for (let r = 3; r <= last; r++) {
        const datum = serialToIso(xlGet(wsPv, r, 1), is1904);
        if (!datum) continue;
        result.pv.push({
          datum,
          zaehler:   xlGet(wsPv, r, 4)  ?? null,
          pv1:       xlGet(wsPv, r, 6)  ?? null,
          bemerkung: xlGet(wsPv, r, 17) || null,
        });
      }
    }

    // Wasser (data from row 3, col: 1=datum, 2=zaehler, 7=ph, 8=haerte, 9=druck, 10=bemerkung)
    const wsWasser = wb.Sheets['Wasser'];
    if (wsWasser) {
      const last = xlLastRow(wsWasser);
      for (let r = 3; r <= last; r++) {
        const datum = serialToIso(xlGet(wsWasser, r, 1), is1904);
        if (!datum) continue;
        result.wasser.push({
          datum,
          zaehler:   xlGet(wsWasser, r, 2)  ?? null,
          ph:        xlGet(wsWasser, r, 7)  ?? null,
          haerte:    xlGet(wsWasser, r, 8)  ?? null,
          druck:     xlGet(wsWasser, r, 9)  ?? null,
          bemerkung: xlGet(wsWasser, r, 10) || null,
        });
      }
    }

    // Heizung Stunden (data from row 3, col: 1=datum, 2=zaehler, 3=druck, 9=bemerkung)
    const wsHeizung = wb.Sheets['Heizung Stunden'];
    if (wsHeizung) {
      const last = xlLastRow(wsHeizung);
      for (let r = 3; r <= last; r++) {
        const datum = serialToIso(xlGet(wsHeizung, r, 1), is1904);
        if (!datum) continue;
        result.heizung.push({
          datum,
          zaehler:   xlGet(wsHeizung, r, 2) ?? null,
          druck:     xlGet(wsHeizung, r, 3) ?? null,
          bemerkung: xlGet(wsHeizung, r, 9) || null,
        });
      }
    }

    // Maschinen (data from row 2, col: 1=datum, 2=kategorie, 3=thema, 4=zaehler, 5=kosten)
    const wsMasch = wb.Sheets['Maschinen'];
    if (wsMasch) {
      const last = xlLastRow(wsMasch);
      for (let r = 2; r <= last; r++) {
        const datum = serialToIso(xlGet(wsMasch, r, 1), is1904);
        if (!datum) continue;
        result.maschinen.push({
          datum,
          kategorie: xlGet(wsMasch, r, 2) || null,
          thema:     xlGet(wsMasch, r, 3) || null,
          zaehler:   xlGet(wsMasch, r, 4) ?? null,
          kosten:    xlGet(wsMasch, r, 5) ?? null,
        });
      }
    }

    const counts = Object.entries(result)
      .filter(([k]) => k !== 'v')
      .map(([k, v]) => `${k}: ${v.length} Einträge`)
      .join('\n');

    if (!confirm(`Migration bereit:\n\n${counts}\n\nDaten jetzt speichern?`)) {
      if (btn) btn.textContent = '📥 Aus Excel migrieren';
      return;
    }

    _data = result;
    await saveData();
    if (btn) btn.textContent = '✅ Migriert!';
    loadRecent();
  } catch (e) {
    alert('Fehler bei Migration: ' + e.message);
    if (btn) btn.textContent = '📥 Aus Excel migrieren';
  }
}

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
          <button class="link-btn" style="color:var(--blue)" onclick="applyUpdate()">🔄 Aktualisieren</button>
          <button id="migrate-btn" class="link-btn" style="color:var(--blue)" onclick="migrateFromExcel()">📥 Aus Excel migrieren</button>
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
  editIdx = null;
  recentCount = 10;
  renderTabBar();
  renderForm();
  loadRecent();
}

// ── Form ─────────────────────────────────────────────────────

function renderForm() {
  const tab    = TABS[tabIdx];
  const today  = new Date().toISOString().slice(0, 10);
  const isEdit = editIdx !== null;

  let html = `
    <div class="field-group">
      <label>Datum</label>
      <input type="date" id="f-datum" value="${today}"${isEdit ? '' : ` max="${today}"`}>
    </div>`;

  for (const f of tab.fields) {
    const reqStar = f.req ? ' <span class="req">*</span>' : '';
    if (f.key === 'kategorie') {
      html += `
        <div class="field-group">
          <label>${f.label}${reqStar}</label>
          <input id="f-kategorie" type="text" autocomplete="off" autocorrect="off"
                 spellcheck="false" oninput="onKatInput(this)">
          <div id="kat-bar" class="suggestion-bar" style="display:none"></div>
        </div>`;
    } else {
      html += `
        <div class="field-group">
          <label>${f.label}${reqStar}</label>
          <input id="f-${f.key}" type="text"
                 inputmode="${f.type === 'decimal' ? 'decimal' : 'text'}"
                 placeholder="${f.req ? '' : 'optional'}">
        </div>`;
    }
  }

  if (isEdit) {
    html += `<div class="edit-banner">✏️ Eintrag #${editIdx + 1} wird bearbeitet</div>`;
    html += `<button class="btn-secondary" onclick="cancelEdit()">Abbrechen</button>`;
  }
  html += `<button class="btn-primary" onclick="onSubmit()">${isEdit ? 'Änderung speichern' : 'Eintragen'}</button>`;
  html += `<div class="status" id="status"></div>`;
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

// ── Render recent table ───────────────────────────────────────

function renderRecentTable(tab, rows, idxs, hasMore) {
  const hdrs   = tab.headers;
  const isMasch = tab.key === 'maschinen';
  const colClass = (i) => {
    if (i === 0) return 'col-date';
    if (isMasch && i === 1) return 'col-kat';
    if (isMasch && i === 2) return 'col-thema';
    if (isMasch && i === 3) return 'col-num col-kosten';
    return 'col-num';
  };

  let html = '<table class="rt"><thead><tr>';
  hdrs.forEach((h, i) => { html += `<th class="${colClass(i)}">${h}</th>`; });
  html += '</tr></thead><tbody>';

  for (let i = rows.length - 1; i >= 0; i--) {
    const row    = rows[i];
    const idx    = idxs[i];
    const active = idx === editIdx;
    html += `<tr class="rt-row${active ? ' rt-editing' : ''}" data-idx="${idx}" onclick="startEdit(${idx})">`;
    row.forEach((v, ci) => { html += `<td class="${colClass(ci)}">${v ?? '–'}</td>`; });
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
    if (!_data) await loadData();
    const entries = _data[tab.key] || [];

    if (tab.key === 'maschinen') {
      katCategories = [...new Set(entries.map(e => e.kategorie).filter(Boolean))].sort();
    }

    const { rows, idxs, hasMore } = recentRowsJson(entries, tab.key, recentCount);

    if (tbl) {
      if (rows.length === 0) {
        tbl.innerHTML = '<p style="color:var(--label);font-size:13px;padding:4px 0">Noch keine Einträge. Daten via „Aus Excel migrieren" importieren.</p>';
      } else {
        tbl.innerHTML = renderRecentTable(tab, rows, idxs, hasMore);
      }
    }
  } catch (e) {
    if (tbl) tbl.textContent = '❌ ' + e.message;
  }
}

function loadMore() {
  recentCount += 10;
  loadRecent();
}

// ── Edit existing entry ──────────────────────────────────────

async function startEdit(idx) {
  const tab    = TABS[tabIdx];
  const status = document.getElementById('status');
  if (status) { status.textContent = '⏳ Lade…'; status.className = 'status'; }
  try {
    if (!_data) await loadData();
    const entry = (_data[tab.key] || [])[idx];
    if (!entry) throw new Error('Eintrag nicht gefunden');

    editIdx = idx;
    renderForm();

    const datumEl = document.getElementById('f-datum');
    if (datumEl && entry.datum) datumEl.value = entry.datum;

    for (const f of tab.fields) {
      const el = document.getElementById('f-' + f.key);
      if (el && entry[f.key] != null) el.value = String(entry[f.key]);
    }

    updateEditHighlight();
    document.getElementById('form-card')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    if (status) { status.textContent = ''; status.className = 'status'; }
  } catch (e) {
    if (status) { status.textContent = '❌ ' + e.message; status.className = 'status err'; }
  }
}

function cancelEdit() {
  editIdx = null;
  renderForm();
  updateEditHighlight();
}

function updateEditHighlight() {
  document.querySelectorAll('.rt-row').forEach(tr => {
    tr.classList.toggle('rt-editing', +tr.dataset.idx === editIdx);
  });
}

// ── Submit ───────────────────────────────────────────────────

async function onSubmit() {
  const tab    = TABS[tabIdx];
  const isEdit = editIdx !== null;
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
  if (!dateStr) { setStatus('⚠️ Datum erforderlich', false); return; }

  const entry = { datum: dateStr };
  for (const f of tab.fields) {
    let v = document.getElementById('f-' + f.key)?.value.trim() ?? '';
    if (f.type === 'decimal' && v) {
      v = v.replace(',', '.');
      const n = parseFloat(v);
      entry[f.key] = isNaN(n) ? null : n;
    } else {
      entry[f.key] = v || null;
    }
  }

  // Strom: compute zaehler_ges based on previous entry
  if (tab.key === 'strom') {
    const arr     = _data?.strom || [];
    const prevIdx = isEdit ? editIdx - 1 : arr.length - 1;
    const prev    = arr[prevIdx];
    entry.zaehler_ges = prev
      ? (prev.zaehler_ges ?? prev.zaehler) + (entry.zaehler - prev.zaehler)
      : entry.zaehler;
  }

  setStatus('⏳ Speichern…');
  try {
    if (!_data) await loadData();
    const arr = _data[tab.key] || (_data[tab.key] = []);

    if (isEdit) {
      arr[editIdx] = entry;
      if (tab.key === 'strom') recalcStromGes(arr, editIdx + 1);
    } else {
      arr.push(entry);
    }

    await saveData();

    const msg = isEdit ? '✅ Eintrag aktualisiert' : '✅ Eintrag gespeichert';
    if (isEdit) {
      editIdx = null;
      renderForm();
      const s = document.getElementById('status');
      if (s) { s.textContent = msg; s.className = 'status ok'; }
    } else {
      for (const f of tab.fields) {
        const el = document.getElementById('f-' + f.key);
        if (el) el.value = '';
      }
      const bar = document.getElementById('kat-bar');
      if (bar) bar.style.display = 'none';
      setStatus(msg);
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
  if (location.search.includes('code=')) await handleCallback();
  if (isConnected()) {
    renderApp();
    loadRecent();
  } else {
    renderSetup();
  }
}

init();
