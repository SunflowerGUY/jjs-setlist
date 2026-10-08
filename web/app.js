/* JJ's Musicians Setlist Organiser - browser edition.
   A port of jjs_setlist.py. Saved setlists use the same JSON format, so they
   move freely between this and the desktop app (File > Import / Export). */
'use strict';

const APP_VERSION = '1.5.B';            // B = browser (W / M / L on the desktop)
const APP_NAME = "JJ's Musicians Setlist Organiser";
const DONATE_URL = 'https://paypal.me/jellyjazzsoftware';
const APP_WEBSITE = 'https://github.com/SunflowerGUY';

const NUM_SETS = 4;
const MAX_SONGS = 16;
const MINUTES_PER_SONG = 3.5;           // for each set's approximate running time
const DEFAULT_FONT = 15, MIN_FONT = 11, MAX_FONT = 26;

// Spreadsheet header names recognised for each field (case-insensitive).
const COLUMN_ALIASES = {
  title: ['title', 'song', 'song title', 'song name', 'name', 'track'],
  artist: ['artist', 'band', 'performer', 'original artist', 'by'],
  style: ['style', 'genre', 'type'],
  vocalist: ['vocalist', 'vocals', 'vocal', 'singer', 'lead vocal'],
  link: ['link', 'url', 'pdf', 'songsheet', 'song sheet', 'sheet', 'chart'],
};
const HYPERLINK_RE = /^\s*=?\s*HYPERLINK\(\s*"([^"]*)"/i;

const DEFAULT_COLOURS = {
  bg: '#003d4b', fg: '#ffffff', used: '#9e9e9e', select: '#3d8bfd',
  lib_frame: '#afafaf', frame: '#67c6c5', active_bg: '#eef5ff',
};
const COLOUR_NAMES = [
  ['bg', 'Songlist Background'], ['fg', 'Song Title'],
  ['used', 'Songs assigned to a Setlist'], ['select', 'Selected Song'],
  ['lib_frame', 'Songlist Frame'], ['frame', 'Active Setlist Frame'],
  ['active_bg', 'Active Setlist Background'],
];
const COLOUR_VARS = {
  bg: '--lib-bg', fg: '--lib-fg', used: '--lib-used', select: '--lib-select',
  lib_frame: '--lib-frame', frame: '--set-frame', active_bg: '--set-active-bg',
};

const IS_MAC = /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
const MOD = IS_MAC ? '⌘' : 'Ctrl+';
// The helper's launcher on this computer (see README: Getting started).
const LAUNCHER = IS_MAC ? "Start JJ's Setlist.command"
  : /Win/.test(navigator.platform || navigator.userAgent) ? "Start JJ's Setlist.bat" : 'start-jjs-setlist.sh';

// ---------------------------------------------------------------- storage
// Things kept in this browser's localStorage. It can be unavailable (private
// windows, blocked site data), so every access is guarded. While someone is
// signed in to an online account, their copies get a separate area of their own
// ("jjs.u.<account>."), cleared again when they sign out.
const BASE_PREFIX = 'jjs.';
const store = {
  prefix: BASE_PREFIX,
  get(key, fallback, prefix = this.prefix) {
    try {
      const v = localStorage.getItem(prefix + key);
      return v === null ? fallback : JSON.parse(v);
    } catch { return fallback; }
  },
  set(key, value, prefix = this.prefix) {
    try { localStorage.setItem(prefix + key, JSON.stringify(value)); return true; }
    catch { return false; }
  },
  /** Remove everything kept under a prefix (e.g. an account's area on sign-out). */
  clear(prefix) {
    try {
      for (let i = localStorage.length - 1; i >= 0; i--) {
        const k = localStorage.key(i);
        if (k && k.startsWith(prefix)) localStorage.removeItem(k);
      }
    } catch { /* nothing to clear */ }
  },
};

// ---------------------------------------------------------------- helpers
const $ = (sel) => document.querySelector(sel);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g,
  (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fold = (s) => String(s ?? '').toLowerCase();
const plural = (n, word) => `${n} ${word}${n === 1 ? '' : 's'}`;

function cellToText(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function songId(song) {
  return fold(song.title) + '\u0000' + fold(song.artist || '');
}

function extrasOf(song) {
  return ['style', 'vocalist'].map((k) => song[k]).filter(Boolean).join(' · ');
}

function songLabel(song) {
  let text = (song.link ? '📄 ' : '') + song.title;
  if (song.artist) text += `  —  ${song.artist}`;
  const extras = extrasOf(song);
  if (extras) text += `   [${extras}]`;
  return text;
}

function rowsToSongs(rows, links) {
  if (!rows.length) return [];
  const header = rows[0].map((h) => fold(cellToText(h)));
  const cols = {};
  for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
    for (const alias of aliases) {
      const idx = header.indexOf(alias);
      if (idx >= 0) { cols[field] = idx; break; }
    }
  }
  if (!('title' in cols)) cols.title = 0;   // no recognised title column: use the first

  const songs = [];
  rows.slice(1).forEach((row, i) => {
    const song = {};
    for (const [field, idx] of Object.entries(cols)) {
      song[field] = idx < row.length ? cellToText(row[idx]) : '';
    }
    if (!song.title) return;
    if (links && !song.link) {
      // Prefer the link on the song-name cell, else any link in the row.
      const rowLinks = links[i + 1] || [];
      song.link = rowLinks[cols.title] || rowLinks.find(Boolean) || '';
    }
    songs.push(song);
  });
  return songs;
}

function readXlsx(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellFormula: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = [], links = [];
  if (!ws || !ws['!ref']) return { rows, links };
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [], rowLinks = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      row.push(cell ? cell.v : null);
      let target = cell && cell.l ? (cell.l.Target || '') : '';
      if (!target && cell && typeof cell.f === 'string') {
        const m = HYPERLINK_RE.exec(cell.f);
        if (m) target = m[1];
      }
      rowLinks.push(target);
    }
    rows.push(row);
    links.push(rowLinks);
  }
  return { rows, links };
}

function parseCsv(text) {
  text = text.replace(/^﻿/, '');
  const rows = [];
  let row = [], field = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false;
      } else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ',') { row.push(field); field = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += ch;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function decodeText(buffer) {
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buffer); }
  catch { return new TextDecoder('windows-1252').decode(buffer); }
}

function csvField(value) {
  const s = String(value ?? '');
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function safeFilename(name) {
  return name.replace(/[<>:"/\\|?*]+/g, '_').replace(/^[\s.]+|[\s.]+$/g, '');
}

function stem(filename) {
  return String(filename || '').split(/[\\/]/).pop().replace(/\.[^.]+$/, '');
}

function todayText() {
  const d = new Date();                      // e.g. "Thursday 08 October 2026"
  const part = (opts) => d.toLocaleDateString('en-GB', opts);
  return `${part({ weekday: 'long' })} ${part({ day: '2-digit' })} ${part({ month: 'long' })} ${d.getFullYear()}`;
}

function download(filename, text, type = 'text/plain') {
  const blob = new Blob([text], { type: type + ';charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

/** Absolute URL for a songsheet link. Relative paths (local PDFs next to the
    spreadsheet) are taken relative to this page. */
function resolveLink(link) {
  link = (link || '').trim();
  if (!link) return '';
  if (/^[a-z][\w+.-]+:/i.test(link)) return link;
  try { return new URL(link.replace(/\\/g, '/'), location.href).href; } catch { return link; }
}

// Plain-text columns (for the text export): Song Name, Artist, Style, Vocalist.
function songTable(songs, numbered, sizeTo) {
  const cols = [['title', 'SONG NAME', 46], ['artist', 'ARTIST', 26],
    ['style', 'STYLE', 18], ['vocalist', 'VOCALIST', 16]];
  const widths = cols.map(([key, head, cap]) =>
    Math.min(cap, Math.max(head.length, ...(sizeTo || songs).map((s) => (s[key] || '').length))));
  const fit = (t, w) => (t.length <= w ? t : t.slice(0, w - 1) + '…').padEnd(w);
  const indent = numbered ? '      ' : '  ';
  const lines = [indent + cols.map(([, h], i) => fit(h, widths[i])).join('  '),
    indent + widths.map((w) => '-'.repeat(w)).join('  ')];
  songs.forEach((s, i) => {
    const prefix = numbered ? `  ${String(i + 1).padStart(2)}. ` : '  ';
    lines.push((prefix + cols.map(([k], c) => fit(s[k] || '', widths[c])).join('  ')).trimEnd());
  });
  return lines;
}

function textHeading(title, subtitle) {
  const line = `${title}  —  ${subtitle}`;
  return [line, '='.repeat(line.length), todayText(), ''];
}

// ---------------------------------------------------------------- state
const state = {
  library: [],
  dbName: '',                                // spreadsheet name without extension
  dbFile: '',                                // file name, saved into setlists
  dbSource: 'file',                          // 'file' (opened here) or 'drive'
  dbLoaded: '',                              // when it was read (ISO date)
  dbBackup: '',                              // set when the backup copy is in use: why
  driveUrl: '',                              // Google Drive share link (File > Song Database Settings)
  driveFirst: true,                          // load the Drive copy each time the app opens
  filtered: [],
  sets: normaliseSets([]),
  name: '',
  savedKey: '',
  setlistDb: '',
  dirty: false,
  active: 0,
  sel: { list: 'lib', index: -1 },           // selected song: 'lib' or set number
  fontSize: DEFAULT_FONT,
  colours: { ...DEFAULT_COLOURS },
};

/** Fill the state from what this browser kept (in the current store area). */
function loadLocalState() {
  const library = store.get('library', null);
  const working = store.get('working', null);
  const settings = store.get('settings', {});
  useLibraryData(library);
  Object.assign(state, {
    driveUrl: settings.driveUrl || '',
    driveFirst: settings.driveFirst !== false,
    sets: normaliseSets(working?.sets),
    name: working?.name || '',
    savedKey: working?.savedKey || '',
    setlistDb: working?.setlistDb || '',
    dirty: !!working?.dirty,
    fontSize: settings.fontSize || DEFAULT_FONT,
    colours: { ...DEFAULT_COLOURS, ...(settings.colours || {}) },
  });
}

/** A kept song library: { name, file, songs, source, loaded } (or null for none). */
function useLibraryData(library) {
  Object.assign(state, {
    library: Array.isArray(library?.songs) ? library.songs : [],
    dbName: library?.name || '',
    dbFile: library?.file || '',
    dbSource: library?.source || 'file',
    dbLoaded: library?.loaded || '',
  });
}

loadLocalState();

function normaliseSets(sets) {
  sets = (Array.isArray(sets) ? sets : []).map((st) => (Array.isArray(st) ? st : st?.songs || []));
  sets = [...sets, ...Array.from({ length: NUM_SETS }, () => [])].slice(0, NUM_SETS);
  return sets.map((st) => st.filter((s) => s && s.title).slice(0, MAX_SONGS));
}

// ---------------------------------------------------------------- where things are saved
// Three ways, chosen when the app opens:
//   'cloud'   signed in to an online account (Firebase): saved in that account,
//             kept separate from everyone else's
//   'files'   the helper is running (started with LAUNCHER): real files in the
//             project folder - config.json and setlists/*.json, as the desktop app
//   'browser' neither: this browser's own storage
let mode = 'browser';
const files = { on: false, folder: '' };     // on: the helper is running (templates also use it)
let savedSetlists = {};                      // cloud / files: setlist name → data

const SETTINGS_KEYS = ['library_colours', 'web_font_size', 'backup_url', 'drive_first'];
function settingsData() {                    // the same names as config.json (and the desktop app)
  return {
    library_colours: state.colours, web_font_size: state.fontSize,
    backup_url: state.driveUrl || null, drive_first: state.driveFirst,
  };
}

/** Call the helper's API (serve.py). */
async function api(method, path, body, timeout = 8000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    const response = await fetch('/api/' + path, {
      method, cache: 'no-store', signal: controller.signal,
      headers: { 'X-JJS': '1', ...(body === undefined ? {} : { 'Content-Type': body instanceof Blob ? body.type : 'application/json' }) },
      body: body === undefined || body instanceof Blob ? body : JSON.stringify(body),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || `${response.status} ${response.statusText}`);
    return data;
  } catch (err) {
    if (err.name === 'AbortError' || err instanceof TypeError) {
      throw new Error("JJ's Setlist's helper isn't answering. Its window has to stay open while you use the app - "
        + `start it again with “${LAUNCHER}”.`);
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

/** Is the helper running? */
async function findHelper() {
  // The helper only ever runs on this computer (http://localhost:8765/), not on a website.
  if (!/^https?:$/.test(location.protocol) || !/^(localhost|127\.0\.0\.1|\[::1\])$/.test(location.hostname)) return false;
  try {
    const info = await api('GET', 'ping', undefined, 1500);
    if (info.app !== 'jjs-setlist') return false;
    files.on = true;
    files.folder = info.folder || '';
    return true;
  } catch {
    return false;
  }
}

/** Use the helper's files: config.json and the setlists folder. */
async function loadHelperFiles() {
  const [config, lists] = await Promise.all([api('GET', 'config'), api('GET', 'setlists')]);
  applyConfig(config);
  savedSetlists = lists;
}

/** Settings from config.json or an online account (same key names as the desktop app). */
function applyConfig(config) {
  // These are the one source of settings: anything not in them is the default,
  // whatever this browser remembered on its own.
  const colours = config.library_colours || {};
  state.colours = { ...DEFAULT_COLOURS };
  for (const key of Object.keys(DEFAULT_COLOURS)) {
    if (/^#[0-9a-f]{6}$/i.test(colours[key] || '')) state.colours[key] = colours[key];
  }
  state.fontSize = Number.isFinite(config.web_font_size)   // the desktop app's font_size is in points
    ? Math.max(MIN_FONT, Math.min(MAX_FONT, config.web_font_size)) : DEFAULT_FONT;
  state.driveUrl = config.backup_url || '';
  state.driveFirst = 'drive_first' in config ? !!config.drive_first : true;
}

/** Re-read the saved setlists (the desktop app, or another computer, may have saved one meanwhile). */
async function refreshSavedSetlists() {
  try {
    if (mode === 'files') savedSetlists = await api('GET', 'setlists');
    else if (mode === 'cloud') savedSetlists = await cloudSetlists();
    else return;
    renderSaved();
  } catch { /* offline, or the helper was closed: saving will say so */ }
}

function setlists() { return mode === 'browser' ? store.get('setlists', {}) : savedSetlists; }

async function putSetlist(key, data) {
  if (mode === 'cloud') {
    await cloudWrite(cloudRoot().collection('setlists').doc(key), data);
    savedSetlists[key] = data;
  } else if (mode === 'files') {
    await api('PUT', 'setlists/' + encodeURIComponent(key), data);
    savedSetlists[key] = data;
  } else {
    const all = setlists();
    all[key] = data;
    if (!store.set('setlists', all)) {
      throw new Error("This browser wouldn't store the setlist (site data may be blocked, or storage is full).");
    }
  }
}

async function removeSetlist(key) {
  if (mode === 'cloud') {
    await cloudWrite(cloudRoot().collection('setlists').doc(key), null);
    delete savedSetlists[key];
  } else if (mode === 'files') {
    await api('DELETE', 'setlists/' + encodeURIComponent(key));
    delete savedSetlists[key];
  } else {
    const all = setlists();
    delete all[key];
    store.set('setlists', all);
  }
}

function whereSaved() {
  return { cloud: `in ${spaceName()}`, files: 'in the setlists folder', browser: 'in this browser' }[mode];
}

function renderStorage() {
  const el = $('#storageLabel');
  if (mode === 'cloud') {
    el.textContent = cloud.band ? `☁ Saving to ${cloud.band.name}` : '☁ Saving to your account';
    el.title = cloud.band
      ? `Signed in as ${accountName()}.\nThe song library and setlists are ${cloud.band.name}'s, shared with its members.\nSwitch with the band list at the top right.`
      : `Signed in as ${accountName()}.\nYour own song library and setlists, kept in your account - nobody else can see them.`;
  } else if (mode === 'files') {
    el.textContent = '💾 Saving to files in the project folder';
    el.title = `Settings: config.json\nSaved setlists: setlists\\\nFolder: ${files.folder}`;
  } else {
    el.textContent = 'Saving in this browser only';
    el.title = cloud.enabled
      ? 'Sign in to keep your settings and setlists in your own account.'
      : `Start JJ's Setlist with “${LAUNCHER}” to save settings and setlists as files in the project folder.`;
  }
  renderAccount();
}

function saveWorking() {
  store.set('working', {
    sets: state.sets, name: state.name, savedKey: state.savedKey,
    setlistDb: state.setlistDb, dirty: state.dirty,
  });
}

function saveSettings() {
  store.set('settings', { fontSize: state.fontSize, colours: state.colours,
    driveUrl: state.driveUrl, driveFirst: state.driveFirst });
  if (mode === 'cloud') {
    // Colours and text size are always mine; the Google Drive link belongs to the open space.
    const mine = { library_colours: state.colours, web_font_size: state.fontSize };
    const drive = { backup_url: state.driveUrl || null, drive_first: state.driveFirst };
    const failed = (err) => messageBox('Could not save settings', `The settings couldn't be saved:\n\n${err.message}`);
    cloudWrite(cloudUserDoc().collection('data').doc('settings'), cloud.band ? mine : { ...mine, ...drive }, { merge: true }).catch(failed);
    if (cloud.band) cloudWrite(cloudRoot().collection('data').doc('settings'), drive, { merge: true }).catch(failed);
  } else if (mode === 'files') {
    api('PUT', 'config', settingsData())
      .catch((err) => messageBox('Could not save settings', `config.json couldn't be saved:\n\n${err.message}`));
  }
}

/** Keep the song library: in the account when signed in, else in this browser. */
async function saveLibrary(library) {
  if (mode === 'cloud') {
    try {
      await cloudWrite(cloudRoot().collection('data').doc('library'), library);
      return true;
    } catch (err) {
      messageBox('Could not save the song library', `It couldn't be saved to ${spaceName()}:\n\n${err.message}`);
      return false;
    }
  }
  return store.set('library', library);
}

// ---------------------------------------------------------------- online accounts (Firebase)
// Turned on by web/firebase-config.js. Sign-up needs an invitation code from the
// administrator; the code also says which band the new member joins. Each person
// has a personal space (users/{id}/) and can switch to any band they're in
// (bands/{id}/). firestore.rules enforces who may read or change what.
const cloud = {
  enabled: false, auth: null, db: null, user: null, emulator: false,
  admin: false,                              // the app's administrator (admins/{uid} exists)
  bands: [],                                 // the bands I can open: [{ id, name }]
  band: null,                                // the band open now, or null for my own space
  busy: false,                               // signing in / up: don't reload on the auth change yet
};
const CLOUD_TIMEOUT = 4000;                  // ms before a save is reported as "will upload later"
const SITE_URL = 'https://sunflowerguy.github.io/jjs-setlist/';

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error(`${src} couldn't be loaded`));
    document.head.appendChild(s);
  });
}

/** Start Firebase, if it's set up, and find out who (if anyone) is signed in. */
async function initCloud() {
  const config = window.JJS_FIREBASE_CONFIG;
  // Testing: http://localhost:…/?emulator uses Firebase's local emulator instead.
  const emulator = /^(localhost|127\.0\.0\.1)$/.test(location.hostname)
    && new URLSearchParams(location.search).has('emulator');
  if (!config && !emulator) return;
  try {
    for (const part of ['app', 'auth', 'firestore']) await loadScript(`lib/firebase/firebase-${part}-compat.js`);
    const app = firebase.initializeApp(emulator
      ? { apiKey: 'demo-key', authDomain: 'localhost', projectId: 'demo-jjs-setlist' } : config);
    cloud.auth = app.auth();
    cloud.db = app.firestore();
    if (emulator) {
      cloud.auth.useEmulator('http://127.0.0.1:9099', { disableWarnings: true });
      cloud.db.useEmulator('127.0.0.1', 8095);
      cloud.emulator = true;
    }
    // Keep a copy in this browser, so saved work survives a dropped connection.
    await cloud.db.enablePersistence({ synchronizeTabs: true }).catch(() => {});
    cloud.enabled = true;
    cloud.user = await new Promise((resolve) => {
      const stop = cloud.auth.onAuthStateChanged((user) => { stop(); resolve(user); });
    });
    // Signing in or out (here, or in another tab) reopens the app in the right mode.
    cloud.auth.onAuthStateChanged((user) => {
      if (cloud.busy) return;
      if ((user?.uid || null) !== (cloud.user?.uid || null)) location.reload();
    });
  } catch (err) {
    console.error(err);
    cloud.enabled = false;
    status(`Online accounts are unavailable just now: ${err.message}`);
  }
}

const now = () => firebase.firestore.FieldValue.serverTimestamp();
function cloudUserDoc() { return cloud.db.collection('users').doc(cloud.user.uid); }
/** Where the open space's library, settings and setlists live. */
function cloudRoot() { return cloud.band ? cloud.db.collection('bands').doc(cloud.band.id) : cloudUserDoc(); }
function spaceName() { return cloud.band ? cloud.band.name : 'your account'; }
/** This account's area in the browser (its spaces are inside it). */
function accountPrefix(uid = cloud.user.uid) { return `${BASE_PREFIX}u.${uid}.`; }

/** Save (data) or delete (null) a document. Offline, it's kept in this browser
    and uploaded when the connection is back - say so rather than wait.
    Returns 'queued' in that case. */
let cloudQueued = false;                     // the last save is waiting for the connection
async function cloudWrite(ref, data, options) {
  const write = data === null ? ref.delete() : ref.set(data, options || {});
  const slow = new Promise((resolve) => setTimeout(() => resolve('queued'), CLOUD_TIMEOUT));
  try {
    const result = await Promise.race([write, slow]);
    cloudQueued = result === 'queued';
    if (cloudQueued) {
      status(`Saved in this browser - it will be uploaded to ${spaceName()} when the connection is back.`);
      write.then(() => { cloudQueued = false; status(`Your changes have now been uploaded to ${spaceName()}.`); }, () => {});
    }
    return result;
  } catch (err) {
    throw new Error(cloudMessage(err));
  }
}

async function cloudSetlists(root = cloudRoot()) {
  const snap = await root.collection('setlists').get();
  const found = {};
  snap.forEach((doc) => { if (Array.isArray(doc.data().sets)) found[doc.id] = doc.data(); });
  return found;
}

/** Is this login allowed in? The administrator, or signed up with an invitation code. */
async function accountStatus(user = cloud.user) {
  const [adminDoc, regDoc] = await Promise.all([
    cloud.db.collection('admins').doc(user.uid).get(),
    cloud.db.collection('registered').doc(user.uid).get(),
  ]);
  return { admin: adminDoc.exists, registered: adminDoc.exists || regDoc.exists };
}

/** Open the signed-in account: its bands, the space chosen last time, and that space's data. */
async function loadCloudData() {
  const { admin, registered } = await accountStatus();
  if (!registered) {
    const err = new Error('not registered');
    err.code = 'jjs/not-registered';
    throw err;
  }
  cloud.admin = admin;
  // My bands. The administrator can open every band.
  const mine = await cloudUserDoc().collection('bands').get();
  const bands = new Map();
  mine.forEach((d) => bands.set(d.id, d.data().name || d.id));
  if (admin) (await cloud.db.collection('bands').get()).forEach((d) => bands.set(d.id, d.data().name || d.id));
  cloud.bands = [...bands].map(([id, name]) => ({ id, name })).sort((a, b) => fold(a.name).localeCompare(fold(b.name)));
  if (admin) pendingRequests().then((n) => { cloud.requests = n; renderAccount(); });

  // The space chosen last time on this computer - if I'm still in that band.
  const wanted = store.get('space', '', accountPrefix());
  cloud.band = cloud.bands.find((b) => b.id === wanted) || null;
  if (cloud.band) {
    try {
      await cloud.db.collection('bands').doc(cloud.band.id).get();
    } catch {                                 // removed from the band since
      const gone = cloud.band;
      cloud.band = null;
      cloud.bands = cloud.bands.filter((b) => b.id !== gone.id);
      cloudUserDoc().collection('bands').doc(gone.id).delete().catch(() => {});
      store.set('space', '', accountPrefix());
      cloud.notice = ['No longer in the band', `You're no longer a member of ${gone.name}, so your own setlists are open instead.`];
    }
  }
  store.prefix = accountPrefix() + (cloud.band ? `b.${cloud.band.id}.` : '');
  loadLocalState();                          // the unsaved setlist on screen, kept per account and space

  const root = cloudRoot();
  const [mySettings, spaceSettings, libraryDoc, lists] = await Promise.all([
    cloudUserDoc().collection('data').doc('settings').get(),
    root.collection('data').doc('settings').get(),
    root.collection('data').doc('library').get(),
    cloudSetlists(root),
  ]);
  savedSetlists = lists;
  // Colours and text size are always mine; the Google Drive link belongs to the space.
  const mine2 = mySettings.exists ? mySettings.data() : {};
  const space = spaceSettings.exists ? spaceSettings.data() : {};
  applyConfig({ library_colours: mine2.library_colours, web_font_size: mine2.web_font_size,
    ...('backup_url' in space ? { backup_url: space.backup_url } : {}),
    ...('drive_first' in space ? { drive_first: space.drive_first } : {}) });
  useLibraryData(libraryDoc.exists ? libraryDoc.data() : null);
  // A brand-new personal space (nothing saved yet).
  return !cloud.band && !mySettings.exists && !libraryDoc.exists && !Object.keys(lists).length;
}

/** Friendly words for Firebase's error codes. */
function cloudMessage(err) {
  const code = err?.code || '';
  return {
    'auth/invalid-credential': 'The email address or password is wrong.',
    'auth/wrong-password': 'The email address or password is wrong.',
    'auth/user-not-found': 'The email address or password is wrong.',
    'auth/invalid-email': "That doesn't look like an email address.",
    'auth/missing-email': 'Please type your email address.',
    'auth/missing-password': 'Please type your password.',
    'auth/email-already-in-use': 'There is already an account with that email address. Sign in instead (or use “Forgot password?”).',
    'auth/weak-password': 'Please choose a longer password - at least 8 characters.',
    'auth/password-does-not-meet-requirements': 'Please choose a longer password - at least 8 characters.',
    'auth/too-many-requests': 'Too many tries. Please wait a few minutes, then try again.',
    'auth/network-request-failed': "Couldn't connect. Check the internet connection and try again.",
    'auth/popup-closed-by-user': 'The Google sign-in window was closed before signing in.',
    'auth/cancelled-popup-request': 'The Google sign-in window was closed before signing in.',
    'auth/popup-blocked': 'The browser blocked the Google sign-in window. Allow pop-ups for this page, then try again.',
    'auth/account-exists-with-different-credential': 'That email address already has an account with a password. Sign in with your email and password.',
    'auth/requires-recent-login': 'For safety, please sign out, sign in again, and then try once more.',
    'auth/user-disabled': 'This account has been switched off.',
    'auth/operation-not-allowed': 'This way of signing in has not been switched on for JJ\'s Setlist yet.',
    'jjs/not-registered': "There's no JJ's Setlist account for that login. New members sign up with an invitation code from the administrator.",
    'permission-denied': "That isn't allowed for your account (if an invitation code was used, it may have been withdrawn).",
    'unavailable': "Couldn't connect. Check the internet connection.",
  }[code] || err?.message || String(err);
}

// ---------------------------------------------------------------- invitation codes
const CODE_LETTERS = '23456789ABCDEFGHJKMNPQRSTUVWXYZ';   // no 0/O, 1/I/L - easy to read out

/** A new random code like JAZZ-7K2Q-M4XP: 12 characters, too many to guess. */
function newInviteCode() {
  const bytes = crypto.getRandomValues(new Uint8Array(12));
  const chars = [...bytes].map((b) => CODE_LETTERS[b % CODE_LETTERS.length]).join('');
  return `${chars.slice(0, 4)}-${chars.slice(4, 8)}-${chars.slice(8)}`;
}

/** Tidy a typed code: capitals, no spaces, dashes every 4 characters. */
function normaliseCode(text) {
  const c = String(text || '').toUpperCase().replace(/[^0-9A-Z]/g, '');
  return c.match(/.{1,4}/g)?.join('-') || '';
}

/** The band an invitation code is for, or null if it isn't a current code. */
async function lookUpCode(text) {
  const code = normaliseCode(text);
  if (code.length !== 14) return null;
  try {
    const doc = await cloud.db.collection('invites').doc(code).get();
    return doc.exists ? { code, bandId: doc.data().bandId, bandName: doc.data().bandName } : null;
  } catch {
    return null;
  }
}

/** A band name's matching key: "The Jelly-Jazz!" and "jelly jazz" are the same band. */
function bandKey(name) {
  return String(name || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/^the\s+/, '').replace(/[^a-z0-9]/g, '');
}

/** Wire a code box: shows which band the code is for as it's typed. */
function watchCodeBox(input, output, onResult) {
  let timer = null, last = '';
  const check = async () => {
    const code = normaliseCode(input.value);
    if (code === last) return;
    last = code;
    if (code.length < 14) { output.textContent = ''; output.className = 'drive-result'; onResult(null); return; }
    output.textContent = 'Checking the code…';
    output.className = 'drive-result';
    const found = await lookUpCode(code);
    if (normaliseCode(input.value) !== code) return;           // typed on meanwhile
    output.textContent = found ? `✓  This code is for: ${found.bandName}` : '✗  That code isn\'t recognised. Check it with whoever sent it.';
    output.className = 'drive-result ' + (found ? 'ok' : 'bad');
    onResult(found);
  };
  input.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(check, 350); });
  input.addEventListener('blur', () => { input.value = normaliseCode(input.value) || input.value; });
  check();
}

/** Signed up with a code: record it, join the band, and open the band next time. */
async function completeSignUp(user, invite) {
  const db = cloud.db, email = user.email || '';
  await db.collection('registered').doc(user.uid).set({ email, code: invite.code, bandId: invite.bandId, created: now() });
  await joinBand(user, invite);
}

async function joinBand(user, invite) {
  const db = cloud.db, email = user.email || '';
  await db.collection('bands').doc(invite.bandId).collection('members').doc(user.uid)
    .set({ email, name: user.displayName || '', code: invite.code, joined: now() });
  await db.collection('users').doc(user.uid).collection('bands').doc(invite.bandId).set({ name: invite.bandName, joined: now() });
  try { localStorage.setItem(accountPrefix(user.uid) + 'space', JSON.stringify(invite.bandId)); } catch { /* no storage */ }
  try { sessionStorage.setItem('jjs.welcome', invite.bandName); } catch { /* no storage */ }
}

// ---------------------------------------------------------------- account: sign in, sign up, menu
function accountName() {
  return cloud.user?.email || cloud.user?.displayName || 'your account';
}

function renderAccount() {
  const el = $('#account');
  if (!el) return;
  el.hidden = !cloud.enabled;
  $('#helpAccountItem').hidden = $('#helpAccountRule').hidden = !cloud.enabled;   // Help ▸ Your Online Account
  renderSpacePicker();
  if (!cloud.enabled) return;
  const badge = cloud.admin && cloud.requests ? ` <span class="badge" title="Invitation requests waiting">${cloud.requests}</span>` : '';
  el.innerHTML = cloud.user
    ? `<button type="button" id="accountBtn" title="Your account">👤 ${esc(accountName())} ▾${badge}</button>`
    : '<button type="button" id="accountBtn" class="primary" title="Sign in, or sign up with an invitation code">Sign in</button>';
  $('#accountBtn').onclick = () => (cloud.user ? accountMenu() : signIn());
}

/** The band switcher next to the account button: my own setlists, or a band. */
function renderSpacePicker() {
  const sel = $('#spacePicker');
  if (!sel) return;
  const show = mode === 'cloud' && cloud.bands.length > 0;
  sel.hidden = !show;
  if (!show) return;
  sel.innerHTML = '<option value="">My own setlists</option>'
    + cloud.bands.map((b) => `<option value="${esc(b.id)}">🎵 ${esc(b.name)}</option>`).join('');
  sel.value = cloud.band?.id || '';
  sel.onchange = async () => {
    const id = sel.value;
    if (!await confirmDiscard()) { sel.value = cloud.band?.id || ''; return; }
    store.set('space', id, accountPrefix());
    status('Opening…');
    location.reload();
  };
}

function accountMenu() {
  const r = $('#accountBtn').getBoundingClientRect();
  const hasPassword = (cloud.user.providerData || []).some((p) => p.providerId === 'password');
  const inBand = mode === 'cloud' && cloud.band;
  showContextMenu(r.left, r.bottom + 4, [
    { label: `Signed in as ${accountName()}${cloud.admin ? '  ·  administrator' : ''}`, note: true, plain: true },
    ...(cloud.admin ? [
      { label: 'Bands & invitation codes…', action: adminBands },
      { label: `Invitation requests${cloud.requests ? ` (${cloud.requests} waiting)` : ''}…`, action: adminRequests },
      '-',
    ] : []),
    { label: 'Join a band with an invitation code…', action: joinWithCode },
    ...(inBand ? [{ label: `Copy my own setlists into ${cloud.band.name}…`, action: copyMineIntoBand }] : []),
    { label: `Copy setlists from this computer into ${inBand ? cloud.band.name : 'my account'}…`, action: () => offerImport(false) },
    ...(inBand && !cloud.admin ? [{ label: `Leave ${cloud.band.name}…`, action: leaveBand }] : []),
    '-',
    ...(hasPassword ? [{ label: 'Change my password…', action: changePassword }] : []),
    { label: 'Sign out', action: signOut },
    { label: 'Delete my account…', action: deleteAccount },
  ]);
}

/** The sign-in window: members sign in; new people sign up with an invitation
    code (which says which band they join), or ask for one. Loops until done. */
async function signIn(view = 'signin', message = '', email = '', invite = null) {
  let tone = 'bad', codeText = invite?.code || '';
  for (;;) {
    const msg = `<p id="acMsg" class="drive-result ${tone}"${message ? '' : ' hidden'}>${esc(message)}</p>`;
    const views = {
      signin: {
        title: 'Sign in to JJ\'s Setlist',
        html: `<div class="account-form">${msg}
          <h4 class="ac-head">Already a member?</h4>
          <label for="acEmail">Email address</label>
          <input id="acEmail" type="email" autocomplete="username" spellcheck="false" value="${esc(email)}">
          <label for="acPassword">Password</label>
          <input id="acPassword" type="password" autocomplete="current-password">
          <p class="small"><button type="button" class="link-button" id="acForgot">Forgot password?</button></p>
          <button type="button" class="google-btn" id="acGoogle"><b>G</b>  Sign in with Google</button>
          <div class="or"><span>New here?</span></div>
          <label for="acCode">Invitation code</label>
          <div class="drive-row"><input id="acCode" type="text" autocomplete="off" spellcheck="false" placeholder="e.g. JAZZ-7K2Q-M4XP" value="${esc(codeText)}">
            <button type="button" id="acSignUp" disabled>Sign up with this code</button></div>
          <div id="acCodeResult" class="drive-result"></div>
          <p class="small muted">No code? <button type="button" class="link-button" id="acRequest">Request an invitation code</button></p>
        </div>`,
        buttons: [{ label: 'Cancel', value: false }, { label: 'Sign In', value: 'go', primary: true, check: () => {
          const e = $('#acEmail').value.trim(), pw = $('#acPassword').value;
          if (!e) { showMsg('Please type your email address.'); $('#acEmail').focus(); return false; }
          if (!pw) { showMsg('Please type your password.'); $('#acPassword').focus(); return false; }
          fields = { email: e, password: pw };
          return true;
        } }],
      },
      signup: {
        title: `Join ${invite?.bandName || 'a band'} on JJ's Setlist`,
        html: `<div class="account-form">${msg}
          <p class="ac-band">🎵  You're joining <b>${esc(invite?.bandName || '')}</b></p>
          <button type="button" class="google-btn" id="acGoogle"><b>G</b>  Sign up with Google</button>
          <div class="or"><span>or with any email address</span></div>
          <label for="acEmail">Email address</label>
          <input id="acEmail" type="email" autocomplete="username" spellcheck="false" value="${esc(email)}">
          <label for="acPassword">Choose a password</label>
          <input id="acPassword" type="password" autocomplete="new-password">
          <label for="acPassword2">Password again</label>
          <input id="acPassword2" type="password" autocomplete="new-password">
          <p class="muted small">At least 8 characters. Any email address works - it doesn't have to be Gmail.
            A message is sent to check the address. You'll also have your own private setlists, besides the band's.</p>
        </div>`,
        buttons: [{ label: 'Back', value: 'back' }, { label: 'Cancel', value: false },
          { label: 'Create My Account', value: 'go', primary: true, check: () => {
            const e = $('#acEmail').value.trim(), pw = $('#acPassword').value;
            if (!e) { showMsg('Please type your email address.'); $('#acEmail').focus(); return false; }
            if (pw.length < 8) { showMsg('Please choose a password of at least 8 characters.'); $('#acPassword').focus(); return false; }
            if (pw !== $('#acPassword2').value) { showMsg("The two passwords don't match."); return false; }
            fields = { email: e, password: pw };
            return true;
          } }],
      },
      request: {
        title: 'Request an invitation code',
        html: `<div class="account-form">${msg}
          <p>JJ's Setlist is by invitation. Send your details and the administrator will be in touch by email
            - usually with an invitation code for your band.</p>
          <label for="rqName">Your name</label>
          <input id="rqName" type="text" autocomplete="name" maxlength="100">
          <label for="rqEmail">Your email address</label>
          <input id="rqEmail" type="email" autocomplete="email" spellcheck="false" maxlength="200" value="${esc(email)}">
          <label for="rqBand">Band name <span class="muted">(if you're joining a band)</span></label>
          <input id="rqBand" type="text" maxlength="100">
          <label for="rqMessage">Message <span class="muted">(optional)</span></label>
          <textarea id="rqMessage" rows="3" maxlength="1000"></textarea>
        </div>`,
        buttons: [{ label: 'Back', value: 'back' }, { label: 'Cancel', value: false },
          { label: 'Send Request', value: 'go', primary: true, check: () => {
            const name = $('#rqName').value.trim(), e = $('#rqEmail').value.trim();
            if (!name) { showMsg('Please type your name.'); $('#rqName').focus(); return false; }
            if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) { showMsg('Please type your email address, so the administrator can reply.'); $('#rqEmail').focus(); return false; }
            fields = { name, email: e, band: $('#rqBand').value.trim(), message: $('#rqMessage').value.trim() };
            return true;
          } }],
      },
    };
    const v = views[view];
    let fields = null;
    const showMsg = (text) => { const m = $('#acMsg'); m.className = 'drive-result bad'; m.textContent = text; m.hidden = false; };
    const p = dialog(v.title, v.html, v.buttons, true);
    const body = $('#dialogBody');
    body.querySelectorAll('input').forEach((inp) => inp.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      if (inp.id === 'acCode') { if (!$('#acSignUp').disabled) $('#acSignUp').click(); return; }
      [...$('#dialogButtons').children].find((b) => b.classList.contains('primary')).click();
    }));
    $('#acGoogle')?.addEventListener('click', () => dialogFinish('google'));
    $('#acForgot')?.addEventListener('click', () => { fields = { email: $('#acEmail').value.trim() }; dialogFinish('forgot'); });
    $('#acRequest')?.addEventListener('click', () => { email = $('#acEmail')?.value.trim() || email; dialogFinish('to-request'); });
    if (view === 'signin') {
      let found = null;
      watchCodeBox($('#acCode'), $('#acCodeResult'), (f) => { found = f; $('#acSignUp').disabled = !f; });
      $('#acSignUp').addEventListener('click', () => { invite = found; codeText = found.code; dialogFinish('to-signup'); });
      (email ? $('#acPassword') : $('#acEmail')).focus();
    } else if (view === 'signup') {
      $('#acEmail').focus();
    } else {
      $('#rqName').focus();
    }

    const ans = await p;
    if (!ans) return;
    if (ans === 'to-signup') { view = 'signup'; message = ''; continue; }
    if (ans === 'to-request') { view = 'request'; message = ''; continue; }
    if (ans === 'back') { view = 'signin'; message = ''; continue; }
    email = fields?.email || email;
    document.body.classList.add('busy');
    cloud.busy = true;                        // finish the checks before the app reopens
    try {
      if (ans === 'forgot') {
        if (!email) { tone = 'bad'; message = 'Type your email address first, then choose “Forgot password?”.'; continue; }
        await cloud.auth.sendPasswordResetEmail(email);
        tone = 'ok';
        message = `If there's an account for ${email}, an email with a link to choose a new password is on its way. `
          + 'Check your spam folder if it doesn\'t arrive.';
        continue;
      }
      if (view === 'request') {
        await cloud.db.collection('requests').add({ ...fields, created: now() });
        cloud.busy = false;
        await messageBox('Request sent', `Thanks, ${fields.name}. Your request has been sent to the administrator, who'll be in touch at ${fields.email}.`);
        return;
      }
      if (view === 'signup') {
        // The code may have been withdrawn while the window was open.
        if (!await lookUpCode(invite.code)) { tone = 'bad'; message = 'That invitation code has just been withdrawn. Ask the administrator for a new one.'; view = 'signin'; continue; }
        status('Creating your account…');
        let user;
        if (ans === 'google') {
          ({ user } = await cloud.auth.signInWithPopup(new firebase.auth.GoogleAuthProvider()));
          if ((await accountStatus(user)).registered) {          // already a member: join the band
            await joinBand(user, invite);
            location.reload();
            return;
          }
        } else {
          ({ user } = await cloud.auth.createUserWithEmailAndPassword(fields.email, fields.password));
        }
        try {
          await completeSignUp(user, invite);
        } catch (err) {
          await user.delete().catch(() => {});                  // don't leave a half-made account
          throw err;
        }
        if (ans !== 'google') await user.sendEmailVerification().catch(() => {});
        status('Account created - opening your band…');
        location.reload();
        return;
      }
      // Signing in: only for accounts that signed up with a code (or the administrator).
      status('Signing in…');
      const result = ans === 'google'
        ? await cloud.auth.signInWithPopup(new firebase.auth.GoogleAuthProvider())
        : await cloud.auth.signInWithEmailAndPassword(fields.email, fields.password);
      if (!(await accountStatus(result.user)).registered) {
        // A Google login made just now, without a code: remove it again.
        if (result.additionalUserInfo?.isNewUser) await result.user.delete().catch(() => {});
        await cloud.auth.signOut().catch(() => {});
        tone = 'bad';
        message = cloudMessage({ code: 'jjs/not-registered' });
        continue;
      }
      status('Signed in - opening your account…');
      location.reload();
      return;
    } catch (err) {
      tone = 'bad';
      message = cloudMessage(err);
      status(message);
    } finally {
      cloud.busy = false;
      document.body.classList.remove('busy');
    }
  }
}

/** Signed in already: join another band with its code. */
async function joinWithCode() {
  let found = null;
  const p = dialog('Join a band', `<div class="account-form">
      <p>Type the invitation code the administrator sent you.</p>
      <label for="jnCode">Invitation code</label>
      <input id="jnCode" type="text" autocomplete="off" spellcheck="false" placeholder="e.g. JAZZ-7K2Q-M4XP">
      <div id="jnResult" class="drive-result"></div></div>`,
  [{ label: 'Cancel', value: false }, { label: 'Join Band', value: true, primary: true, check: () => !!found }], true);
  watchCodeBox($('#jnCode'), $('#jnResult'), (f) => { found = f; });
  $('#jnCode').focus();
  if (!await p || !found) return;
  if (cloud.bands.some((b) => b.id === found.bandId) && !cloud.admin) {
    messageBox('Join a band', `You're already in ${found.bandName}. Choose it in the band list at the top right.`);
    return;
  }
  if (!await confirmDiscard()) return;
  try {
    await joinBand(cloud.user, found);
    location.reload();
  } catch (err) {
    messageBox('Could not join the band', cloudMessage(err));
  }
}

async function leaveBand() {
  const band = cloud.band;
  if (!await confirmBox('Leave the band', `Leave ${band.name}?\n\nYou won't see its songs and setlists any more. To come back, you'd need a current invitation code from the administrator.`, 'Leave Band')) return;
  try {
    await cloud.db.collection('bands').doc(band.id).collection('members').doc(cloud.user.uid).delete();
    await cloudUserDoc().collection('bands').doc(band.id).delete();
    store.clear(store.prefix);
    store.set('space', '', accountPrefix());
    location.reload();
  } catch (err) {
    messageBox('Could not leave the band', cloudMessage(err));
  }
}

/** Copy my personal setlists (and, if the band has none, my song library) into the band. */
async function copyMineIntoBand() {
  const band = cloud.band;
  const mine = await cloudSetlists(cloudUserDoc());
  const keys = Object.keys(mine).filter((k) => !(k in savedSetlists));
  const lib = state.library.length ? null : await cloudUserDoc().collection('data').doc('library').get();
  const takeLibrary = lib?.exists && lib.data().songs?.length;
  if (!keys.length && !takeLibrary) { messageBox('Copy my setlists', `All your own setlists are in ${band.name} already.`); return; }
  const what = [keys.length ? `   •  ${plural(keys.length, 'setlist')}` : '', takeLibrary ? `   •  your song library “${lib.data().name}”` : ''].filter(Boolean).join('\n');
  if (!await confirmBox('Copy my setlists', `Copy these into ${band.name}, for all its members?\n\n${what}\n\nYour own copies stay as they are. ${band.name}'s setlists with the same names aren't changed.`, 'Copy Them')) return;
  document.body.classList.add('busy');
  try {
    for (const k of keys) await putSetlist(k, mine[k]);
    if (takeLibrary) { useLibraryData(lib.data()); await saveLibrary(lib.data()); }
    renderAll();
    renderSaved();
    status(`Copied ${plural(keys.length, 'setlist')}${takeLibrary ? ' and your song library' : ''} into ${band.name}.`);
  } catch (err) {
    messageBox('Could not copy everything', err.message);
  } finally {
    document.body.classList.remove('busy');
  }
}

async function signOut() {
  if (!await confirmDiscard()) return;
  // Clear this account's copies from this browser (it may be a shared computer).
  store.clear(accountPrefix());
  status('Signing out…');
  await cloud.auth.signOut();                  // the sign-in listener reopens the app signed out
}

async function changePassword() {
  try {
    await cloud.auth.sendPasswordResetEmail(cloud.user.email);
    messageBox('Change password', `An email with a link to choose a new password has been sent to:\n\n   ${cloud.user.email}\n\nCheck your spam folder if it doesn't arrive.`);
  } catch (err) {
    messageBox('Change password', cloudMessage(err));
  }
}

async function deleteAccount() {
  const mine = await cloudSetlists(cloudUserDoc()).catch(() => ({}));
  const bands = cloud.bands.filter((b) => !cloud.admin);
  if (!await dialog('Delete my account',
    `Delete the account ${accountName()} and everything in it?\n\n`
    + `   •  your own ${plural(Object.keys(mine).length, 'saved setlist')}, song library and settings\n`
    + (bands.length ? `   •  your membership of ${bands.map((b) => b.name).join(', ')} (the bands' setlists stay, for the other members)\n` : '')
    + '\nThis cannot be undone. Export any setlists you want to keep first (File ▸ Export This Setlist).',
  [{ label: 'Cancel', value: false }, { label: 'Delete My Account', value: true, primary: true }])) return;
  document.body.classList.add('busy');
  try {
    const user = cloudUserDoc(), db = cloud.db, uid = cloud.user.uid;
    const batch = db.batch();
    (await user.collection('setlists').get()).forEach((d) => batch.delete(d.ref));
    (await user.collection('bands').get()).forEach((d) => {
      batch.delete(d.ref);
      if (!cloud.admin) batch.delete(db.collection('bands').doc(d.id).collection('members').doc(uid));
    });
    batch.delete(user.collection('data').doc('settings'));
    batch.delete(user.collection('data').doc('library'));
    await batch.commit();
    await db.collection('registered').doc(uid).delete().catch(() => {});
    store.clear(accountPrefix());
    await cloud.user.delete();               // the sign-in listener reopens the app signed out
    status('Your account has been deleted.');
  } catch (err) {
    messageBox('Could not delete the account', cloudMessage(err));
  } finally {
    document.body.classList.remove('busy');
  }
}

// ---------------------------------------------------------------- administrator: bands, codes, requests
async function pendingRequests() {
  try { return (await cloud.db.collection('requests').get()).size; } catch { return 0; }
}

function inviteText(bandName, code, name = '') {
  return `Hi${name ? ` ${name}` : ''},\n\nHere's your invitation to join ${bandName} on JJ's Setlist:\n\n`
    + `    Invitation code:  ${code}\n\n`
    + `1. Open ${SITE_URL}\n`
    + '2. Click "Sign in" (top right).\n'
    + '3. Under "New here?", type the invitation code, then click "Sign up with this code".\n'
    + '4. Choose your email address and a password (or use Google).\n\n'
    + `You'll then see ${bandName}'s songs and setlists. The confirmation email may land in your spam folder.\n`;
}

async function copyText(text, what) {
  try {
    await navigator.clipboard.writeText(text);
    status(`${what} copied - paste it into an email or message.`);
    return true;
  } catch {
    await dialog(what, `<p>Copy this (select it, then Ctrl+C):</p><pre class="copy-box">${esc(text)}</pre>`, undefined, true);
    return false;
  }
}

/** Bands & invitation codes: create bands, make / revoke codes, see and remove members. */
async function adminBands(message = '') {
  for (;;) {
    let bands = [];
    try {
      const snap = await cloud.db.collection('bands').get();
      bands = await Promise.all(snap.docs.map(async (d) => ({
        id: d.id, ...d.data(), members: (await d.ref.collection('members').get()).size,
      })));
    } catch (err) {
      messageBox('Bands & invitation codes', cloudMessage(err));
      return;
    }
    bands.sort((a, b) => fold(a.name).localeCompare(fold(b.name)));
    const rows = bands.map((b) => `<div class="band-row">
        <div class="band-name">🎵 ${esc(b.name)} <span class="muted">· ${plural(b.members, 'member')}</span></div>
        <div class="band-code">${b.inviteCode
          ? `Invitation code: <code>${esc(b.inviteCode)}</code>
             <button type="button" data-act="copy" data-band="${esc(b.id)}">Copy invitation</button>
             <button type="button" data-act="revoke" data-band="${esc(b.id)}">Revoke code</button>`
          : `<span class="muted">No invitation code - nobody can join.</span>
             <button type="button" data-act="new" data-band="${esc(b.id)}">Make a code</button>`}
          <button type="button" data-act="members" data-band="${esc(b.id)}">Members…</button></div>
      </div>`).join('') || '<p class="muted">No bands yet.</p>';
    const p = dialog('Bands & invitation codes', `<div class="admin">
        ${message ? `<p class="drive-result ok">${esc(message)}</p>` : ''}
        ${rows}
        <p class="small muted">Send a band's invitation to its members. Everyone can use the same code until you revoke it.
          Revoking stops new sign-ups; members already in keep their access.</p></div>`,
    [{ label: bands.length ? 'Create Another Band…' : 'Create a Band…', value: { act: 'create' } },
      { label: 'Close', value: false, primary: true }], true);
    $('#dialogBody').querySelectorAll('[data-act]').forEach((btn) => btn.addEventListener('click', () => dialogFinish({ act: btn.dataset.act, id: btn.dataset.band })));
    const ans = await p;
    if (!ans) return;
    const band = bands.find((b) => b.id === ans.id);
    message = '';
    try {
      if (ans.act === 'create') message = await adminCreateBand(bands) || '';
      else if (ans.act === 'copy') await copyText(inviteText(band.name, band.inviteCode), 'Invitation');
      else if (ans.act === 'new') message = await setBandCode(band, newInviteCode());
      else if (ans.act === 'revoke') {
        if (await confirmBox('Revoke the invitation code', `Revoke ${band.name}'s code ${band.inviteCode}?\n\nNobody new can sign up with it any more. Members already in keep their access. You can make a new code at any time.`, 'Revoke Code')) {
          message = await setBandCode(band, null);
        }
      } else if (ans.act === 'members') await adminMembers(band);
    } catch (err) {
      message = '';
      await messageBox('Bands & invitation codes', cloudMessage(err));
    }
  }
}

/** Give a band a new code (or none): the old code stops working at once. */
async function setBandCode(band, code) {
  const db = cloud.db, batch = db.batch();
  if (band.inviteCode) batch.delete(db.collection('invites').doc(band.inviteCode));
  if (code) batch.set(db.collection('invites').doc(code), { bandId: band.id, bandName: band.name, created: now() });
  batch.update(db.collection('bands').doc(band.id), { inviteCode: code });
  await batch.commit();
  return code ? `${band.name}'s new invitation code is ${code}.` : `${band.name}'s invitation code has been revoked.`;
}

async function adminCreateBand(bands) {
  let name = '', note = '';
  for (;;) {
    const p = dialog('Create a band', `<div class="account-form">
        ${note ? `<p class="drive-result bad">${esc(note)}</p>` : ''}
        <label for="bnName">Band name</label>
        <input id="bnName" type="text" maxlength="60" value="${esc(name)}" placeholder="e.g. Jelly Jazz">
        <p class="small muted">An invitation code is made for it straight away, to send to the band's members.</p></div>`,
    [{ label: 'Cancel', value: false }, { label: 'Create Band', value: true, primary: true, check: () => !!$('#bnName').value.trim() }], true);
    $('#bnName').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); [...$('#dialogButtons').children].pop().click(); } });
    $('#bnName').focus();
    if (!await p) return null;
    name = $('#bnName').value.trim().replace(/\s+/g, ' ');
    const key = bandKey(name);
    if (!key) { note = 'Please use some letters or numbers in the name.'; continue; }
    const same = bands.find((b) => (b.key || bandKey(b.name)) === key);
    if (same) { note = `There's already a band called “${same.name}” - that's the same name, apart from spaces, capitals or punctuation.`; continue; }
    const close = bands.find((b) => spellingSimilarity(key, b.key || bandKey(b.name)) >= 0.8);
    if (close && !await confirmBox('Create a band', `“${name}” looks very like the band “${close.name}”.\n\nCreate “${name}” as a separate band anyway?`, 'Create It')) { note = ''; continue; }
    const code = newInviteCode(), db = cloud.db, batch = db.batch();
    batch.set(db.collection('bands').doc(key), { name, key, inviteCode: code, created: now(), createdBy: cloud.user.uid });
    batch.set(db.collection('invites').doc(code), { bandId: key, bandName: name, created: now() });
    batch.set(cloudUserDoc().collection('bands').doc(key), { name, joined: now() });
    await batch.commit();
    cloud.bands.push({ id: key, name });
    cloud.bands.sort((a, b) => fold(a.name).localeCompare(fold(b.name)));
    renderSpacePicker();
    return `Created ${name}. Its invitation code is ${code} - use “Copy invitation” to send it to the members.`;
  }
}

async function adminMembers(band) {
  for (;;) {
    const snap = await cloud.db.collection('bands').doc(band.id).collection('members').get();
    const members = snap.docs.map((d) => ({ uid: d.id, ...d.data() }))
      .sort((a, b) => fold(a.email).localeCompare(fold(b.email)));
    const joined = (m) => (m.joined?.toDate ? dateTimeText(m.joined.toDate().toISOString()) : '');
    const p = dialog(`${band.name} - members`, `<div class="admin">${members.map((m) => `<div class="band-row member-row">
        <div>${esc(m.email || m.name || m.uid)} <span class="muted small">· joined ${esc(joined(m))}</span></div>
        <button type="button" data-uid="${esc(m.uid)}">Remove</button></div>`).join('') || '<p class="muted">No members yet.</p>'}</div>`,
    [{ label: 'Close', value: false, primary: true }], true);
    $('#dialogBody').querySelectorAll('[data-uid]').forEach((btn) => btn.addEventListener('click', () => dialogFinish(btn.dataset.uid)));
    const uid = await p;
    if (!uid) return;
    const m = members.find((x) => x.uid === uid);
    if (await confirmBox('Remove member', `Remove ${m.email || 'this member'} from ${band.name}?\n\nThey won't see the band's songs and setlists any more. Their own account and setlists stay.`, 'Remove')) {
      await cloud.db.collection('bands').doc(band.id).collection('members').doc(uid).delete();
    }
  }
}

/** Invitation requests: reply by email with a band's code, or delete. */
async function adminRequests() {
  for (;;) {
    let requests = [], bands = [];
    try {
      const [rq, bd] = await Promise.all([cloud.db.collection('requests').get(), cloud.db.collection('bands').get()]);
      requests = rq.docs.map((d) => ({ id: d.id, ...d.data() }))
        .sort((a, b) => (b.created?.seconds || 0) - (a.created?.seconds || 0));
      bands = bd.docs.map((d) => ({ id: d.id, ...d.data() })).filter((b) => b.inviteCode)
        .sort((a, b) => fold(a.name).localeCompare(fold(b.name)));
    } catch (err) {
      messageBox('Invitation requests', cloudMessage(err));
      return;
    }
    cloud.requests = requests.length;
    renderAccount();
    const when = (r) => (r.created?.toDate ? dateTimeText(r.created.toDate().toISOString()) : '');
    const options = bands.map((b) => `<option value="${esc(b.id)}">${esc(b.name)}</option>`).join('');
    const rows = requests.map((r) => `<div class="band-row request-row">
        <div><b>${esc(r.name)}</b> &lt;${esc(r.email)}&gt; <span class="muted small">· ${esc(when(r))}</span></div>
        ${r.band ? `<div class="small">Band: ${esc(r.band)}</div>` : ''}
        ${r.message ? `<div class="small request-msg">${esc(r.message)}</div>` : ''}
        <div class="band-code">${bands.length ? `<select data-for="${esc(r.id)}">${options}</select>
          <button type="button" data-act="email" data-id="${esc(r.id)}">Email them the code…</button>`
          : '<span class="muted small">Make a band code first (Bands & invitation codes).</span>'}
          <button type="button" data-act="delete" data-id="${esc(r.id)}">Delete request</button></div>
      </div>`).join('') || '<p class="muted">No requests waiting.</p>';
    const p = dialog('Invitation requests', `<div class="admin">${rows}
        <p class="small muted">“Email them the code” opens an email to them in your email program, with the band's invitation code and steps filled in. Check who they are before sending.</p></div>`,
    [{ label: 'Close', value: false, primary: true }], true);
    // Suggest the band they asked for.
    requests.forEach((r) => {
      const sel = $('#dialogBody').querySelector(`select[data-for="${CSS.escape(r.id)}"]`);
      const match = sel && bands.find((b) => bandKey(b.name) === bandKey(r.band));
      if (match) sel.value = match.id;
    });
    $('#dialogBody').querySelectorAll('[data-act]').forEach((btn) => btn.addEventListener('click', () => dialogFinish({
      act: btn.dataset.act, id: btn.dataset.id,
      band: $('#dialogBody').querySelector(`select[data-for="${CSS.escape(btn.dataset.id)}"]`)?.value,
    })));
    const ans = await p;
    if (!ans) return;
    const r = requests.find((x) => x.id === ans.id);
    try {
      if (ans.act === 'email') {
        const band = bands.find((b) => b.id === ans.band);
        const subject = `Your invitation to ${band.name} on JJ's Setlist`;
        window.location.href = `mailto:${encodeURIComponent(r.email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(inviteText(band.name, band.inviteCode, r.name))}`;
        if (await confirmBox('Invitation requests', `An email to ${r.email} should now be open in your email program.\n\nOnce it's sent, delete this request?`, 'Delete Request')) {
          await cloud.db.collection('requests').doc(r.id).delete();
        }
      } else if (ans.act === 'delete') {
        if (await confirmBox('Delete request', `Delete the request from ${r.name} <${r.email}>?`, 'Delete')) {
          await cloud.db.collection('requests').doc(r.id).delete();
        }
      }
    } catch (err) {
      await messageBox('Invitation requests', cloudMessage(err));
    }
  }
}

/** What this computer has that could go into the account: the helper's files,
    or what this browser kept when not signed in. */
async function localDataForImport() {
  let lists = {}, settings = null, from;
  if (files.on) {
    try {
      [lists, settings] = await Promise.all([api('GET', 'setlists'), api('GET', 'config')]);
      from = 'the setlists folder';
    } catch { /* helper gone */ }
  }
  if (!from) {
    lists = store.get('setlists', {}, BASE_PREFIX);
    const s = store.get('settings', null, BASE_PREFIX);
    settings = s && { library_colours: s.colours, web_font_size: s.fontSize, backup_url: s.driveUrl || null, drive_first: s.driveFirst };
    from = 'this browser';
  }
  const library = store.get('library', null, BASE_PREFIX);
  return { lists, settings, library: library?.songs?.length ? library : null, from };
}

/** Copy setlists (and, for a new account, the song library and settings) into the open space. */
async function offerImport(newAccount) {
  const { lists, settings, library, from } = await localDataForImport();
  const into = cloud.band ? cloud.band.name : 'your account';
  const keys = Object.keys(lists).filter((k) => !(k in savedSetlists));
  const bring = [];
  if (keys.length) bring.push(`${plural(keys.length, 'saved setlist')} (from ${from})`);
  const takeLibrary = library && !state.library.length;
  if (takeLibrary) bring.push(`the song library “${library.name}” (${plural(library.songs.length, 'song')})`);
  const takeSettings = newAccount && settings;
  if (takeSettings) bring.push('your colours, text size and Google Drive link');
  if (!bring.length) {
    if (newAccount) {
      saveSettings();                          // marks the account as set up
      messageBox('Welcome to JJ\'s Setlist', `Your account is ready: ${accountName()}.\n\nOpen your song spreadsheet (File ▸ Open Song Database), or add its Google Drive link in File ▸ Song Database Settings. Setlists you save are kept in your account.`);
    } else {
      messageBox('Copy setlists', `There are no setlists on this computer that aren't in ${into} already.`);
    }
    return;
  }
  const ok = await dialog(newAccount ? 'Welcome to JJ\'s Setlist' : `Copy setlists into ${into}`,
    `${newAccount ? `Your account is ready: ${accountName()}.\n\n` : ''}Copy these into ${into}?\n\n${bring.map((b) => `   •  ${b}`).join('\n')}\n\n`
    + `They stay on this computer too. Setlists already in ${into} aren't changed.`,
  [{ label: 'Not Now', value: false }, { label: 'Copy Them', value: true, primary: true }]);
  if (!ok) { if (newAccount) saveSettings(); return; }
  document.body.classList.add('busy');
  let copied = 0;
  try {
    for (const key of keys) { await putSetlist(key, lists[key]); copied++; }
    if (takeLibrary) {
      useLibraryData(library);
      await saveLibrary(library);
    }
    if (takeSettings) applyConfig(settings);
    if (newAccount || takeSettings) saveSettings();
    applyLook();
    renderAll();
    renderSaved();
    status(`Copied ${plural(copied, 'setlist')}${takeLibrary ? ' and the song library' : ''} into ${into}.`);
  } catch (err) {
    messageBox('Could not copy everything', `${plural(copied, 'setlist')} copied, then:\n\n${err.message}`);
  } finally {
    document.body.classList.remove('busy');
  }
}

// ---------------------------------------------------------------- DOM
const libList = $('#libList');
const setsEl = $('#sets');
const setPanels = [], setLists = [], setTitles = [];

function buildSets() {
  for (let n = 0; n < NUM_SETS; n++) {
    const panel = document.createElement('section');
    panel.className = 'panel set';
    panel.dataset.set = n;
    panel.innerHTML = `
      <h2></h2>
      <ul class="songs set-list" tabindex="0" data-list="${n}" data-set="${n}"></ul>
      <div class="panel-buttons">
        <button data-set-action="up">▲ Up</button>
        <button data-set-action="down">▼ Down</button>
        <button data-set-action="remove">Remove</button>
        <span class="spacer"></span>
        <button data-set-action="clear">Clear Set</button>
      </div>`;
    setsEl.appendChild(panel);
    setPanels.push(panel);
    setLists.push(panel.querySelector('ul'));
    setTitles.push(panel.querySelector('h2'));
    panel.addEventListener('pointerdown', () => { if (state.active !== n) setActive(n); });
    panel.querySelectorAll('[data-set-action]').forEach((btn) => btn.addEventListener('click', () => {
      setActive(n);
      ({ up: () => move(n, -1), down: () => move(n, 1), remove: () => remove(n),
        clear: () => clearSet(n) })[btn.dataset.setAction]();
    }));
  }
}

function songHtml(song, number, missing) {
  const extras = extrasOf(song);
  return (number ? `<span class="num">${number}.</span>` : '')
    + (song.link ? '<span class="sheet" data-sheet title="Open songsheet">📄</span>'
      : '<span class="sheet none"></span>')
    + '<span class="label">'
    + (missing ? '⚠ ' : '')
    + esc(song.title)
    + (song.artist ? `  —  ${esc(song.artist)}` : '')
    + (extras ? `<span class="extras">   [${esc(extras)}]</span>` : '')
    + '</span>';
}

function usedIds() {
  return new Set(state.sets.flat().map(songId));
}

function renderLibrary() {
  const q = fold($('#search').value.trim());
  state.filtered = state.library.filter((s) => !q || fold(
    Object.entries(s).filter(([k]) => k !== 'link').map(([, v]) => v).join(' ')).includes(q));
  const used = usedIds();
  const top = libList.scrollTop;
  libList.innerHTML = state.filtered.map((s, i) => {
    const cls = [used.has(songId(s)) ? 'used' : '',
      state.sel.list === 'lib' && state.sel.index === i ? 'selected' : ''].join(' ');
    return `<li data-i="${i}" class="${cls}" title="${esc(songLabel(s))}">${songHtml(s)}</li>`;
  }).join('');
  libList.scrollTop = top;
  $('#libCount').textContent = state.library.length
    ? `${state.filtered.length} of ${state.library.length} songs` : '';
  $('#getStarted').hidden = state.library.length > 0;
  const links = state.library.filter((s) => s.link).length;
  const label = $('#dbLabel');
  let text = 'No song database loaded';
  if (state.library.length) {
    text = `${state.dbName}  (${plural(state.library.length, 'song')}, ${plural(links, 'songsheet')})`;
    if (state.dbBackup) text = `⚠ BACKUP: ${text}  ·  copy from ${dateTimeText(state.dbLoaded)}`;
    else if (state.dbSource === 'drive') text += `  ·  Google Drive, ${dateTimeText(state.dbLoaded)}`;
  }
  label.textContent = text;
  label.classList.toggle('backup', !!state.dbBackup);
  label.title = state.dbBackup || text;
}

function dateTimeText(iso) {
  const d = new Date(iso);
  if (isNaN(d)) return 'earlier';
  return d.toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })
    .replace(',', '');
}

function isMissing(song) {
  if (!state.library.length) return false;
  if (!isMissing.cache || isMissing.cache.lib !== state.library) {
    isMissing.cache = { lib: state.library, ids: new Set(state.library.map(songId)) };
  }
  return !isMissing.cache.ids.has(songId(song));
}

function renderSet(n) {
  const songs = state.sets[n];
  setLists[n].innerHTML = songs.map((s, i) => {
    const missing = isMissing(s);
    const cls = [missing ? 'missing' : '',
      state.sel.list === n && state.sel.index === i ? 'selected' : ''].join(' ');
    return `<li data-i="${i}" class="${cls}" title="${esc(songLabel(s))}">${songHtml(s, i + 1, missing)}</li>`;
  }).join('');
  let title = `Set ${n + 1}   —   ${songs.length}/${MAX_SONGS} songs`;
  if (songs.length) title += `   —   Approx ${Math.floor(songs.length * MINUTES_PER_SONG + 0.5)} minutes`;
  setTitles[n].textContent = title;
  setPanels[n].classList.toggle('active', n === state.active);
}

function renderTotals() {
  const count = state.sets.reduce((a, s) => a + s.length, 0);
  $('#total').textContent = `Total: ${plural(count, 'song')}` + (state.dirty ? '   (unsaved changes)' : '');
  document.title = state.name ? `JJ's Setlist — ${state.name}` : "JJ's Setlist";
}

function renderSaved() {
  const sel = $('#saved');
  const keys = Object.keys(setlists()).sort((a, b) => fold(a).localeCompare(fold(b)));
  sel.innerHTML = '<option value=""></option>'
    + keys.map((k) => `<option value="${esc(k)}">${esc(k)}</option>`).join('');
  sel.value = keys.includes(state.savedKey) ? state.savedKey : '';
}

function renderAll() {
  renderLibrary();
  for (let n = 0; n < NUM_SETS; n++) renderSet(n);
  renderTotals();
}

function status(text) { $('#status').textContent = text; }

function setActive(n) {
  state.active = n;
  setPanels.forEach((p, i) => p.classList.toggle('active', i === n));
  $('#addBtn').textContent = `Add to Set ${n + 1} ▶`;
}

function select(list, index, scroll = true) {
  state.sel = { list, index };
  for (const ul of [libList, ...setLists]) {
    ul.querySelectorAll('li.selected').forEach((li) => li.classList.remove('selected'));
  }
  const ul = list === 'lib' ? libList : setLists[list];
  const li = ul.children[index];
  if (li) {
    li.classList.add('selected');
    if (scroll) li.scrollIntoView({ block: 'nearest' });
  }
}

function applyLook() {
  const root = document.documentElement.style;
  root.setProperty('--fs', state.fontSize + 'px');
  for (const [key, cssVar] of Object.entries(COLOUR_VARS)) root.setProperty(cssVar, state.colours[key]);
}

// ---------------------------------------------------------------- editing
function changed(selectList, selectIndex) {
  state.dirty = true;
  if (selectList !== undefined) state.sel = { list: selectList, index: selectIndex };
  saveWorking();
  renderAll();
}

function insertSong(n, index, song) {
  if (state.sets[n].length >= MAX_SONGS) {
    status(`Set ${n + 1} is full (${MAX_SONGS} songs).`);
    renderAll();
    return false;
  }
  const where = state.sets.map((st, i) => (st.some((s) => songId(s) === songId(song)) ? `Set ${i + 1}` : null))
    .filter(Boolean);
  state.sets[n].splice(index, 0, { ...song });
  changed(n, index);
  status(`Added “${song.title}” to Set ${n + 1}.` + (where.length ? `  (Note: also in ${where.join(', ')})` : ''));
  return true;
}

function addSelected() {
  if (state.sel.list !== 'lib' || !state.filtered[state.sel.index]) {
    status('Select a song in the library first.');
    return;
  }
  const n = state.active;
  insertSong(n, state.sets[n].length, state.filtered[state.sel.index]);
}

function selectedIndex(n) {
  return state.sel.list === n ? state.sel.index : -1;
}

function move(n, delta) {
  const i = selectedIndex(n), j = i + delta;
  const songs = state.sets[n];
  if (i < 0 || j < 0 || j >= songs.length) return;
  [songs[i], songs[j]] = [songs[j], songs[i]];
  changed(n, j);
  select(n, j);
}

function remove(n) {
  const i = selectedIndex(n);
  if (i < 0 || !state.sets[n][i]) { status('Select a song in the set first.'); return; }
  const [song] = state.sets[n].splice(i, 1);
  changed(n, Math.min(i, state.sets[n].length - 1));
  status(`Removed “${song.title}” from Set ${n + 1}.`);
}

async function clearSet(n) {
  if (!state.sets[n].length) return;
  if (!await confirmBox('Clear set', `Remove all ${state.sets[n].length} songs from Set ${n + 1}?`, 'Clear Set')) return;
  state.sets[n] = [];
  changed();
  status(`Cleared Set ${n + 1}.`);
}

async function clearAll() {
  if (!state.sets.flat().length) return;
  if (!await confirmBox('Clear all sets', 'Remove every song from all four sets?', 'Clear All')) return;
  state.sets = normaliseSets([]);
  changed();
  status('Cleared all sets.');
}

function selectedSong() {
  const { list, index } = state.sel;
  return list === 'lib' ? state.filtered[index] : state.sets[list]?.[index];
}

function openSheet(song = selectedSong()) {
  if (!song) { status('Select a song first.'); return; }
  if (!song.link) { status(`“${song.title}” has no songsheet link.`); return; }
  window.open(resolveLink(song.link), '_blank', 'noopener');
  status(`Opened the songsheet for “${song.title}”.`);
}

// ---------------------------------------------------------------- right-click menus
const ctxMenu = document.createElement('div');
ctxMenu.className = 'context-menu';
ctxMenu.hidden = true;
ctxMenu.setAttribute('role', 'menu');

function closeContextMenu() { ctxMenu.hidden = true; }

/** items: [{ label, action, disabled, note }] or '-' for a separator. */
function showContextMenu(x, y, items) {
  ctxMenu.innerHTML = '';
  for (const item of items) {
    if (item === '-') { ctxMenu.appendChild(document.createElement('hr')); continue; }
    const el = document.createElement(item.note ? 'div' : 'button');
    el.type = 'button';
    el.textContent = item.label;
    if (item.note) el.className = item.plain ? 'note plain' : 'note';
    else if (item.disabled) el.disabled = true;
    else el.addEventListener('click', () => { closeContextMenu(); item.action(); });
    ctxMenu.appendChild(el);
  }
  ctxMenu.hidden = false;
  // Keep it on screen.
  const r = ctxMenu.getBoundingClientRect();
  ctxMenu.style.left = Math.max(4, Math.min(x, innerWidth - r.width - 4)) + 'px';
  ctxMenu.style.top = Math.max(4, Math.min(y, innerHeight - r.height - 4)) + 'px';
  ctxMenu.querySelector('button:not(:disabled)')?.focus();
}

function libraryMenu(x, y, i) {
  const song = state.filtered[i];
  showContextMenu(x, y, [
    { label: '📄 Open Songsheet', action: () => openSheet(song), disabled: !song.link },
    '-',
    ...state.sets.map((st, n) => ({
      label: `Add to Set ${n + 1}` + (st.length >= MAX_SONGS ? '  (full)' : ''),
      disabled: st.length >= MAX_SONGS,
      action: () => { setActive(n); insertSong(n, st.length, song); },
    })),
  ]);
}

function setMenu(x, y, n, i) {
  const song = state.sets[n][i];
  showContextMenu(x, y, [
    ...(isMissing(song) ? [{ label: '⚠ Not in the song database', note: true }] : []),
    { label: '📄 Open Songsheet', action: () => openSheet(song), disabled: !song.link },
    '-',
    { label: '▲ Move Up', action: () => move(n, -1), disabled: i === 0 },
    { label: '▼ Move Down', action: () => move(n, 1), disabled: i === state.sets[n].length - 1 },
    '-',
    { label: 'Replace with…', action: () => replaceSong(n, i) },
    { label: 'Remove', action: () => remove(n) },
  ]);
}

/** Right-click (or the keyboard's menu key / Shift+F10) on a song. */
function openSongMenu(e, list) {
  const li = e.target.closest('li');
  let i = li ? +li.dataset.i : (state.sel.list === list ? state.sel.index : -1);
  if (i < 0) return;
  e.preventDefault();
  select(list, i, false);
  if (list !== 'lib') setActive(list);
  let { clientX: x, clientY: y } = e;
  if (e.type === 'keydown' || (!x && !y)) {      // from the keyboard: next to the selected song
    const r = (list === 'lib' ? libList : setLists[list]).children[i].getBoundingClientRect();
    x = r.left + 40;
    y = r.bottom;
  }
  if (list === 'lib') libraryMenu(x, y, i); else setMenu(x, y, list, i);
}

// ---------------------------------------------------------------- replace with…
/** 0..1: how alike two titles are spelt (1 - edit distance / length). */
function spellingSimilarity(a, b) {
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  let prev = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return 1 - prev[b.length] / Math.max(a.length, b.length);
}

const SMALL_WORDS = new Set(['the', 'a', 'of', 'in', 'my']);
const titleWords = (t) => new Set((fold(t).match(/[\p{L}\p{N}]+/gu) || []).filter((w) => !SMALL_WORDS.has(w)));

/** Library songs, most likely replacements for `old` first: same artist, then
    shared words in the title (e.g. a renamed song), then similar spelling. */
function rankReplacements(old) {
  const oldTitle = fold(old.title), oldWords = titleWords(old.title), oldArtist = fold(old.artist);
  return state.library.map((song) => {
    const words = titleWords(song.title);
    return {
      song,
      sameArtist: !!oldArtist && fold(song.artist) === oldArtist ? 1 : 0,
      shared: [...oldWords].filter((w) => words.has(w)).length,
      spelling: spellingSimilarity(oldTitle, fold(song.title)),
    };
  }).sort((a, b) => b.sameArtist - a.sameArtist || b.shared - a.shared || b.spelling - a.spelling)
    .map((r) => r.song);
}

/** Swap a set song for a library song, keeping its place in the set. */
async function replaceSong(n, i) {
  const old = state.sets[n][i];
  if (!old) return;
  if (!state.library.length) { await messageBox('Replace song', 'Open a song database first.'); return; }
  const ranked = rankReplacements(old);
  let shown = [], chosen = 0;
  const html = `<p class="replace-what">Replace <b>“${esc(old.title)}”</b>${old.artist ? `  —  ${esc(old.artist)}` : ''}
      <br><span class="muted">Set ${n + 1}, position ${i + 1}${isMissing(old) ? ' · ⚠ not in the song database' : ''}</span></p>
    <p class="muted small">with this song (the closest matches are first):</p>
    <div class="search"><label for="replSearch">Search:</label>
      <input id="replSearch" type="search" autocomplete="off" placeholder="Song, artist, style or vocalist"></div>
    <ul id="replList" class="songs lib-list replace-list" tabindex="0"></ul>`;
  const p = dialog('Replace song', html, [
    { label: 'Cancel', value: false },
    { label: 'Replace', value: true, primary: true, check: () => !!shown[chosen] },
  ], true);
  const search = $('#replSearch'), list = $('#replList');
  const pick = (k, scroll = true) => {
    if (!shown.length) return;
    chosen = Math.max(0, Math.min(shown.length - 1, k));
    list.querySelector('li.selected')?.classList.remove('selected');
    const li = list.children[chosen];
    li.classList.add('selected');
    if (scroll) li.scrollIntoView({ block: 'nearest' });
  };
  const replaceButton = () => [...document.querySelectorAll('#dialogButtons button')].find((b) => b.textContent === 'Replace');
  const refresh = () => {
    const q = fold(search.value.trim());
    shown = ranked.filter((s) => !q || fold(Object.entries(s).filter(([k]) => k !== 'link').map(([, v]) => v).join(' ')).includes(q));
    const used = usedIds();
    list.innerHTML = shown.map((s, k) => `<li data-i="${k}" class="${used.has(songId(s)) ? 'used' : ''}" title="${esc(songLabel(s))}">${songHtml(s)}</li>`).join('')
      || '<li class="none">No songs match.</li>';
    list.scrollTop = 0;
    chosen = 0;
    pick(0);
  };
  search.addEventListener('input', refresh);
  search.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); replaceButton().click(); }                  // not "close as Cancel"
    if (e.key === 'ArrowDown') { e.preventDefault(); list.focus(); pick(chosen + 1); }
  });
  list.addEventListener('click', (e) => { const li = e.target.closest('li[data-i]'); if (li) pick(+li.dataset.i, false); });
  list.addEventListener('dblclick', (e) => { if (e.target.closest('li[data-i]')) replaceButton().click(); });
  list.addEventListener('keydown', (e) => {
    const step = { ArrowDown: 1, ArrowUp: -1, PageDown: 10, PageUp: -10 }[e.key];
    if (step) { e.preventDefault(); pick(chosen + step); }
    if (e.key === 'Home' || e.key === 'End') { e.preventDefault(); pick(e.key === 'Home' ? 0 : shown.length - 1); }
    if (e.key === 'Enter') { e.preventDefault(); replaceButton().click(); }
  });
  refresh();
  search.focus();
  if (!await p) return;
  const song = shown[chosen];
  if (!song || state.sets[n][i] !== old) return;   // the set changed meanwhile
  const where = state.sets.map((st, k) => (st.some((s, j) => (k !== n || j !== i) && songId(s) === songId(song)) ? `Set ${k + 1}` : null))
    .filter(Boolean);
  state.sets[n][i] = { ...song };
  changed(n, i);
  setActive(n);
  status(`Replaced “${old.title}” with “${song.title}” in Set ${n + 1}.`
    + (where.length ? `  (Note: also in ${where.join(', ')})` : ''));
}

// ---------------------------------------------------------------- drag and drop
function setupDragging() {
  // forceFallback: Sortable's own pointer dragging rather than the browser's
  // native drag and drop, which looks and behaves the same in every browser.
  const common = { animation: 120, forceFallback: true, fallbackTolerance: 4, filter: '.sheet[data-sheet]', preventOnFilter: false };
  Sortable.create(libList, {
    ...common,
    group: { name: 'songs', pull: 'clone', put: true },
    sort: false,
    onAdd(evt) {                                   // dragged back from a set: remove it
      const n = +evt.from.dataset.set;
      evt.item.remove();
      const [song] = state.sets[n].splice(evt.oldIndex, 1);
      changed();
      if (song) status(`Removed “${song.title}” from Set ${n + 1}.`);
    },
  });
  setLists.forEach((ul, n) => Sortable.create(ul, {
    ...common,
    group: { name: 'songs', pull: true, put: true },
    onAdd(evt) {
      evt.item.remove();
      setActive(n);
      if (evt.from === libList) {
        const song = state.filtered[+evt.item.dataset.i];
        if (song) insertSong(n, evt.newIndex, song);
        return;
      }
      const from = +evt.from.dataset.set;
      if (state.sets[n].length >= MAX_SONGS) {
        status(`Set ${n + 1} is full (${MAX_SONGS} songs).`);
        renderAll();
        return;
      }
      const [song] = state.sets[from].splice(evt.oldIndex, 1);
      state.sets[n].splice(evt.newIndex, 0, song);
      changed(n, evt.newIndex);
      status(`Moved “${song.title}” to Set ${n + 1}.`);
    },
    onUpdate(evt) {                                // reordered within the set
      const [song] = state.sets[n].splice(evt.oldIndex, 1);
      state.sets[n].splice(evt.newIndex, 0, song);
      changed(n, evt.newIndex);
    },
  }));
}

// ---------------------------------------------------------------- song database
const NO_SONGS_HELP = 'Row 1 should hold the headings (e.g. SONG NAME, Artist, Style, Vocalist), with one song per row below.';

/** Songs from a spreadsheet's contents (.xlsx / .xlsm, else CSV), sorted by title. */
function parseDatabase(buffer, filename) {
  let songs;
  if (/\.(xlsx|xlsm)$/i.test(filename)) {
    const { rows, links } = readXlsx(buffer);
    songs = rowsToSongs(rows, links);
  } else {
    songs = rowsToSongs(parseCsv(decodeText(buffer)));
  }
  return songs.sort((a, b) => fold(a.title).localeCompare(fold(b.title)));
}

/** Show a freshly loaded song database and remember it for next time. */
async function useSongs(songs, { file, source }) {
  state.library = songs;
  state.dbName = stem(file);
  state.dbFile = file;
  state.dbSource = source;
  state.dbLoaded = new Date().toISOString();
  state.dbBackup = '';
  state.sel = { list: 'lib', index: -1 };
  const kept = await saveLibrary({ name: state.dbName, file, songs, source, loaded: state.dbLoaded });
  // Refresh set songs with the latest details (e.g. new songsheet links).
  state.sets = state.sets.map((st) => st.map(latestDetails));
  saveWorking();
  renderAll();
  const links = songs.filter((s) => s.link).length;
  status(`Loaded ${plural(songs.length, 'song')} (${plural(links, 'songsheet link')}) from `
    + (source === 'drive' ? 'Google Drive.' : `“${file}”.`)
    + (kept ? '' : "  (This browser won't keep it for next time.)"));
  warnMissing();
}

async function openDatabaseFile(file) {
  try {
    const songs = parseDatabase(await file.arrayBuffer(), file.name);
    if (!songs.length) {
      await messageBox('No songs found', `“${file.name}” has no songs in its first worksheet.\n\n${NO_SONGS_HELP}`);
      return;
    }
    await useSongs(songs, { file: file.name, source: 'file' });
  } catch (err) {
    console.error(err);
    await messageBox('Could not open the song database', `“${file.name}” couldn't be read.\n\n${err.message || err}`);
  }
}

// ---------------------------------------------------------------- Google Drive copy
const DRIVE_TIMEOUT = 15000;                 // ms
const DRIVE_SHARING = 'Check the internet connection, and that the file is shared in Google Drive as “Anyone with the link”.';
// Google refuses downloads into a page opened straight from a file
// (double-clicking index.html); it needs a web address (http://localhost or https://).
const OPENED_AS_FILE = location.protocol === 'file:';
const DRIVE_FILE_PAGE = "Google Drive doesn't allow downloads into a page opened straight from a file "
  + '(double-clicking index.html). Open JJ\'s Setlist from a web address instead, e.g. http://localhost - see the README.';

/** The file id in a Google Drive / Google Sheets share link, or ''. */
function driveFileId(url) {
  const m = /\/d\/([\w-]{20,})/.exec(url || '') || /[?&]id=([\w-]{20,})/.exec(url || '');
  return m ? m[1] : '';
}

/** Download the spreadsheet behind a Drive share link as .xlsx: { songs, file }.
    Uses Google Sheets' export address, which (unlike Drive's own download
    link) lets a web page read the reply. Works for uploaded Excel files and
    for Google Sheets. */
async function downloadDrive(url) {
  const id = driveFileId(url);
  if (!id) throw new Error("That doesn't look like a Google Drive link.");
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), DRIVE_TIMEOUT);
  let response, buffer;
  try {
    response = await fetch(`https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx`,
      { cache: 'no-store', signal: controller.signal });
    buffer = await response.arrayBuffer();
  } catch (err) {
    if (err.name === 'AbortError') throw new Error(`Google Drive didn't answer within ${DRIVE_TIMEOUT / 1000} seconds.`);
    throw new Error(OPENED_AS_FILE ? DRIVE_FILE_PAGE : `Couldn't reach Google Drive. ${DRIVE_SHARING}`);
  } finally {
    clearTimeout(timer);
  }
  if (response.status === 404) throw new Error("Google Drive couldn't find that file. Check the link.");
  if (!response.ok) throw new Error(`Google Drive said “${response.status} ${response.statusText}”. ${DRIVE_SHARING}`);
  const bytes = new Uint8Array(buffer, 0, Math.min(2, buffer.byteLength));
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {           // "PK": a real .xlsx (zip) file
    throw new Error(`Google sent a web page instead of the spreadsheet. ${DRIVE_SHARING}`);
  }
  // The spreadsheet's own name, e.g. "JELLY JAZZ SETLIST - with links.xlsx"
  const cd = response.headers.get('content-disposition') || '';
  let file = 'Google Drive song database.xlsx';
  const utf8 = /filename\*=UTF-8''([^;]+)/i.exec(cd), plain = /filename="([^"]+)"/i.exec(cd);
  try { if (utf8) file = decodeURIComponent(utf8[1]); else if (plain) file = plain[1]; } catch { /* keep the default */ }
  if (!/\.(xlsx|xlsm)$/i.test(file)) file += '.xlsx';
  const songs = parseDatabase(buffer, file);
  if (!songs.length) throw new Error(`There are no songs in its first worksheet. ${NO_SONGS_HELP}`);
  return { songs, file };
}

let driveBusy = false;

/** Load the song database from Google Drive. On failure the song database
    already here stays in use - marked as the backup if Drive should have
    replaced it. startup: quieter messages when the app opens. */
async function loadFromDrive({ startup = false } = {}) {
  if (!state.driveUrl) {
    status('Add a Google Drive link first: File ▸ Song Database Settings…');
    if (!startup) databaseSettings();
    return false;
  }
  if (driveBusy) return false;
  driveBusy = true;
  document.body.classList.add('busy');
  status('Loading the song database from Google Drive…');
  try {
    const { songs, file } = await downloadDrive(state.driveUrl);
    await useSongs(songs, { file, source: 'drive' });
    return true;
  } catch (err) {
    document.body.classList.remove('busy');  // not while the message is showing
    if (state.library.length) {
      state.dbBackup = `Google Drive couldn't be reached: ${err.message}`;
      renderLibrary();
      status(`⚠ Using the BACKUP song database: the copy from ${dateTimeText(state.dbLoaded)}.`);
      const again = await dialog('Using the backup song database',
        `The song database couldn't be loaded from Google Drive:\n\n   ${err.message}\n\n`
        + `Using the copy kept in this browser instead:\n\n   “${state.dbName}”, from ${dateTimeText(state.dbLoaded)}\n\n`
        + 'It may be missing recent changes.',
        [{ label: 'OK', value: false }, { label: 'Try Again', value: true, primary: true }]);
      if (again) { driveBusy = false; return await loadFromDrive({ startup }); }
    } else {
      status("Couldn't load the song database from Google Drive.");
      await messageBox('Could not load the song database',
        `The song database couldn't be loaded from Google Drive:\n\n   ${err.message}\n\nCheck File ▸ Song Database Settings, or open a spreadsheet from this computer.`);
    }
    return false;
  } finally {
    driveBusy = false;
    document.body.classList.remove('busy');
  }
}

/** File > Song Database Settings: the Google Drive link and when to use it. */
function databaseSettings() {
  const html = (OPENED_AS_FILE ? `<p class="drive-result bad">⚠ ${esc(DRIVE_FILE_PAGE)}</p>` : '') + `
    <p>Keep your song spreadsheet on Google Drive, and every computer can load the latest version, with no file to copy around.</p>
    <label for="driveUrl"><strong>Google Drive link to the song spreadsheet:</strong></label>
    <div class="drive-row">
      <input id="driveUrl" type="text" spellcheck="false" autocomplete="off"
        placeholder="https://drive.google.com/file/d/…  or  https://docs.google.com/spreadsheets/d/…">
      <button type="button" id="driveTest">Test link</button>
    </div>
    <div id="driveResult" class="drive-result"></div>
    <p><strong>Which to use:</strong></p>
    <label class="choice"><input type="radio" name="driveFirst" value="1">
      <span><b>The Google Drive copy first</b> - always the latest. It's loaded each time the app opens.
      If Google Drive can't be reached, the copy kept in this browser is used instead (marked ⚠ BACKUP).</span></label>
    <label class="choice"><input type="radio" name="driveFirst" value="0">
      <span><b>The spreadsheet I open on this computer.</b> Google Drive is used when no song database is loaded,
      or when you choose File ▸ Load Song Database from Google Drive.</span></label>
    <p class="muted small">Tips: in Google Drive, share the file as “Anyone with the link”. To update it, use
    Manage versions ▸ Upload new version, so the link stays the same. Leave the link empty to stop using Google Drive.</p>`;
  const p = dialog('Song Database Settings', html, [
    { label: 'Cancel', value: false },
    { label: 'Save', value: true, primary: true, check: () => {
      const link = $('#driveUrl').value.trim();
      if (link && !driveFileId(link)) {
        showDriveResult(false, "That doesn't look like a Google Drive link. Copy it from Share ▸ Copy link in Google Drive.");
        return false;
      }
      return true;
    } },
  ], true);
  $('#driveUrl').value = state.driveUrl;
  document.querySelector(`input[name=driveFirst][value="${state.driveFirst ? 1 : 0}"]`).checked = true;
  $('#driveTest').addEventListener('click', async () => {
    const btn = $('#driveTest');
    btn.disabled = true;
    showDriveResult(null, 'Downloading…');
    try {
      const { songs, file } = await downloadDrive($('#driveUrl').value.trim());
      showDriveResult(true, `Downloaded OK: “${stem(file)}”, ${plural(songs.length, 'song')}, `
        + `${plural(songs.filter((s) => s.link).length, 'songsheet')}.`);
    } catch (err) {
      showDriveResult(false, err.message);
    } finally {
      btn.disabled = false;
    }
  });
  $('#driveUrl').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); $('#driveTest').click(); }   // not "close the window"
  });
  $('#driveUrl').focus();
  p.then(async (ok) => {
    if (!ok) return;
    const link = $('#driveUrl').value.trim();
    const changed = link !== state.driveUrl;
    state.driveUrl = link;
    state.driveFirst = document.querySelector('input[name=driveFirst]:checked')?.value === '1';
    saveSettings();
    if (!link) { status('Google Drive is no longer used for the song database.'); return; }
    status('Song Database Settings saved.');
    // Like the desktop app: load from Drive straight away when it's the first
    // choice, the link is new, or nothing is loaded yet.
    if (state.driveFirst || changed || !state.library.length) await loadFromDrive();
  });
}

function showDriveResult(ok, text) {
  const el = $('#driveResult');
  el.className = 'drive-result' + (ok === true ? ' ok' : ok === false ? ' bad' : '');
  el.textContent = (ok === true ? '✓  ' : ok === false ? '✗  ' : '') + text;
}

/** The library's copy of a song: title + artist, else the title alone when
    only one song has it (so a corrected artist carries through). */
function latestDetails(song) {
  const exact = state.library.find((s) => songId(s) === songId(song));
  if (exact) return { ...exact };
  const same = state.library.filter((s) => fold(s.title) === fold(song.title));
  return same.length === 1 ? { ...same[0] } : song;
}

function isFound(song) {
  return state.library.some((s) => songId(s) === songId(song))
    || state.library.filter((s) => fold(s.title) === fold(song.title)).length === 1;
}

function missingSongs() {
  return state.sets.flat().filter(isMissing).map((s) => s.title);
}

function otherDb(name) {
  return !!(name && state.dbName && fold(name) !== fold(state.dbName));
}

function listed(titles, max) {
  let text = titles.slice(0, max).map((t) => `   •  ${t}`).join('\n');
  if (titles.length > max) text += `\n   …and ${titles.length - max} more`;
  return text;
}

function warnMissing() {
  const missing = missingSongs();
  if (!missing.length) return;
  const n = missing.length;
  if (otherDb(state.setlistDb)) {
    status(`⚠ ${plural(n, 'song')} not in this song database - the setlist was made with “${state.setlistDb}”.`);
    messageBox('Different song database',
      `This setlist was made with the song database:\n\n   “${state.setlistDb}”\n\nbut the one loaded now is:\n\n   “${state.dbName}”\n\n`
      + `${n} of its song${n === 1 ? ' is' : 's are'} not in “${state.dbName}”:\n\n${listed(missing, 8)}\n\n`
      + `They're marked ⚠ in orange. To fix it, use File ▸ Open Song Database and open “${state.setlistDb}”.`);
    return;
  }
  status(`⚠ ${plural(n, 'song')} in this setlist no longer in the song database (marked in orange).`);
  messageBox('Songs not in the song database',
    `${n} song${n === 1 ? ' is' : 's are'} in this setlist but no longer in the song database:\n\n${listed(missing, 12)}\n\n`
    + "They're kept in the setlist (marked ⚠ in orange) with the details saved in the setlist, which may be out of date.\n\n"
    + 'Right-click one to Remove it, or Replace with… another song (e.g. if it was renamed).');
}

// ---------------------------------------------------------------- saved setlists
function setlistData() {
  return {
    name: state.name.trim(),
    saved: new Date().toISOString().slice(0, 19),
    database: state.dbFile || '',
    database_name: state.dbName || '',
    sets: state.sets.map((songs, i) => ({ name: `Set ${i + 1}`, songs })),
  };
}

async function save() {
  state.name = $('#setName').value.trim();
  if (!state.name) {
    await messageBox('Save setlist', 'Please type a name for the setlist first.');
    $('#setName').focus();
    return false;
  }
  const key = safeFilename(state.name);
  if (!key) { await messageBox('Save setlist', "That name can't be used as a setlist name."); return false; }
  const all = setlists();
  if (all[key] && state.savedKey !== key
      && !await confirmBox('Overwrite?', `A setlist called “${key}” already exists.\nReplace it?`, 'Replace')) {
    return false;
  }
  try {
    await putSetlist(key, setlistData());
  } catch (err) {
    await messageBox('Could not save setlist', `${err.message}\n\nUse File ▸ Export This Setlist (.json) to keep a copy.`);
    return false;
  }
  state.savedKey = key;
  state.dirty = false;
  state.setlistDb = state.dbName;
  saveWorking();
  renderSaved();
  renderTotals();
  status({
    cloud: cloudQueued
      ? `Saved setlist “${state.name}” in this browser - it will be uploaded to ${spaceName()} when the connection is back.`
      : `Saved setlist “${state.name}” in ${spaceName()}.`,
    files: `Saved setlist “${state.name}”  →  setlists\\${key}.json`,
    browser: `Saved setlist “${state.name}” in this browser.  (File ▸ Export makes a copy for the desktop app or another computer.)`,
  }[mode]);
  return true;
}

/** Unsaved changes? Save / Don't Save / Cancel. True to carry on. */
async function confirmDiscard() {
  if (!state.dirty || !state.sets.flat().length) return true;
  const ans = await dialog('Unsaved changes', 'Save the current setlist first?',
    [{ label: 'Cancel', value: null }, { label: "Don't Save", value: false }, { label: 'Save', value: true, primary: true }]);
  if (ans === null) return false;
  if (ans) return save();
  return true;
}

async function loadSelected() {
  const key = $('#saved').value;
  if (!key) { status('Choose a setlist in the Saved setlists list first.'); return; }
  if (!await confirmDiscard()) { renderSaved(); return; }
  const data = setlists()[key];
  if (!data) return;
  loadSetlistData(data, key);
}

function setlistDbName(data) {
  let saved = (data.database_name || '').trim();
  if (!saved) {                                    // older setlists only kept the file path
    const s = stem(data.database);
    if (s !== 'drive_copy') saved = s;
  }
  return saved;
}

async function loadSetlistData(data, key) {
  const sets = normaliseSets((data.sets || []).map((st) => st.songs || []));
  const notFound = sets.flat().filter((s) => !isFound(s)).map((s) => s.title);
  const savedDb = setlistDbName(data);
  if (state.library.length && notFound.length && otherDb(savedDb)) {
    const ok = await dialog('Different song database',
      `This setlist was made with the song database:\n\n   “${savedDb}”\n\nbut the one loaded now is:\n\n   “${state.dbName}”\n\n`
      + `${notFound.length} of its song${notFound.length === 1 ? ' is' : 's are'} not in “${state.dbName}”:\n\n${listed(notFound, 8)}\n\n`
      + `Open “${savedDb}” first (File ▸ Open Song Database), then load the setlist again - or load it anyway.`,
      [{ label: 'Cancel', value: false }, { label: 'Load Anyway', value: true, primary: true }]);
    if (!ok) { renderSaved(); return; }
  }
  state.sets = sets.map((st) => st.map(latestDetails));
  state.setlistDb = savedDb;
  state.name = data.name || key;
  state.savedKey = key;
  state.dirty = false;
  state.sel = { list: 'lib', index: -1 };
  $('#setName').value = state.name;
  saveWorking();
  renderAll();
  renderSaved();
  status(`Loaded setlist “${state.name}”.`);
  warnMissing();
}

async function deleteSaved() {
  const key = $('#saved').value;
  if (!key) { status('Choose a setlist in the Saved setlists list first.'); return; }
  if (!await confirmBox('Delete setlist',
    ({ cloud: `Delete the saved setlist “${key}” from your account?`,
      files: `Delete the saved setlist file “${key}.json” from the setlists folder?`,
      browser: `Delete the saved setlist “${key}” from this browser?` }[mode])
    + '\nThe songs currently on screen are not affected.', 'Delete')) return;
  try {
    await removeSetlist(key);
  } catch (err) {
    await messageBox('Could not delete setlist', err.message);
    return;
  }
  if (state.savedKey === key) { state.savedKey = ''; saveWorking(); }
  renderSaved();
  status(`Deleted “${key}”.`);
}

async function newSetlist() {
  if (!await confirmDiscard()) return;
  state.sets = normaliseSets([]);
  state.name = '';
  state.savedKey = '';
  state.setlistDb = '';
  state.dirty = false;
  state.sel = { list: 'lib', index: -1 };
  $('#setName').value = '';
  saveWorking();
  renderAll();
  renderSaved();
  status('Started a new setlist.');
}

async function importSetlistFiles(files) {
  const all = setlists();
  const added = [], replaced = [], bad = [];
  const incoming = [];
  for (const file of files) {
    try {
      const data = JSON.parse(await file.text());
      if (!Array.isArray(data.sets)) throw new Error('not a setlist');
      const key = safeFilename(data.name || stem(file.name)) || stem(file.name);
      incoming.push([key, data]);
    } catch { bad.push(file.name); }
  }
  const clashes = incoming.filter(([k]) => all[k]).map(([k]) => k);
  let overwrite = false;
  if (clashes.length) {
    overwrite = await dialog('Setlists already here',
      `${plural(clashes.length, 'setlist')} with the same name already ${clashes.length === 1 ? 'is' : 'are'} saved ${whereSaved()}:\n\n${listed(clashes, 8)}\n\nReplace ${clashes.length === 1 ? 'it' : 'them'} with the imported cop${clashes.length === 1 ? 'y' : 'ies'}?`,
      [{ label: 'Keep Existing', value: false }, { label: 'Replace', value: true, primary: true }]);
  }
  for (const [key, data] of incoming) {
    if (all[key] && !overwrite) continue;
    const existed = !!all[key];
    try {
      await putSetlist(key, data);
      (existed ? replaced : added).push(key);
    } catch (err) {
      await messageBox('Could not import setlist', `“${key}” couldn't be saved:\n\n${err.message}`);
      break;
    }
  }
  renderSaved();
  let msg = `Imported ${plural(added.length + replaced.length, 'setlist')}.`;
  if (bad.length) msg += `  Not setlists: ${bad.join(', ')}`;
  status(msg);
  if (added.length + replaced.length === 1 && !state.sets.flat().length) {
    $('#saved').value = (added[0] || replaced[0]);
  }
}

// ---------------------------------------------------------------- export & print
function exportName(ext) {
  return `${safeFilename(state.name.trim() || 'Setlist')}.${ext}`;
}

function setlistText() {
  const name = $('#setName').value.trim() || 'Setlist';
  const all = state.sets.flat();
  const lines = textHeading(name, plural(all.length, 'song'));
  state.sets.forEach((songs, n) => {
    if (!songs.length) return;
    lines.push(`SET ${n + 1}   (${plural(songs.length, 'song')})`, ...songTable(songs, true, all), '');
  });
  return lines.join('\n');
}

function exportJson() {
  state.name = $('#setName').value.trim();
  download(exportName('json'), JSON.stringify(setlistData(), null, 2), 'application/json');
  status(`Exported “${exportName('json')}”. Put it in the desktop app's setlists folder, or Import it on another device.`);
}

function exportText() {
  download(exportName('txt'), setlistText());
}

function exportCsv() {
  const rows = [['Set', '#', 'Song Name', 'Artist', 'Style', 'Vocalist', 'Songsheet']];
  state.sets.forEach((songs, n) => songs.forEach((s, i) =>
    rows.push([n + 1, i + 1, s.title, s.artist, s.style, s.vocalist, s.link])));
  download(exportName('csv'), '﻿' + rows.map((r) => r.map(csvField).join(',')).join('\r\n'), 'text/csv');
}

function exportDbCsv() {
  if (!state.library.length) { status('Open a song database first.'); return; }
  const rows = [['SONG NAME', 'Artist', 'Style', 'Vocalist', 'Songsheet']];
  state.library.forEach((s) => rows.push([s.title, s.artist, s.style, s.vocalist, s.link]));
  download(`${state.dbName || 'Song Database'} (with links).csv`,
    '﻿' + rows.map((r) => r.map(csvField).join(',')).join('\r\n'), 'text/csv');
}

function printTable(songs, numbered) {
  return '<table class="print-table"><thead><tr>'
    + (numbered ? '<th class="n">#</th>' : '')
    + '<th class="t">SONG NAME</th><th class="a">ARTIST</th><th>STYLE</th><th>VOCALIST</th></tr></thead><tbody>'
    + songs.map((s, i) => '<tr>' + (numbered ? `<td class="n">${i + 1}.</td>` : '')
      + `<td class="t">${esc(s.title)}</td><td class="a">${esc(s.artist)}</td><td>${esc(s.style)}</td><td>${esc(s.vocalist)}</td></tr>`).join('')
    + '</tbody></table>';
}

function printSetlist() {
  const all = state.sets.flat();
  if (!all.length) { status('There are no songs in the setlist to print.'); return; }
  const name = $('#setName').value.trim() || 'Setlist';
  $('#printArea').innerHTML = `<div class="print-head"><h1>${esc(name)}  —  ${plural(all.length, 'song')}</h1><p>${todayText()}</p></div>`
    + state.sets.map((songs, n) => (songs.length
      ? `<div class="print-set"><h2>SET ${n + 1}   (${plural(songs.length, 'song')})</h2>${printTable(songs, true)}</div>` : '')).join('');
  const before = document.title;
  document.title = name;                     // suggested file name for Save as PDF
  window.print();
  document.title = before;
}

function printSonglist() {
  if (!state.filtered.length) { status('There are no songs to print.'); return; }
  const q = $('#search').value.trim();
  const sub = plural(state.filtered.length, 'song') + (q ? ` matching “${q}”` : '');
  $('#printArea').innerHTML = `<div class="print-head"><h1>${esc(state.dbName || 'Song List')}  —  ${esc(sub)}</h1><p>${todayText()}</p></div>`
    + printTable(state.filtered, false);
  window.print();
}

// ---------------------------------------------------------------- dialogs
const dlg = $('#dialog');

function dialog(title, body, buttons = [{ label: 'OK', value: true, primary: true }], html = false) {
  return new Promise((resolve) => {
    $('#dialogTitle').textContent = title;
    const bodyEl = $('#dialogBody');
    bodyEl.classList.toggle('html', html);
    if (html) bodyEl.innerHTML = body; else bodyEl.textContent = body;
    const box = $('#dialogButtons');
    box.innerHTML = '';
    const cancelValue = buttons.find((b) => b.value === null || b.value === false)?.value ?? null;
    // Answer straight from the button, rather than waiting for the dialog's
    // "close" event: that arrives later (or, in some cases, not at all), and a
    // late one could otherwise answer the next dialog.
    let settled = false;
    const finish = (value) => {
      if (settled) return;
      settled = true;
      if (dlg.open) dlg.close();
      resolve(value);
    };
    dialogFinish = finish;                     // for buttons inside the body (e.g. "Continue with Google")
    // Enter in a one-box form would "submit" it, closing it as Cancel: use the main button instead.
    dlg.querySelector('form').onsubmit = (e) => { e.preventDefault(); box.querySelector('.primary')?.click(); };
    buttons.forEach((b) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = b.label;
      if (b.primary) btn.className = 'primary';
      btn.addEventListener('click', () => {
        if (b.check && !b.check()) return;     // e.g. a setting that needs fixing first
        finish(b.value);
      });
      box.appendChild(btn);
    });
    dlg.oncancel = () => finish(cancelValue);  // Esc
    dlg.onclose = () => { if (!dlg.open) finish(cancelValue); };
    dlg.showModal();
    (box.querySelector('.primary') || box.lastChild)?.focus();
  });
}

let dialogFinish = null;                     // closes the open dialog with a value
const messageBox = (title, text) => dialog(title, text);
const confirmBox = (title, text, okLabel = 'OK') =>
  dialog(title, text, [{ label: 'Cancel', value: false }, { label: okLabel, value: true, primary: true }]);

function coloursDialog() {
  const original = { ...state.colours };
  const rows = COLOUR_NAMES.map(([key, label]) =>
    `<label for="c_${key}">${label}</label><input type="color" id="c_${key}" value="${state.colours[key]}"><code id="h_${key}">${state.colours[key]}</code>`).join('');
  const html = `<div class="colour-grid">${rows}</div><div class="contrast" id="contrast"></div>`;
  const p = dialog('Songlist & Setlist Colours', html, [
    { label: 'Defaults', value: 'reset' }, { label: 'Cancel', value: false }, { label: 'OK', value: true, primary: true }], true);
  const update = () => {
    for (const [key] of COLOUR_NAMES) {
      state.colours[key] = $(`#c_${key}`).value;
      $(`#h_${key}`).textContent = state.colours[key];
    }
    applyLook();
    const c = state.colours;
    const checks = [['Song titles', contrast(c.fg, c.bg)], ['Songs in a set', contrast(c.used, c.bg)],
      ['Selected song', contrast('#ffffff', c.select)]];
    $('#contrast').innerHTML = 'Readability: ' + checks.map(([what, r]) =>
      `<span class="${r < 3 ? 'warn' : ''}">${what} ${r.toFixed(1)}:1${r < 3 ? ' ⚠ hard to read' : ''}</span>`).join(' · ');
  };
  document.querySelectorAll('.colour-grid input').forEach((inp) => inp.addEventListener('input', update));
  update();
  p.then((ans) => {
    if (ans === 'reset') { state.colours = { ...DEFAULT_COLOURS }; applyLook(); saveSettings(); status('Colours back to the defaults.'); return; }
    if (ans === true) { saveSettings(); status('Songlist & Setlist colours saved.'); return; }
    state.colours = original;
    applyLook();
  });
}

function contrast(fg, bg) {
  const lum = (hex) => {
    const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
      .map((v) => (v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4));
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  };
  const [a, b] = [lum(fg), lum(bg)].sort((x, y) => y - x);
  return (a + 0.05) / (b + 0.05);
}

function setFont(size) {
  state.fontSize = Math.max(MIN_FONT, Math.min(MAX_FONT, size));
  applyLook();
  saveSettings();
  status(`Text size ${state.fontSize}.`);
}

// ---------------------------------------------------------------- help guides
// Each guide is a list of [kind, text]: h2 / h3 headings, p (paragraph),
// b (bullet), n (numbered step), code, tip, warn. Ctrl+ becomes Cmd+ on a Mac.
const GUIDES = {
  basics: {
    title: 'How to Use - The Basics',
    content: () => [
      ['h2', 'Building sets'],
      ['b', 'Click a set to make it the current one (coloured frame and blue title).'],
      ['b', 'Double-click a library song (or select it and press Enter) to add it to the current set. Or drag songs from the library into any set, at any position.'],
      ['b', 'Drag within a set to reorder, or to another set to move a song. Drag a song back to the library to remove it.'],
      ['b', 'Alt+↑ / Alt+↓ (or the ▲ Up / ▼ Down buttons) move the selected song; Delete removes it.'],
      ['b', 'Alt+1 … Alt+4 jump to a set, Alt+L to the library, Ctrl+F to the Search box.'],
      ['b', 'Songs already in the setlist are greyed out in the library. Each set shows roughly how long it runs (about 3½ minutes a song).'],
      ['h2', 'Right-click a song'],
      ['p', 'Right-click a song (or press Shift+F10) for its menu: Open Songsheet, Move Up / Down, Replace with…, Remove - or, in the library, Add to Set 1–4.'],
      ['b', 'A song marked ⚠ in orange is no longer in the song database (e.g. it was renamed). Right-click it and choose Replace with… - the closest matches are listed first - or Remove.'],
      ['h2', 'Songsheets'],
      ['b', 'Songs with a songsheet link show 📄. Click the 📄, or double-click a song in a set, to open its songsheet.'],
      ['b', 'Links are added in your spreadsheet - see Help ▸ Setting Up Your Spreadsheet.'],
      ['h2', 'Saving and loading setlists'],
      ['b', 'Type a name (include the venue and date, e.g. “The Local Pub - October 2026”) and press Save Setlist (Ctrl+S, or Enter in the name box).'],
      ['b', 'Load a saved setlist from the Saved setlists list. New starts an empty one.'],
      ['b', 'File ▸ Export This Setlist (.json) makes a copy you can e-mail, or open in the desktop app. File ▸ Import Saved Setlists brings copies in (or drag .json files onto the window).'],
      ['h2', 'Printing'],
      ['b', 'File ▸ Print / Save as PDF (Ctrl+P): landscape, aligned columns, a set is never split across pages. Choose the printer - or “Save as PDF” to e-mail it - in the print window.'],
      ['b', 'File ▸ Print / Save as PDF: Song List prints the whole library, or just the songs matching a search.'],
      ['h2', 'Where your setlists are kept'],
      ['p', 'The label next to the song total says which is in use:'],
      ['b', '☁ Saving to your account (or ☁ Saving to your band) - signed in: your setlists and songs are online, on any computer; a band\'s are shared by its members. Switch between your own setlists and your band at the top right. See Help ▸ Your Online Account & Bands.'],
      ['b', `💾 Saving to files in the project folder - started with “${LAUNCHER}”: settings in config.json and setlists in the setlists folder, the same files as the desktop app.`],
      ['b', 'Saving in this browser only - neither of those. Use Export to keep copies of setlists you care about.'],
      ['tip', 'Tip: the app remembers the setlist on screen, even unsaved, so closing the tab or reloading the page loses nothing.'],
    ],
  },

  spreadsheet: {
    title: 'Setting Up Your Song Spreadsheet',
    actions: [['Create Template Spreadsheet', 'xlsx']],
    content: () => [
      ['p', 'The song library comes from a spreadsheet: an Excel workbook (.xlsx) - or a CSV file, see Help ▸ Using a CSV File Instead. Each row is one song, and each song name can carry a link to its songsheet (PDF).'],
      ['tip', 'Quickest start: click “Create Template Spreadsheet” below, fill it in, then open it with File ▸ Open Song Database.'],
      ['h2', '1.  The column headings (row 1)'],
      ['p', 'Put these headings in the first row of the first worksheet. Capitals don\'t matter and the columns can be in any order.'],
      ['code', 'SONG NAME      Artist      Style      Vocalist'],
      ['b', 'SONG NAME - required. The title as you want it shown.'],
      ['b', 'Artist - optional. Original artist or band.'],
      ['b', 'Style - optional. e.g. Ballad, 60\'s  ·  Rock  ·  Latin'],
      ['b', 'Vocalist - optional. Who sings it.'],
      ['p', 'Other accepted heading names:'],
      ['b', 'SONG NAME:  Title, Song, Song Title, Name, Track'],
      ['b', 'Artist:  Band, Performer, Original Artist'],
      ['b', 'Style:  Genre, Type'],
      ['b', 'Vocalist:  Singer, Vocals, Vocal, Lead Vocal'],
      ['p', 'Any other columns (notes, keys, etc.) are simply ignored, so you can keep extra information in the sheet.'],
      ['h2', '2.  One song per row'],
      ['b', 'Start on row 2, one song per row. Empty rows are skipped.'],
      ['b', 'If you have two versions of a song, make the names different, e.g. “Scarborough Fair [v1]” and “Scarborough Fair [v2]”.'],
      ['b', 'Style and Vocalist are searchable - type “Ballad” or a singer\'s name in the Search box.'],
      ['h2', '3.  Adding songsheet links'],
      ['p', 'Attach the link to the SONG NAME cell. Songs with a link show 📄 in the app.'],
      ['h3', 'Method A - Insert Link (recommended)'],
      ['n', '1.  Click the song-name cell.'],
      ['n', '2.  Press Ctrl+K  (or Insert ▸ Link).'],
      ['n', '3.  In “Address”, paste the web address of the songsheet.'],
      ['n', '4.  Click OK. The name turns blue and underlined.'],
      ['h3', 'Method B - HYPERLINK formula'],
      ['code', '=HYPERLINK("https://drive.google.com/file/d/…/view", "El Paso")'],
      ['p', 'The first part is the link, the second is the song name shown.'],
      ['h3', 'Method C - a separate link column'],
      ['p', 'Add a column headed Songsheet (or Link, URL, PDF) and type the web address in it. This also works in CSV files.'],
      ['h2', '4.  Getting a Google Drive link for a PDF'],
      ['n', '1.  In Google Drive, right-click the PDF ▸ Share ▸ Copy link.'],
      ['n', '2.  If bandmates will use the songsheets too, set General access to “Anyone with the link”.'],
      ['n', '3.  Paste the link into the spreadsheet with Ctrl+K, as in Method A.'],
      ['warn', 'Web links (e.g. Google Drive) work everywhere. Links to PDF files on your computer only open when JJ\'s Setlist is started with its launcher, and the PDFs are inside its folder (e.g. a Songsheets folder in it) - use a link like Songsheets\\El Paso.pdf. They don\'t work in the online version.'],
      ['h2', '5.  Editing a linked cell (in Excel)'],
      ['b', 'Clicking a linked cell opens the link. To edit the name instead, select the cell with the arrow keys (or click and hold) and press F2.'],
      ['b', 'Editing the text keeps the link. To change the link, press Ctrl+K.'],
      ['h2', '6.  Saving - important!'],
      ['warn', 'Always save as an Excel Workbook (.xlsx): saving as CSV throws away every link behind a song name.'],
      ['b', 'Use Ctrl+S. If Excel asks about the format, choose Excel Workbook.'],
      ['b', 'Only the first worksheet is read. Formatting (fonts, colours, column widths) doesn\'t matter.'],
      ['h2', '7.  Loading it into JJ\'s Setlist'],
      ['b', 'File ▸ Open Song Database… (Ctrl+O) and choose the .xlsx. The songs are remembered (in your account when signed in), so next time the library is already there.'],
      ['b', 'After editing the spreadsheet, open it again the same way - a browser can\'t re-read a file on its own. (A Google Drive copy, below, is re-read automatically.)'],
      ['b', 'The top bar shows how many songs and songsheets were found - if the songsheet count is 0, the file was probably saved as CSV.'],
      ['tip', 'You can also drag the spreadsheet file straight onto the JJ\'s Setlist window.'],
      ['h2', '8.  Keeping it on Google Drive'],
      ['p', 'Keep the spreadsheet on Google Drive and JJ\'s Setlist can load the latest version each time it opens - on any computer, and for bandmates too.'],
      ['n', '1.  Upload the .xlsx to Google Drive (or make it in Google Sheets) and share it as “Anyone with the link”.'],
      ['n', '2.  Right-click it in Drive ▸ Share ▸ Copy link.'],
      ['n', '3.  In JJ\'s Setlist: File ▸ Song Database Settings, paste the link, click Test link, choose “The Google Drive copy first”, then Save.'],
      ['b', 'If Google Drive can\'t be reached (e.g. no internet), the copy kept from last time is used, and the top bar shows ⚠ BACKUP in orange.'],
      ['warn', 'To update a Drive copy, use Manage versions ▸ Upload new version in Google Drive. Deleting it and uploading a new file gives it a new link.'],
      ['warn', 'Loading from Google Drive needs JJ\'s Setlist opened from a web address - online, or started with its launcher - not by double-clicking index.html.'],
    ],
  },

  csv: {
    title: 'Using a CSV File Instead',
    actions: [['Create Template CSV', 'csv']],
    content: () => [
      ['p', 'The song library doesn\'t have to be an Excel workbook. A CSV file (comma-separated values) works too - handy if you keep your list in Google Sheets, Numbers, LibreOffice or a plain text editor.'],
      ['warn', 'A CSV can\'t hold links hidden behind a song name: Excel drops them all when it saves as CSV. Put the links in their own column instead (see step 3).'],
      ['tip', 'Quickest start: click “Create Template CSV” below - it already has the right headings and a link column.'],
      ['tip', 'Already have an Excel sheet with links behind the names? Open it, then use File ▸ Export Song Database as CSV (with web links): every link is written into a Songsheet column, so none are lost.'],
      ['h2', '1.  The column headings (first line)'],
      ['p', 'The same headings as the Excel version - any order, capitals don\'t matter, and extra columns are ignored:'],
      ['code', 'SONG NAME,Artist,Style,Vocalist,Songsheet'],
      ['b', 'SONG NAME - required. Also accepted: Title, Song, Name, Track'],
      ['b', 'Artist, Style, Vocalist - optional (same alternatives as Excel).'],
      ['b', 'Songsheet - optional link column. Also accepted: Link, URL, PDF, Sheet, Chart'],
      ['h2', '2.  One song per line'],
      ['code', 'El Paso,Marty Robbins,Country,Adrian,https://drive.google.com/…\nBlue Moon,Nat King Cole,Ballad,Gary,Songsheets\\Blue Moon.pdf\nMoon River,Henry Mancini,Ballad,,'],
      ['b', 'Leave a value empty by putting nothing between the commas (Moon River has no vocalist or songsheet above).'],
      ['b', 'If a value contains a comma, wrap it in double quotes, e.g. "Ballad, 60\'s". Spreadsheet programs do this for you.'],
      ['b', 'Extra commas at the end of lines are fine - they\'re ignored.'],
      ['h2', '3.  Songsheet links in a CSV'],
      ['p', 'Type (or paste) the full link into the Songsheet column:'],
      ['b', 'A web address, e.g. a Google Drive link (Drive: right-click the PDF ▸ Share ▸ Copy link). These work everywhere.'],
      ['b', 'Or a file inside JJ\'s Setlist\'s folder, e.g.  Songsheets\\El Paso.pdf  - this only opens when JJ\'s Setlist is started with its launcher, not online.'],
      ['p', 'Songs with a link show 📄 in the app, exactly as with Excel.'],
      ['h2', '4.  Saving a CSV'],
      ['b', 'Excel: File ▸ Save As ▸ “CSV UTF-8 (Comma delimited) (*.csv)”. UTF-8 keeps accented names (Hasta Mañana) correct.'],
      ['b', 'Google Sheets: File ▸ Download ▸ Comma-separated values (.csv).'],
      ['b', 'Text editor: save with a .csv extension (UTF-8 if offered).'],
      ['h2', '5.  Loading it into JJ\'s Setlist'],
      ['b', 'File ▸ Open Song Database… (Ctrl+O) and choose the .csv - or drag it onto the window.'],
      ['b', 'After editing, open it again the same way.'],
      ['b', 'Check the top bar: it shows how many songs and songsheets were found.'],
      ['h2', 'Excel or CSV?'],
      ['b', 'Excel (.xlsx): links can hide behind the song name, and your formatting is kept.'],
      ['b', 'CSV: simple plain text that any program can edit, but links must go in their own column and there\'s no formatting.'],
      ['p', 'Both work fully in JJ\'s Setlist - choose whichever suits the way you keep your song list.'],
    ],
  },

  account: {
    title: 'Your Online Account & Bands',
    content: () => [
      ['p', 'Sign in and your song library, settings and setlists are kept online - there on any computer. If you\'re in a band, you also share the band\'s song library and setlists with its other members.'],
      ['h2', 'Joining (signing up)'],
      ['p', 'JJ\'s Setlist is by invitation: you need an invitation code from the administrator. The code also says which band you\'re joining.'],
      ['n', '1.  Click Sign in (top right). Under “New here?”, type the invitation code - it shows which band it\'s for.'],
      ['n', '2.  Click Sign up with this code.'],
      ['n', '3.  Type your email address - any address works, it doesn\'t have to be Gmail - and a password of at least 8 characters. Or click Sign up with Google.'],
      ['n', '4.  Click Create My Account. A message is sent to check your address: open it and click its link (look in the spam folder if it doesn\'t arrive).'],
      ['b', 'No code yet? Click Request an invitation code on the Sign in window, and send your name and email address. The administrator will be in touch.'],
      ['h2', 'Your band, and your own setlists'],
      ['b', 'In a band, its members share one song library and one set of saved setlists. Everyone has full access: anything a member saves, the others see.'],
      ['b', 'You also have your own private setlists. Switch between them and your band(s) with the list at the top right, next to your email address.'],
      ['b', 'The label next to the song total shows which is open, e.g. ☁ Saving to Jelly Jazz, or ☁ Saving to your account.'],
      ['b', 'Colours and text size are always your own; a band\'s Google Drive link (File ▸ Song Database Settings) is shared by its members.'],
      ['b', 'Already signed up, and given a code for another band? Use Join a band with an invitation code in your account menu.'],
      ['h2', 'Your account menu'],
      ['p', 'Click your email address (top right):'],
      ['b', 'Copy my own setlists into the band - shares setlists you made yourself with the band (yours stay as they are).'],
      ['b', 'Copy setlists from this computer - brings in setlists saved before you signed in, or in the setlists folder.'],
      ['b', 'Leave the band - you stop seeing its songs and setlists.'],
      ['b', 'Change my password - sends you an email with a link to choose a new one.'],
      ['b', 'Sign out - also clears your copies from this browser, so it\'s safe on a shared computer.'],
      ['b', 'Delete my account - removes your account and your own setlists for good (a band\'s setlists stay, for its other members). Export any you want to keep first.'],
      ['h2', 'For the administrator'],
      ['b', 'Bands & invitation codes: create a band (similar or duplicate names are caught), copy its invitation to send to the members, revoke the code once everyone has joined, make a new one, and see or remove members.'],
      ['b', 'Everyone in a band can use the same code until it\'s revoked. Revoking stops new sign-ups; members already in keep their access.'],
      ['b', 'Invitation requests: see who asked for a code (a number on your account button shows how many are waiting), email them the band\'s invitation, or delete the request.'],
      ['h2', 'Forgotten your password?'],
      ['p', 'On the Sign in window, type your email address and click Forgot password? - an email with a link to choose a new password is sent to you.'],
      ['h2', 'No internet?'],
      ['p', 'Changes are kept in this browser and uploaded when the connection is back - the status bar says so.'],
    ],
  },
};

/** A guide as HTML: consecutive bullets / steps are grouped into lists. */
function guideHtml(content) {
  let html = '', list = null;
  const close = () => { if (list) { html += `</${list}>`; list = null; } };
  for (const [kind, raw] of content) {
    const text = esc(forPlatform(raw));
    if (kind === 'b' || kind === 'n') {
      const want = kind === 'b' ? 'ul' : 'ol';
      if (list !== want) { close(); html += `<${want}>`; list = want; }
      html += `<li>${kind === 'n' ? text.replace(/^\d+\.\s+/, '') : text}</li>`;
      continue;
    }
    close();
    html += {
      h2: `<h4>${text}</h4>`,
      h3: `<h5>${text}</h5>`,
      p: `<p>${text}</p>`,
      code: `<pre>${text}</pre>`,
      tip: `<div class="guide-tip">💡 ${text}</div>`,
      warn: `<div class="guide-warn">⚠ ${text}</div>`,
    }[kind] || '';
  }
  close();
  return `<div class="guide">${html}</div>`;
}

async function showGuide(name) {
  const guide = GUIDES[name];
  const shown = dialog(guide.title, guideHtml(guide.content()), [
    ...(guide.actions || []).map(([label, value]) => ({ label, value })),
    { label: 'Close', value: false, primary: true },
  ], true);
  $('#dialogBody').scrollTop = 0;                // start at the top, not where the last guide was
  const ans = await shown;
  if (ans === 'xlsx' || ans === 'csv') createTemplate(ans);
}

function helpShortcuts() {
  const rows = [
    ['Open song database', `${MOD}O`], ['Save setlist', `${MOD}S`],
    ['Print setlist / Save as PDF', `${MOD}P`], ['Find song', `${MOD}F`],
    ['Go to Set 1–4 / library', 'Alt+1…4 / Alt+L'], ['Move song up / down', 'Alt+↑ / Alt+↓'],
    ['Add library song to current set', 'Enter, or double-click'],
    ['Open songsheet (set song)', 'Enter, or double-click'],
    ['Remove song from set', 'Delete'], ['Move the selection', '↑ / ↓'],
    ['Song menu (Replace with…, Move, Remove…)', 'Right-click, or Shift+F10'],
    ['Text size', 'View menu, or the browser zoom'],
    ['How to use - the basics', 'F1'],
  ];
  dialog('Keyboard Shortcuts', '<table>' + rows.map(([a, k]) => `<tr><td>${esc(k)}</td><td>${esc(a)}</td></tr>`).join('') + '</table>'
    + '<p class="muted">The desktop app uses Ctrl+1…4 and Ctrl+L. In a browser those switch tabs and go to the address bar, so they use Alt here.</p>',
  undefined, true);
}

// ---------------------------------------------------------------- templates
// Help > Create Template Spreadsheet / CSV: a ready-to-fill song list with the
// right headings and example rows (the same as the desktop app makes).
const TEMPLATE_STEPS = [
  'Open the template in Excel.',
  'Replace the example rows with your own songs - one song per row. '
    + 'Only SONG NAME is required; Artist, Style and Vocalist are optional.',
  'Optional - songsheet links: click a song name, press Ctrl+K, paste the web address '
    + 'of its PDF (e.g. a Google Drive link) and click OK.',
  'Save with Ctrl+S. Keep it as an Excel Workbook (.xlsx) - saving as CSV loses the links.',
  "Back in JJ's Setlist, choose File ▸ Open Song Database and pick this spreadsheet - "
    + 'or upload it to Google Drive and add its link in File ▸ Song Database Settings.',
];
const CSV_TEMPLATE_STEPS = [
  'Open the template in Excel (or any spreadsheet program).',
  'Replace the example rows with your own songs - one song per row. '
    + 'Only SONG NAME is required; the other columns are optional.',
  'Songsheet links: put the web address of each PDF (e.g. a Google Drive link) in the Songsheet column.',
  'Save it, keeping the CSV format.',
  "Back in JJ's Setlist, choose File ▸ Open Song Database and pick this file.",
];

const forPlatform = (text) => (IS_MAC ? text.replace(/Ctrl\+/g, 'Cmd+') : text);

function xlsxTemplate() {
  const songs = XLSX.utils.aoa_to_sheet([
    ['SONG NAME', 'Artist', 'Style', 'Vocalist'],
    ['Example Song (with a web link)', 'Example Artist', "Ballad, 60's", 'Lead singer'],
    ['Another Song (no link yet)', 'Another Artist', 'Rock', ''],
  ]);
  // A songsheet link embedded behind the song name, as Ctrl+K does in Excel.
  songs.A2.l = { Target: 'https://example.com/songsheets/example-song.pdf', Tooltip: 'Open the songsheet' };
  songs['!cols'] = [{ wch: 42 }, { wch: 28 }, { wch: 20 }, { wch: 14 }];
  // A second tab with the steps - the app only reads the first tab.
  const how = XLSX.utils.aoa_to_sheet([
    [`How to build your song list for ${APP_NAME}`], [''],
    ...TEMPLATE_STEPS.map((step, i) => [`${i + 1}.  ${forPlatform(step)}`]), [''],
    ['The app only reads the first tab (Songs). Extra columns are ignored, so you can keep notes in the sheet.'],
    ["More help in the app: Help ▸ Setting Up Your Spreadsheet."],
  ]);
  how['!cols'] = [{ wch: 110 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, songs, 'Songs');
  XLSX.utils.book_append_sheet(wb, how, 'How to use');
  return new Uint8Array(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
}

function csvTemplate() {
  const rows = [
    ['SONG NAME', 'Artist', 'Style', 'Vocalist', 'Songsheet'],
    ['Example Song (web link)', 'Example Artist', "Ballad, 60's", 'Adrian', 'https://example.com/songsheets/example-song.pdf'],
    ['Another Song (file link)', 'Another Artist', 'Rock', '', 'Songsheets\\Another Song.pdf'],
    ['A Third Song (no link yet)', 'Third Artist', '', '', ''],
  ];
  // UTF-8 with BOM so Excel opens accented names correctly.
  return '﻿' + rows.map((r) => r.map(csvField).join(',')).join('\r\n') + '\r\n';
}

async function createTemplate(kind = 'xlsx') {
  const csv = kind === 'csv';
  const filename = csv ? 'JJ-SETLIST-TEMPLATE.csv' : 'JJ-SETLIST-TEMPLATE.xlsx';
  const type = csv ? 'text/csv' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  let data;
  try {
    data = csv ? csvTemplate() : xlsxTemplate();
  } catch (err) {
    await messageBox('Create template', `The template couldn't be made:\n\n${err.message}`);
    return;
  }
  let saved = null;                          // with the helper: { file, path } in the project folder
  if (files.on) {
    try {
      saved = await api('PUT', 'template/' + encodeURIComponent(filename), new Blob([data], { type }));
    } catch (err) {
      await messageBox('Create template', `The template couldn't be saved in the project folder:\n\n${err.message}\n\nIt will be downloaded instead.`);
    }
  }
  if (!saved) download(filename, data, type);
  status(saved ? `Template created: ${saved.path}` : `Template downloaded: ${filename}`);
  showTemplateSteps(csv, saved, filename);
}

async function showTemplateSteps(csv, saved, filename) {
  const steps = (csv ? CSV_TEMPLATE_STEPS : TEMPLATE_STEPS).map(forPlatform);
  const where = saved
    ? `Saved in the project folder:<br><code>${esc(saved.path)}</code>`
    : `Downloaded as <code>${esc(filename)}</code> - look in your <b>Downloads</b> folder. Move it somewhere handy, e.g. the JJ's Setlist folder.`;
  const html = `<p class="template-done">✓  Template created</p>
    <p class="muted">${where}</p>
    <p><strong>Now build your own song list:</strong></p>
    <ol class="steps">${steps.map((s) => `<li>${esc(s)}</li>`).join('')}</ol>
    ${csv ? '' : '<p class="muted small">These steps are also on the “How to use” tab inside the template.</p>'}`;
  const buttons = [
    ...(saved ? [{ label: 'Open it in Excel', value: 'open', primary: true }] : []),
    { label: 'Full guide', value: 'guide' },
    { label: 'Done', value: false, primary: !saved },
  ];
  const ans = await dialog('Your song spreadsheet - next steps', html, buttons, true);
  if (ans === 'open') {
    try {
      await api('PUT', 'open/' + encodeURIComponent(saved.file));
      status(`Opened “${saved.file}”.`);
    } catch (err) {
      await messageBox('Open template', `“${saved.file}” couldn't be opened:\n\n${err.message}`);
    }
  } else if (ans === 'guide') {
    spreadsheetGuide();
  }
}

function spreadsheetGuide() { return showGuide('spreadsheet'); }

function about() {
  dialog('About', `<div style="text-align:center"><img src="logo.png" width="110" alt=""><h3 style="margin:8px 0 2px">${APP_NAME}</h3>
    <p class="muted" style="margin:0 0 12px">Version ${APP_VERSION} (browser edition)</p>
    <p>Brought to you by JELLY JAZZ.<br>Vibe Coding by Adrian Newington.</p>
    <p><a href="${APP_WEBSITE}" target="_blank" rel="noopener">${APP_WEBSITE.replace('https://', '')}</a> ·
    <a href="${DONATE_URL}" target="_blank" rel="noopener">Buy us a coffee ☕</a></p></div>`, undefined, true);
}

// ---------------------------------------------------------------- events
const ACTIONS = {
  openDb: () => $('#dbFile').click(),
  dbSettings: () => databaseSettings(),
  loadDrive: () => loadFromDrive(),
  importSetlists: () => $('#setlistFiles').click(),
  newSetlist, save, load: loadSelected, deleteSaved,
  exportJson, exportText, exportCsv, exportDbCsv, printSetlist, printSonglist,
  addSelected, openSheet: () => openSheet(), clearAll,
  set0: () => focusSet(0), set1: () => focusSet(1), set2: () => focusSet(2), set3: () => focusSet(3),
  focusLibrary: () => libList.focus(),
  fontBigger: () => setFont(state.fontSize + 1), fontSmaller: () => setFont(state.fontSize - 1),
  fontNormal: () => setFont(DEFAULT_FONT), colours: coloursDialog,
  helpBasics: () => showGuide('basics'), helpShortcuts,
  helpSpreadsheet: () => showGuide('spreadsheet'),
  helpCsv: () => showGuide('csv'),
  helpAccount: () => showGuide('account'),
  templateXlsx: () => createTemplate('xlsx'),
  templateCsv: () => createTemplate('csv'),
  support: () => dialog('Support', `<p>${APP_NAME} is free and open source. If it's useful to you, you can buy us a coffee - any amount, entirely optional:</p>`
    + `<p><a href="${DONATE_URL}" target="_blank" rel="noopener">${DONATE_URL.replace('https://', '')}</a></p>`, undefined, true),
  about,
};

function focusSet(n) {
  setActive(n);
  setLists[n].focus();
  if (state.sets[n].length && state.sel.list !== n) select(n, 0);
}

function closeMenus() {
  document.querySelectorAll('.menu.open').forEach((m) => m.classList.remove('open'));
}

function wireEvents() {
  // Buttons and menu items
  document.addEventListener('click', (e) => {
    const title = e.target.closest('.menu-title');
    if (title) {
      const menu = title.parentElement;
      const wasOpen = menu.classList.contains('open');
      closeMenus();
      if (!wasOpen) menu.classList.add('open');
      return;
    }
    const btn = e.target.closest('[data-action]');
    if (!e.target.closest('.menu-items') || btn) closeMenus();
    if (btn && ACTIONS[btn.dataset.action]) ACTIONS[btn.dataset.action]();
  });
  document.querySelectorAll('.menu').forEach((menu) => menu.addEventListener('mouseenter', () => {
    if (document.querySelector('.menu.open') && !menu.classList.contains('open')) {
      closeMenus();
      menu.classList.add('open');
    }
  }));
  document.querySelectorAll('kbd[data-keys]').forEach((k) => {
    k.textContent = k.dataset.keys.replace('Ctrl+', MOD).replace('Alt+', IS_MAC ? '⌥' : 'Alt+');
  });

  // Files
  $('#dbFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (file) await openDatabaseFile(file);
  });
  $('#setlistFiles').addEventListener('change', async (e) => {
    const files = [...e.target.files];
    e.target.value = '';
    if (files.length) await importSetlistFiles(files);
  });
  // Drop a spreadsheet or saved setlists anywhere on the window.
  document.addEventListener('dragover', (e) => {
    if (e.dataTransfer?.types?.includes('Files')) e.preventDefault();
  });
  document.addEventListener('drop', async (e) => {
    const files = [...(e.dataTransfer?.files || [])];
    if (!files.length) return;
    e.preventDefault();
    const sheet = files.find((f) => /\.(xlsx|xlsm|csv)$/i.test(f.name));
    const jsons = files.filter((f) => /\.json$/i.test(f.name));
    if (sheet) await openDatabaseFile(sheet);
    if (jsons.length) await importSetlistFiles(jsons);
  });

  // Search and name
  $('#search').addEventListener('input', () => { state.sel = { list: 'lib', index: -1 }; renderLibrary(); });
  $('#search').addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { e.target.value = ''; renderLibrary(); }
    if (e.key === 'ArrowDown' || e.key === 'Enter') {
      e.preventDefault();
      libList.focus();
      if (state.filtered.length) select('lib', 0);
    }
  });
  $('#setName').addEventListener('input', (e) => {
    state.name = e.target.value;
    state.dirty = true;
    saveWorking();
    renderTotals();
  });
  $('#setName').addEventListener('keydown', (e) => { if (e.key === 'Enter') save(); });

  // Lists: select, double-click, + and 📄 buttons
  for (const ul of [libList, ...setLists]) {
    const list = ul.dataset.list === 'lib' ? 'lib' : +ul.dataset.list;
    ul.addEventListener('click', (e) => {
      const li = e.target.closest('li');
      if (!li) return;
      const i = +li.dataset.i;
      select(list, i, false);
      if (list !== 'lib') setActive(list);
      const song = selectedSong();
      if (e.target.closest('[data-sheet]')) openSheet(song);
    });
    ul.addEventListener('dblclick', (e) => {
      if (e.target.closest('[data-sheet]') || !e.target.closest('li')) return;
      if (list === 'lib') addSelected(); else openSheet();
    });
    ul.addEventListener('keydown', (e) => listKey(e, list));
    ul.addEventListener('contextmenu', (e) => openSongMenu(e, list));
  }

  // Right-click menu: close it when clicking elsewhere, scrolling, or with Esc.
  document.body.appendChild(ctxMenu);
  document.addEventListener('pointerdown', (e) => { if (!ctxMenu.contains(e.target)) closeContextMenu(); }, true);
  document.addEventListener('scroll', closeContextMenu, true);
  window.addEventListener('blur', closeContextMenu);
  window.addEventListener('resize', closeContextMenu);
  ctxMenu.addEventListener('keydown', (e) => {
    const buttons = [...ctxMenu.querySelectorAll('button:not(:disabled)')];
    const k = buttons.indexOf(document.activeElement);
    if (e.key === 'Escape') { e.preventDefault(); closeContextMenu(); }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      buttons[(k + (e.key === 'ArrowDown' ? 1 : buttons.length - 1)) % buttons.length]?.focus();
    }
  });

  // Global shortcuts
  document.addEventListener('keydown', (e) => {
    if (dlg.open) return;
    const mod = IS_MAC ? e.metaKey : e.ctrlKey;
    const key = e.key.toLowerCase();
    if (mod && !e.altKey) {
      const fn = { o: ACTIONS.openDb, s: save, p: printSetlist, f: () => $('#search').focus() }[key];
      if (fn && !e.shiftKey) { e.preventDefault(); fn(); }
      return;
    }
    if (e.altKey && !mod) {
      if (/^[1-4]$/.test(e.key) || /^Digit[1-4]$/.test(e.code)) {
        e.preventDefault();
        focusSet(+(e.code.slice(-1) || e.key) - 1);
      } else if (e.code === 'KeyL') { e.preventDefault(); libList.focus(); }
    }
    if (e.key === 'F1') { e.preventDefault(); showGuide('basics'); }
    if (e.key === 'Escape') closeMenus();
  });

  // Close menus when clicking outside them
  document.addEventListener('pointerdown', (e) => { if (!e.target.closest?.('.menu')) closeMenus(); });
}

function listKey(e, list) {
  const length = list === 'lib' ? state.filtered.length : state.sets[list].length;
  const current = state.sel.list === list ? state.sel.index : -1;
  if (e.altKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown') && list !== 'lib') {
    e.preventDefault();
    move(list, e.key === 'ArrowUp' ? -1 : 1);
    setLists[list].focus();
    return;
  }
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault();
    if (!length) return;
    const next = Math.max(0, Math.min(length - 1, current + (e.key === 'ArrowDown' ? 1 : -1)));
    select(list, current < 0 ? 0 : next);
  } else if (e.key === 'Home' || e.key === 'End') {
    e.preventDefault();
    if (length) select(list, e.key === 'Home' ? 0 : length - 1);
  } else if (e.key === 'Enter') {
    e.preventDefault();
    if (list === 'lib') addSelected(); else openSheet();
  } else if ((e.key === 'Delete' || e.key === 'Backspace') && list !== 'lib') {
    e.preventDefault();
    remove(list);
    setLists[list].focus();
  }
}

// ---------------------------------------------------------------- start
async function start() {
  if (typeof XLSX === 'undefined' || typeof Sortable === 'undefined') {
    document.body.insertAdjacentHTML('afterbegin',
      '<div style="background:#d35400;color:#fff;padding:8px 12px">Some parts of the app couldn\'t load - please reload the page.</div>');
  }
  // Where to save: a signed-in online account, else the helper's files, else this browser.
  await Promise.all([initCloud(), findHelper()]);
  let newAccount = false;
  if (cloud.user) {
    try {
      newAccount = await loadCloudData();
      mode = 'cloud';
    } catch (err) {
      console.error(err);
      store.prefix = BASE_PREFIX;
      loadLocalState();
      if (err.code === 'jjs/not-registered') {
        // A login that didn't come with an invitation code: it can't use the app.
        messageBox('Sign-up needs an invitation code', `${cloudMessage(err)}\n\nYou've been signed out. If you have a code, choose Sign in ▸ “New here?”.`)
          .then(() => cloud.auth.signOut());
      } else {
        messageBox('Could not open your account', `Your account's data couldn't be loaded:\n\n${cloudMessage(err)}\n\nFor now, changes are kept in this browser only. Reload the page to try again.`);
      }
    }
  } else if (files.on) {
    try {
      await loadHelperFiles();
      mode = 'files';
    } catch { /* the helper vanished: browser storage */ }
  }
  buildSets();
  applyLook();
  renderStorage();
  // Pick up setlists saved meanwhile (by the desktop app, or on another computer) when coming back.
  window.addEventListener('focus', refreshSavedSetlists);
  $('#setName').value = state.name;
  wireEvents();
  if (typeof Sortable !== 'undefined') setupDragging();
  renderAll();
  renderSaved();
  setActive(0);
  if (state.library.length) {
    status(`Song database “${state.dbName}” (remembered from last time).`);
  } else {
    status('Open your song spreadsheet to begin: File ▸ Open Song Database…');
  }
  // Just joined a band: say hello. A brand-new personal space: offer to bring in what's on this computer.
  let welcome = null;
  try { welcome = sessionStorage.getItem('jjs.welcome'); sessionStorage.removeItem('jjs.welcome'); } catch { /* no storage */ }
  if (cloud.notice) await messageBox(...cloud.notice);   // e.g. removed from a band - say so first
  if (welcome && mode === 'cloud') {
    messageBox(`Welcome to ${welcome}`, `You're now a member of ${welcome}, with full access to its song library and setlists.\n\n`
      + 'Switch between the band and your own setlists with the list at the top right, next to your email address.');
  } else if (newAccount) {
    offerImport(true);
  }
  // Google Drive copy: each time the app opens when it's the first choice,
  // or as the backup when no song database is loaded here.
  if (state.driveUrl && (state.driveFirst || !state.library.length)) {
    if (OPENED_AS_FILE) status('Google Drive copy not loaded: it needs JJ\'s Setlist opened from a web address (see the README).');
    else loadFromDrive({ startup: true });
  }
}

start();
