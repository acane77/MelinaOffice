/* Shared by the VS Code extension and its tests: which app opens which file, and how a
 * save is split into messages the webview can carry (code-server posts JSON, so bytes
 * travel as base64). Keep the extension lists in step with VO.APPS in public/common/suite.js. */
'use strict';

const APPS = {
  quire: ['docx', 'docm', 'dotx', 'dotm', 'doc', 'rtf', 'htm', 'html', 'txt', 'md', 'markdown', 'zip'],
  ledger: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xls', 'csv', 'tsv', 'prn', 'xml'],
  lectern: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm', 'ppt'],
};

/** Office Open XML opens in the editor on a double-click; the other types are Open With. */
const DEFAULT_EXT = {
  quire: ['docx', 'docm', 'dotx', 'dotm'],
  ledger: ['xlsx', 'xlsm', 'xltx', 'xltm'],
  lectern: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm'],
};

/** File ▸ New writes this type, whichever file the editor tab was opened on. */
const SAVE_EXT = { quire: 'docx', ledger: 'xlsx', lectern: 'pptx' };
const SAVE_STEM = { quire: 'Document', ledger: 'Book', lectern: 'Presentation' };

function extname(name) {
  const m = /\.([^./\\]+)$/.exec(String(name || ''));
  return m ? m[1].toLowerCase() : '';
}

function appFor(name) {
  const e = extname(name);
  return Object.keys(APPS).find((a) => APPS[a].includes(e)) || null;
}

function viewTypeFor(name) {
  const app = appFor(name);
  return app ? 'vibeoffice.' + app : null;
}

/** the last path segment, with any directory stripped */
function baseName(name) {
  const s = String(name || '').replace(/\\/g, '/');
  const i = s.lastIndexOf('/');
  return i >= 0 ? s.slice(i + 1) : s;
}

/** a save name the editor may write: no directory, so a page cannot aim outside the file's folder */
function safeName(name) {
  const base = baseName(name);
  if (!base || base === '.' || base === '..') return '';
  return base;
}

function sameFile(savedName, openName) {
  const a = safeName(savedName), b = baseName(openName);
  return !!a && a.toLowerCase() === b.toLowerCase();
}

/** the name File ▸ New will write. A suffix the user typed is kept, and the editor's own suffix is added after it:
 *  新建文档.txt with docx becomes 新建文档.txt.docx. A name that already ends with that suffix is left as it is. */
function ensureExt(name, ext, stem) {
  const e = String(ext || '').replace(/^\./, '').toLowerCase();
  let base = safeName(name);
  if (!base) base = (stem || 'Document') + (e ? '.' + e : '');
  if (!e || base.toLowerCase().endsWith('.' + e)) return base;
  return base + '.' + e;
}

function toBytes(b) {
  if (b instanceof Uint8Array) return b;
  if (typeof Buffer !== 'undefined' && Buffer.isBuffer && Buffer.isBuffer(b)) return new Uint8Array(b);
  if (typeof b === 'string') return new Uint8Array(Buffer.from(b, 'base64'));
  if (Array.isArray(b)) return Uint8Array.from(b);
  if (b && b.type === 'Buffer' && Array.isArray(b.data)) return Uint8Array.from(b.data);
  return new Uint8Array();
}

/** a folder File ▸ Save As may write into. An empty ask is the open file's own folder.
 *  Anything else must be that folder, or one the user picked in this editor (allowed is a Set of URI strings). */
function saveFolder(openDir, asked, allowed) {
  const home = String(openDir || '');
  const want = String(asked || '');
  if (!home) return '';
  if (!want || want === home) return home;
  if (allowed && allowed.has && allowed.has(want)) return want;
  return '';
}

/** reassemble chunk messages. push() returns { id, name, purpose, uri, folder, draftKey, bytes } once every part has arrived. */
function createAssembler() {
  const pending = new Map();
  return {
    push(msg) {
      if (!msg || msg.type !== 'chunk' || msg.id == null) return null;
      let rec = pending.get(msg.id);
      if (!rec) {
        rec = { n: msg.n || 1, name: msg.name || '', purpose: msg.purpose || 'save', uri: msg.uri || '', folder: msg.folder || '', draftKey: msg.draftKey || '', path: msg.path || '', ability: msg.ability || '', open: !!msg.open, parts: [], got: 0 };
        pending.set(msg.id, rec);
      }
      if (msg.uri) rec.uri = msg.uri;
      if (msg.folder) rec.folder = msg.folder;
      if (msg.draftKey) rec.draftKey = msg.draftKey;
      if (msg.path) rec.path = msg.path;
      if (msg.ability) rec.ability = msg.ability;
      if (msg.open) rec.open = true;
      if (rec.parts[msg.i] == null) {
        rec.parts[msg.i] = toBytes(msg.b64 != null ? msg.b64 : msg.bytes);
        rec.got++;
      }
      if (rec.got < rec.n) return null;
      pending.delete(msg.id);
      let size = 0;
      for (let i = 0; i < rec.n; i++) size += (rec.parts[i] ? rec.parts[i].length : 0);
      const out = new Uint8Array(size);
      let o = 0;
      for (let i = 0; i < rec.n; i++) {
        const p = rec.parts[i];
        if (p) { out.set(p, o); o += p.length; }
      }
      return { id: msg.id, name: rec.name, purpose: rec.purpose, uri: rec.uri || '', folder: rec.folder || '', draftKey: rec.draftKey || '', path: rec.path || '', ability: rec.ability || '', open: !!rec.open, bytes: out };
    },
  };
}

/** One custom-editor save. VS Code's tab dot is this session plus the edit it fired, not the
 *  page's "Saved" line. A second save joins the first. A dirty notice during the save is held
 *  and, if the document is still dirty when the save ends, becomes one new edit so the dot
 *  does not stick and a real keystroke during the save is not forgotten. Dispose after the
 *  bytes are on disk resolves the save; dispose before that rejects it once. */
function createSaveSession() {
  const s = {
    dirty: false,
    saving: false,
    skip: false,
    wrote: false,
    closed: false,
    dirtyAfter: false,
    runs: 0,
    chain: null,
    _wait: null,
  };
  function end(ok, value, error) {
    const w = s._wait;
    if (!w) return false;
    s._wait = null;
    const again = !!(ok && s.dirtyAfter && !s.closed);
    s.dirtyAfter = false;
    s.saving = false;
    if (ok) {
      s.dirty = again;
      w.resolve(value || { ok: true });
    } else {
      w.reject(error instanceof Error ? error : new Error(String(error || 'Save was cancelled.')));
    }
    return again;
  }
  return {
    get dirty() { return s.dirty; },
    set dirty(v) { s.dirty = !!v; },
    get saving() { return s.saving; },
    get wrote() { return s.wrote; },
    get closed() { return s.closed; },
    get chain() { return s.chain; },
    get runs() { return s.runs; },
    noteWritten() { if (s.saving) s.wrote = true; },
    armSkip() { s.skip = true; },
    /** drop a skip that this save did not consume, so the next real save still writes */
    disarmSkip() { s.skip = false; },
    begin() {
      if (s.skip) {
        s.skip = false;
        s.dirty = false;
        return { action: 'skip', promise: Promise.resolve() };
      }
      if (s.chain) return { action: 'join', promise: s.chain };
      s.runs++;
      s.saving = true;
      s.wrote = false;
      s.dirtyAfter = false;
      let resolve, reject;
      const inner = new Promise((res, rej) => { resolve = res; reject = rej; });
      s._wait = { resolve, reject };
      const chain = inner.finally(() => {
        s.saving = false;
        if (s.chain === chain) s.chain = null;
      });
      /* dispose can reject this before saveCustomDocument returns it */
      chain.catch(() => {});
      s.chain = chain;
      return { action: 'run', promise: chain };
    },
    succeed(value) { return end(true, value); },
    fail(error) { return end(false, null, error); },
    /** the custom document or its webview went away */
    close() {
      s.closed = true;
      if (!s._wait) return false;
      if (s.wrote) return end(true, { ok: true, inplace: true });
      return end(false, null, new Error('The editor was closed.'));
    },
    /** @returns {'mark'|'ignore'} */
    onDirty(dirty) {
      if (s.closed) return 'ignore';
      if (s.saving) {
        if (dirty) s.dirtyAfter = true;
        return 'ignore';
      }
      if (!dirty || s.dirty) return 'ignore';
      s.dirty = true;
      return 'mark';
    },
    /** @returns {'join'|'start'|'closed'} */
    onRequest() {
      if (s.closed) return 'closed';
      if (s.chain) return 'join';
      return 'start';
    },
  };
}

/** File ▸ Enable Auto-Save. Off until setEnabled(true). note(key, dirty, save) waits delayMs
 *  after the latest edit, then calls save once. A clean document, or turning it off, cancels
 *  the wait. A later edit starts the wait again, so typing does not save on every key. */
function createAutoSave(delayMs) {
  const delay = delayMs == null ? 1000 : delayMs;
  let enabled = false;
  const timers = new Map();
  function clear(key) {
    const t = timers.get(key);
    if (t) clearTimeout(t);
    timers.delete(key);
  }
  return {
    get enabled() { return enabled; },
    setEnabled(v) {
      enabled = !!v;
      if (!enabled) { for (const k of [...timers.keys()]) clear(k); }
    },
    note(key, dirty, save) {
      clear(key);
      if (!enabled || !dirty || typeof save !== 'function') return;
      timers.set(key, setTimeout(() => {
        timers.delete(key);
        if (!enabled) return;
        try { save(); } catch (e) { /* the caller reports a failed save */ }
      }, delay));
    },
    clear,
    pending(key) { return timers.has(key); },
  };
}

module.exports = {
  APPS, DEFAULT_EXT, SAVE_EXT, SAVE_STEM, extname, appFor, viewTypeFor, baseName, safeName, sameFile, ensureExt, saveFolder, toBytes, createAssembler, createSaveSession, createAutoSave,
};
