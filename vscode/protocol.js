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
        rec = { n: msg.n || 1, name: msg.name || '', purpose: msg.purpose || 'save', uri: msg.uri || '', folder: msg.folder || '', draftKey: msg.draftKey || '', parts: [], got: 0 };
        pending.set(msg.id, rec);
      }
      if (msg.uri) rec.uri = msg.uri;
      if (msg.folder) rec.folder = msg.folder;
      if (msg.draftKey) rec.draftKey = msg.draftKey;
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
      return { id: msg.id, name: rec.name, purpose: rec.purpose, uri: rec.uri || '', folder: rec.folder || '', draftKey: rec.draftKey || '', bytes: out };
    },
  };
}

module.exports = {
  APPS, DEFAULT_EXT, SAVE_EXT, SAVE_STEM, extname, appFor, viewTypeFor, baseName, safeName, sameFile, ensureExt, saveFolder, toBytes, createAssembler,
};
