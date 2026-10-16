/* VibeOffice as a VS Code extension. code-server runs this in its Node extension host and
 * draws the same app pages in a webview, so a file in the workspace opens in Quire, Ledger
 * or Lectern. Save writes that file; another name is written beside it and opened too.
 * File ▸ New and File ▸ Open use VS Code's own dialogs and open the result in a new tab.
 * Docs: docs/vscode.md. */
'use strict';

const fs = require('fs');
const path = require('path');
const vscode = require('vscode');
const proto = require('./protocol');
const page = require('./page');

const VIEW = {
  quire: 'vibeoffice.quire',
  ledger: 'vibeoffice.ledger',
  lectern: 'vibeoffice.lectern',
};

const SAVE_LABEL = { quire: 'Document', ledger: 'Workbook', lectern: 'Presentation' };

/** bytes for a Recent Files entry that should open in place of the file on disk, keyed by the file's URI */
const pendingOpen = new Map();

function publicRoot(extensionPath) {
  const candidates = [path.join(extensionPath, 'public'), path.join(extensionPath, '..', 'public')];
  for (const c of candidates) {
    if (fs.existsSync(path.join(c, 'quire', 'index.html'))) return c;
  }
  return null;
}

function platform() {
  if (process.platform === 'win32') return 'windows';
  if (process.platform === 'darwin') return 'macos';
  return 'linux';
}

function b64(bytes) {
  return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString('base64');
}

function parentUri(uri) {
  return vscode.Uri.joinPath(uri, '..');
}

function extsFromAccept(accept, fallback) {
  if (!accept) return fallback;
  const out = [];
  for (const part of String(accept).split(',')) {
    const m = /^\*?\.?([a-z0-9]+)$/i.exec(part.trim());
    if (m) out.push(m[1].toLowerCase());
  }
  return out.length ? out : fallback;
}

/** a VibeOffice tab already showing this file, if there is one */
function openTab(uri) {
  const key = uri.toString();
  for (const group of vscode.window.tabGroups.all) {
    for (const tab of group.tabs) {
      const input = tab.input;
      if (!input || !input.uri || input.uri.toString() !== key) continue;
      if (input.viewType && String(input.viewType).indexOf('vibeoffice.') === 0) return input.viewType;
    }
  }
  return '';
}

async function reveal(uri, viewType) {
  const view = openTab(uri);
  if (!view) return false;
  await vscode.commands.executeCommand('vscode.openWith', uri, view || viewType);
  return true;
}

/** a free name in the file's folder: the requested name, or "Name (1).ext" when that one exists */
async function placeBeside(dir, name, openUri) {
  const safe = proto.safeName(name) || 'Document';
  const ext = (safe.match(/\.[^.]+$/) || [''])[0];
  const stem = safe.slice(0, safe.length - ext.length) || 'Document';
  let uri = vscode.Uri.joinPath(dir, safe);
  for (let n = 0; n < 50; n++) {
    if (openUri && uri.toString() === openUri.toString()) return uri;
    try { await vscode.workspace.fs.stat(uri); }
    catch { return uri; }
    uri = vscode.Uri.joinPath(dir, stem + ' (' + (n + 1) + ')' + ext);
  }
  return uri;
}

class OfficeDocument {
  constructor(uri, viewType) {
    this.uri = uri;
    this.viewType = viewType;
    this.panel = null;
    this.pending = new Map();
    this.dirty = false;
    this.savingFromHost = false;
    this.skipSave = false;
    this.backupId = '';
  }
  dispose() {
    for (const p of this.pending.values()) p.reject(new Error('The editor was closed.'));
    this.pending.clear();
  }
}

class OfficeEditorProvider {
  constructor(context, app) {
    this._context = context;
    this._app = app;
    this._public = publicRoot(context.extensionPath);
    this._assembler = proto.createAssembler();
    this._onDidChange = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this._onDidChange.event;
  }

  async openCustomDocument(uri, openContext) {
    const doc = new OfficeDocument(uri, VIEW[this._app]);
    doc.backupId = openContext && openContext.backupId || '';
    return doc;
  }

  async resolveCustomEditor(doc, panel) {
    doc.panel = panel;
    const root = this._public ? vscode.Uri.file(this._public) : null;
    const roots = [];
    if (root) roots.push(root);
    roots.push(parentUri(doc.uri));
    panel.webview.options = { enableScripts: true, localResourceRoots: roots };
    panel.webview.onDidReceiveMessage((msg) => this._onMessage(doc, msg));
    panel.webview.html = this._html(doc);
  }

  async saveCustomDocument(doc) {
    if (doc.skipSave) { doc.skipSave = false; doc.dirty = false; return; }
    doc.savingFromHost = true;
    try {
      const res = await this._request(doc, 'save-result', { type: 'do-save' });
      if (!res || !res.ok) throw new Error((res && res.error) || 'Save was cancelled.');
      doc.dirty = false;
    } finally {
      doc.savingFromHost = false;
    }
  }

  async saveCustomDocumentAs(doc, destination) {
    const bytes = await this._backupBytes(doc);
    await vscode.workspace.fs.writeFile(destination, bytes);
    doc.dirty = false;
  }

  async revertCustomDocument(doc) {
    doc.dirty = false;
    doc.backupId = '';
    if (doc.panel) doc.panel.webview.html = this._html(doc);
  }

  async backupCustomDocument(doc, context) {
    const bytes = await this._backupBytes(doc);
    await vscode.workspace.fs.writeFile(context.destination, bytes);
    return {
      id: context.destination.toString(),
      delete: () => vscode.workspace.fs.delete(context.destination),
    };
  }

  _html(doc) {
    if (!this._public) return page.missingPage('VibeOffice could not find its pages next to the extension.');
    const file = path.join(this._public, this._app, 'index.html');
    let html;
    try { html = fs.readFileSync(file, 'utf8'); }
    catch (e) { return page.missingPage('VibeOffice could not read ' + this._app + '. ' + (e && e.message ? e.message : '')); }
    const base = doc.panel.webview.asWebviewUri(vscode.Uri.file(path.join(this._public, this._app))).toString();
    return page.renderPage({
      html,
      baseHref: base,
      cspSource: doc.panel.webview.cspSource,
      fileName: proto.baseName(doc.uri.path),
      platform: platform(),
      lang: this._context.globalState.get('vibeoffice.lang') || '',
    });
  }

  _request(doc, expect, msg) {
    const id = Math.random().toString(36).slice(2);
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        doc.pending.delete(id);
        reject(new Error('The editor did not answer.'));
      }, 300000);
      doc.pending.set(id, {
        expect,
        resolve: (v) => { clearTimeout(timer); doc.pending.delete(id); resolve(v); },
        reject: (e) => { clearTimeout(timer); doc.pending.delete(id); reject(e); },
      });
      doc.panel.webview.postMessage(Object.assign({ id }, msg));
    });
  }

  async _backupBytes(doc) {
    try {
      const id = Math.random().toString(36).slice(2);
      const bytes = await new Promise((resolve, reject) => {
        const timer = setTimeout(() => { doc.pending.delete(id); reject(new Error('backup timed out')); }, 300000);
        doc.pending.set(id, {
          expect: 'bytes',
          resolve: (v) => { clearTimeout(timer); doc.pending.delete(id); resolve(v); },
          reject: (e) => { clearTimeout(timer); doc.pending.delete(id); reject(e); },
        });
        doc.panel.webview.postMessage({ type: 'backup', id });
      });
      if (bytes && bytes.length) return bytes;
    } catch (e) { /* the page has nothing to give yet: keep the file that is on disk */ }
    if (doc.backupId) {
      try { return await vscode.workspace.fs.readFile(vscode.Uri.parse(doc.backupId)); }
      catch (e) { /* fall through to the workspace file */ }
    }
    return vscode.workspace.fs.readFile(doc.uri);
  }

  _markDirty(doc) {
    if (doc.dirty) return;
    doc.dirty = true;
    this._onDidChange.fire({
      document: doc,
      undo() { doc.dirty = false; },
      redo() { doc.dirty = true; },
    });
  }

  async _onMessage(doc, msg) {
    if (!msg || typeof msg !== 'object') return;
    if (msg.type === 'ready') { await this._sendOpen(doc); return; }
    if (msg.type === 'read-fallback') { await this._sendBytes(doc, null, proto.baseName(doc.uri.path)); return; }
    if (msg.type === 'dirty') {
      if (msg.dirty) this._markDirty(doc);
      else if (!doc.savingFromHost && doc.dirty) {
        doc.skipSave = true;
        await vscode.commands.executeCommand('workbench.action.files.save');
      }
      return;
    }
    if (msg.type === 'close') {
      /* the page already asked about unsaved changes; clear the tab's dot without writing them again */
      try {
        if (doc.dirty) {
          doc.skipSave = true;
          await vscode.commands.executeCommand('workbench.action.files.save');
        }
      } catch (e) { /* close the tab either way */ }
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
      return;
    }
    if (msg.type === 'pick-new') { await this._pickNew(doc, msg); return; }
    if (msg.type === 'pick-open') { await this._pickOpen(doc, msg); return; }
    if (msg.type === 'open-uri') { await this._openUri(doc, msg); return; }
    if (msg.type === 'lang') {
      const id = String(msg.lang || '');
      await this._context.globalState.update('vibeoffice.lang', id);
      /* assigning the same HTML does nothing, so the new language has to be part of the page */
      if (msg.reload && doc.panel) doc.panel.webview.html = this._html(doc);
      return;
    }
    if (msg.type === 'open-url') {
      try {
        const u = vscode.Uri.parse(String(msg.url || ''));
        if (u.scheme === 'http' || u.scheme === 'https' || u.scheme === 'mailto') await vscode.env.openExternal(u);
      } catch (e) { /* a page link that is not a web or mail address stays put */ }
      return;
    }
    if (msg.type === 'save-result' || (msg.type === 'wrote' && msg.ok === false)) {
      const p = doc.pending.get(msg.id);
      if (!p) return;
      if (msg.type === 'wrote' && msg.ok === false) p.reject(new Error(msg.error || 'The file could not be saved.'));
      else p.resolve(msg);
      return;
    }
    if (msg.type === 'chunk') {
      const done = this._assembler.push(msg);
      if (!done) return;
      if (done.purpose === 'backup') {
        const p = doc.pending.get(done.id);
        if (doc.panel) doc.panel.webview.postMessage({ type: 'wrote', id: done.id, ok: true, inplace: false, name: done.name });
        if (p) p.resolve(done.bytes);
        return;
      }
      if (done.purpose === 'create' || done.purpose === 'reopen') {
        await this._acceptBytes(doc, done);
        return;
      }
      try {
        const inplace = proto.sameFile(done.name, doc.uri.path);
        const target = inplace ? doc.uri : await placeBeside(parentUri(doc.uri), done.name, doc.uri);
        await vscode.workspace.fs.writeFile(target, done.bytes);
        const wroteInPlace = target.toString() === doc.uri.toString();
        if (!wroteInPlace) await vscode.commands.executeCommand('vscode.openWith', target, doc.viewType);
        if (doc.panel) doc.panel.webview.postMessage({
          type: 'wrote', id: done.id, ok: true, inplace: wroteInPlace, name: proto.baseName(target.path),
        });
      } catch (e) {
        const error = e && e.message ? e.message : String(e);
        if (doc.panel) doc.panel.webview.postMessage({ type: 'wrote', id: done.id, ok: false, error });
        vscode.window.showErrorMessage('VibeOffice could not save the file. ' + error);
      }
    }
  }

  _reply(doc, id, body) {
    if (doc.panel) doc.panel.webview.postMessage(Object.assign({ type: 'reply', id }, body));
  }

  /** the next free Document1.docx / Book1.xlsx / Presentation1.pptx beside the open file */
  async _freeName(doc) {
    const ext = proto.SAVE_EXT[this._app] || 'docx';
    const stem = proto.SAVE_STEM[this._app] || 'Document';
    const dir = parentUri(doc.uri);
    for (let n = 1; n < 100; n++) {
      const name = stem + n + '.' + ext;
      try { await vscode.workspace.fs.stat(vscode.Uri.joinPath(dir, name)); }
      catch { return name; }
    }
    return stem + Date.now() + '.' + ext;
  }

  _withExt(uri, ext) {
    const name = proto.ensureExt(proto.baseName(uri.path), ext, proto.SAVE_STEM[this._app]);
    return vscode.Uri.joinPath(parentUri(uri), name);
  }

  async _exists(uri) {
    try { await vscode.workspace.fs.stat(uri); return true; }
    catch { return false; }
  }

  /** File ▸ New: VS Code's save dialog, then the page writes a new document of this editor's type */
  async _pickNew(doc, msg) {
    try {
      const ext = String(msg.ext || proto.SAVE_EXT[this._app] || 'docx').replace(/^\./, '').toLowerCase();
      const label = SAVE_LABEL[this._app] || 'Document';
      const seed = msg.name ? proto.ensureExt(msg.name, ext, proto.SAVE_STEM[this._app]) : await this._freeName(doc);
      const filters = {};
      filters[label] = [ext];
      let uri = await vscode.window.showSaveDialog({
        defaultUri: vscode.Uri.joinPath(parentUri(doc.uri), seed),
        filters,
        saveLabel: 'Save',
      });
      if (!uri) { this._reply(doc, msg.id, { ok: false }); return; }
      /* a typed suffix stays, and this editor's suffix is appended. If that file is already there, ask again
         so the replace prompt names the file that will actually be written. */
      for (let i = 0; i < 2; i++) {
        const fixed = this._withExt(uri, ext);
        if (fixed.toString() === uri.toString()) break;
        if (!(await this._exists(fixed))) { uri = fixed; break; }
        const again = await vscode.window.showSaveDialog({ defaultUri: fixed, filters, saveLabel: 'Save' });
        if (!again) { this._reply(doc, msg.id, { ok: false }); return; }
        uri = again;
      }
      uri = this._withExt(uri, ext);
      const view = proto.viewTypeFor(uri.path) || doc.viewType;
      if (await reveal(uri, view)) { this._reply(doc, msg.id, { ok: true, alreadyOpen: true }); return; }
      if (!doc.allowedCreates) doc.allowedCreates = new Set();
      doc.allowedCreates.add(uri.toString());
      this._reply(doc, msg.id, { ok: true, uri: uri.toString(), name: proto.baseName(uri.path) });
    } catch (e) {
      const error = e && e.message ? e.message : String(e);
      this._reply(doc, msg.id, { ok: false, error });
      vscode.window.showErrorMessage('VibeOffice could not create the file. ' + error);
    }
  }

  /** File ▸ Open: VS Code's open dialog, then a tab for each file. Nothing is uploaded from the browser. */
  async _pickOpen(doc, msg) {
    try {
      const exts = extsFromAccept(msg.accept, proto.APPS[this._app] || []);
      const label = SAVE_LABEL[this._app] || 'Document';
      const filters = {};
      filters[label] = exts;
      const picked = await vscode.window.showOpenDialog({
        canSelectMany: this._app === 'ledger',
        canSelectFiles: true,
        defaultUri: parentUri(doc.uri),
        filters,
        openLabel: 'Open',
      });
      if (!picked || !picked.length) { this._reply(doc, msg.id, { ok: true, cancelled: true }); return; }
      for (const uri of picked) {
        /* the dialog belongs to this editor, so a shared type such as .txt opens here rather than in another app */
        await vscode.commands.executeCommand('vscode.openWith', uri, doc.viewType);
      }
      this._reply(doc, msg.id, { ok: true });
    } catch (e) {
      const error = e && e.message ? e.message : String(e);
      this._reply(doc, msg.id, { ok: false, error });
      vscode.window.showErrorMessage('VibeOffice could not open the file. ' + error);
    }
  }

  /** a Recent Files entry whose path was remembered: show that file, or say it is gone */
  async _openUri(doc, msg) {
    try {
      const uri = vscode.Uri.parse(String(msg.uri || ''));
      const view = proto.viewTypeFor(uri.path) || doc.viewType;
      if (await reveal(uri, view)) { this._reply(doc, msg.id, { ok: true, alreadyOpen: true }); return; }
      if (!(await this._exists(uri))) { this._reply(doc, msg.id, { ok: false, error: 'missing' }); return; }
      await vscode.commands.executeCommand('vscode.openWith', uri, view);
      this._reply(doc, msg.id, { ok: true });
    } catch (e) {
      this._reply(doc, msg.id, { ok: false, error: e && e.message ? e.message : String(e) });
    }
  }

  /** a new file the page built, or a Recent Files copy opened onto its original path */
  async _acceptBytes(doc, done) {
    const fail = (error) => {
      if (doc.panel) doc.panel.webview.postMessage({ type: 'wrote', id: done.id, ok: false, error });
      if (done.purpose === 'create') vscode.window.showErrorMessage('VibeOffice could not create the file. ' + error);
    };
    try {
      const target = vscode.Uri.parse(String(done.uri || ''));
      const view = proto.viewTypeFor(target.path) || doc.viewType;
      if (done.purpose === 'create' && !(doc.allowedCreates && doc.allowedCreates.has(done.uri))) {
        fail('That location was not chosen.');
        return;
      }
      if (await reveal(target, view)) {
        if (doc.allowedCreates) doc.allowedCreates.delete(done.uri);
        if (doc.panel) doc.panel.webview.postMessage({ type: 'wrote', id: done.id, ok: true, inplace: false, alreadyOpen: true, name: proto.baseName(target.path) });
        return;
      }
      if (done.purpose === 'reopen') {
        if (!(await this._exists(target))) { fail('missing'); return; }
        pendingOpen.set(target.toString(), { bytes: done.bytes, draftKey: done.draftKey || '' });
        try { await vscode.commands.executeCommand('vscode.openWith', target, view); }
        catch (e) { pendingOpen.delete(target.toString()); throw e; }
      } else {
        await vscode.workspace.fs.writeFile(target, done.bytes);
        if (doc.allowedCreates) doc.allowedCreates.delete(target.toString());
        await vscode.commands.executeCommand('vscode.openWith', target, view);
      }
      if (doc.panel) doc.panel.webview.postMessage({ type: 'wrote', id: done.id, ok: true, inplace: false, name: proto.baseName(target.path) });
    } catch (e) {
      fail(e && e.message ? e.message : String(e));
    }
  }

  async _sendOpen(doc) {
    const pending = pendingOpen.get(doc.uri.toString());
    if (pending) {
      pendingOpen.delete(doc.uri.toString());
      await this._sendBytes(doc, pending.bytes, proto.baseName(doc.uri.path), { draft: true, draftKey: pending.draftKey });
      return;
    }
    if (doc.backupId) {
      try {
        const bytes = await vscode.workspace.fs.readFile(vscode.Uri.parse(doc.backupId));
        await this._sendBytes(doc, bytes, proto.baseName(doc.uri.path));
        return;
      } catch (e) { /* the backup is gone: open the workspace file */ }
    }
    if (!doc.panel) return;
    const url = doc.panel.webview.asWebviewUri(doc.uri).toString();
    doc.panel.webview.postMessage({ type: 'open', name: proto.baseName(doc.uri.path), url, uri: doc.uri.toString() });
  }

  async _sendBytes(doc, bytes, name, extra) {
    extra = extra || {};
    const data = bytes || await vscode.workspace.fs.readFile(doc.uri);
    const CHUNK = 256 * 1024;
    const n = Math.max(1, Math.ceil(data.length / CHUNK));
    const uri = doc.uri.toString();
    for (let i = 0; i < n; i++) {
      const slice = data.subarray(i * CHUNK, Math.min(data.length, (i + 1) * CHUNK));
      doc.panel.webview.postMessage({
        type: 'open-chunk', i, n, name, b64: b64(slice), uri,
        draft: !!extra.draft, draftKey: extra.draftKey || '',
      });
    }
  }
}

function activate(context) {
  for (const app of Object.keys(VIEW)) {
    const provider = new OfficeEditorProvider(context, app);
    context.subscriptions.push(vscode.window.registerCustomEditorProvider(VIEW[app], provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: false,
    }));
    context.subscriptions.push(provider._onDidChange);
  }
  context.subscriptions.push(vscode.commands.registerCommand('vibeoffice.open', async (uri) => {
    if (!uri) {
      const picked = await vscode.window.showOpenDialog({
        canSelectMany: false,
        filters: { 'Office files': [...new Set(Object.values(proto.APPS).flat())] },
      });
      uri = picked && picked[0];
    }
    if (!uri) return;
    const view = proto.viewTypeFor(uri.path);
    if (!view) {
      vscode.window.showInformationMessage('VibeOffice opens Word, Excel and PowerPoint files (.docx, .xlsx, .pptx and the other formats the apps open).');
      return;
    }
    await vscode.commands.executeCommand('vscode.openWith', uri, view);
  }));
}

function deactivate() {}

module.exports = { activate, deactivate };
