/* VibeOffice as a VS Code extension. code-server runs this in its Node extension host and
 * draws the same app pages in a webview, so a file in the workspace opens in Quire, Ledger
 * or Lectern. Save writes that file; another name is written beside it and opened too.
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
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
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

  async _sendOpen(doc) {
    if (doc.backupId) {
      try {
        const bytes = await vscode.workspace.fs.readFile(vscode.Uri.parse(doc.backupId));
        await this._sendBytes(doc, bytes, proto.baseName(doc.uri.path));
        return;
      } catch (e) { /* the backup is gone: open the workspace file */ }
    }
    if (!doc.panel) return;
    const url = doc.panel.webview.asWebviewUri(doc.uri).toString();
    doc.panel.webview.postMessage({ type: 'open', name: proto.baseName(doc.uri.path), url });
  }

  async _sendBytes(doc, bytes, name) {
    const data = bytes || await vscode.workspace.fs.readFile(doc.uri);
    const CHUNK = 256 * 1024;
    const n = Math.max(1, Math.ceil(data.length / CHUNK));
    for (let i = 0; i < n; i++) {
      const slice = data.subarray(i * CHUNK, Math.min(data.length, (i + 1) * CHUNK));
      doc.panel.webview.postMessage({ type: 'open-chunk', i, n, name, b64: b64(slice) });
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
