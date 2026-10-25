/* Melina Office as a VS Code extension. code-server runs this in its Node extension host and
 * draws the same app pages in a webview, so a file in the workspace opens in Ecrivaine, Matheuse
 * or Oratrice. Save writes that file. Save As writes the chosen name into a folder the user
 * picked (the open file's folder, unless they choose another) and opens a different file too.
 * File ▸ New and File ▸ Open use VS Code's own dialogs and open the result in a new tab.
 * New File… (file/newFile) creates a blank .docx, .xlsx or .pptx and opens it here.
 * An agent in the editor runs VB6 on the active tab through the MCP server in vscode/mcp.js.
 * Docs: docs/vscode.md. */
'use strict';

const fs = require('fs');
const path = require('path');
const vscode = require('vscode');
const proto = require('./protocol');
const page = require('./page');
const docs = require('./documents');
const bridgeApi = require('./mcp-bridge');

const VIEW = {
  ecrivaine: 'melinaoffice.ecrivaine',
  matheuse: 'melinaoffice.matheuse',
  oratrice: 'melinaoffice.oratrice',
};

const SAVE_LABEL = { ecrivaine: 'Document', matheuse: 'Workbook', oratrice: 'Presentation' };

/** bytes for a Recent Files entry that should open in place of the file on disk, keyed by the file's URI */
const pendingOpen = new Map();
/** a file just created from New File…: the page builds a blank document and writes it here */
const pendingBlank = new Set();
/** webviews that can run VB, keyed by the file URI. The MCP bridge reads this. */
const openDocs = new Map();
/** localhost bridge the MCP process calls. Null until activate finishes listening. */
let mcpBridge = null;

function publicRoot(extensionPath) {
  const candidates = [path.join(extensionPath, 'public'), path.join(extensionPath, '..', 'public')];
  for (const c of candidates) {
    if (fs.existsSync(path.join(c, 'ecrivaine', 'index.html'))) return c;
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

/** the folder of the file this editor tab has open, as a URI string and the path the editor shows */
function folderInfo(doc) {
  const dir = parentUri(doc.uri);
  return { uri: dir.toString(), label: dir.fsPath };
}

const HOST_ALWAYS = 'melinaoffice.hostAlways';

function alwaysAbilities(context) {
  const list = context.workspaceState.get(HOST_ALWAYS);
  return Array.isArray(list) ? list.slice() : [];
}

function pathOutside(fsPath) {
  const folders = vscode.workspace.workspaceFolders || [];
  const norm = String(fsPath || '').replace(/\\/g, '/').replace(/\/+$/, '').toLowerCase();
  if (!norm || !folders.length) return true;
  return !folders.some((f) => {
    const root = f.uri.fsPath.replace(/\\/g, '/').replace(/\/+$/, '').toLowerCase();
    return norm === root || norm.startsWith(root + '/');
  });
}

function hostGranted(doc, ability) {
  return !!(doc.hostGrants && ability && doc.hostGrants.has(ability));
}

function pauseVbs(doc) {
  for (const pending of doc.pending.values()) {
    if (pending.expect === 'vbs-result' && pending.pause) pending.pause();
  }
}

function resumeVbs(doc) {
  for (const pending of doc.pending.values()) {
    if (pending.expect === 'vbs-result' && pending.resume) pending.resume();
  }
}

function globToReg(glob) {
  const body = String(glob || '*').replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
  return new RegExp('^' + body + '$', 'i');
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

/** a Melina Office tab already showing this file, if there is one */
function openTab(uri) {
  const key = uri.toString();
  for (const group of vscode.window.tabGroups.all) {
    for (const tab of group.tabs) {
      const input = tab.input;
      if (!input || !input.uri || input.uri.toString() !== key) continue;
      if (input.viewType && String(input.viewType).indexOf('melinaoffice.') === 0) return input.viewType;
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
    this.save = proto.createSaveSession();
    this.backupId = '';
  }
  dispose() {
    this.save.close();
    this._settlePending();
  }
  /** Wake every request the page or the host is waiting on. A save whose bytes are already
   *  on disk resolves; anything still waiting rejects once, so the flag cannot stay stuck. */
  _settlePending() {
    for (const p of [...this.pending.values()]) {
      if (p.expect === 'save-result' && this.save.wrote) p.resolve({ ok: true, inplace: true });
      else p.reject(new Error('The editor was closed.'));
    }
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
    const key = doc.uri.toString();
    openDocs.set(key, {
      uri: key,
      name: proto.baseName(doc.uri.path),
      app: this._app,
      doc,
      ask: (msg) => this._request(doc, 'vbs-result', msg, docs.VBS_TIMEOUT_MS),
    });
    const root = this._public ? vscode.Uri.file(this._public) : null;
    const roots = [];
    if (root) roots.push(root);
    roots.push(parentUri(doc.uri));
    panel.webview.options = { enableScripts: true, localResourceRoots: roots };
    panel.webview.onDidReceiveMessage((msg) => this._onMessage(doc, msg));
    panel.onDidDispose(() => {
      if (doc.panel === panel) doc.panel = null;
      const cur = openDocs.get(doc.uri.toString());
      if (cur && cur.doc === doc) openDocs.delete(doc.uri.toString());
      doc._settlePending();
    });
    panel.webview.html = this._html(doc);
  }

  async saveCustomDocument(doc) {
    const started = doc.save.begin();
    if (started.action === 'skip') return;
    if (started.action === 'join') return started.promise;
    let again = false;
    try {
      const res = await this._request(doc, 'save-result', { type: 'do-save' });
      if (!res || !res.ok) again = doc.save.fail(new Error((res && res.error) || 'Save was cancelled.'));
      else again = doc.save.succeed(res);
    } catch (e) {
      again = doc.save.fail(e);
    }
    /* a keystroke during the save stays an edit after this promise resolves, which is when
       VS Code drops the edits it had when the save started */
    if (again) this._fireDirty(doc);
    return started.promise;
  }

  async saveCustomDocumentAs(doc, destination) {
    const bytes = await this._backupBytes(doc);
    await vscode.workspace.fs.writeFile(destination, bytes);
    doc.save.dirty = false;
  }

  async revertCustomDocument(doc) {
    doc.save.dirty = false;
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
    if (!this._public) return page.missingPage('Melina Office could not find its pages next to the extension.');
    const file = path.join(this._public, this._app, 'index.html');
    let html;
    try { html = fs.readFileSync(file, 'utf8'); }
    catch (e) { return page.missingPage('Melina Office could not read ' + this._app + '. ' + (e && e.message ? e.message : '')); }
    const base = doc.panel.webview.asWebviewUri(vscode.Uri.file(path.join(this._public, this._app))).toString();
    return page.renderPage({
      html,
      baseHref: base,
      cspSource: doc.panel.webview.cspSource,
      fileName: proto.baseName(doc.uri.path),
      platform: platform(),
      lang: this._context.globalState.get('melinaoffice.lang') || '',
    });
  }

  _request(doc, expect, msg, timeoutMs) {
    if (!doc.panel || !doc.panel.webview) return Promise.reject(new Error('The editor was closed.'));
    const id = Math.random().toString(36).slice(2);
    return new Promise((resolve, reject) => {
      const budget = timeoutMs || 300000;
      let remaining = budget;
      let started = Date.now();
      let timer = null;
      let paused = false;
      const fire = () => {
        doc.pending.delete(id);
        reject(new Error('The editor did not answer.'));
      };
      const arm = () => { timer = setTimeout(fire, remaining); };
      const finish = (fn, value) => {
        if (!doc.pending.has(id)) return;
        if (timer) clearTimeout(timer);
        doc.pending.delete(id);
        fn(value);
      };
      doc.pending.set(id, {
        expect,
        pause() {
          if (paused || !timer) return;
          clearTimeout(timer);
          timer = null;
          remaining -= Date.now() - started;
          if (remaining < 1) remaining = 1;
          paused = true;
        },
        resume() {
          if (!paused) return;
          paused = false;
          started = Date.now();
          arm();
        },
        resolve: (v) => finish(resolve, v),
        reject: (e) => finish(reject, e),
      });
      arm();
      try { doc.panel.webview.postMessage(Object.assign({ id }, msg)); }
      catch (e) { finish(reject, e); }
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

  _fireDirty(doc) {
    this._onDidChange.fire({
      document: doc,
      undo() { doc.save.dirty = false; },
      redo() { doc.save.dirty = true; },
    });
  }

  async _onMessage(doc, msg) {
    if (!msg || typeof msg !== 'object') return;
    if (msg.type === 'ready') { await this._sendOpen(doc); return; }
    if (msg.type === 'read-fallback') { await this._sendBytes(doc, null, proto.baseName(doc.uri.path)); return; }
    if (msg.type === 'dirty') {
      /* dirty:false does not save. The tab clears only when saveCustomDocument resolves. */
      if (msg.dirty && doc.save.onDirty(true) === 'mark') this._fireDirty(doc);
      return;
    }
    if (msg.type === 'request-save') { await this._onRequestSave(doc, msg); return; }
    if (msg.type === 'close') {
      /* the page already asked about unsaved changes; clear the tab's dot without writing them again */
      try {
        if (doc.save.dirty) {
          doc.save.armSkip();
          try { await vscode.workspace.save(doc.uri); }
          finally { doc.save.disarmSkip(); }
        }
      } catch (e) { /* close the tab either way */ }
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
      return;
    }
    if (msg.type === 'pick-new') { await this._pickNew(doc, msg); return; }
    if (msg.type === 'pick-open') { await this._pickOpen(doc, msg); return; }
    if (msg.type === 'pick-folder') { await this._pickFolder(doc, msg); return; }
    if (msg.type === 'document-folder') {
      const folder = folderInfo(doc);
      this._reply(doc, msg.id, { ok: true, uri: folder.uri, label: folder.label });
      return;
    }
    if (msg.type === 'open-uri') { await this._openUri(doc, msg); return; }
    if (msg.type === 'lang') {
      const id = String(msg.lang || '');
      await this._context.globalState.update('melinaoffice.lang', id);
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
    if (msg.type === 'host-approval') { await this._approve(doc, msg); return; }
    if (msg.type === 'host-outside') {
      this._reply(doc, msg.id, { ok: true, outside: pathOutside(msg.path) });
      return;
    }
    if (msg.type === 'host-read') { await this._hostRead(doc, msg); return; }
    if (msg.type === 'host-dir') { await this._hostDir(doc, msg); return; }
    if (msg.type === 'host-open') { await this._hostOpen(doc, msg); return; }
    if (msg.type === 'host-pick') { await this._hostPick(doc, msg); return; }
    if (msg.type === 'vbs-result') {
      const p = doc.pending.get(msg.id);
      if (p) p.resolve(msg);
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
      if (done.purpose === 'host') {
        try {
          await this._writeHost(doc, done);
          if (doc.panel) doc.panel.webview.postMessage({ type: 'wrote', id: done.id, ok: true, inplace: false, name: proto.baseName(done.path) });
        } catch (e) {
          const error = e && e.message ? e.message : String(e);
          if (doc.panel) doc.panel.webview.postMessage({ type: 'wrote', id: done.id, ok: false, error, code: e && e.code });
        }
        return;
      }
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
        const home = parentUri(doc.uri);
        const chosen = proto.saveFolder(home.toString(), done.folder, doc.allowedFolders);
        if (!chosen) throw new Error('That folder was not chosen.');
        const dir = chosen === home.toString() ? home : vscode.Uri.parse(chosen);
        const st = await vscode.workspace.fs.stat(dir);
        if (!(st.type & vscode.FileType.Directory)) throw new Error('That location is not a folder.');
        const inplace = dir.toString() === home.toString() && proto.sameFile(done.name, doc.uri.path);
        const target = inplace ? doc.uri : await placeBeside(dir, done.name, doc.uri);
        await vscode.workspace.fs.writeFile(target, done.bytes);
        pendingBlank.delete(target.toString());
        const wroteInPlace = target.toString() === doc.uri.toString();
        if (wroteInPlace) doc.save.noteWritten();
        if (!wroteInPlace) await vscode.commands.executeCommand('vscode.openWith', target, doc.viewType);
        try {
          if (doc.panel) doc.panel.webview.postMessage({
            type: 'wrote', id: done.id, ok: true, inplace: wroteInPlace, name: proto.baseName(target.path),
          });
        } catch (err) { /* the bytes are written; a webview that already closed does not undo that */ }
      } catch (e) {
        const error = e && e.message ? e.message : String(e);
        if (doc.panel) doc.panel.webview.postMessage({ type: 'wrote', id: done.id, ok: false, error });
        vscode.window.showErrorMessage('Melina Office could not save the file. ' + error);
      }
    }
  }

  _reply(doc, id, body) {
    if (doc.panel) doc.panel.webview.postMessage(Object.assign({ type: 'reply', id }, body));
  }

  async _whileDialog(doc, fn) {
    pauseVbs(doc);
    try { return await fn(); }
    finally { resumeVbs(doc); }
  }

  _remember(ability) {
    const list = alwaysAbilities(this._context);
    if (list.includes(ability)) return Promise.resolve();
    list.push(ability);
    return this._context.workspaceState.update(HOST_ALWAYS, list);
  }

  _allow(doc, ability) {
    if (!doc.hostGrants) doc.hostGrants = new Set();
    doc.hostGrants.add(ability);
  }

  async _approve(doc, msg) {
    const ability = String(msg.ability || '');
    const tier = msg.tier || 'RequestApproval';
    if (tier === 'AlwaysDeny') {
      this._reply(doc, msg.id, { ok: false, code: 'E-DENIED', error: 'The editor refused this operation.' });
      return;
    }
    const outside = msg.path ? pathOutside(msg.path) : false;
    const ask = (tier === 'RequestApproval' || (tier === 'AlwaysAllow' && outside)) && doc.vbsAgent;
    if (!ask || alwaysAbilities(this._context).includes(ability)) {
      this._allow(doc, ability);
      this._reply(doc, msg.id, { ok: true });
      return;
    }
    await this._whileDialog(doc, async () => {
      let choice;
      let timer;
      try {
        choice = await Promise.race([
          vscode.window.showWarningMessage(
            'A script wants to use ' + ability + (msg.detail ? ' (' + msg.detail + ')' : '') + '.',
            { modal: true },
            'Allow',
            'Always allow',
            'Deny',
          ),
          new Promise((resolve) => { timer = setTimeout(() => resolve(undefined), 120000); }),
        ]);
      } finally { if (timer) clearTimeout(timer); }
      if (choice === 'Always allow') await this._remember(ability);
      if (choice === 'Allow' || choice === 'Always allow') {
        this._allow(doc, ability);
        this._reply(doc, msg.id, { ok: true });
      } else {
        this._reply(doc, msg.id, { ok: false, code: 'E-APPROVAL', error: 'The editor did not allow this operation.' });
      }
    });
  }

  _needsGrant(doc, ability, filePath) {
    if (!doc.vbsAgent || !pathOutside(filePath)) return false;
    if (alwaysAbilities(this._context).includes(ability)) return false;
    return !hostGranted(doc, ability);
  }

  async _hostRead(doc, msg) {
    try {
      const filePath = String(msg.path || '');
      if (this._needsGrant(doc, msg.ability || '', filePath)) {
        const e = new Error('The editor did not allow this operation.');
        e.code = 'E-APPROVAL';
        throw e;
      }
      const data = await vscode.workspace.fs.readFile(vscode.Uri.file(filePath));
      if (data.length > 32 * 1024 * 1024) throw new Error('The file is too large.');
      this._reply(doc, msg.id, { ok: true, b64: Buffer.from(data).toString('base64') });
    } catch (e) {
      this._reply(doc, msg.id, { ok: false, code: e && e.code, error: e && e.message ? e.message : String(e) });
    }
  }

  async _hostDir(doc, msg) {
    try {
      const pattern = String(msg.pattern || '');
      const slash = Math.max(pattern.lastIndexOf('/'), pattern.lastIndexOf('\\'));
      const dir = slash >= 0 ? pattern.slice(0, slash) : pattern;
      const glob = slash >= 0 ? pattern.slice(slash + 1) : '*';
      if (this._needsGrant(doc, 'Dir', dir)) {
        const e = new Error('The editor did not allow this operation.');
        e.code = 'E-APPROVAL';
        throw e;
      }
      const entries = await vscode.workspace.fs.readDirectory(vscode.Uri.file(dir));
      const re = globToReg(glob);
      const names = entries.filter((e) => (e[1] & vscode.FileType.File) && re.test(e[0])).map((e) => e[0]);
      this._reply(doc, msg.id, { ok: true, names });
    } catch (e) {
      this._reply(doc, msg.id, { ok: false, code: e && e.code, error: e && e.message ? e.message : String(e) });
    }
  }

  async _hostOpen(doc, msg) {
    try {
      const filePath = String(msg.path || '');
      if (this._needsGrant(doc, 'Open', filePath)) {
        const e = new Error('The editor did not allow this operation.');
        e.code = 'E-APPROVAL';
        throw e;
      }
      const uri = vscode.Uri.file(filePath);
      const view = proto.viewTypeFor(filePath);
      if (!view) throw new Error('That file type is not available.');
      await vscode.commands.executeCommand('vscode.openWith', uri, view);
      this._reply(doc, msg.id, { ok: true });
    } catch (e) {
      this._reply(doc, msg.id, { ok: false, code: e && e.code, error: e && e.message ? e.message : String(e) });
    }
  }

  async _hostPick(doc, msg) {
    await this._whileDialog(doc, async () => {
      try {
        const kind = msg.kind || 'open';
        let picked;
        if (kind === 'save') {
          picked = await vscode.window.showSaveDialog({});
          this._reply(doc, msg.id, { ok: true, paths: picked ? [picked.fsPath] : [] });
          return;
        }
        picked = await vscode.window.showOpenDialog({
          canSelectFiles: kind !== 'folder',
          canSelectFolders: kind === 'folder',
          canSelectMany: !!msg.multi,
          openLabel: kind === 'folder' ? 'Select folder' : 'Open',
        });
        this._reply(doc, msg.id, { ok: true, paths: picked ? picked.map((u) => u.fsPath) : [] });
      } catch (e) {
        this._reply(doc, msg.id, { ok: false, error: e && e.message ? e.message : String(e) });
      }
    });
  }

  async _writeHost(doc, done) {
    const filePath = String(done.path || '');
    if (!filePath) throw new Error('The file has no path.');
    if (this._needsGrant(doc, done.ability || '', filePath)) {
      const e = new Error('The editor did not allow this operation.');
      e.code = 'E-APPROVAL';
      throw e;
    }
    const uri = vscode.Uri.file(filePath);
    await vscode.workspace.fs.writeFile(uri, done.bytes);
    if (done.open) {
      const view = proto.viewTypeFor(filePath);
      if (view) await vscode.commands.executeCommand('vscode.openWith', uri, view);
    }
  }

  /** The page's Save, Ctrl+S and a save before close. One write, through saveCustomDocument,
   *  so the tab dot and the bytes stay the same save. A save already running is that write. */
  async _onRequestSave(doc, msg) {
    const reply = (ok, error) => this._reply(doc, msg.id, { ok: !!ok, error: error || '' });
    try {
      if (doc.vbsAgent && !hostGranted(doc, 'Save') && !alwaysAbilities(this._context).includes('Save')) {
        reply(false, 'The editor did not allow this operation.');
        return;
      }
      const how = doc.save.onRequest();
      if (how === 'closed') { reply(false, 'The editor was closed.'); return; }
      if (how === 'join') {
        await doc.save.chain;
        reply(true);
        return;
      }
      if (!doc.save.dirty) {
        doc.save.dirty = true;
        this._fireDirty(doc);
      }
      const before = doc.save.runs;
      const saved = await vscode.workspace.save(doc.uri);
      if (!saved || (doc.save.runs === before && !doc.save.chain)) {
        reply(false, 'Save was cancelled.');
        return;
      }
      reply(true);
    } catch (e) {
      /* wrote-then-closed resolves saveCustomDocument, so it does not arrive here */
      reply(false, e && e.message ? e.message : String(e));
    }
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
      vscode.window.showErrorMessage('Melina Office could not create the file. ' + error);
    }
  }

  /** File ▸ Save As ▸ Browse: VS Code's folder dialog. Only a folder the user picks may be written to. */
  async _pickFolder(doc, msg) {
    try {
      const home = parentUri(doc.uri);
      const asked = proto.saveFolder(home.toString(), msg.uri, doc.allowedFolders);
      const start = asked && asked !== home.toString() ? vscode.Uri.parse(asked) : home;
      const picked = await vscode.window.showOpenDialog({
        canSelectFiles: false,
        canSelectFolders: true,
        canSelectMany: false,
        defaultUri: start,
        openLabel: 'Select Folder',
        title: 'Select Folder',
      });
      if (!picked || !picked.length) { this._reply(doc, msg.id, { ok: true, cancelled: true }); return; }
      const uri = picked[0];
      const st = await vscode.workspace.fs.stat(uri);
      if (!(st.type & vscode.FileType.Directory)) {
        this._reply(doc, msg.id, { ok: false, error: 'That location is not a folder.' });
        return;
      }
      if (!doc.allowedFolders) doc.allowedFolders = new Set();
      doc.allowedFolders.add(uri.toString());
      this._reply(doc, msg.id, { ok: true, uri: uri.toString(), label: uri.fsPath });
    } catch (e) {
      const error = e && e.message ? e.message : String(e);
      this._reply(doc, msg.id, { ok: false, error });
      vscode.window.showErrorMessage('Melina Office could not choose a folder. ' + error);
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
        canSelectMany: this._app === 'matheuse',
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
      vscode.window.showErrorMessage('Melina Office could not open the file. ' + error);
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
      if (done.purpose === 'create') vscode.window.showErrorMessage('Melina Office could not create the file. ' + error);
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

  /** a file this command just created, or a zero-byte .docx / .xlsx / .pptx: the page writes a blank document */
  async _freshBlank(doc) {
    if (pendingBlank.has(doc.uri.toString())) return true;
    if (proto.extname(doc.uri.path) !== (proto.SAVE_EXT[this._app] || '')) return false;
    try {
      const st = await vscode.workspace.fs.stat(doc.uri);
      return st.size === 0;
    } catch (e) { return false; }
  }

  async _sendOpen(doc) {
    const pending = pendingOpen.get(doc.uri.toString());
    if (pending) {
      pendingOpen.delete(doc.uri.toString());
      await this._sendBytes(doc, pending.bytes, proto.baseName(doc.uri.path), { draft: true, draftKey: pending.draftKey });
      return;
    }
    if (await this._freshBlank(doc)) {
      /* keep the mark until the page's write lands, so a reload before that still starts blank.
         A zero-byte file (New File… with a typed name) is the same case: there is nothing to open. */
      if (doc.panel) {
        const folder = folderInfo(doc);
        doc.panel.webview.postMessage({ type: 'blank', name: proto.baseName(doc.uri.path), uri: doc.uri.toString(), dir: folder.uri, dirLabel: folder.label, fsPath: doc.uri.fsPath });
      }
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
    const folder = folderInfo(doc);
    doc.panel.webview.postMessage({ type: 'open', name: proto.baseName(doc.uri.path), url, uri: doc.uri.toString(), dir: folder.uri, dirLabel: folder.label, fsPath: doc.uri.fsPath });
  }

  async _sendBytes(doc, bytes, name, extra) {
    extra = extra || {};
    const data = bytes || await vscode.workspace.fs.readFile(doc.uri);
    const CHUNK = 256 * 1024;
    const n = Math.max(1, Math.ceil(data.length / CHUNK));
    const uri = doc.uri.toString();
    const folder = folderInfo(doc);
    for (let i = 0; i < n; i++) {
      const slice = data.subarray(i * CHUNK, Math.min(data.length, (i + 1) * CHUNK));
      doc.panel.webview.postMessage({
        type: 'open-chunk', i, n, name, b64: b64(slice), uri,
        dir: folder.uri, dirLabel: folder.label, fsPath: doc.uri.fsPath,
        draft: !!extra.draft, draftKey: extra.draftKey || '',
      });
    }
  }
}

async function fileExists(uri) {
  try { await vscode.workspace.fs.stat(uri); return true; }
  catch { return false; }
}

/** the next free Document1.docx / Book1.xlsx / Presentation1.pptx in a folder */
async function nextFree(dir, stem, ext) {
  for (let n = 1; n < 100; n++) {
    const uri = vscode.Uri.joinPath(dir, stem + n + '.' + ext);
    if (!(await fileExists(uri))) return uri;
  }
  return vscode.Uri.joinPath(dir, stem + Date.now() + '.' + ext);
}

/** New File…: an empty file of this app's type, then the editor. The page fills it with a blank document. */
async function createOfficeFile(context, app) {
  const ext = proto.SAVE_EXT[app];
  const stem = proto.SAVE_STEM[app];
  const view = VIEW[app];
  if (!ext || !view) return;
  if (!publicRoot(context.extensionPath)) {
    vscode.window.showErrorMessage('Melina Office could not find its pages next to the extension.');
    return;
  }
  let uri;
  const folders = vscode.workspace.workspaceFolders;
  if (folders && folders.length) {
    uri = await nextFree(folders[0].uri, stem, ext);
  } else {
    const filters = {};
    filters[SAVE_LABEL[app] || 'Document'] = [ext];
    const base = process.env.HOME || process.env.USERPROFILE || process.cwd();
    const picked = await vscode.window.showSaveDialog({
      defaultUri: vscode.Uri.joinPath(vscode.Uri.file(base), stem + '1.' + ext),
      filters,
      saveLabel: 'Create',
    });
    if (!picked) return;
    uri = vscode.Uri.joinPath(parentUri(picked), proto.ensureExt(proto.baseName(picked.path), ext, stem));
  }
  if (await reveal(uri, view)) return;
  const existed = await fileExists(uri);
  try {
    await vscode.workspace.fs.writeFile(uri, new Uint8Array());
    pendingBlank.add(uri.toString());
    await vscode.commands.executeCommand('vscode.openWith', uri, view);
  } catch (e) {
    pendingBlank.delete(uri.toString());
    if (!existed) { try { await vscode.workspace.fs.delete(uri); } catch (err) { /* the empty file stays if it cannot be removed */ } }
    const error = e && e.message ? e.message : String(e);
    vscode.window.showErrorMessage('Melina Office could not create the file. ' + error);
  }
}

/** the active editor group's tabs, as the snapshot selectDocument expects */
function editorTabs() {
  const group = vscode.window.tabGroups.activeTabGroup;
  if (!group) return [];
  const active = group.activeTab;
  return group.tabs.map((tab) => {
    const input = tab && tab.input;
    return {
      active: tab === active,
      viewType: input && input.viewType ? String(input.viewType) : '',
      uri: input && input.uri ? input.uri.toString() : '',
    };
  });
}

function openList() {
  return [...openDocs.values()].map((d) => ({ uri: d.uri, name: d.name, app: d.app }));
}

/** one script at a time on a document, so Debug.Print from two calls cannot mix */
function enqueue(doc, fn) {
  const prev = doc.vbsChain || Promise.resolve();
  const run = prev.then(fn, fn);
  doc.vbsChain = run.then(() => {}, () => {});
  return run;
}

async function onMcpCall(body) {
  return docs.dispatch(body, editorTabs(), openList(), async (picked, call) => {
    const entry = openDocs.get(picked.uri);
    if (!entry || !entry.doc || !entry.doc.panel) {
      return { ok: false, app: picked.app, fileName: picked.name, prints: [], value: null, error: { message: docs.NOT_READY } };
    }
    entry.doc.vbsAgent = true;
    if (!entry.doc.hostGrants) entry.doc.hostGrants = new Set();
    entry.doc.hostGrants.clear();
    try {
      return await enqueue(entry.doc, () => entry.ask({
        type: 'run-vbs',
        agent: true,
        source: call.source,
        mode: call.mode,
        entry: call.entry,
        args: call.args,
      }));
    } catch (e) {
      const timedOut = e && e.message === 'The editor did not answer.';
      return {
        ok: false,
        app: picked.app,
        fileName: picked.name,
        prints: [],
        value: null,
        error: { message: timedOut ? 'The script did not finish in time.' : (e && e.message ? e.message : String(e)) },
      };
    } finally {
      entry.doc.vbsAgent = false;
    }
  });
}

function mcpProcessEnv() {
  return {
    ELECTRON_RUN_AS_NODE: '1',
    MELINAOFFICE_BRIDGE_URL: mcpBridge.url,
    MELINAOFFICE_BRIDGE_TOKEN: mcpBridge.token,
  };
}

function stdioServerDefinition(label, command, args, env) {
  const Ctor = vscode.McpStdioServerDefinition;
  if (Ctor.length <= 1) return new Ctor({ label, command, args, env, version: '1.0.3' });
  return new Ctor(label, command, args, env, '1.0.3');
}

function registerMcp(context) {
  const script = path.join(context.extensionPath, 'vscode', 'mcp.js');
  const env = mcpProcessEnv();
  const lm = vscode.lm;
  if (lm && typeof lm.registerMcpServerDefinitionProvider === 'function' && vscode.McpStdioServerDefinition) {
    try {
      context.subscriptions.push(lm.registerMcpServerDefinitionProvider('melinaoffice', {
        provideMcpServerDefinitions() {
          return [stdioServerDefinition('Melina Office', process.execPath, [script], env)];
        },
        resolveMcpServerDefinition(server) { return server; },
      }));
    } catch (e) { /* the host advertised the API and then refused the provider */ }
  }
  const cursorMcp = vscode.cursor && vscode.cursor.mcp;
  if (cursorMcp && typeof cursorMcp.registerServer === 'function') {
    try {
      cursorMcp.registerServer({
        name: 'melinaoffice',
        server: { command: process.execPath, args: [script], env },
      });
      context.subscriptions.push({
        dispose() {
          if (typeof cursorMcp.unregisterServer === 'function') cursorMcp.unregisterServer('melinaoffice');
        },
      });
    } catch (e) { /* this host is not Cursor */ }
  }
}

async function activate(context) {
  for (const app of Object.keys(VIEW)) {
    const provider = new OfficeEditorProvider(context, app);
    context.subscriptions.push(vscode.window.registerCustomEditorProvider(VIEW[app], provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: false,
    }));
    context.subscriptions.push(provider._onDidChange);
  }
  context.subscriptions.push(vscode.commands.registerCommand('melinaoffice.open', async (uri) => {
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
      vscode.window.showInformationMessage('Melina Office opens Word, Excel and PowerPoint files (.docx, .xlsx, .pptx and the other formats the apps open).');
      return;
    }
    await vscode.commands.executeCommand('vscode.openWith', uri, view);
  }));
  for (const [command, app] of [['melinaoffice.newDocx', 'ecrivaine'], ['melinaoffice.newXlsx', 'matheuse'], ['melinaoffice.newPptx', 'oratrice']]) {
    context.subscriptions.push(vscode.commands.registerCommand(command, () => createOfficeFile(context, app)));
  }
  try {
    mcpBridge = await bridgeApi.createBridge(onMcpCall);
    context.subscriptions.push({ dispose: () => { if (mcpBridge) mcpBridge.close(); } });
    registerMcp(context);
  } catch (e) { /* the editors still open if this machine will not listen on localhost */ }
}

function deactivate() {}

module.exports = { activate, deactivate };
