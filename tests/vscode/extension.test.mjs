/* The VS Code extension's page builder and save-message assembly. The extension host itself
 * needs the vscode module, so this checks the pieces that run without it, and that package.json
 * offers every file type the protocol knows. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const proto = require(path.join(root, 'vscode/protocol.js'));
const page = require(path.join(root, 'vscode/page.js'));

test('each Office file maps to one editor', () => {
  assert.equal(proto.appFor('Budget.xlsx'), 'ledger');
  assert.equal(proto.viewTypeFor('Notes.DOCX'), 'vibeoffice.quire');
  assert.equal(proto.viewTypeFor('Deck.ppsx'), 'vibeoffice.lectern');
  assert.equal(proto.appFor('readme.txt'), 'quire');
  assert.equal(proto.appFor('photo.png'), null);
});

test('a save name cannot leave the file’s folder', () => {
  assert.equal(proto.safeName('../secret.xlsx'), 'secret.xlsx');
  assert.equal(proto.safeName('..'), '');
  assert.equal(proto.sameFile('Report.DOCX', '/work/Report.docx'), true);
  assert.equal(proto.sameFile('Report.pdf', '/work/Report.docx'), false);
});

test('a new file keeps a typed suffix and appends the editor default', () => {
  assert.equal(proto.SAVE_EXT.quire, 'docx');
  assert.equal(proto.SAVE_EXT.ledger, 'xlsx');
  assert.equal(proto.SAVE_EXT.lectern, 'pptx');
  assert.equal(proto.ensureExt('新建文档.docx', 'docx'), '新建文档.docx');
  assert.equal(proto.ensureExt('新建文档.DOCX', 'docx'), '新建文档.DOCX');
  assert.equal(proto.ensureExt('新建文档.txt', 'docx'), '新建文档.txt.docx');
  assert.equal(proto.ensureExt('../x.txt', 'xlsx'), 'x.txt.xlsx');
  assert.equal(proto.ensureExt('', 'pptx', 'Presentation'), 'Presentation.pptx');
});

test('chunks reassemble in any order', () => {
  const a = proto.createAssembler();
  const raw = new Uint8Array([1, 2, 3, 4, 5]);
  assert.equal(a.push({ type: 'chunk', id: 's', i: 1, n: 2, name: 'A.docx', purpose: 'save', b64: Buffer.from(raw.subarray(3)).toString('base64') }), null);
  const done = a.push({ type: 'chunk', id: 's', i: 0, n: 2, name: 'A.docx', purpose: 'save', uri: 'file:///work/A.docx', draftKey: 'quire:A.docx', b64: Buffer.from(raw.subarray(0, 3)).toString('base64') });
  assert.deepEqual([...done.bytes], [1, 2, 3, 4, 5]);
  assert.equal(done.name, 'A.docx');
  assert.equal(done.purpose, 'save');
  assert.equal(done.uri, 'file:///work/A.docx');
  assert.equal(done.folder, '');
  assert.equal(done.draftKey, 'quire:A.docx');
});

test('Save As may write only the open file’s folder or one the user picked', () => {
  const home = 'file:///work';
  assert.equal(proto.saveFolder(home, '', null), home);
  assert.equal(proto.saveFolder(home, home, null), home);
  assert.equal(proto.saveFolder(home, 'file:///other', null), '');
  assert.equal(proto.saveFolder(home, 'file:///other', new Set(['file:///other'])), 'file:///other');
  const a = proto.createAssembler();
  const done = a.push({ type: 'chunk', id: 'f', i: 0, n: 1, name: 'B.docx', purpose: 'save', folder: 'file:///other', b64: Buffer.from('x').toString('base64') });
  assert.equal(done.folder, 'file:///other');
});

test('the webview page points at the app folder and names the vscode host', () => {
  const html = '<!doctype html><html><head><meta charset="utf-8">\n<script>look</script>\n<link rel="manifest" href="../manifest.webmanifest">\n<link rel="stylesheet" href="../common/looks.css">\n</head><body></body></html>';
  const out = page.renderPage({
    html,
    baseHref: 'https://webview.example/quire',
    cspSource: 'https://webview.example',
    fileName: '</script><script>alert(1)',
    platform: 'linux',
  });
  const baseAt = out.indexOf('<base href="https://webview.example/quire/">');
  const cssAt = out.indexOf('href="../common/looks.css"');
  assert.ok(baseAt > 0 && baseAt < cssAt);
  assert.equal(out.includes('rel="manifest"'), false);
  assert.match(out, /Content-Security-Policy/);
  assert.match(out, /acquireVsCodeApi\(\)/);
  assert.match(out, /"kind":"vscode"/);
  assert.equal(out.includes('</script><script>alert(1)'), false);
  assert.match(out, /\\u003c\/script>/);
});

test('package.json offers every extension the protocol lists', () => {
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  assert.equal(pkg.main, './vscode/extension.js');
  const offered = new Map();
  for (const ed of pkg.contributes.customEditors) {
    const app = ed.viewType.split('.')[1];
    for (const sel of ed.selector) {
      const m = /^\*\.([a-z0-9]+)$/.exec(sel.filenamePattern);
      assert.ok(m, sel.filenamePattern);
      if (!offered.has(app)) offered.set(app, new Set());
      offered.get(app).add(m[1]);
    }
  }
  for (const [app, exts] of Object.entries(proto.APPS)) {
    for (const e of exts) assert.ok(offered.get(app).has(e), app + ' missing ' + e);
    for (const e of proto.DEFAULT_EXT[app]) {
      const entry = pkg.contributes.customEditors.find((ed) => ed.viewType === 'vibeoffice.' + app && ed.priority === 'default');
      assert.ok(entry.selector.some((s) => s.filenamePattern === '*.' + e));
    }
  }
  assert.ok(fs.existsSync(path.join(root, 'public/quire/index.html')));
  assert.ok(fs.existsSync(path.join(root, 'public/ledger/index.html')));
  assert.ok(fs.existsSync(path.join(root, 'public/lectern/index.html')));
});

test('New File offers a blank document, workbook and presentation', () => {
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  const cmds = new Map(pkg.contributes.commands.map((c) => [c.command, c]));
  const menu = pkg.contributes.menus['file/newFile'];
  for (const [id, title] of [
    ['vibeoffice.newDocx', 'Word Document (.docx)'],
    ['vibeoffice.newXlsx', 'Excel Workbook (.xlsx)'],
    ['vibeoffice.newPptx', 'PowerPoint Presentation (.pptx)'],
  ]) {
    assert.equal(cmds.get(id).title, title);
    assert.equal(cmds.get(id).category, 'VibeOffice');
    assert.ok(pkg.activationEvents.includes('onCommand:' + id));
    assert.ok(menu.some((item) => item.command === id && item.group === 'file'));
  }
  const suite = fs.readFileSync(path.join(root, 'public/common/suite.js'), 'utf8');
  assert.match(suite, /type === 'blank'/);
  assert.match(suite, /voBlank/);
  assert.match(suite, /adoptBlank/);
  const ext = fs.readFileSync(path.join(root, 'vscode/extension.js'), 'utf8');
  assert.match(ext, /pendingBlank/);
  assert.match(ext, /_freshBlank/);
  assert.match(ext, /st\.size === 0/);
  assert.match(ext, /vibeoffice\.newDocx/);
  assert.match(ext, /nextFree/);
  for (const app of ['quire', 'ledger', 'lectern']) {
    assert.match(fs.readFileSync(path.join(root, 'public', app, 'js/app.js'), 'utf8'), /adoptBlank\(name\)/);
  }
});

test('the pages recognise the vscode host', () => {
  const suite = fs.readFileSync(path.join(root, 'public/common/suite.js'), 'utf8');
  const core = fs.readFileSync(path.join(root, 'public/common/core.js'), 'utf8');
  assert.match(suite, /kind === 'vscode'/);
  assert.match(suite, /selectall/);
  assert.match(suite, /stopPropagation\(\)/);
  assert.match(suite, /writeFile\(name, blob, opts\)/);
  assert.match(suite, /l\.vscode/);
  assert.match(suite, /createDocument/);
  assert.match(suite, /pick-new/);
  assert.match(suite, /pick-open/);
  assert.match(suite, /pickFolder/);
  assert.match(suite, /document-folder/);
  assert.match(core, /VO\.host\.writeFile/);
  const ext = fs.readFileSync(path.join(root, 'vscode/extension.js'), 'utf8');
  const ui = fs.readFileSync(path.join(root, 'public/common/ui.js'), 'utf8');
  assert.match(ext, /ensureExt/);
  assert.match(ext, /showSaveDialog/);
  assert.match(ext, /showOpenDialog/);
  assert.match(ext, /pick-folder/);
  assert.match(ext, /canSelectFolders: true/);
  assert.match(ui, /saveLocation/);
  assert.match(ui, /kind !== 'vscode'/);
  for (const file of ['public/quire/js/dialogs3.js', 'public/lectern/js/dialogs.js', 'public/ledger/js/dialogs.js']) {
    assert.match(fs.readFileSync(path.join(root, file), 'utf8'), /saveLocation/);
  }
});
