/* The editor's MCP server: which tab a script may touch, the guide an agent reads,
 * and the stdio tools. The extension host itself needs the vscode module, so this
 * stays on the pieces that run without it. */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const docs = require(path.join(root, 'vscode/documents.js'));
const guide = require(path.join(root, 'vscode/vbs-guide.js'));
const mcp = require(path.join(root, 'vscode/mcp.js'));
const bridgeApi = require(path.join(root, 'vscode/mcp-bridge.js'));

const quire = { uri: 'file:///work/Report.docx', name: 'Report.docx', app: 'quire' };
const ledger = { uri: 'file:///work/Budget.xlsx', name: 'Budget.xlsx', app: 'ledger' };

test('only the active VibeOffice tab can be scripted', () => {
  const none = docs.selectDocument([], []);
  assert.equal(none.ok, false);
  assert.equal(none.error, docs.NO_DOCUMENT);

  const other = docs.selectDocument(
    [{ active: true, viewType: 'default', uri: 'file:///work/notes.txt' }],
    [quire, ledger],
  );
  assert.equal(other.ok, false);
  assert.match(other.error, /Report\.docx/);
  assert.match(other.error, /Budget\.xlsx/);

  const hit = docs.selectDocument(
    [{ active: true, viewType: 'vibeoffice.ledger', uri: ledger.uri }],
    [quire, ledger],
  );
  assert.equal(hit.ok, true);
  assert.equal(hit.app, 'ledger');
  assert.equal(hit.name, 'Budget.xlsx');

  const early = docs.selectDocument(
    [{ active: true, viewType: 'vibeoffice.quire', uri: quire.uri }],
    [ledger],
  );
  assert.equal(early.ok, false);
  assert.equal(early.error, docs.NOT_READY);
});

test('a VB object is not returned, and a block of cells is', () => {
  assert.equal(docs.plainValue({ handle: 1 }), undefined);
  assert.deepEqual(docs.plainValue([[1, 'a'], [false, null]]), [[1, 'a'], [false, null]]);
  const dropped = docs.normalizeResult({ ok: true, value: { handle: 4 }, prints: ['12'] }, quire);
  assert.equal(dropped.ok, true);
  assert.equal(dropped.value, null);
  assert.match(dropped.note, /Debug\.Print/);
  assert.deepEqual(dropped.prints, ['12']);
});

test('both tools refuse a window with no VibeOffice document', async () => {
  for (const tool of ['vibeoffice_dom_guide', 'vibeoffice_run_vbs']) {
    const result = await docs.dispatch({ tool, source: 'Public Sub Main\nEnd Sub\n' }, [], [], async () => {
      throw new Error('should not run');
    });
    assert.equal(result.ok, false);
    assert.equal(result.error.message, docs.NO_DOCUMENT);
  }
});

test('the active tab decides the guide, and a gap comes back as an error', async () => {
  const tabs = [{ active: true, viewType: 'vibeoffice.quire', uri: quire.uri }];
  const shown = await docs.dispatch({ tool: 'vibeoffice_dom_guide' }, tabs, [quire], async () => {
    throw new Error('a guide does not run VB');
  });
  assert.equal(shown.ok, true);
  assert.equal(shown.app, 'quire');
  assert.match(shown.guide, /G-STORY/);
  assert.match(shown.guide, /MsgBox/);

  const wrong = await docs.dispatch(
    { tool: 'vibeoffice_run_vbs', source: 'Public Sub Main\nEnd Sub\n' },
    [{ active: true, viewType: 'default', uri: 'file:///work/notes.txt' }],
    [quire],
    async () => { throw new Error('should not run'); },
  );
  assert.equal(wrong.ok, false);
  assert.match(wrong.error.message, /Report\.docx/);

  const gap = await docs.dispatch(
    { tool: 'vibeoffice_run_vbs', source: 'Public Sub Main\nEnd Sub\n' },
    tabs,
    [quire],
    async () => ({
      ok: false,
      app: 'quire',
      fileName: 'Report.docx',
      prints: [],
      value: null,
      error: { message: 'Range', number: 438, code: 'G-STORY' },
    }),
  );
  assert.equal(gap.ok, false);
  assert.equal(gap.error.number, 438);
  assert.equal(gap.error.code, 'G-STORY');
});

test('the guide for each app names MsgBox and the audit codes', () => {
  for (const app of ['quire', 'ledger', 'lectern']) {
    const text = guide.forApp(app);
    assert.match(text, /MsgBox/);
    assert.match(text, /E-APPROVAL/);
    assert.match(text, /E-DENIED/);
    assert.match(text, /vbaProject\.bin is kept and is not run|vbaProject/);
  }
  assert.match(guide.forApp('quire'), /G-STORY/);
  assert.match(guide.forApp('lectern'), /G-TEXT/);
  assert.match(guide.forApp('ledger'), /no G-STORY and no G-TEXT/);
  assert.match(guide.INSTRUCTIONS, /vibeoffice_dom_guide/);
  assert.equal(guide.forApp('slides'), '');
});

test('package.json names the MCP provider and keeps the old engine', () => {
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  assert.equal(pkg.engines.vscode, '^1.75.0');
  const providers = pkg.contributes.mcpServerDefinitionProviders;
  assert.equal(providers.length, 1);
  assert.equal(providers[0].id, 'vibeoffice');
});

test('stdio initialize, tools/list, and tools/call use the bridge', async () => {
  let calls = 0;
  const bridge = await bridgeApi.createBridge(async (body) => {
    calls += 1;
    if (body.tool === 'vibeoffice_dom_guide') {
      return { ok: true, app: 'lectern', fileName: 'Deck.pptx', guide: 'G-TEXT and no MsgBox' };
    }
    return {
      ok: false,
      app: 'lectern',
      fileName: 'Deck.pptx',
      prints: [],
      value: null,
      error: { message: 'Characters', number: 438, code: 'G-TEXT' },
    };
  });
  const child = spawn(process.execPath, [path.join(root, 'vscode/mcp.js')], {
    env: Object.assign({}, process.env, {
      ELECTRON_RUN_AS_NODE: '1',
      VIBEOFFICE_BRIDGE_URL: bridge.url,
      VIBEOFFICE_BRIDGE_TOKEN: bridge.token,
    }),
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  let buf = '';
  let err = '';
  const inbox = [];
  const waiters = [];
  const push = (msg) => {
    const waiter = waiters.shift();
    if (waiter) waiter(msg);
    else inbox.push(msg);
  };
  child.stdout.on('data', (chunk) => {
    buf += chunk.toString('utf8');
    let nl;
    while ((nl = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, nl);
      buf = buf.slice(nl + 1);
      if (line.trim()) push(JSON.parse(line));
    }
  });
  child.stderr.on('data', (chunk) => { err += chunk.toString('utf8'); });
  const next = () => {
    if (inbox.length) return Promise.resolve(inbox.shift());
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('no MCP reply' + (err ? ': ' + err : ''))), 5000);
      waiters.push((msg) => { clearTimeout(timer); resolve(msg); });
    });
  };
  const send = async (method, params, id) => {
    const msg = { jsonrpc: '2.0', method };
    if (id != null) msg.id = id;
    if (params) msg.params = params;
    child.stdin.write(JSON.stringify(msg) + '\n');
    return id == null ? null : next();
  };
  try {
    const init = await send('initialize', { protocolVersion: '2025-06-18', capabilities: {} }, 1);
    assert.equal(init.result.protocolVersion, '2025-06-18');
    assert.match(init.result.instructions, /vibeoffice_dom_guide/);
    assert.equal(init.result.serverInfo.name, 'vibeoffice');
    await send('notifications/initialized');

    const listed = await send('tools/list', {}, 2);
    const names = listed.result.tools.map((tool) => tool.name).sort();
    assert.deepEqual(names, ['vibeoffice_dom_guide', 'vibeoffice_run_vbs']);
    const guideTool = listed.result.tools.find((tool) => tool.name === 'vibeoffice_dom_guide');
    assert.equal(guideTool.annotations.readOnlyHint, true);
    const runTool = listed.result.tools.find((tool) => tool.name === 'vibeoffice_run_vbs');
    assert.notEqual(runTool.annotations.readOnlyHint, true);

    const guideReply = await send('tools/call', { name: 'vibeoffice_dom_guide', arguments: {} }, 3);
    const guideBody = JSON.parse(guideReply.result.content[0].text);
    assert.equal(guideReply.result.isError, false);
    assert.equal(guideBody.app, 'lectern');
    assert.match(guideBody.guide, /G-TEXT/);

    const gapReply = await send('tools/call', {
      name: 'vibeoffice_run_vbs',
      arguments: { source: 'Public Sub Main\nEnd Sub\n' },
    }, 4);
    assert.equal(gapReply.result.isError, true);
    const gapBody = JSON.parse(gapReply.result.content[0].text);
    assert.equal(gapBody.error.code, 'G-TEXT');
    assert.equal(gapBody.error.number, 438);

    const huge = await send('tools/call', {
      name: 'vibeoffice_run_vbs',
      arguments: { source: 'x'.repeat(docs.SOURCE_MAX + 1) },
    }, 5);
    assert.equal(huge.result.isError, true);
    assert.match(JSON.parse(huge.result.content[0].text).error.message, /64 KiB/);
    assert.equal(calls, 2);

    const denied = await fetch(bridge.url, { method: 'POST', body: '{}' });
    assert.equal(denied.status, 403);
  } finally {
    child.kill();
    await bridge.close();
  }
});
