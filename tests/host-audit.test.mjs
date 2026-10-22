import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';

const context = { window: {}, console };
context.window = context;
context.globalThis = context;
vm.createContext(context);
vm.runInContext(fs.readFileSync(new URL('../public/common/host-audit.js', import.meta.url), 'utf8'), context);
const { L, VO } = context;

test('an unknown ability and a browser host are false', () => {
  VO.host = { kind: 'web' };
  assert.equal(L.hostAudit.has('Path'), false);
  assert.equal(L.hostAudit.has('NotAThing'), false);
  L.APP_ID = 'quire';
  VO.host = { kind: 'vscode', documentPath: () => '' };
  assert.equal(L.hostAudit.has('Path'), false);
  assert.equal(L.hostAudit.has('SaveAs'), true);
  assert.equal(L.hostAudit.has('InsertFile'), true);
  assert.equal(L.hostAudit.has('Export'), false);
  assert.equal(L.hostAudit.has('SheetSaveAs'), false);
  L.APP_ID = 'lectern';
  assert.equal(L.hostAudit.has('Export'), true);
  assert.equal(L.hostAudit.has('InsertFile'), false);
  L.APP_ID = 'ledger';
  VO.host.documentPath = () => '/work/Book1.xlsx';
  assert.equal(L.hostAudit.has('Path'), true);
  assert.equal(L.hostAudit.has('SheetSaveAs'), true);
  assert.equal(L.hostAudit.has('Shell'), true);
});

test('AlwaysDeny and a denied request do not become G-HOST', async () => {
  VO.host = { kind: 'vscode', documentPath: () => '/work/Book1.xlsx' };
  VO.audit = { origin: 'agent' };
  const denied = await L.hostAudit.gate({ tier: 'AlwaysDeny', ability: 'Shell' });
  assert.equal(denied.ok, false);
  assert.equal(denied.code, 'E-DENIED');

  VO.audit.decide = async () => ({ ok: false });
  L.hostAudit.beginScript();
  const approval = await L.hostAudit.gate({ tier: 'RequestApproval', ability: 'SaveAs' });
  assert.equal(approval.code, 'E-APPROVAL');

  VO.audit.origin = 'user';
  L.hostAudit.beginScript();
  const user = await L.hostAudit.gate({ tier: 'RequestApproval', ability: 'SaveAs' });
  assert.equal(user.ok, true);

  assert.throws(() => L.hostAudit.shell(), (e) => e.code === 'E-DENIED');
});

test('XPS fails before a PDF format is accepted', () => {
  L.APP_ID = 'quire';
  assert.equal(L.hostAudit.isXps(18), true);
  assert.equal(L.hostAudit.isPdf(17), true);
  assert.equal(L.hostAudit.isXps(null), false);
  L.APP_ID = 'ledger';
  assert.equal(L.hostAudit.isXps(1), true);
  assert.equal(L.hostAudit.isPdf(0), true);
  assert.equal(L.hostAudit.isPdf(null), true);
});

test('the catalog keeps the audit token off the return type', () => {
  const host = { window: context.window, L, console };
  vm.createContext(host);
  vm.runInContext(fs.readFileSync(new URL('../public/common/core.js', import.meta.url), 'utf8'), host);
  vm.runInContext('L.dom = L.dom || {};', host);
  vm.runInContext(fs.readFileSync(new URL('../public/quire/js/dom-ir.js', import.meta.url), 'utf8'), host);
  const save = host.L.domIr.member('Document', 'Save');
  assert.equal(save.audit, 'RequestApproval');
  assert.equal(save.ability, 'Save');
  assert.equal(save.kind, 'call');
  assert.equal(save.ret, 'empty');
  const name = host.L.domIr.member('Document', 'FullName');
  assert.equal(name.audit, 'AlwaysAllow');
  assert.equal(name.ability, 'Path');
  assert.equal(name.kind, 'get');
  const shell = host.L.domIr.member('Application', 'HostShell');
  assert.equal(shell.audit, 'AlwaysDeny');
  const open = host.L.domIr.member('Application', 'GetOpenFilename');
  assert.equal(open.fileDialog, true);
  assert.equal(open.audit, 'RequestApproval');
});
