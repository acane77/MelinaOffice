/* Saving a new workbook that has a macro. It has no original package, so the part is written, not claimed. */
const { test } = require('node:test');
const assert = require('node:assert/strict');

test('a macro on a new workbook survives two saves', async () => {
  global.window = global;
  const L = require('./load.js');
  L.layout = { invalidate() {}, touchMerges() {} };
  if (!L.t) L.t = (s, ...args) => s.replace(/\{(\d+)\}/g, (_, i) => (args[i] != null ? String(args[i]) : ''));
  require('../../public/ledger/js/macros.js');
  const wb = new L.model.Workbook();
  wb.addSheet('Sheet1');
  wb.macros = [{ name: 'Fill', source: 'Public Sub Fill()\nEnd Sub\n', created: '2026-10-22', author: 'Ledger' }];
  const once = await L.xlsxWrite.bytes(wb);
  const twice = await L.xlsxWrite.bytes(wb);
  const pkg = await L.opc.open(twice);
  const rel = pkg.rels(pkg.main).find((r) => /\/customXml$/.test(r.type));
  assert.ok(rel, 'workbook relationship');
  const back = L.macros.parse(pkg.text(rel.part));
  assert.equal(back[0].name, 'Fill');
  assert.equal(back[0].source, 'Public Sub Fill()\nEnd Sub\n');
  assert.ok(once.length > 0);
});
