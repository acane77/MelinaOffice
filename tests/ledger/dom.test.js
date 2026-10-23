/* Ledger DOM without the VB engine. */
const assert = require('node:assert/strict');
const test = require('node:test');

function boot() {
  global.window = global;
  require('../../public/common/core.js');
  global.L.N = (s) => s;
  require('../../public/common/numfmt.js');
  require('../../public/ledger/js/formula.js');
  require('../../public/ledger/js/model.js');
  require('../../public/ledger/js/calc.js');
  require('../../public/ledger/js/fn-core.js');
  require('../../public/ledger/js/fn-lookup.js');
  require('../../public/ledger/js/layout.js');
  require('../../public/ledger/js/ops.js');
  require('../../public/ledger/js/xmlss.js');
  require('../../public/ledger/js/fninfo.js');
  require('../../public/ledger/js/dom.js');
  require('../../public/ledger/js/dom-ir.js');
}

boot();

function fresh() {
  L.dom.reset();
  return L.dom.Application;
}

test('a cell value, a formula and a block array', () => {
  const app = fresh();
  const sh = app.ActiveSheet;
  assert.equal(sh.Name, 'Sheet1');
  sh.Range('A1').Value = 'Item';
  sh.Range('B2').Value = 12.5;
  sh.Range('B3').Value = 7.5;
  sh.Range('B4').Formula = '=SUM(B2:B3)';
  assert.equal(sh.Range('A1').Value, 'Item');
  assert.equal(sh.Range('B2').Value, 12.5);
  assert.equal(sh.Range('B4').Value, 20);
  const block = sh.Range('B2:B3').Value;
  assert.deepEqual(block, [[12.5], [7.5]]);
  sh.Range('A1:B1').Font.Bold = -1;
  assert.equal(sh.Range('A1').Font.Bold, -1);
  sh.Range('B2:B4').NumberFormat = '0.00';
  assert.equal(sh.Range('B2').NumberFormat, '0.00');
  app.ActiveWorkbook.Names.Add('Total', '=Sheet1!$B$4');
  assert.equal(app.ActiveWorkbook.Names.Item('Total').RefersTo, '=Sheet1!$B$4');
  assert.equal(app.WorksheetFunction.Sum(sh.Range('B2:B3')), 20);
});

test('gaps and a name that is not in the inventory', () => {
  const app = fresh();
  assert.throws(() => app.ActiveWorkbook.FullName, (e) => e.code === 'G-HOST');
  assert.throws(() => app.ActiveWorkbook.Creator, (e) => e.code === 'G-ENGINE');
  assert.throws(() => app.Run('Macro1'), (e) => e.code === 'G-HOST');
  const s = L.domIr.open();
  const book = s.exec({ op: 'get', handle: s.root.handle, member: 'ActiveWorkbook' });
  const full = s.exec({ op: 'get', handle: book.handle, member: 'FullName' });
  assert.equal(full.ok, false);
  assert.equal(full.code, 'G-HOST');
  const absent = s.exec({ op: 'get', handle: book.handle, member: 'NotAMember' });
  assert.equal(absent.code, 'absent');
  const sh = s.exec({ op: 'get', handle: s.root.handle, member: 'ActiveSheet' });
  const range = s.exec({ op: 'call', handle: sh.handle, member: 'Range', args: ['A1'] });
  s.exec({ op: 'set', handle: range.handle, member: 'Value', value: 'Hi' });
  const got = s.exec({ op: 'get', handle: range.handle, member: 'Value' });
  assert.deepEqual(got, { ok: true, kind: 'scalar', value: 'Hi' });
  const block = s.exec({ op: 'call', handle: sh.handle, member: 'Range', args: ['A1:B1'] });
  s.exec({ op: 'set', handle: block.handle, member: 'Value', value: [['Hi', 2]] });
  const arr = s.exec({ op: 'get', handle: block.handle, member: 'Value' });
  assert.equal(arr.kind, 'array');
  assert.deepEqual(arr.value, [['Hi', 2]]);
  s.release();
  assert.equal(s.exec({ op: 'get', handle: s.root.handle, member: 'ActiveSheet' }).code, 'released');
});

test('implemented calls publish VBA parameter names', () => {
  const bad = [];
  for (const typeName in L.domIr.types) {
    const t = L.domIr.types[typeName];
    for (const member in t.members) {
      const desc = t.members[member];
      if (desc.kind === 'gap') {
        if (desc.params) bad.push(typeName + '.' + member + ' gap');
        continue;
      }
      if (!desc.params) continue;
      const seen = new Set();
      for (const p of desc.params) {
        if (!/^[A-Za-z][A-Za-z0-9]*$/.test(p.name) || /^Arg\d+$/.test(p.name)) bad.push(typeName + '.' + member + ' ' + p.name);
        const key = p.name.toLowerCase();
        if (seen.has(key)) bad.push(typeName + '.' + member + ' duplicate ' + p.name);
        seen.add(key);
      }
    }
  }
  assert.deepEqual(bad, []);
  function names(type, member) {
    return L.domIr.member(type, member).params.map((p) => p.name);
  }
  assert.deepEqual(names('Range', 'Item'), ['RowIndex', 'ColumnIndex']);
  assert.deepEqual(names('Worksheet', 'Range'), ['Cell1', 'Cell2']);
  assert.deepEqual(names('Areas', 'Item'), ['Index']);
  assert.deepEqual(names('Borders', 'Item'), ['Index']);
  assert.deepEqual(names('Workbooks', 'Open'), ['Filename', 'UpdateLinks', 'ReadOnly', 'Format', 'Password']);
  assert.equal(names('Shapes', 'AddPicture').length, 7);
  assert.ok(L.domIr.member('Worksheet', 'Range').params.every((p) => p.optional));
});
