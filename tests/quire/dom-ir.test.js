const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/quire/js/dmodel.js');
require('../../public/quire/js/ops.js');
require('../../public/quire/js/dom.js');
require('../../public/quire/js/dom-ir.js');
const { D } = L;

function fresh() {
  D.doc = D.newDoc();
  D.resetHistory();
  if (L.ed) L.ed.sel = null;
  ActiveDocument;
}

function ok(res, what) {
  assert.equal(res.ok, true, (what || 'op') + ' ' + (res.message || res.code || ''));
  return res;
}

test('catalog covers the specification tables', () => {
  const spec = fs.readFileSync(path.join(__dirname, '../../docs/quire_dom_api_specifications.md'), 'utf8');
  const sections = spec.split(/^## /m).slice(1);
  const missing = [];
  function has(type, member) {
    return !!L.domIr.member(type, member);
  }
  function checkPair(type, member) {
    if (!L.domIr.types[type]) {
      missing.push(type + ' (type)');
      return;
    }
    if (!has(type, member)) missing.push(type + '.' + member);
    return L.domIr.member(type, member);
  }
  for (const section of sections) {
    const title = section.slice(0, section.indexOf('\n')).trim();
    if (title === 'Enumerations' || title === 'Worked calls') continue;
    const lines = section.split('\n');
    let header = '';
    for (const line of lines) {
      if (line.startsWith('| Member |') || line.startsWith('| Object |')) header = line;
      if (!line.startsWith('|') || line.startsWith('| Member |') || line.startsWith('| Object |') || line.startsWith('|---') || line.startsWith('| Status |') || line.startsWith('| Code |') || line.startsWith('| Implemented |') || line.startsWith('| Not implemented |') || line.startsWith('| G-')) continue;
      const cells = line.split('|').slice(1, -1).map((s) => s.trim());
      if (!cells.length || cells[0] === '---') continue;
      if (header.startsWith('| Object |')) {
        const ticks = [...cells[0].matchAll(/`([A-Za-z][A-Za-z0-9]*)`/g)].map((m) => m[1]);
        const names = ticks.length ? ticks : cells[0].replace(/\([^)]*\)/g, '').split(',').map((s) => s.trim()).filter((s) => /^[A-Z][A-Za-z0-9]*$/.test(s));
        const code = (cells[1] || '').indexOf('G-HOST') >= 0 ? 'G-HOST' : 'G-ENGINE';
        for (const name of names) {
          const t = L.domIr.types[name];
          if (!t || t.gap !== code) missing.push(name + ' gap ' + code);
        }
        continue;
      }
      if (!header.startsWith('| Member |')) continue;
      let carry = '';
      for (const piece of cells[0].split(',')) {
        const bit = piece.trim();
        const three = bit.match(/^([A-Za-z][A-Za-z0-9]*)\.([A-Za-z][A-Za-z0-9]*)\.([A-Za-z][A-Za-z0-9]*)/);
        const two = bit.match(/^([A-Za-z][A-Za-z0-9]*)\.([A-Za-z][A-Za-z0-9]*)/);
        const one = bit.match(/^([A-Za-z][A-Za-z0-9]*)$/);
        if (three && L.domIr.types[three[1]]) {
          const mid = checkPair(three[1], three[2]);
          carry = three[1];
          if (mid && mid.ret && L.domIr.types[mid.ret] && !has(mid.ret, three[3])) missing.push(mid.ret + '.' + three[3]);
        } else if (two && L.domIr.types[two[1]]) {
          checkPair(two[1], two[2]);
          carry = two[1];
        } else if (one && carry) {
          checkPair(carry, one[1]);
        }
      }
      const sig = cells[1] || '';
      for (const m of sig.matchAll(/([A-Za-z][A-Za-z0-9]*)\.([A-Za-z][A-Za-z0-9]*)/g)) {
        if (L.domIr.types[m[1]] && !has(m[1], m[2])) missing.push(m[1] + '.' + m[2]);
      }
    }
  }
  assert.deepEqual(missing, []);
});

test('worked calls match the live object model', () => {
  fresh();
  Documents.Add();
  const ir = L.domIr.open();
  const doc = ok(ir.exec({ op: 'get', handle: ir.root.handle, member: 'ActiveDocument' }));
  const paras = ok(ir.exec({ op: 'get', handle: doc.handle, member: 'Paragraphs' }));
  const para = ok(ir.exec({ op: 'item', handle: paras.handle, args: [1] }));
  const range = ok(ir.exec({ op: 'get', handle: para.handle, member: 'Range' }));
  const before = ok(ir.exec({ op: 'get', handle: range.handle, member: 'Text' }));
  assert.equal(before.kind, 'scalar');
  assert.equal(before.value, ActiveDocument.Paragraphs(1).Range.Text);
  ok(ir.exec({ op: 'set', handle: range.handle, member: 'Text', value: 'Hello' }));
  assert.equal(ActiveDocument.Paragraphs(1).Range.Text, 'Hello\r');
  const readBack = ok(ir.exec({ op: 'get', handle: range.handle, member: 'Text' }));
  assert.equal(readBack.value, ActiveDocument.Paragraphs(1).Range.Text);

  const sel = ok(ir.exec({ op: 'get', handle: ir.root.handle, member: 'Selection' }));
  ok(ir.exec({ op: 'call', handle: sel.handle, member: 'TypeText', args: [' '] }));
  assert.equal(D.ptext(ActiveDocument.Paragraphs(1)._block), ' Hello');
  const font = ok(ir.exec({ op: 'get', handle: range.handle, member: 'Font' }));
  ok(ir.exec({ op: 'set', handle: font.handle, member: 'Bold', value: -1 }));
  assert.equal(ActiveDocument.Paragraphs(1)._block.runs[0].rPr.b, true);
  assert.equal(ok(ir.exec({ op: 'get', handle: font.handle, member: 'Bold' })).value, ActiveDocument.Paragraphs(1).Range.Font.Bold);
  const selFont = ok(ir.exec({ op: 'get', handle: sel.handle, member: 'Font' }));
  ok(ir.exec({ op: 'set', handle: selFont.handle, member: 'Bold', value: -1 }));
  assert.equal(ok(ir.exec({ op: 'get', handle: selFont.handle, member: 'Bold' })).value, Selection.Font.Bold);

  ok(ir.exec({ op: 'set', handle: para.handle, member: 'Alignment', value: wdAlignParagraphCenter }));
  assert.equal(ActiveDocument.Paragraphs(1).Alignment, wdAlignParagraphCenter);

  const tables = ok(ir.exec({ op: 'get', handle: doc.handle, member: 'Tables' }));
  ok(ir.exec({ op: 'call', handle: tables.handle, member: 'Add', args: [{ handle: range.handle }, 2, 2] }));
  const table = ok(ir.exec({ op: 'item', handle: tables.handle, args: [1] }));
  const cell = ok(ir.exec({ op: 'call', handle: table.handle, member: 'Cell', args: [1, 1] }));
  const cellRange = ok(ir.exec({ op: 'get', handle: cell.handle, member: 'Range' }));
  assert.equal(ok(ir.exec({ op: 'get', handle: cellRange.handle, member: 'Text' })).value, ActiveDocument.Tables(1).Cell(1, 1).Range.Text);

  const marks = ok(ir.exec({ op: 'get', handle: doc.handle, member: 'Bookmarks' }));
  const selRange = ok(ir.exec({ op: 'get', handle: sel.handle, member: 'Range' }));
  ok(ir.exec({ op: 'call', handle: marks.handle, member: 'Add', args: ['Here', { handle: selRange.handle }] }));
  assert.equal(ActiveDocument.Bookmarks.Exists('Here'), true);

  const find = ok(ir.exec({ op: 'get', handle: sel.handle, member: 'Find' }));
  const found = ok(ir.exec({
    op: 'call', handle: find.handle, member: 'Execute',
    args: ['Hello', -1, 0, 0, 0, 0, -1, wdFindContinue, 0, 'Hi', wdReplaceOne],
  }));
  assert.equal(found.value, true);
  assert.equal(ActiveDocument.Bookmarks.Exists('Here'), true);
  ir.release();
  assert.equal(ir.exec({ op: 'get', handle: doc.handle, member: 'Name' }).code, 'released');
});

test('a promised call is waited on before the result is wrapped', async () => {
  fresh();
  const ir = L.domIr.open();
  const docs = ir.exec({ op: 'get', handle: ir.root.handle, member: 'Documents' });
  const real = docs.ok && L.dom.Application.Documents.Add;
  L.dom.Application.Documents.Add = function () { return Promise.resolve('later'); };
  try {
    const pending = ir.exec({ op: 'call', handle: docs.handle, member: 'Add' });
    assert.equal(typeof pending.then, 'function');
    const done = await pending;
    assert.equal(done.kind, 'scalar');
    assert.equal(done.value, 'later');
  } finally {
    L.dom.Application.Documents.Add = real;
  }
});

test('gaps stay on the specification codes', () => {
  fresh();
  Documents.Add();
  const ir = L.domIr.open();
  const doc = ok(ir.exec({ op: 'get', handle: ir.root.handle, member: 'ActiveDocument' }));
  const range = ok(ir.exec({ op: 'call', handle: doc.handle, member: 'Range', args: [0, 5] }));
  assert.equal(range.type, 'Range');
  assert.equal(ir.exec({ op: 'get', handle: doc.handle, member: 'Creator' }).code, 'G-ENGINE');
  assert.equal(ir.exec({ op: 'get', handle: doc.handle, member: 'FullName' }).code, 'G-HOST');
  assert.equal(ir.exec({ op: 'get', handle: doc.handle, member: 'NotAMember' }).code, 'absent');
  const added = ok(ir.exec({ op: 'call', handle: ok(ir.exec({ op: 'get', handle: ir.root.handle, member: 'Documents' })).handle, member: 'Add', args: [L.domIr.Missing] }));
  assert.equal(added.type, 'Document');
  const controls = ok(ir.named('ContentControls'));
  assert.equal(ir.exec({ op: 'call', handle: controls.handle, member: 'Add' }).code, 'G-ENGINE');
  const windows = ok(ir.named('Windows'));
  assert.equal(ir.exec({ op: 'get', handle: windows.handle, member: 'Count' }).code, 'G-HOST');
  assert.equal(ok(ir.exec({ op: 'get', handle: windows.handle, member: 'Application' })).type, 'Application');
});

test('implemented calls publish VBA parameter names', () => {
  const bad = [];
  for (const typeName in L.domIr.types) {
    const t = L.domIr.types[typeName];
    for (const member in t.members) {
      const desc = t.members[member];
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
  assert.deepEqual(names('Tables', 'Add'), ['Range', 'NumRows', 'NumColumns', 'DefaultTableBehavior', 'AutoFitBehavior']);
  assert.equal(names('Tables', 'Add').length, 5);
  assert.equal(names('Selection', 'ConvertToTable').length, 16);
  assert.deepEqual(names('Paragraphs', 'Item'), ['Index']);
  assert.deepEqual(names('BuiltinDialogs', 'Item'), ['Index']);
  assert.deepEqual(names('Documents', 'Open'), ['FileName', 'ConfirmConversions', 'ReadOnly', 'AddToRecentFiles', 'PasswordDocument']);
  assert.equal(L.domIr.member('Selection', 'TypeParagraph').params, undefined);
  assert.deepEqual(names('Selection', 'MoveLeft'), ['Unit', 'Count', 'Extend']);
  assert.deepEqual(names('Characters', 'Item'), ['Index']);
  assert.deepEqual(names('Application', 'GetSaveAsFilename'), ['InitialFilename', 'FileFilter', 'FilterIndex', 'Title', 'ButtonText']);
  assert.equal(names('Document', 'PrintOut').length, 19);
  assert.deepEqual(names('Comments', 'AddReply'), ['Comment', 'Text']);
  assert.deepEqual(names('List', 'CountNumberedItems'), ['NumberType', 'Level']);
  assert.equal(L.domIr.member('Document', 'UndoClear').params, undefined);
  assert.ok(L.domIr.member('Tables', 'Add').params.every((p) => p.optional));
  assert.ok(L.domIr.member('Selection', 'MoveLeft').params.every((p) => p.optional));
});
