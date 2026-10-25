const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/common/geometry.js');
require('../../public/oratrice/js/model.js');
require('../../public/oratrice/js/textedit.js');
require('../../public/oratrice/js/dom.js');
require('../../public/oratrice/js/dom-ir.js');

function fresh() {
  L.pres = L.model.newPresentation();
  L.dom.reset();
}

test('the presentation IR reads and writes the open deck', () => {
  fresh();
  const s = L.domIr.open();
  const pres = s.exec({ op: 'get', handle: s.root.handle, member: 'ActivePresentation' });
  assert.equal(pres.ok, true);
  assert.equal(pres.type, 'Presentation');
  const name = s.exec({ op: 'get', handle: pres.handle, member: 'Name' });
  assert.equal(name.value, ActivePresentation.Name);
  assert.equal(s.exec({ op: 'get', handle: pres.handle, member: 'FullName' }).code, 'G-HOST');
  assert.equal(s.exec({ op: 'get', handle: pres.handle, member: 'Creator' }).code, 'G-ENGINE');
  assert.equal(s.exec({ op: 'set', handle: pres.handle, member: 'Name', value: 'Z' }).code, 'readonly');

  const slides = s.exec({ op: 'get', handle: pres.handle, member: 'Slides' });
  const slide = s.exec({ op: 'item', handle: slides.handle, args: [1] });
  const shapes = s.exec({ op: 'get', handle: slide.handle, member: 'Shapes' });
  const title = s.exec({ op: 'get', handle: shapes.handle, member: 'Title' });
  const frame = s.exec({ op: 'get', handle: title.handle, member: 'TextFrame' });
  const range = s.exec({ op: 'get', handle: frame.handle, member: 'TextRange' });
  assert.equal(s.exec({ op: 'set', handle: range.handle, member: 'Text', value: 'Hello' }).ok, true);
  assert.equal(s.exec({ op: 'get', handle: range.handle, member: 'Text' }).value, 'Hello');
  const chars = s.exec({ op: 'get', handle: range.handle, member: 'Characters' });
  assert.equal(s.exec({ op: 'get', handle: chars.handle, member: 'Count' }).value, 5);
  assert.equal(s.exec({ op: 'item', handle: chars.handle, args: [1] }).code, 'G-TEXT');
  assert.equal(s.exec({ op: 'get', handle: slide.handle, member: 'TimeLine' }).code, 'G-ENGINE');

  const named = s.named('TimeLine');
  assert.equal(named.ok, true);
  assert.equal(s.exec({ op: 'get', handle: named.handle, member: 'MainSequence' }).code, 'G-ENGINE');
  assert.equal(s.named('NotAThing').code, 'absent');
  s.release();
  assert.equal(s.exec({ op: 'get', handle: pres.handle, member: 'Name' }).code, 'released');
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
  assert.deepEqual(names('Slides', 'Add'), ['Index', 'Layout']);
  assert.deepEqual(names('Slides', 'Item'), ['Index']);
  assert.deepEqual(names('Shapes', 'AddPicture'), ['FileName', 'LinkToFile', 'SaveWithDocument', 'Left', 'Top', 'Width', 'Height']);
  assert.deepEqual(names('Presentations', 'Open'), ['FileName', 'ReadOnly', 'Untitled', 'WithWindow']);
  assert.deepEqual(names('Characters', 'Item'), ['Start', 'Length']);
  assert.deepEqual(names('TextRange', 'Find'), ['FindWhat', 'After', 'MatchCase', 'WholeWords']);
  assert.equal(names('Presentation', 'PrintOut').length, 5);
  assert.deepEqual(names('Application', 'GetSaveAsFilename'), ['InitialFilename', 'FileFilter', 'FilterIndex', 'Title', 'ButtonText']);
  assert.equal(L.domIr.member('Presentation', 'NewWindow').params, undefined);
  assert.ok(L.domIr.member('Slides', 'Add').params.every((p) => p.optional));
});
