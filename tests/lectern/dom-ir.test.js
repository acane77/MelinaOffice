const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/common/geometry.js');
require('../../public/lectern/js/model.js');
require('../../public/lectern/js/textedit.js');
require('../../public/lectern/js/dom.js');
require('../../public/lectern/js/dom-ir.js');

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
