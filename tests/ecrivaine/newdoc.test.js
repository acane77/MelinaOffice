const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/ecrivaine/js/dmodel.js');

test('a new document is laid out as Microsoft Word 2013 and later', () => {
  assert.equal(L.D.newDoc().settings.compat, 15);
  assert.equal(L.D.newDoc({ modern: true }).settings.compat, 15);
});
