/* Shared start for the Matheuse VB tests. Not a test file. */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

async function boot() {
  global.window = global;
  require('../../public/common/host-audit.js');
  require('../../public/common/core.js');
  L.N = (s) => s;
  require('../../public/common/numfmt.js');
  require('../../public/matheuse/js/formula.js');
  require('../../public/matheuse/js/model.js');
  require('../../public/matheuse/js/calc.js');
  require('../../public/matheuse/js/fn-core.js');
  require('../../public/matheuse/js/fn-lookup.js');
  require('../../public/matheuse/js/layout.js');
  require('../../public/matheuse/js/ops.js');
  require('../../public/matheuse/js/xmlss.js');
  require('../../public/matheuse/js/fninfo.js');
  require('../../public/matheuse/js/dom.js');
  require('../../public/matheuse/js/dom-ir.js');
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  const out = buildEngine();
  vm.runInThisContext(fs.readFileSync(out, 'utf8'), { filename: out });
  require('../../public/common/vb-forms.js');
  require('../../public/matheuse/js/vb-host.js');
  return out;
}

function fresh() {
  L.dom.reset();
}

async function expectGap(host, statement, code) {
  await assert.rejects(() => host.eval(statement), (err) => {
    assert.equal(err && err.code, code, statement + ': ' + (err && err.message));
    return true;
  });
}

module.exports = { boot, fresh, expectGap, fs, path };
