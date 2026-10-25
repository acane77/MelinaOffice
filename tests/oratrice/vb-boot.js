/* Shared start for the Oratrice VB tests. Not a test file. */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

async function boot() {
  globalThis.window = globalThis;
  require('../../public/common/host-audit.js');
  require('../../public/common/core.js');
  require('../../public/common/geometry.js');
  require('../../public/oratrice/js/model.js');
  require('../../public/oratrice/js/textedit.js');
  require('../../public/oratrice/js/dom.js');
  require('../../public/oratrice/js/dom-ir.js');
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  const out = buildEngine();
  vm.runInThisContext(fs.readFileSync(out, 'utf8'), { filename: out });
  require('../../public/common/vb-forms.js');
  require('../../public/oratrice/js/vb-host.js');
  return out;
}

function fresh() {
  L.pres = L.model.newPresentation();
  L.dom.reset();
}

async function expectGap(host, statement, code) {
  await assert.rejects(() => host.eval(statement), (err) => {
    assert.equal(err && err.code, code, statement + ': ' + (err && err.message));
    return true;
  });
}

module.exports = { boot, fresh, expectGap, fs, path };
