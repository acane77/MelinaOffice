/* Shared start for the Ledger VB tests. Not a test file. */
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
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  const out = buildEngine();
  vm.runInThisContext(fs.readFileSync(out, 'utf8'), { filename: out });
  require('../../public/ledger/js/vb-host.js');
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
