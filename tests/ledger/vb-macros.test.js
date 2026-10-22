/* Current Document macros and All templates and documents. */
const test = require('node:test');
const assert = require('node:assert/strict');
const { boot } = require('./vb-boot.js');

const bookSrc = 'Public Sub Fill()\n  ActiveSheet.Range("A1").Value = "from book"\nEnd Sub\n';
const allSrc = 'Public Sub Fill()\n  ActiveSheet.Range("A1").Value = "from all"\nEnd Sub\n';

test('workbook macros and the browser-wide list stay separate', { timeout: 120000 }, async () => {
  await boot();
  if (!L.t) L.t = (s, ...args) => s.replace(/\{(\d+)\}/g, (_, i) => (args[i] != null ? String(args[i]) : ''));
  require('../../public/common/xml.js');
  require('../../public/ledger/js/macros.js');
  const wb = new L.model.Workbook();
  wb.addSheet('Sheet1');
  L.app = { books: [{ wb: wb, name: 'Book1' }], cur: 0 };
  wb.extra.vba = new Uint8Array([1]);
  assert.equal(L.macros.commit('Fill', bookSrc), true);
  await L.macros.run('Fill');
  assert.equal(L.dom.Application.ActiveSheet.Range('A1').Value, 'from book');
  assert.deepEqual(L.macroHost.names('document'), ['Fill']);
  assert.deepEqual(L.macroHost.names('all'), []);

  const bag = new Map();
  global.localStorage = {
    getItem(k) { return bag.has(k) ? bag.get(k) : null; },
    setItem(k, v) { bag.set(k, String(v)); },
    removeItem(k) { bag.delete(k); },
  };
  assert.equal(L.macros.commit('Fill', allSrc, { store: 'all' }), true);
  await L.macros.run('Fill', 'all');
  assert.equal(L.dom.Application.ActiveSheet.Range('A1').Value, 'from all');
  await L.macros.run('Fill');
  assert.equal(L.dom.Application.ActiveSheet.Range('A1').Value, 'from book');
  assert.deepEqual(L.macroHost.names('all'), ['Fill']);
  assert.equal(L.macroHost.describe('Fill', 'all').indexOf('Macro created') >= 0, true);

  const xml = L.macros.serialize([{ name: 'A&B', source: 'x < y & z\n', created: '2026-10-22', author: 'Ledger' }]);
  const back = L.macros.parse(xml);
  assert.equal(back[0].name, 'A&B');
  assert.equal(back[0].source, 'x < y & z\n');
  try { localStorage.removeItem('vibeoffice.ledger.allMacros'); } catch (e) { /* the list is already gone */ }
});
