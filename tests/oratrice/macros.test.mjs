import test from 'node:test';
import assert from 'node:assert/strict';

globalThis.window = globalThis;
await import('../../public/common/core.js');
await import('../../public/common/xml.js');
await import('../../public/common/opc.js');
await import('../../public/common/zip.js');
await import('../../public/oratrice/js/model.js');
if (!L.t) L.t = (s, ...args) => s.replace(/\{(\d+)\}/g, (m, i) => (args[i] != null ? String(args[i]) : m));
await import('../../public/oratrice/js/macros.js');

const K = L.opc;
const rel = (id, type, target) => `<Relationship Id="${id}" Type="${type}" Target="${K.esc(target)}"/>`;
const rels = (content) => `<Relationships xmlns="${K.NS.pkg}">${content}</Relationships>`;
const types = '<Types xmlns="' + K.NS.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>';
const text = (data) => (typeof data === 'string' ? data : new TextDecoder().decode(data));
const ITEM = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
const PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps';

async function pkg(parts) {
  const blob = await L.zip.write(Object.entries({ '[Content_Types].xml': types, ...parts }).map(([name, data]) => ({ name, data })));
  return K.open(new Uint8Array(await blob.arrayBuffer()));
}

test('a macro name is one Visual Basic identifier', () => {
  const ok = L.macros.vbIdentifier;
  assert.equal(ok('Main'), true);
  assert.equal(ok('你好'), true);
  assert.equal(ok('[Sub]'), true);
  assert.equal(ok('A!B'), false);
  assert.equal(ok('1abc'), false);
  assert.equal(ok(''), false);
});

test('macro source round-trips through the presentation custom XML part', async () => {
  const source = 'x < y & z\n  ';
  const xml = L.macros.serialize([{ name: 'A&B', source, created: '2026-10-21', author: 'Miyuki' }]);
  const back = L.macros.parse(xml);
  assert.equal(back[0].name, 'A&B');
  assert.equal(back[0].source, source);
  assert.equal(back[0].author, 'Miyuki');

  const props = '<ds:datastoreItem ds:itemID="{11111111-1111-4111-8111-111111111111}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="https://acane.net/melinaoffice/macros"/></ds:schemaRefs></ds:datastoreItem>';
  const other = '<store xmlns="urn:example:other">kept</store>';
  const body = '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>';
  const p = await pkg({
    '_rels/.rels': rels(rel('rId1', K.NS.rel + '/officeDocument', 'ppt/presentation.xml')),
    'ppt/presentation.xml': body,
    'ppt/_rels/presentation.xml.rels': rels(rel('rId1', ITEM, '../customXml/item1.xml') + rel('rId2', ITEM, '../customXml/item2.xml')),
    'customXml/item1.xml': xml,
    'customXml/_rels/item1.xml.rels': rels(rel('rId1', PROPS, 'itemProps1.xml')),
    'customXml/itemProps1.xml': props,
    'customXml/item2.xml': other,
    'customXml/_rels/item2.xml.rels': rels(rel('rId1', PROPS, 'itemProps2.xml')),
    'customXml/itemProps2.xml': '<ds:datastoreItem ds:itemID="{22222222-2222-4222-8222-222222222222}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>',
  });
  const pres = { macros: [], keep: {}, losses: [] };
  K.attach(pres, p);
  L.macros.adopt(pres);
  assert.equal(pres.macros[0].source, source);
  assert.equal(pres.keep.vibeMacros.item, 'customXml/item1.xml');
  assert.equal(pres.keep.vibeMacros.itemID, '{11111111-1111-4111-8111-111111111111}');
  pres.macros[0].source = 'changed <';
  const out = K.output(p, { doc: pres, consumes: (base, rel) => L.macros.ownsRel(pres, rel) });
  out.bind('ppt/presentation.xml', p.main);
  out.put('ppt/presentation.xml', body, 'application/xml');
  L.macros.finish(pres, out);
  const saved = new Map(out.finish().files.map((f) => [f.name, text(f.data)]));
  assert.equal(L.macros.parse(saved.get('customXml/item1.xml'))[0].source, 'changed <');
  assert.match(saved.get('customXml/item2.xml'), /kept/);
  assert.match(saved.get('ppt/_rels/presentation.xml.rels'), /item1\.xml/);
  assert.match(saved.get('ppt/_rels/presentation.xml.rels'), /item2\.xml/);

  pres.macros = [];
  const out2 = K.output(p, { doc: pres, consumes: (base, rel) => L.macros.ownsRel(pres, rel) });
  out2.bind('ppt/presentation.xml', p.main);
  out2.put('ppt/presentation.xml', body, 'application/xml');
  L.macros.finish(pres, out2);
  const saved2 = new Map(out2.finish().files.map((f) => [f.name, text(f.data)]));
  assert.equal(saved2.has('customXml/item1.xml'), false);
  assert.match(saved2.get('customXml/item2.xml'), /kept/);
  assert.doesNotMatch(saved2.get('ppt/_rels/presentation.xml.rels'), /item1\.xml/);
});

test('saving and deleting a macro can be undone', () => {
  L.pres = L.model.newPresentation();
  L.hist.clear();
  assert.equal(L.macros.commit('Hello', 'Public Sub Hello()\nEnd Sub\n', { created: '2026-10-21', author: 'Miyuki' }), true);
  assert.equal(L.pres.macros[0].name, 'Hello');
  assert.equal(L.hist.lastLabel, 'Save Macro');
  L.hist.doUndo();
  assert.deepEqual(L.pres.macros, []);
  L.hist.doRedo();
  assert.equal(L.pres.macros[0].source, 'Public Sub Hello()\nEnd Sub\n');
  L.macros.remove('hello');
  assert.equal(L.pres.macros.length, 0);
  assert.equal(L.hist.lastLabel, 'Delete Macro');
  L.hist.doUndo();
  assert.equal(L.pres.macros[0].name, 'Hello');
});

test('all templates and documents is stored apart from the current presentation', () => {
  const mem = new Map();
  const prev = L.pres;
  const prevStorage = globalThis.localStorage;
  globalThis.localStorage = {
    getItem: (k) => (mem.has(k) ? mem.get(k) : null),
    setItem: (k, v) => { mem.set(k, String(v)); },
  };
  try {
    L.pres = L.model.newPresentation();
    L.pres.macros = [{ name: 'Hello', source: 'presentation body', created: '2026-10-21', author: 'Doc' }];
    assert.equal(L.macros.commit('Hello', 'stored body', { store: 'all', created: '2026-10-21', author: 'Share' }), true);
    assert.equal(L.macros.commit('Only', 'only body', { store: 'all', created: '2026-10-22', author: 'Share' }), true);
    assert.equal(L.pres.macros.length, 1);
    assert.equal(L.pres.macros[0].source, 'presentation body');
    assert.deepEqual(L.macroHost.names('document').slice().sort(), ['Hello']);
    assert.deepEqual(L.macroHost.names('all').slice().sort(), ['Hello', 'Only']);
    assert.equal(L.macroHost.describe('Hello', 'all'), 'Macro created 2026-10-21 by Share');
    assert.equal(L.macroHost.describe('Hello'), 'Macro created 2026-10-21 by Doc');
    assert.equal(mem.has('melinaoffice.oratrice.allMacros'), true);
    assert.equal(mem.has('melinaoffice.ecrivaine.allMacros'), false);
    assert.equal(L.macroHost.remove('Only', 'all'), true);
    assert.deepEqual(L.macroHost.names('all'), ['Hello']);
    assert.equal(L.pres.macros[0].source, 'presentation body');
    globalThis.localStorage.setItem = () => { throw new Error('full'); };
    assert.equal(L.macros.commit('Nope', 'x', { store: 'all' }), false);
    assert.deepEqual(L.macroHost.names('all'), ['Hello']);
  } finally {
    L.pres = prev;
    if (prevStorage === undefined) delete globalThis.localStorage;
    else globalThis.localStorage = prevStorage;
  }
});
