import test from 'node:test';
import assert from 'node:assert/strict';

globalThis.window = globalThis;
await import('../public/common/core.js');
await import('../public/common/xml.js');
await import('../public/common/opc.js');
await import('../public/common/zip.js');
await import('../public/quire/js/dmodel.js');
if (!L.t) L.t = (s, ...args) => s.replace(/\{(\d+)\}/g, (m, i) => (args[i] != null ? String(args[i]) : m));
await import('../public/common/macros.js');
await import('../public/quire/js/macros.js');

const { D } = L;
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
  assert.equal(ok('A你好1'), true);
  assert.equal(ok('_A1'), true);
  assert.equal(ok('Name$'), true);
  assert.equal(ok('A!'), true);
  assert.equal(ok('é'), true);
  assert.equal(ok('１'), true);
  assert.equal(ok('[Sub]'), true);
  assert.equal(ok('[a b]'), true);
  assert.equal(ok('🙂'), true);
  assert.equal(ok(''), false);
  assert.equal(ok('1abc'), false);
  assert.equal(ok('A B'), false);
  assert.equal(ok('A-B'), false);
  assert.equal(ok('A.B'), false);
  assert.equal(ok('A!B'), false);
  assert.equal(ok('Name$%'), false);
  assert.equal(ok('[]'), false);
  assert.equal(ok('[a]b'), false);
  assert.equal(ok('['), false);
  assert.equal(ok(' Main'), false);
  assert.equal(ok('Main '), false);
  assert.equal(ok('$'), false);
});

test('macro buttons follow the typed name', () => {
  const names = ['dsdsdc'];
  assert.deepEqual(L.macroButtonState('', names, true), { run: false, step: false, edit: false, create: false, del: false, organizer: false, cancel: true, primary: null });
  assert.deepEqual(L.macroButtonState('dsdsdc', names, true).primary, 'run');
  assert.equal(L.macroButtonState('DSDSDC', names, true).run, true);
  assert.equal(L.macroButtonState('DSDSDC', names, true).create, false);
  const fresh = L.macroButtonState('112', names, true);
  assert.equal(fresh.create, true);
  assert.equal(fresh.run, false);
  assert.equal(fresh.del, false);
  assert.equal(fresh.primary, 'create');
  assert.equal(fresh.step, false);
  assert.equal(fresh.organizer, false);
  const off = L.macroButtonState('dsdsdc', names, false);
  assert.equal(off.run, false);
  assert.equal(off.create, false);
  assert.equal(off.cancel, true);
  assert.equal(off.primary, null);
});

test('macro source round-trips through the custom XML part', async () => {
  const source = 'x < y & z\n  ';
  const xml = L.macros.serialize([{ name: 'A&B', source, created: '2026-10-21', author: 'Miyuki' }]);
  const back = L.macros.parse(xml);
  assert.equal(back[0].name, 'A&B');
  assert.equal(back[0].source, source);
  assert.equal(back[0].author, 'Miyuki');
  assert.equal(L.macros.parse(L.macros.serialize([{ name: 'S', source: '   ', created: '', author: '' }]))[0].source, '   ');

  const props = '<ds:datastoreItem ds:itemID="{11111111-1111-4111-8111-111111111111}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="urn:vibeoffice:macros"/></ds:schemaRefs></ds:datastoreItem>';
  const other = '<store xmlns="urn:example:other">kept</store>';
  const body = `<w:document xmlns:w="${K.NS.w}"><w:body><w:p/></w:body></w:document>`;
  const p = await pkg({
    '_rels/.rels': rels(rel('rId1', K.NS.rel + '/officeDocument', 'word/document.xml')),
    'word/document.xml': body,
    'word/_rels/document.xml.rels': rels(rel('rId1', ITEM, '../customXml/item1.xml') + rel('rId2', ITEM, '../customXml/item2.xml')),
    'customXml/item1.xml': xml,
    'customXml/_rels/item1.xml.rels': rels(rel('rId1', PROPS, 'itemProps1.xml')),
    'customXml/itemProps1.xml': props,
    'customXml/item2.xml': other,
    'customXml/_rels/item2.xml.rels': rels(rel('rId1', PROPS, 'itemProps2.xml')),
    'customXml/itemProps2.xml': '<ds:datastoreItem ds:itemID="{22222222-2222-4222-8222-222222222222}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>',
  });
  const doc = { macros: [], keep: {}, losses: [] };
  K.attach(doc, p);
  L.macros.adopt(doc);
  assert.equal(doc.macros[0].source, source);
  assert.equal(doc.keep.vibeMacros.item, 'customXml/item1.xml');
  assert.equal(doc.keep.vibeMacros.itemID, '{11111111-1111-4111-8111-111111111111}');
  doc.macros[0].source = 'changed <';
  const out = K.output(p, { doc, consumes: (base, rel) => L.macros.ownsRel(doc, rel) });
  out.bind('word/document.xml', p.main);
  out.put('word/document.xml', body, 'application/xml');
  L.macros.finish(doc, out);
  const saved = new Map(out.finish().files.map((f) => [f.name, text(f.data)]));
  assert.equal(L.macros.parse(saved.get('customXml/item1.xml'))[0].source, 'changed <');
  assert.match(saved.get('customXml/item2.xml'), /kept/);
  assert.match(saved.get('word/_rels/document.xml.rels'), /item1\.xml/);
  assert.match(saved.get('word/_rels/document.xml.rels'), /item2\.xml/);

  doc.macros = [];
  const out2 = K.output(p, { doc, consumes: (base, rel) => L.macros.ownsRel(doc, rel) });
  out2.bind('word/document.xml', p.main);
  out2.put('word/document.xml', body, 'application/xml');
  L.macros.finish(doc, out2);
  const saved2 = new Map(out2.finish().files.map((f) => [f.name, text(f.data)]));
  assert.equal(saved2.has('customXml/item1.xml'), false);
  assert.match(saved2.get('customXml/item2.xml'), /kept/);
});

test('saving and deleting a macro can be undone', () => {
  D.doc = D.newDoc();
  D.resetHistory();
  assert.equal(L.macros.commit('Hello', 'Public Sub Hello()\nEnd Sub\n', { created: '2026-10-21', author: 'Miyuki' }), true);
  assert.equal(D.doc.macros[0].name, 'Hello');
  assert.equal(D.undoLabel(), 'Save Macro');
  D.undo();
  assert.deepEqual(D.doc.macros, []);
  D.redo();
  assert.equal(D.doc.macros[0].source, 'Public Sub Hello()\nEnd Sub\n');
  L.macros.remove('hello');
  assert.equal(D.doc.macros.length, 0);
  assert.equal(D.undoLabel(), 'Delete Macro');
  D.undo();
  assert.equal(D.doc.macros[0].name, 'Hello');
});

test('all templates and documents is stored apart from the current document', () => {
  const mem = new Map();
  const prevDoc = D.doc;
  const prevStorage = globalThis.localStorage;
  globalThis.localStorage = {
    getItem: (k) => (mem.has(k) ? mem.get(k) : null),
    setItem: (k, v) => { mem.set(k, String(v)); },
  };
  try {
    D.doc = D.newDoc();
    D.doc.macros = [{ name: 'Hello', source: 'document body', created: '2026-10-21', author: 'Doc' }];
    assert.equal(L.macros.commit('Hello', 'stored body', { store: 'all', created: '2026-10-21', author: 'Share' }), true);
    assert.equal(L.macros.commit('Only', 'only body', { store: 'all', created: '2026-10-22', author: 'Share' }), true);
    assert.equal(D.doc.macros.length, 1);
    assert.equal(D.doc.macros[0].source, 'document body');
    assert.deepEqual(L.macroHost.names('document').slice().sort(), ['Hello']);
    assert.deepEqual(L.macroHost.names('all').slice().sort(), ['Hello', 'Only']);
    assert.equal(L.macroHost.describe('Hello', 'all'), 'Macro created 2026-10-21 by Share');
    assert.equal(L.macroHost.describe('Hello'), 'Macro created 2026-10-21 by Doc');
    assert.equal(L.macroHost.remove('Only', 'all'), true);
    assert.deepEqual(L.macroHost.names('all'), ['Hello']);
    assert.equal(D.doc.macros[0].source, 'document body');
    globalThis.localStorage.setItem = () => { throw new Error('full'); };
    assert.equal(L.macros.commit('Nope', 'x', { store: 'all' }), false);
    assert.deepEqual(L.macroHost.names('all'), ['Hello']);
  } finally {
    D.doc = prevDoc;
    if (prevStorage === undefined) delete globalThis.localStorage;
    else globalThis.localStorage = prevStorage;
  }
});
