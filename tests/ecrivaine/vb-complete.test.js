const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

async function boot() {
  global.window = global;
  require('../../public/common/core.js');
  require('../../public/ecrivaine/js/dmodel.js');
  require('../../public/ecrivaine/js/ops.js');
  require('../../public/ecrivaine/js/dom.js');
  require('../../public/ecrivaine/js/dom-ir.js');
  const { buildEngine } = await import('../../tools/vb6/build-engine.mjs');
  const out = buildEngine();
  vm.runInThisContext(fs.readFileSync(out, 'utf8'), { filename: out });
  require('../../public/common/vb-complete.js');
  L.N = (s) => s;
  require('../../public/common/vb-dom-ref.js');
  require('../../public/ecrivaine/js/vb-dom-ref.js');
}

function names(result) {
  return result.items.map((item) => item.name);
}

function has(result, name) {
  return names(result).some((item) => item.toLowerCase() === name.toLowerCase());
}

const host = [
  'Public Property Get ActiveDocument() As Object',
  'End Property',
  'Public Sub Shown()',
  'End Sub',
  'Private Sub HostOnly()',
  'End Sub',
].join('\n');

const code = [
  'Private Sub LocalOnly()',
  '  Dim hiddenLocal As Long',
  'End Sub',
  'Public Sub Main(n As Long)',
  '  Dim count As Long',
  '  Lef',
  'End Sub',
  'Dim shared As Long',
  'cou',
].join('\n');

function complete(at, text) {
  const source = text == null ? code : text;
  return L.vbComplete({
    engine: VB6Engine,
    code: source,
    offset: at,
    moduleName: 'Module1',
    modules: [{ name: 'EcrivaineDOM', code: host, host: true }],
    types: { ActiveDocument: 'Document' },
    member: L.domIr.member,
    domTypes: L.domIr.types,
    intrinsics: VB6Engine.intrinsicCatalog(),
    files: false,
    purpose: L.vbDomRef && L.vbDomRef.purpose,
  });
}

function inspect(at, text) {
  const source = text == null ? code : text;
  return L.vbInspect({
    engine: VB6Engine,
    code: source,
    offset: at,
    moduleName: 'Module1',
    modules: [{ name: 'EcrivaineDOM', code: host, host: true }],
    types: { ActiveDocument: 'Document' },
    member: L.domIr.member,
    domTypes: L.domIr.types,
    intrinsics: VB6Engine.intrinsicCatalog(),
    files: false,
    purpose: L.vbDomRef && L.vbDomRef.purpose,
  });
}

test('completion follows the procedure and the virtual machine', async () => {
  await boot();
  const inside = complete(code.indexOf('Lef') + 3);
  assert.equal(has(inside, 'Left'), true);
  assert.equal(has(inside, 'Left$'), true);
  assert.equal(has(inside, 'count'), false);

  const inMain = code.indexOf('  Lef');
  const locals = complete(inMain);
  assert.equal(has(locals, 'count'), true);
  assert.equal(has(locals, 'n'), true);
  assert.equal(has(locals, 'hiddenLocal'), false);
  assert.equal(has(locals, 'shared'), true);
  assert.equal(has(locals, 'Shown'), true);
  assert.equal(has(locals, 'HostOnly'), false);
  assert.equal(has(locals, 'MsgBox'), true);
  assert.equal(has(locals, 'Shell'), false);
  assert.equal(has(locals, 'OpenDatabase'), false);
  assert.equal(has(locals, 'App'), false);
  assert.equal(has(locals, 'Dir'), false);

  const outside = complete(code.indexOf('\ncou') + 1);
  assert.equal(has(outside, 'count'), false);
  assert.equal(has(outside, 'hiddenLocal'), false);
  assert.equal(has(outside, 'shared'), true);

  const moduleDot = complete(code.indexOf('Lef'));
  const quoted = code.replace('  Lef', '  EcrivaineDOM.');
  const onModule = complete(quoted.indexOf('EcrivaineDOM.') + 'EcrivaineDOM.'.length, quoted);
  assert.equal(has(onModule, 'ActiveDocument'), true);
  assert.equal(has(onModule, 'Shown'), true);
  assert.equal(has(onModule, 'HostOnly'), false);
  assert.equal(has(onModule, 'LocalOnly'), false);
  assert.equal(moduleDot.start >= 0, true);

  const documentDot = code.replace('  Lef', '  ActiveDocument.');
  const members = complete(documentDot.indexOf('ActiveDocument.') + 'ActiveDocument.'.length, documentDot);
  assert.equal(has(members, 'Name'), true);
  assert.equal(has(members, 'Paragraphs'), true);
  const saveAs2 = members.items.find((item) => item.name === 'SaveAs2');
  assert.equal(saveAs2.params, '[FileName], [FileFormat], [LockComments], [Password]');
  const shapes = code.replace('  Lef', '  ActiveDocument.InlineShapes.');
  const shapeMembers = complete(shapes.indexOf('ActiveDocument.InlineShapes.') + 'ActiveDocument.InlineShapes.'.length, shapes);
  const addPicture = shapeMembers.items.find((item) => item.name === 'AddPicture');
  assert.equal(addPicture.params, '[FileName], [LinkToFile], [SaveWithDocument], [Range]');
  assert.equal(has(members, 'Application'), true);
  assert.equal(has(members, 'Parent'), true);
  assert.equal(has(members, 'Range'), true);
  assert.equal(has(members, 'Characters'), true);
  assert.equal(has(members, 'Creator'), false);

  const typed = 'Sub Main()\n  Dim x As Lon\nEnd Sub\n';
  const types = complete(typed.indexOf('Lon') + 3, typed);
  assert.deepEqual(names(types), ['Long']);

  const comment = "Sub Main()\n  ' Lef\nEnd Sub\n";
  assert.equal(complete(comment.indexOf('Lef') + 3, comment).items.length, 0);
  assert.equal(complete('#D'.length, '#D').items.length, 0);
  assert.equal(complete('Print "D'.length, 'Print "D').items.length, 0);
  assert.equal(complete('Print "say ""D'.length, 'Print "say ""D').items.length, 0);
  assert.equal(complete('Print "Documents.'.length, 'Print "Documents.').items.length, 0);

  const docHost = host + '\nPublic Property Get Documents() As Object\nEnd Property\n';
  const docRequest = {
    engine: VB6Engine,
    moduleName: 'Module1',
    modules: [{ name: 'EcrivaineDOM', code: docHost, host: true }],
    types: { ActiveDocument: 'Document', Documents: 'Documents' },
    member: L.domIr.member,
    domTypes: L.domIr.types,
    intrinsics: VB6Engine.intrinsicCatalog(),
    files: false,
    purpose: L.vbDomRef && L.vbDomRef.purpose,
  };
  const first = L.vbComplete(Object.assign({ code: 'D', offset: 1 }, docRequest));
  const dim = first.items.find((item) => item.name === 'Dim');
  assert.equal(!!dim, true);
  assert.equal(dim.kind, 'keyword');
  assert.equal(has(first, 'Documents'), true);
  assert.equal(has(first, 'Do'), true);
  const afterQuote = L.vbComplete(Object.assign({ code: 'Print "ok" D', offset: 'Print "ok" D'.length }, docRequest));
  assert.equal(afterQuote.items.find((item) => item.name === 'Dim').kind, 'keyword');
  assert.equal(has(afterQuote, 'Debug'), true);
  const endIf = L.vbComplete(Object.assign({ code: 'End I', offset: 'End I'.length }, docRequest));
  assert.equal(endIf.items.find((item) => item.name === 'If').kind, 'keyword');
  const chain = L.vbComplete(Object.assign({ code: 'Documents.', offset: 'Documents.'.length }, docRequest));
  assert.equal(has(chain, 'Count'), true);
  assert.equal(has(chain, 'Dim'), false);
  assert.equal(has(chain, 'If'), false);

  const framed = L.vbComplete({
    engine: VB6Engine,
    code: 'coun',
    offset: 4,
    moduleName: 'Module1',
    source: code,
    modules: [{ name: 'EcrivaineDOM', code: host, host: true }],
    types: { ActiveDocument: 'Document' },
    member: L.domIr.member,
    domTypes: L.domIr.types,
    intrinsics: VB6Engine.intrinsicCatalog(),
    files: false,
    purpose: L.vbDomRef && L.vbDomRef.purpose,
    frame: { procedure: 'Main', module: 'Module1', locals: [{ name: 'count', type: 'Long' }, { name: 'temp', type: 'Long' }] },
  });
  assert.equal(has(framed, 'count'), true);
  assert.equal(has(framed, 'temp'), false);
  assert.equal(has(framed, 'hiddenLocal'), false);
  assert.equal(framed.items.find((item) => item.name === 'count').type, 'Long');
  const paused = L.vbComplete({
    engine: VB6Engine,
    code: '',
    offset: 0,
    moduleName: 'Module1',
    source: code,
    modules: [{ name: 'EcrivaineDOM', code: host, host: true }],
    types: { ActiveDocument: 'Document' },
    member: L.domIr.member,
    domTypes: L.domIr.types,
    intrinsics: VB6Engine.intrinsicCatalog(),
    files: false,
    purpose: L.vbDomRef && L.vbDomRef.purpose,
    frame: { procedure: 'Main', module: 'Module1', locals: [{ name: 'count', type: 'Long' }, { name: 'temp', type: 'Long' }] },
  });
  assert.equal(has(paused, 'count'), true);
  assert.equal(has(paused, 'temp'), true);
  assert.equal(has(paused, 'n'), true);
  assert.equal(has(paused, 'hiddenLocal'), false);

  const atCount = code.indexOf('count');
  const countInfo = inspect(atCount);
  assert.equal(countInfo.name, 'count');
  assert.equal(countInfo.type, 'Long');
  assert.equal(countInfo.local, true);
  assert.equal(countInfo.line, 5);
  const shown = inspect(1, 'Shown');
  assert.equal(shown.name, 'Shown');
  assert.equal(shown.local, false);
  const saveSrc = 'ActiveDocument.SaveAs';
  const saveAs = inspect(saveSrc.length - 2, saveSrc);
  assert.equal(saveAs.name, 'SaveAs');
  assert.equal(saveAs.local, false);
  assert.equal(saveAs.params.indexOf('FileName') >= 0, true);
  assert.equal(saveAs.purpose, 'On the VS Code host, writes that path and opens it. Elsewhere G-HOST.');
});

test('a procedure carries the reference note, and a local function its return type', async () => {
  await boot();
  const src = 'ActiveDocument.\n';
  const members = complete(src.indexOf('ActiveDocument.') + 'ActiveDocument.'.length, src);
  const activate = members.items.find((item) => item.name === 'Activate');
  assert.equal(activate.purpose, 'Switches the session to this document.');
  assert.equal(activate.params, undefined);
  const saveAs = members.items.find((item) => item.name === 'SaveAs');
  assert.equal(saveAs.purpose, 'On the VS Code host, writes that path and opens it. Elsewhere G-HOST.');
  assert.equal(saveAs.params.indexOf('FileName') >= 0, true);
  assert.equal(saveAs.type, undefined);
  assert.equal(saveAs.unavailable, true);
  assert.equal(saveAs.why, 'platform');
  assert.equal(activate.unavailable, undefined);
  const select = members.items.find((item) => item.name === 'Select');
  assert.equal(select.unavailable, undefined);
  L.hostAudit = { has: (name) => String(name).toLowerCase() === 'saveas' };
  try {
    const onHost = complete(src.indexOf('ActiveDocument.') + 'ActiveDocument.'.length, src);
    assert.equal(onHost.items.find((item) => item.name === 'SaveAs').unavailable, undefined);
    assert.equal(onHost.items.find((item) => item.name === 'Save').why, 'platform');
    assert.equal(onHost.items.find((item) => item.name === 'SaveAs').why, undefined);
    assert.equal(onHost.items.find((item) => item.name === 'FullName').unavailable, true);
    assert.equal(onHost.items.find((item) => item.name === 'Activate').unavailable, undefined);
  } finally {
    delete L.hostAudit;
  }
  const fn = 'Function Foo(n As Long) As Integer\n  Fo\nEnd Function\n';
  const foo = complete(fn.indexOf('Fo') + 2, fn).items.find((item) => item.name === 'Foo');
  assert.equal(foo.type, 'Integer');
  assert.equal(foo.params, 'n');
  assert.equal(foo.purpose, undefined);
});

test('intrinsic catalog is a subset of the virtual machine library', async () => {
  await boot();
  const catalog = VB6Engine.intrinsicCatalog();
  const program = VB6Engine.compileProject({
    schema: 1,
    id: 't',
    name: 'T',
    startup: '',
    modules: [{ id: 'm', name: 'Module1', kind: 'module', code: 'Sub Main()\nEnd Sub\n' }],
    settings: {},
    references: [],
    assets: {},
  });
  assert.equal(program.valid, true, JSON.stringify(program.diagnostics));
  const machine = new VB6Engine.VirtualMachine(program, { print() {}, error() {} });
  try {
    for (const proc of catalog.procedures) {
      assert.equal(machine.library.has(proc.name.toLowerCase()), true, proc.name);
    }
    assert.equal(catalog.procedures.some((proc) => proc.name.toLowerCase() === 'opendatabase'), false);
    assert.equal(catalog.procedures.some((proc) => proc.name === 'Left$'), true);
    assert.equal(machine.library.has('debug'), true);
    assert.equal(machine.library.has('err'), true);
  } finally {
    machine.stop();
  }
});
