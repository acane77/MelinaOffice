/* VB6 forms on the suite's dialogs. Node only: no browser.
   node --test tests/vb-forms.test.mjs */
import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';
import vm from 'node:vm';

const require = createRequire(import.meta.url);
const file = fileURLToPath(import.meta.url);
const hostName = process.argv.includes('--host') ? process.argv[process.argv.indexOf('--host') + 1] : '';

const FRM = [
  'VERSION 5.00',
  'Begin VB.Form Form1',
  '   Caption         =   "Hello"',
  '   ClientWidth     =   4800',
  '   BeginProperty Font',
  '      Name            =   "Tahoma"',
  '      Size            =   9',
  '      Weight          =   700',
  '   EndProperty',
  '   Begin VB.CommandButton Command1',
  '      Caption         =   "&Go"',
  '      Left            =   120',
  '   End',
  '   Begin VB.Menu mnuFile',
  '      Caption         =   "&File"',
  '   End',
  'End',
  'Attribute VB_Name = "Form1"',
  'Attribute VB_GlobalNameSpace = False',
  'Private Sub Command1_Click()',
  'End Sub',
  '',
].join('\n');

function sampleProject() {
  return {
    schema: 1,
    startup: 'Form1',
    modules: [
      {
        name: 'Form1',
        kind: 'form',
        code: [
          'Private Sub Form_Load()',
          '    Text1.Text = "Loaded"',
          'End Sub',
          'Private Sub Command1_Click()',
          '    Text1.Text = "Clicked"',
          '    Answer = MsgBox("Go?", vbYesNo)',
          'End Sub',
          '',
        ].join('\n'),
        form: {
          type: 'Form',
          properties: { Caption: 'Form1', ClientWidth: 4000, ClientHeight: 2000 },
          controls: [
            { name: 'Text1', type: 'TextBox', properties: { Left: 120, Top: 120, Width: 2000, Height: 300, Text: '' } },
            { name: 'Command1', type: 'CommandButton', properties: { Left: 120, Top: 600, Width: 1200, Height: 360, Caption: 'Go' } },
          ],
          menus: [],
        },
      },
      { name: 'Mod1', kind: 'module', code: 'Public Answer As Variant\n' },
    ],
  };
}

function loadForms() {
  globalThis.window = globalThis;
  require('../public/common/vb-forms.js');
  return globalThis.VO.vbForms;
}

async function bootQuire() {
  globalThis.window = globalThis;
  require('../public/common/host-audit.js');
  require('../public/common/core.js');
  require('../public/quire/js/dmodel.js');
  require('../public/quire/js/ops.js');
  require('../public/quire/js/dom.js');
  require('../public/quire/js/dom-ir.js');
  const { buildEngine } = await import('../tools/vb6/build-engine.mjs');
  const out = buildEngine();
  vm.runInThisContext(fs.readFileSync(out, 'utf8'), { filename: out });
  require('../public/common/vb-forms.js');
  require('../public/quire/js/vb-host.js');
  const { D } = globalThis.L;
  D.doc = D.newDoc();
  D.resetHistory();
}

async function runHost(name) {
  if (name === 'quire') await bootQuire();
  else if (name === 'ledger') {
    const boot = require('./ledger/vb-boot.js');
    await boot.boot();
    boot.fresh();
  } else if (name === 'lectern') {
    const boot = require('./lectern/vb-boot.js');
    await boot.boot();
    boot.fresh();
  } else throw new Error('Unknown host ' + name);
  const L = globalThis.L;
  const forms = globalThis.VO.vbForms;
  L.ui = L.ui || {};
  L.ui.msg = function () { return 0; };
  L.ui.prompt = function () { return null; };
  const host = await L.vb.open();
  try {
    assert.equal(host.createForm, null);
    assert.equal(host.beginModal, null);
    assert.equal(host.msgBox, null);
    assert.equal(host.inputBox, null);
    assert.equal(host.openFile, null);
    assert.equal(host.saveFile, null);
    assert.equal(typeof host.loadProject, 'function');
    await host.loadProject(sampleProject());
    assert.equal(typeof host.createForm, 'function');
    assert.equal(host.openFile, null);
    assert.equal(host.saveFile, null);
    const loaded = host.formMachine.instances.get('form1');
    assert.equal(loaded.formObject.controlMap.get('text1').Text, 'Loaded');
    assert.equal(await host.eval('?Form1.Text1.Text'), 'Loaded');
    const form = host.formMachine.instances.get('form1');
    const button = form.formObject.controlMap.get('command1');
    forms.element(button).click();
    await button._lastEvent;
    assert.equal(await host.eval('?Form1.Text1.Text'), 'Clicked');
    assert.equal(await host.eval('?Answer'), 6);
    const keep = L.ui.msg;
    L.ui.msg = function (text) {
      assert.match(String(text), /do you want to stop it/);
      return 1;
    };
    await assert.rejects(() => host.run('Public Sub Main()\nEnd Sub'), (err) => !!(err && err.quiet));
    assert.equal(forms.running(host), true);
    L.ui.msg = function () { return 0; };
    await host.formMachine.requestUnload(host.formMachine.instances.get('form1'));
    assert.equal(forms.running(host), false);
    assert.equal(host.formMachine, null);
    L.ui.msg = keep;
  } finally {
    host.close();
  }
  assert.equal(host.createForm, null);
  assert.equal(host.msgBox, null);
  assert.equal(host.openFile, null);
}

if (hostName) {
  runHost(hostName).then(function () { process.exit(0); }, function (error) {
    console.error(error);
    process.exit(1);
  });
} else {
  test('a .frm becomes a form module, and a .frx picture is left out', () => {
    const forms = loadForms();
    const parsed = forms.parseFrm(FRM);
    assert.equal(parsed.name, 'Form1');
    assert.match(parsed.code, /^Private Sub Command1_Click\(\)/);
    assert.doesNotMatch(parsed.code, /Attribute/);
    assert.equal(parsed.form.properties.Caption, 'Hello');
    assert.equal(parsed.form.properties.FontName, 'Tahoma');
    assert.equal(parsed.form.properties.FontBold, -1);
    assert.equal(parsed.form.controls[0].name, 'Command1');
    assert.equal(parsed.form.controls[0].properties.Caption, '&Go');
    assert.equal(parsed.form.menus[0].name, 'mnuFile');
    const built = forms.filesToProject([
      { name: 'Form1.frm', text: FRM },
      { name: 'Form1.frx', text: 'picture' },
      { name: 'Mod1.bas', text: 'Attribute VB_Name = "Mod1"\nPublic Answer As Variant\n' },
    ]);
    assert.equal(built.losses.length, 1);
    assert.equal(built.losses[0].notify, false);
    assert.equal(built.losses[0].action, 'drop');
    assert.equal(built.project.startup, 'Form1');
    assert.equal(built.project.modules.length, 2);
  });

  test('a button click is dispatched, and MsgBox and InputBox use the suite dialogs', async () => {
    const forms = loadForms();
    const seen = [];
    const vmStub = { dispatch(instance, name) { seen.push(name); return Promise.resolve('sent'); } };
    const form = forms.create({
      type: 'Form',
      name: 'Form1',
      properties: { Caption: 'Form1' },
      controls: [
        { name: 'Text1', type: 'TextBox', properties: { Text: '' } },
        { name: 'Command1', type: 'CommandButton', properties: { Caption: 'Go' } },
        { name: 'Picture1', type: 'PictureBox', properties: { Tag: 'kept' } },
      ],
      menus: [{ name: 'mnuFile', properties: { Caption: '&File' } }],
    }, {}, vmStub);
    form.Show(0);
    assert.equal(form.shown, true);
    const button = form.controlMap.get('command1');
    assert.equal(button.type, 'CommandButton');
    forms.element(button).click();
    assert.deepEqual(seen, ['Command1_Click']);
    const grouped = forms.create({
      type: 'Form', name: 'Form1', properties: {},
      controls: [
        { name: 'Opt', type: 'OptionButton', properties: { Index: 0, Caption: 'A' } },
        { name: 'Opt', type: 'OptionButton', properties: { Index: 1, Caption: 'B' } },
      ],
      menus: [],
    }, {}, vmStub);
    const opts = grouped.controlMap.get('opt');
    assert.equal(opts.__type, 'ControlArray');
    assert.equal(opts.Count, 2);
    assert.equal(opts.Item(1).Caption, 'B');
    assert.equal(form.controlMap.get('picture1').Tag, 'kept');
    assert.equal(form.controlMap.get('picture1').type, 'PictureBox');
    const host = {};
    globalThis.L.ui = {
      msg: function () { return 0; },
      prompt: function () { return null; },
    };
    forms.install(host);
    assert.equal(await host.msgBox('Go?', 4, 'Form1'), 6);
    assert.equal(await host.inputBox('Name?', 'Form1', 'x'), '');
    form.Hide();
    assert.equal(form.shown, false);
    const kept = { formMachine: { state: 'running' } };
    globalThis.L.ui.msg = function (text) {
      assert.match(String(text), /do you want to stop it/);
      return 1;
    };
    assert.equal(await forms.block(kept), true);
    assert.equal(forms.running(kept), true);
    globalThis.L.ui.msg = function () { return 0; };
    assert.equal(await forms.block(kept), false);
    assert.equal(forms.running(kept), false);
  });

  test('each host runs a form only after loadProject', { timeout: 180000 }, async () => {
    for (const name of ['quire', 'ledger', 'lectern']) {
      await new Promise((resolve, reject) => {
        const child = spawn(process.execPath, [file, '--host', name], { stdio: ['ignore', 'inherit', 'inherit'] });
        child.on('exit', (code) => code === 0 ? resolve() : reject(new Error(name + ' form host failed: ' + code)));
      });
    }
  });
}
