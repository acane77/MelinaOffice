/* Run tests/ledger/vb-playground.bas against the open sheet. */
const test = require('node:test');
const assert = require('node:assert/strict');
const { boot, fresh, expectGap, fs, path } = require('./vb-boot.js');

test('playground script writes the sheet', async () => {
  await boot();
  fresh();
  const host = await L.vb.open();
  const src = fs.readFileSync(path.join(__dirname, 'vb-playground.bas'), 'utf8');
  await host.run(src);
  const sh = L.dom.Application.ActiveSheet;
  assert.equal(sh.Range('A1').Value, 'Item');
  assert.equal(sh.Range('B1').Value, 'Amount');
  assert.equal(sh.Range('A2').Value, 'Pens');
  assert.equal(sh.Range('B2').Value, 12.5);
  assert.equal(sh.Range('B4').Value, 20);
  assert.equal(sh.Range('A1').Font.Bold, -1);
  assert.equal(sh.Range('B2').NumberFormat, '0.00');
  assert.equal(sh.Range('A6').Value, 20);
  await host.eval('ActiveSheet.Range(Cell1:="C2", Cell2:="C3").Value = ActiveSheet.Range(Cell1:="B2", Cell2:="B3").Value');
  assert.equal(sh.Range('C2').Value, 12.5);
  assert.equal(sh.Range('C3').Value, 7.5);
  assert.equal(L.dom.Application.ActiveWorkbook.Names.Item('Total').RefersTo, '=Sheet1!$B$4');
  await expectGap(host, 'Application.Run Macro:="Nope"', 'G-HOST');
  await expectGap(host, '?ActiveWorkbook.FullName', 'G-HOST');
  await expectGap(host, '?ActiveWorkbook.Creator', 'G-ENGINE');
  host.close();
});
