const test = require('node:test');
const assert = require('node:assert/strict');
const { boot, fresh, expectGap } = require('./vb-boot.js');

test('M5 gap-only objects are named and reject', { timeout: 120000 }, async () => {
  await boot();
  fresh();
  const host = await L.vb.open();
  try {
    for (const name of ['TimeLine', 'Sequence', 'Effect', 'Chart', 'Diagram', 'OLEFormat', 'MediaFormat', 'CustomerData', 'NotesPage', 'Comments', 'Tags', 'PrintOptions', 'Panes', 'CommandBars']) {
      assert.equal(await host.eval('?' + name + '.Parent Is Application'), -1, name);
    }
    await expectGap(host, '?TimeLine.MainSequence', 'G-ENGINE');
    await expectGap(host, '?Chart.ChartData', 'G-ENGINE');
    await expectGap(host, '?NotesPage.Shapes', 'G-ENGINE');
    await expectGap(host, '?PrintOptions.OutputType', 'G-ENGINE');
    await expectGap(host, '?CommandBars.Count', 'G-HOST');
    await expectGap(host, '?Panes.Count', 'G-HOST');
    await expectGap(host, '?ActivePresentation.Slides(1).Comments', 'G-ENGINE');
    await expectGap(host, '?ActivePresentation.SlideMaster.Shapes', 'G-ENGINE');
    await expectGap(host, '?ActiveWindow.Caption', 'G-HOST');
  } finally {
    host.close();
  }
});
