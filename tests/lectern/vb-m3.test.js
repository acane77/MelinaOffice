const test = require('node:test');
const assert = require('node:assert/strict');
const { boot, fresh, expectGap } = require('./vb-boot.js');

test('M3 text, font and paragraph format are callable from VB', { timeout: 120000 }, async () => {
  await boot();
  fresh();
  const host = await L.vb.open();
  try {
    const range = 'ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange';
    await host.eval(range + '.Text = "Hello"');
    assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, 'Hello');
    await host.eval(range + '.Font.Name = "Arial"');
    await host.eval(range + '.Font.Size = 32');
    await host.eval(range + '.Font.Bold = msoTrue');
    await host.eval(range + '.ParagraphFormat.Alignment = ppAlignLeft');
    const live = ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange;
    assert.equal(live.Font.Name, 'Arial');
    assert.equal(live.Font.Size, 32);
    assert.equal(live.Font.Bold, -1);
    assert.equal(live.ParagraphFormat.Alignment, ppAlignLeft);
    await host.eval(range + '.InsertAfter "!"');
    assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, 'Hello!');
    assert.equal(await host.eval('?' + range + '.Characters.Count'), 6);
    await host.eval(range + '.ParagraphFormat.Bullet.Type = ppBulletUnnumbered');
    assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.ParagraphFormat.Bullet.Type, ppBulletUnnumbered);
    assert.equal(await host.eval('?ActivePresentation.Slides(1).Shapes.Title.PlaceholderFormat.Type'), ppPlaceholderCenterTitle);
    await expectGap(host, range + '.Characters(1)', 'G-TEXT');
    await expectGap(host, '?' + range + '.Start', 'G-TEXT');
    await expectGap(host, '?' + range + '.Length', 'G-TEXT');
    await expectGap(host, range + '.Find "H"', 'G-TEXT');
  } finally {
    host.close();
  }
});
