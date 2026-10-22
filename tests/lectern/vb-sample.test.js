const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { boot, fresh } = require('./vb-boot.js');

test('the Run Script sample matches the milestone checks', { timeout: 120000 }, async () => {
  await boot();
  fresh();
  const source = fs.readFileSync(path.join(__dirname, 'vb-milestones-sample.vbs'), 'utf8');
  const host = await L.vb.open();
  try {
    const before = Presentations.Count;
    await host.run(source);
    assert.deepEqual(host.output, ['ok']);
    assert.equal(Presentations.Count, before + 1);
    const title = ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange;
    assert.equal(title.Text, 'Hello from Visual Basic');
    assert.equal(title.Font.Name, 'Arial');
    assert.equal(title.Font.Bold, -1);
    assert.equal(title.Hyperlinks(1).Address, 'https://vibeoffice.work');
    assert.equal(ActivePresentation.Slides(2).Shapes(1).TextFrame.TextRange.Text, 'A text box');
    assert.equal(ActivePresentation.Slides(2).Shapes(2).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text, 'Name');
    assert.equal(ActivePresentation.Slides(2).SlideShowTransition.EntryEffect, ppEffectFadeSmoothly);
  } finally {
    host.close();
  }
});
