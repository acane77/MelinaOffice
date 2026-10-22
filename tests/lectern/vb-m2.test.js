const test = require('node:test');
const assert = require('node:assert/strict');
const { boot, fresh } = require('./vb-boot.js');

test('M2 slides and shapes are callable from VB', { timeout: 120000 }, async () => {
  await boot();
  fresh();
  const host = await L.vb.open();
  try {
    assert.equal(await host.eval('?ActivePresentation.Slides.Count'), 1);
    await host.eval('ActivePresentation.Slides.Add 2, ppLayoutBlank');
    assert.equal(ActivePresentation.Slides.Count, 2);
    assert.equal(ActivePresentation.Slides(2).Layout, ppLayoutBlank);
    await host.eval('ActivePresentation.Slides(2).Name = "Notes"');
    assert.equal(ActivePresentation.Slides(2).Name, 'Notes');
    await host.eval('ActivePresentation.Slides(2).MoveTo 1');
    assert.equal(ActivePresentation.Slides(1).Name, 'Notes');
    await host.eval('ActivePresentation.Slides(1).Duplicate');
    assert.equal(ActivePresentation.Slides.Count, 3);

    await host.eval('ActivePresentation.Slides(1).Shapes.AddTextbox ppTextOrientationHorizontal, 72, 36, 200, 40');
    const box = ActivePresentation.Slides(1).Shapes(ActivePresentation.Slides(1).Shapes.Count);
    assert.equal(box.Left, 72);
    assert.equal(box.Type, msoTextBox);
    await host.eval('ActivePresentation.Slides(1).Shapes.AddShape msoShapeRectangle, 10, 20, 40, 30');
    const rect = ActivePresentation.Slides(1).Shapes(ActivePresentation.Slides(1).Shapes.Count);
    assert.equal(rect.AutoShapeType, msoShapeRectangle);
    await host.eval('ActivePresentation.Slides(1).Shapes(' + ActivePresentation.Slides(1).Shapes.Count + ').Left = 15');
    assert.equal(rect.Left, 15);
    await host.eval('ActivePresentation.Slides(1).Shapes(' + ActivePresentation.Slides(1).Shapes.Count + ').Select');
    assert.equal(ActiveWindow.Selection.Type, ppSelectionShapes);
    await host.eval('ActiveWindow.Selection.ShapeRange.Left = 18');
    assert.equal(rect.Left, 18);
    const before = ActivePresentation.Slides(1).Shapes.Count;
    await host.eval('ActiveWindow.Selection.ShapeRange.Delete');
    assert.equal(ActivePresentation.Slides(1).Shapes.Count, before - 1);
    assert.equal(box.Left, 72);
  } finally {
    host.close();
  }
});
