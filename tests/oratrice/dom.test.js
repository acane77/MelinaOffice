const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../public/common/core.js');
require('../../public/common/geometry.js');
require('../../public/oratrice/js/model.js');
require('../../public/oratrice/js/textedit.js');
require('../../public/oratrice/js/dom.js');

function fresh() {
  L.pres = L.model.newPresentation();
  L.dom.reset();
}

test('PowerPoint object model edits the open presentation', () => {
  fresh();
  const title = ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange;
  title.Text = 'Hello';
  assert.equal(title.Text, 'Hello');
  assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, 'Hello');
  title.Font.Bold = -1;
  title.Font.Size = 32;
  title.Font.Name = 'Arial';
  assert.equal(title.Font.Bold, -1);
  assert.equal(title.Font.Size, 32);
  assert.equal(title.Font.Name, 'Arial');
  title.ParagraphFormat.Alignment = ppAlignLeft;
  assert.equal(title.ParagraphFormat.Alignment, ppAlignLeft);

  const slide = ActivePresentation.Slides.Add(2, ppLayoutBlank);
  assert.equal(slide.Layout, ppLayoutBlank);
  assert.equal(ActivePresentation.Slides.Count, 2);
  const box = slide.Shapes.AddTextbox(ppTextOrientationHorizontal, 72, 36, 200, 40);
  box.TextFrame.TextRange.Text = 'Box';
  assert.equal(box.Left, 72);
  box.Left = 80;
  assert.equal(box.Left, 80);
  box.Fill.ForeColor.RGB = 255;
  assert.equal(box.Fill.ForeColor.RGB, 255);
  const rect = slide.Shapes.AddShape(msoShapeRectangle, 10, 10, 40, 20);
  assert.equal(rect.AutoShapeType, msoShapeRectangle);
  assert.equal(rect.Type, msoAutoShape);

  const tableShape = slide.Shapes.AddTable(2, 2, 20, 80, 200);
  tableShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = 'A1';
  assert.equal(tableShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text, 'A1');

  ActivePresentation.Slides(2).SlideShowTransition.EntryEffect = ppEffectFadeSmoothly;
  assert.equal(ActivePresentation.Slides(2).SlideShowTransition.EntryEffect, ppEffectFadeSmoothly);
  ActivePresentation.SlideMaster.HeadersFooters.Footer.Text = 'Foot';
  assert.equal(ActivePresentation.SlideMaster.HeadersFooters.Footer.Visible, -1);

  ActivePresentation.BuiltInDocumentProperties('Title').Value = 'Deck';
  assert.equal(ActivePresentation.BuiltInDocumentProperties('Title').Value, 'Deck');

  rect.Select();
  assert.equal(ActiveWindow.Selection.Type, ppSelectionShapes);
  ActiveWindow.Selection.ShapeRange.Left = 15;
  assert.equal(rect.Left, 15);

  assert.equal(ActivePresentation.Undo(), -1);
  assert.equal(rect.Left, 10);
  ActivePresentation.Redo();
  assert.equal(rect.Left, 15);

  assert.throws(() => ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Characters(1), (e) => e.code === 'G-TEXT');
  assert.throws(() => ActivePresentation.FullName, (e) => e.code === 'G-HOST');
  assert.throws(() => ActivePresentation.Slides(1).TimeLine, (e) => e.code === 'G-ENGINE');
  assert.equal(title.Characters.Count, 5);

  const before = Presentations.Count;
  Presentations.Add();
  assert.equal(Presentations.Count, before + 1);
  ActivePresentation.Close();
  assert.equal(Presentations.Count, before);
  assert.equal(ActivePresentation.Slides.Count, 2);
});
