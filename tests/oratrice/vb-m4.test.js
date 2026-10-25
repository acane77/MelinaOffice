const test = require('node:test');
const assert = require('node:assert/strict');
const { boot, fresh, expectGap } = require('./vb-boot.js');

test('M4 tables, fill, links, transition and the show are callable from VB', { timeout: 120000 }, async () => {
  await boot();
  fresh();
  const host = await L.vb.open();
  try {
    await host.eval('ActivePresentation.Slides.Add 2, ppLayoutBlank');
    await host.eval('ActivePresentation.Slides(2).Shapes.AddTable 2, 2, 20, 80, 200');
    const n = ActivePresentation.Slides(2).Shapes.Count;
    await host.eval('ActivePresentation.Slides(2).Shapes(' + n + ').Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "A1"');
    assert.equal(ActivePresentation.Slides(2).Shapes(n).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text, 'A1');

    await host.eval('ActivePresentation.Slides(2).Shapes.AddShape msoShapeOval, 10, 10, 40, 40');
    const oval = ActivePresentation.Slides(2).Shapes.Count;
    await host.eval('ActivePresentation.Slides(2).Shapes(' + oval + ').Fill.ForeColor.RGB = 255');
    await host.eval('ActivePresentation.Slides(2).Shapes(' + oval + ').Line.Weight = 2');
    const shape = ActivePresentation.Slides(2).Shapes(oval);
    assert.equal(shape.Fill.ForeColor.RGB, 255);
    assert.equal(shape.Fill.Type, msoFillSolid);
    assert.equal(shape.Line.Weight, 2);

    const title = 'ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange';
    await host.eval(title + '.Text = "Site"');
    await host.eval(title + '.Hyperlinks.Add Nothing, "https://acane.net/melinaoffice", "", "", "Site"');
    assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Hyperlinks.Count, 1);
    assert.equal(ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Hyperlinks(1).Address, 'https://acane.net/melinaoffice');

    await host.eval('ActivePresentation.SlideMaster.HeadersFooters.Footer.Text = "Foot"');
    assert.equal(ActivePresentation.SlideMaster.HeadersFooters.Footer.Visible, -1);
    assert.equal(ActivePresentation.SlideMaster.HeadersFooters.Footer.Text, 'Foot');

    const design = L.model.DESIGNS[1];
    await host.eval('ActivePresentation.Slides(2).Design = "' + design.name + '"');
    assert.equal(ActivePresentation.Slides(2).Design.Name, design.name);

    await host.eval('ActivePresentation.Slides(2).SlideShowTransition.EntryEffect = ppEffectFadeSmoothly');
    await host.eval('ActivePresentation.Slides(2).SlideShowTransition.Speed = ppTransitionSpeedSlow');
    assert.equal(ActivePresentation.Slides(2).SlideShowTransition.EntryEffect, ppEffectFadeSmoothly);
    assert.equal(ActivePresentation.Slides(2).SlideShowTransition.Speed, ppTransitionSpeedSlow);

    await host.eval('ActivePresentation.SlideShowSettings.Run');
    assert.equal(ActivePresentation.SlideShowWindows.Count, 1);
    await host.eval('ActivePresentation.SlideShowWindows(1).View.Next');
    assert.equal(ActivePresentation.SlideShowWindows(1).View.CurrentShowPosition, 2);
    await host.eval('ActivePresentation.SlideShowWindows(1).View.GotoSlide 1');
    assert.equal(ActivePresentation.SlideShowWindows(1).View.CurrentShowPosition, 1);
    await host.eval('ActivePresentation.SlideShowWindows(1).View.Exit');
    assert.equal(ActivePresentation.SlideShowWindows.Count, 0);

    await expectGap(host, 'ActivePresentation.Slides(2).Shapes.AddPicture "x", msoFalse, msoFalse, 0, 0', 'G-HOST');
    await expectGap(host, 'ActivePresentation.Slides(2).Shapes.AddChart 1, 0, 0, 100, 100', 'G-ENGINE');
    await expectGap(host, '?ActivePresentation.Slides(1).TimeLine', 'G-ENGINE');
  } finally {
    host.close();
  }
});
