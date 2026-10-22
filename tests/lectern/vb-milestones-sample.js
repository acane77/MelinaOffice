/* Paste into the Lectern DOM Console with the language set to JavaScript, then:
     await runVbMilestones()
   Save and Presentations.Open are left to the Node tests: in the browser they open a dialog.
   Each check starts from a blank presentation. After the checks, the open presentation is rebuilt
   so the page shows the title, the text box, the table and the link.
   Run Script uses Sub Main; these lines call the host. */
async function runVbMilestones() {
  const failures = [];
  function fail(message) { failures.push(String(message)); }
  function same(label, actual, expected) {
    if (actual !== expected) fail(label + ' = ' + actual + ', live ' + expected);
  }
  async function gap(statement, code) {
    try {
      await host.eval(statement);
      fail(statement + ' did not fail');
    } catch (err) {
      if (!err || err.code !== code) fail(statement + ' -> ' + (err && err.code) + ' ' + (err && err.message));
    }
  }
  function reset() {
    L.pres = L.model.newPresentation();
    L.dom.reset();
    if (L.ed && L.ed.render) L.ed.render();
  }
  async function present() {
    reset();
    await host.eval('ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "Hello from Visual Basic"');
    await host.eval('ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Font.Name = "Arial"');
    await host.eval('ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Font.Size = 32');
    await host.eval('ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Font.Bold = msoTrue');
    await host.eval('ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Hyperlinks.Add Nothing, "https://vibeoffice.work"');
    await host.eval('ActivePresentation.Slides.Add 2, ppLayoutBlank');
    await host.eval('ActivePresentation.Slides(2).Shapes.AddTextbox ppTextOrientationHorizontal, 72, 36, 240, 40');
    await host.eval('ActivePresentation.Slides(2).Shapes(1).TextFrame.TextRange.Text = "A text box"');
    await host.eval('ActivePresentation.Slides(2).Shapes.AddTable 2, 2, 72, 100, 240');
    await host.eval('ActivePresentation.Slides(2).Shapes(2).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Name"');
    await host.eval('ActivePresentation.Slides(2).Shapes(2).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Value"');
    await host.eval('ActivePresentation.Slides(2).SlideShowTransition.EntryEffect = ppEffectFadeSmoothly');
    if (L.ed && L.ed.render) L.ed.render();
  }

  delete window.VB6Engine;
  document.querySelectorAll('script[src*="vb6-engine"]').forEach(function (node) { node.remove(); });
  reset();
  const host = await L.vb.open();
  try {
    for (const key of Object.keys(L.dom)) {
      if (/^(pp|mso)[A-Za-z][A-Za-z0-9]*$/.test(key) && Number.isInteger(L.dom[key])) same(key, await host.eval('?' + key), L.dom[key]);
    }
    same('Name', await host.eval('?ActivePresentation.Name'), ActivePresentation.Name);
    same('Presentations.Count', await host.eval('?Presentations.Count'), Presentations.Count);
    same('Parent', await host.eval('?ActivePresentation.Parent Is Presentations'), -1);
    same('Application', await host.eval('?ActivePresentation.Application Is Application'), -1);
    const name = ActivePresentation.Name;
    try { await host.eval('ActivePresentation.Name = "Z"'); fail('Name assignment did not fail'); }
    catch (err) { if (ActivePresentation.Name !== name) fail('Name changed'); }
    await host.eval('ActivePresentation.Saved = False');
    same('Saved', ActivePresentation.Saved, false);
    await host.eval('ActivePresentation.Saved = True');
    await host.eval('ActivePresentation.BuiltInDocumentProperties("Title").Value = "M1"');
    same('Title', ActivePresentation.BuiltInDocumentProperties('Title').Value, 'M1');
    await host.eval('ActivePresentation.PageSetup.SlideWidth = 720');
    same('SlideWidth', ActivePresentation.PageSetup.SlideWidth, 720);
    await gap('?ActivePresentation.Creator', 'G-ENGINE');
    await gap('?ActivePresentation.FullName', 'G-HOST');
    await gap('Presentations.Open "x"', 'G-HOST');
    await gap('ActivePresentation.SaveAs "x"', 'G-HOST');
    await gap('ActivePresentation.PrintOut', 'G-HOST');
    await gap('?ActiveWindow.View', 'G-HOST');

    reset();
    await host.eval('ActivePresentation.Slides.Add 2, ppLayoutBlank');
    same('slides', await host.eval('?ActivePresentation.Slides.Count'), 2);
    await host.eval('ActivePresentation.Slides(2).Shapes.AddShape msoShapeRectangle, 10, 20, 40, 30');
    await host.eval('ActivePresentation.Slides(2).Shapes(1).Left = 15');
    same('Left', ActivePresentation.Slides(2).Shapes(1).Left, 15);
    await host.eval('ActivePresentation.Slides(2).Shapes(1).Select');
    same('selection', await host.eval('?ActiveWindow.Selection.Type'), ppSelectionShapes);
    await host.eval('ActiveWindow.Selection.ShapeRange.Left = 18');
    same('ShapeRange.Left', ActivePresentation.Slides(2).Shapes(1).Left, 18);

    reset();
    const range = 'ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange';
    await host.eval(range + '.Text = "Hello"');
    await host.eval(range + '.Font.Bold = msoTrue');
    await host.eval(range + '.ParagraphFormat.Alignment = ppAlignLeft');
    same('text', ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text, 'Hello');
    same('bold', ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Font.Bold, -1);
    same('count', await host.eval('?' + range + '.Characters.Count'), 5);
    await gap(range + '.Characters(1)', 'G-TEXT');
    await gap('?' + range + '.Start', 'G-TEXT');
    await gap(range + '.Find "H"', 'G-TEXT');

    reset();
    await host.eval('ActivePresentation.Slides.Add 2, ppLayoutBlank');
    await host.eval('ActivePresentation.Slides(2).Shapes.AddTable 2, 2, 20, 80, 200');
    await host.eval('ActivePresentation.Slides(2).Shapes(1).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "A1"');
    same('cell', ActivePresentation.Slides(2).Shapes(1).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text, 'A1');
    await host.eval('ActivePresentation.Slides(2).Shapes.AddShape msoShapeOval, 10, 10, 40, 40');
    await host.eval('ActivePresentation.Slides(2).Shapes(2).Fill.ForeColor.RGB = 255');
    same('fill', ActivePresentation.Slides(2).Shapes(2).Fill.ForeColor.RGB, 255);
    await host.eval(range + '.Text = "Site"');
    await host.eval(range + '.Hyperlinks.Add Nothing, "https://vibeoffice.work"');
    same('link', ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Hyperlinks(1).Address, 'https://vibeoffice.work');
    await host.eval('ActivePresentation.Slides(2).SlideShowTransition.EntryEffect = ppEffectFadeSmoothly');
    same('effect', ActivePresentation.Slides(2).SlideShowTransition.EntryEffect, ppEffectFadeSmoothly);
    await host.eval('ActivePresentation.SlideShowSettings.Run');
    await host.eval('ActivePresentation.SlideShowWindows(1).View.Next');
    same('show', ActivePresentation.SlideShowWindows(1).View.CurrentShowPosition, 2);
    await host.eval('ActivePresentation.SlideShowWindows(1).View.Exit');
    await gap('ActivePresentation.Slides(2).Shapes.AddPicture "x", msoFalse, msoFalse, 0, 0', 'G-HOST');
    await gap('?ActivePresentation.Slides(1).TimeLine', 'G-ENGINE');

    const types = L.domIr.types;
    for (const typeName of Object.keys(types)) {
      if (!types[typeName].gap) continue;
      same(typeName + '.Parent', await host.eval('?' + typeName + '.Parent Is Application'), -1);
      await gap('?' + typeName + '.Creator', 'G-ENGINE');
    }
    if (host.createForm !== null || host.loadProject !== null) fail('form slots are set');
    await present();
  } finally {
    host.close();
  }
  return { ok: failures.length === 0, failures: failures };
}
