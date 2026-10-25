/* Paste into the Ecrivaine DOM Console (it is still JavaScript), then:
     await runVbMilestones()
   Save and Documents.Add() are left to the Node tests: in the browser they open a dialog or a new window.
   Each check starts from a blank document. After the checks, the open document is rebuilt
   so the page shows the paragraph, the table, the link and the note.
   The DOM Console does not speak VB until P5. These lines call the host. */
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
    if (L.app && L.app.loadDoc) L.app.loadDoc(L.D.newDoc(), 'Milestones');
    else { L.D.doc = L.D.newDoc(); L.D.resetHistory(); }
  }
  async function present() {
    reset();
    await host.eval('ActiveDocument.TrackRevisions = False');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Hello world."');
    await host.eval('ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphCenter');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Font.Name = "Arial"');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Font.Size = 16');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Font.Bold = True');
    await host.eval('ActiveDocument.Bookmarks.Add("Here", ActiveDocument.Paragraphs(1).Range)');
    await host.eval('ActiveDocument.Comments.Add(ActiveDocument.Paragraphs(1).Range, "A note")');
    await host.eval('ActiveDocument.Footnotes.Add(ActiveDocument.Paragraphs(1).Range, "Foot")');
    await host.eval('Selection.TypeParagraph');
    await host.eval('Selection.TypeText("link")');
    await host.eval('ActiveDocument.Hyperlinks.Add(Selection.Range, "https://acane.net/melinaoffice", "", "", "Site")');
    await host.eval('ActiveDocument.Tables.Add(Selection.Range, 2, 2)');
    await host.eval('ActiveDocument.Tables(1).Cell(1, 1).Range.Text = "Name"');
    await host.eval('ActiveDocument.Tables(1).Cell(1, 2).Range.Text = "Value"');
    await host.eval('ActiveDocument.Tables(1).Cell(2, 1).Range.Text = "DOM"');
    await host.eval('ActiveDocument.Tables(1).Cell(2, 2).Range.Text = "Word object model"');
    await host.eval('ActiveDocument.Tables(1).Rows(1).HeadingFormat = True');
    await host.eval('ActiveDocument.Tables(1).Cell(1, 1).Shading.BackgroundPatternColor = 65535');
    await host.eval('ActiveDocument.Tables(1).Cell(1, 2).Shading.BackgroundPatternColor = 65535');
    await host.eval('ActiveDocument.Tables(1).Cell(1, 1).Borders(-1).LineStyle = 1');
  }

  delete window.VB6Engine;
  document.querySelectorAll('script[src*="vb6-engine"]').forEach(function (node) { node.remove(); });
  reset();
  const host = await L.vb.open();
  try {
    for (const key of Object.keys(L.dom)) {
      if (/^wd[A-Za-z][A-Za-z0-9]*$/.test(key) && Number.isInteger(L.dom[key])) same(key, await host.eval('?' + key), L.dom[key]);
    }
    same('Name', await host.eval('?ActiveDocument.Name'), ActiveDocument.Name);
    same('Documents.Count', await host.eval('?Documents.Count'), Documents.Count);
    same('Parent', await host.eval('?ActiveDocument.Parent Is Documents'), -1);
    same('Application', await host.eval('?ActiveDocument.Application Is Application'), -1);
    const name = ActiveDocument.Name;
    try { await host.eval('ActiveDocument.Name = "Z"'); fail('Name assignment did not fail'); }
    catch (err) { if (ActiveDocument.Name !== name) fail('Name changed'); }
    await host.eval('ActiveDocument.Saved = False');
    same('Saved', ActiveDocument.Saved, false);
    await host.eval('ActiveDocument.Saved = True');
    await host.eval('ActiveDocument.BuiltInDocumentProperties("Title").Value = "M1"');
    same('Title', ActiveDocument.BuiltInDocumentProperties('Title').Value, 'M1');
    same('Paragraphs stat', await host.eval('?ActiveDocument.ComputeStatistics(wdStatisticParagraphs)'), ActiveDocument.ComputeStatistics(wdStatisticParagraphs));
    await host.eval('ActiveDocument.Select');
    await gap('?ActiveDocument.Creator', 'G-ENGINE');
    await gap('?ActiveDocument.FullName', 'G-HOST');
    await gap('ActiveDocument.Range(0, 5)', 'G-STORY');
    await gap('Documents.Add("x")', 'G-ENGINE');
    await gap('Documents.Open("x")', 'G-HOST');
    await gap('Documents.Save', 'G-HOST');
    await gap('ActiveDocument.SaveAs("x")', 'G-HOST');
    await gap('ActiveDocument.PrintOut', 'G-HOST');
    await gap('ActiveDocument.Close(wdPromptToSaveChanges)', 'G-HOST');

    reset();
    const nothing = await host.eval('?ActiveDocument.Paragraphs(1).Next');
    if (!nothing || nothing.__nothing !== true) fail('Paragraph.Next was not Nothing');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Hello world."');
    same('paragraph', await host.eval('?ActiveDocument.Paragraphs(1).Range.Text'), ActiveDocument.Paragraphs(1).Range.Text);
    same('words', await host.eval('?ActiveDocument.Words.Count'), ActiveDocument.Words.Count);
    await gap('ActiveDocument.Paragraphs(1).Range.Characters(1)', 'G-STORY');
    await gap('ActiveDocument.Sentences.First', 'missing-live');
    await host.eval('ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphCenter');
    same('align', ActiveDocument.Paragraphs(1).Alignment, wdAlignParagraphCenter);
    await host.eval('ActiveDocument.Paragraphs(1).Range.Bold = True');
    same('bold', ActiveDocument.Paragraphs(1).Range.Bold, -1);
    await host.eval('Selection.TypeText("!")');
    if (String(ActiveDocument.Content.Text).indexOf('!') < 0) fail('TypeText did not appear');
    await gap('?Selection.Start', 'G-STORY');
    await gap('Selection.InsertFile("x")', 'G-HOST');
    await gap('Selection.InsertRows', 'G-ENGINE');
    const pages = document.querySelector('.pages');
    if (pages && String(pages.innerText).indexOf('Hello') < 0) fail('page text has no Hello');

    reset();
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Hello"');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Font.Name = "Arial"');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Font.Size = 16');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Font.Bold = True');
    same('font', ActiveDocument.Paragraphs(1).Range.Font.Name, 'Arial');
    same('size', ActiveDocument.Paragraphs(1).Range.Font.Size, 16);
    await gap('?ActiveDocument.Paragraphs(1).Range.Font.Borders', 'G-ENGINE');
    await gap('?ActiveDocument.Paragraphs(1).Range.Font.Animation', 'G-ENGINE');
    await host.eval('ActiveDocument.Paragraphs(1).Format.SpaceBefore = 8');
    same('space', ActiveDocument.Paragraphs(1).Format.SpaceBefore, 8);
    await gap('?ActiveDocument.Paragraphs(1).Format.Borders', 'missing-live');
    await host.eval('ActiveDocument.Paragraphs(1).Format.TabStops.Add(72)');
    same('tabs', ActiveDocument.Paragraphs(1).Format.TabStops.Count, 1);
    await host.eval('ActiveDocument.Styles.Add("BrowserM", wdStyleTypeParagraph)');
    same('style', ActiveDocument.Styles('BrowserM').NameLocal, 'BrowserM');
    await host.eval('ActiveDocument.Styles("BrowserM").Delete');
    await gap('ActiveDocument.Styles(1).Delete', 'G-ENGINE');
    await host.eval('ActiveDocument.Paragraphs(1).Range.ListFormat.ApplyNumberDefault');
    if (ActiveDocument.Lists.Count < 1) fail('list was not applied');
    await gap('?ActiveDocument.Paragraphs(1).Range.ListFormat.ListType', 'missing-live');
    await host.eval('ActiveDocument.Paragraphs(1).Range.ListFormat.RemoveNumbers');
    await host.eval('ActiveDocument.Sections(1).PageSetup.PageWidth = 600');
    same('page width', ActiveDocument.Sections(1).PageSetup.PageWidth, 600);
    await host.eval('ActiveDocument.Sections(1).PageSetup.PageWidth = 612');
    same('header', await host.eval('?ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Exists'), ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Exists);
    await gap('?ActiveDocument.Sections(1).PageSetup.PaperSize', 'G-ENGINE');

    reset();
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Hello world."');
    const replaced = await host.eval('?Selection.Find.Execute("Hello", False, False, False, False, False, True, wdFindContinue, False, "Hi", wdReplaceOne)');
    if (replaced !== true && replaced !== -1) fail('Find.Execute returned ' + replaced);
    if (String(ActiveDocument.Paragraphs(1).Range.Text).indexOf('Hi') < 0) fail('replace missed');
    await gap('Selection.Find.ClearFormatting', 'missing-live');
    await gap('Selection.Find.Execute("Hi", False, False, False, False, False, True, 2)', 'G-HOST');
    await gap('?Selection.Find.Replacement.Font', 'G-ENGINE');
    await host.eval('ActiveDocument.Bookmarks.Add("Here", ActiveDocument.Paragraphs(1).Range)');
    same('bookmark', await host.eval('?ActiveDocument.Bookmarks.Exists("Here")'), ActiveDocument.Bookmarks.Exists('Here'));
    await gap('?ActiveDocument.Bookmarks("Here").Start', 'G-STORY');
    await host.eval('Selection.TypeParagraph');
    await host.eval('Selection.TypeText("link")');
    await host.eval('ActiveDocument.Hyperlinks.Add(Selection.Range, "https://acane.net/melinaoffice", "", "", "Site")');
    if (ActiveDocument.Hyperlinks.Count < 1) fail('hyperlink missing');
    same('address', await host.eval('?ActiveDocument.Hyperlinks(1).Address'), 'https://acane.net/melinaoffice');
    await gap('?ActiveDocument.Hyperlinks(1).Range.Start', 'G-STORY');
    await gap('ActiveDocument.Hyperlinks(1).Follow', 'G-HOST');
    await host.eval('ActiveDocument.Comments.Add(ActiveDocument.Paragraphs(1).Range, "A note")');
    same('initial', await host.eval('?ActiveDocument.Comments(1).Initial'), ActiveDocument.Comments(1).Initials);
    await gap('?ActiveDocument.Comments(1).Scope', 'missing-live');
    await gap('ActiveDocument.Comments(1).Edit', 'G-HOST');
    await host.eval('ActiveDocument.Comments(1).Delete');
    await host.eval('ActiveDocument.Footnotes.Add(ActiveDocument.Paragraphs(1).Range, "Foot")');
    same('note', await host.eval('?ActiveDocument.Footnotes.Count'), ActiveDocument.Footnotes.Count);
    await gap('?ActiveDocument.Footnotes(1).Reference', 'G-STORY');
    await gap('ActiveDocument.Footnotes.Convert', 'missing-live');
    await host.eval('ActiveDocument.Endnotes.Add(ActiveDocument.Paragraphs(1).Range, "End")');
    await gap('?ActiveDocument.Endnotes(1).Reference', 'G-STORY');
    await host.eval('ActiveDocument.Tables.Add(ActiveDocument.Paragraphs(1).Range, 2, 2)');
    same('tables', await host.eval('?ActiveDocument.Tables.Count'), 1);
    await host.eval('ActiveDocument.Tables(1).Cell(1, 1).Range.Text = "Name"');
    if (String(ActiveDocument.Tables(1).Cell(1, 1).Range.Text).indexOf('Name') < 0) fail('cell text');
    await host.eval('ActiveDocument.Tables(1).Cell(1, 1).Shading.BackgroundPatternColor = 65535');
    same('shade', ActiveDocument.Tables(1).Cell(1, 1).Shading.BackgroundPatternColor, 65535);
    await host.eval('ActiveDocument.Tables(1).Cell(1, 1).Borders(-1).LineStyle = 1');
    same('border', ActiveDocument.Tables(1).Cell(1, 1).Borders(-1).LineStyle, 1);
    await host.eval('ActiveDocument.Tables(1).Rows(1).HeadingFormat = True');
    same('heading', await host.eval('?ActiveDocument.Tables(1).Rows(1).HeadingFormat'), -1);
    await gap('?ActiveDocument.Tables(1).Uniform', 'G-ENGINE');
    await host.eval('ActiveDocument.Tables(1).Rows.Add');
    await gap('ActiveDocument.Tables(1).Rows.Delete', 'missing-live');
    await host.eval('ActiveDocument.Tables(1).ConvertToText');
    same('converted', ActiveDocument.Tables.Count, 0);
    await host.eval('Selection.TypeParagraph');
    await host.eval('Selection.TypeText("field")');
    await host.eval('ActiveDocument.Fields.Add(Selection.Range, 1, "QUOTE hi")');
    if (ActiveDocument.Fields.Count < 1) fail('field missing');
    await gap('?ActiveDocument.Fields(1).Code.Start', 'G-STORY');
    await gap('?ActiveDocument.Fields(1).Locked', 'missing-live');
    await gap('ActiveDocument.Fields(1).DoClick', 'G-HOST');
    await gap('ActiveDocument.Fields.ToggleShowCodes', 'G-ENGINE');

    reset();
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Shape"');
    same('shapes', await host.eval('?ActiveDocument.Shapes.Count'), ActiveDocument.Shapes.Count);
    await gap('ActiveDocument.InlineShapes.AddPicture("picture.png")', 'G-HOST');
    await gap('ActiveDocument.Shapes.AddChart', 'G-ENGINE');
    try { await host.eval('ActiveDocument.Shapes.AddShape(1, 10, 20, 72, 36)'); }
    catch (err) { if (!err || err.code !== 'G-ENGINE') fail('AddShape -> ' + (err && err.code)); }
    if (ActiveDocument.Shapes.Count >= 1) {
      await host.eval('ActiveDocument.Shapes(1).Width = 90');
      same('shape width', ActiveDocument.Shapes(1).Width, 90);
      await gap('?ActiveDocument.Shapes(1).Left', 'missing-live');
      await gap('?ActiveDocument.Shapes(1).OLEFormat', 'G-ENGINE');
      await host.eval('ActiveDocument.Shapes(1).Delete');
    }
    await host.eval('ActiveDocument.TrackRevisions = True');
    await host.eval('ActiveDocument.Paragraphs(1).Range.Text = "Tracked"');
    same('revisions', await host.eval('?ActiveDocument.Revisions.Count'), ActiveDocument.Revisions.Count);
    if (ActiveDocument.Revisions.Count >= 1) await gap('?ActiveDocument.Revisions(1).Author', 'missing-live');
    await host.eval('ActiveDocument.Revisions.AcceptAll');
    await host.eval('ActiveDocument.Revisions.RejectAll');
    const types = L.domIr.types;
    for (const typeName of Object.keys(types)) {
      if (!types[typeName].gap) continue;
      same(typeName + '.Application', await host.eval('?' + typeName + '.Application Is Application'), -1);
      same(typeName + '.Parent', await host.eval('?' + typeName + '.Parent Is Application'), -1);
      await gap('?' + typeName + '.Creator', 'G-ENGINE');
      await gap('?' + typeName + '.Count', types[typeName].gap);
    }
    await gap('Windows.Item(1)', 'G-HOST');
    await gap('ContentControls.Add', 'G-ENGINE');
    if (host.createForm !== null) fail('form slots are set');
    await present();
  } finally {
    host.close();
  }
  return { ok: failures.length === 0, failures: failures };
}
