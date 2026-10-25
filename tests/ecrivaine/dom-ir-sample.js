/*
Run this in Ecrivaine. It builds the same document as tests/ecrivaine/dom-sample.js,
with every document call going through the DOM IR (L.domIr).

  1. Open Ecrivaine.
  2. Either press F12, or open Help → Ecrivaine DOM Console.
  3. Paste this file.
  4. Run: await buildSampleIr()

The page must have loaded js/dom-ir.js (it is included after js/dom.js).
 */
async function buildSampleIr() {
  const ir = L.domIr.open();

  async function go(req) {
    const res = await ir.exec(req);
    if (!res || res.ok === false) throw new Error((res && res.member ? res.member + ': ' : '') + ((res && res.message) || 'IR call failed'));
    return res;
  }
  const at = (handle, member) => go({ op: 'get', handle: handle, member: member });
  const put = (handle, member, value) => go({ op: 'set', handle: handle, member: member, value: value });
  const run = (handle, member, args) => go({ op: 'call', handle: handle, member: member, args: args || [] });
  const nth = (handle, index) => go({ op: 'item', handle: handle, args: [index] });
  const h = (res) => res.handle;
  const n = (res) => res.value;

  const documents = await at(ir.root.handle, 'Documents');
  await run(h(documents), 'Add');
  const doc = await at(ir.root.handle, 'ActiveDocument');

  const props = await at(h(doc), 'BuiltInDocumentProperties');
  const titleProp = await nth(h(props), 'Title');
  await put(h(titleProp), 'Value', 'DOM Test');

  const paragraphs = await at(h(doc), 'Paragraphs');
  const title = await nth(h(paragraphs), 1);
  await put(h(title), 'Style', 'Heading 1');
  await put(h(title), 'Alignment', wdAlignParagraphCenter);
  const titleRange = await at(h(title), 'Range');
  await put(h(titleRange), 'Text', 'DOM Test 标题');
  const titleFont = await at(h(titleRange), 'Font');
  await put(h(titleFont), 'Color', wdColorDarkBlue);
  await put(h(titleFont), 'Name', 'Times New Roman');
  await put(h(titleFont), 'NameFarEast', '宋体');

  await run(h(titleRange), 'Select');
  const selection = await at(ir.root.handle, 'Selection');
  await run(h(selection), 'Collapse', [wdCollapseEnd]);
  await run(h(selection), 'TypeParagraph');
  const selFont = await at(h(selection), 'Font');
  await put(h(selFont), 'Name', 'Times New Roman');
  await put(h(selFont), 'NameFarEast', '宋体');
  await run(h(selection), 'TypeText', ['Ecrivaine exposes a Word object model. 中文用宋体。']);
  await put(h(selFont), 'Bold', true);
  await run(h(selection), 'TypeText', ['This sentence is bold.']);
  await put(h(selFont), 'Bold', false);
  await run(h(selection), 'TypeParagraph');

  for (let i = 1; i <= 10; i++) {
    await run(h(selection), 'TypeText', ['Hello world ' + i]);
    await run(h(selection), 'TypeParagraph');
  }

  const body = await nth(h(paragraphs), 2);
  await put(h(body), 'LeftIndent', 18);
  await put(h(body), 'SpaceAfter', 8);
  const bodyRange = await at(h(body), 'Range');
  const links = await at(h(bodyRange), 'Hyperlinks');
  await run(h(links), 'Add', [{ handle: h(bodyRange) }, 'https://acane.net/melinaoffice', '', 'Melina Office']);

  await run(h(selection), 'TypeText', ['A numbered item from ListFormat']);
  const selRange = await at(h(selection), 'Range');
  const listFormat = await at(h(selRange), 'ListFormat');
  await run(h(listFormat), 'ApplyNumberDefault');
  await run(h(selection), 'TypeParagraph');
  const cleared = await at(h(selection), 'Range');
  await run(h(await at(h(cleared), 'ListFormat')), 'RemoveNumbers');

  const tables = await at(h(doc), 'Tables');
  const selRange2 = await at(h(selection), 'Range');
  const table = await run(h(tables), 'Add', [{ handle: h(selRange2) }, 2, 2]);
  const tail = await at(h(paragraphs), 'Last');
  await run(h(await at(h(await at(h(tail), 'Range')), 'ListFormat')), 'ApplyNumberDefault');
  async function cellText(row, column, text) {
    const cell = await run(h(table), 'Cell', [row, column]);
    const range = await at(h(cell), 'Range');
    await put(h(range), 'Text', text);
    return range;
  }
  await cellText(1, 1, 'Name');
  await cellText(1, 2, 'Value');
  await cellText(2, 1, 'DOM');
  await cellText(2, 2, 'Word object model 对象模型');
  const rows = await at(h(table), 'Rows');
  const header = await nth(h(rows), 1);
  await put(h(header), 'HeadingFormat', true);
  for (let r = 1; r <= 2; r++) {
    for (let c = 1; c <= 2; c++) {
      const cell = await run(h(table), 'Cell', [r, c]);
      const range = await at(h(cell), 'Range');
      const font = await at(h(range), 'Font');
      await put(h(font), 'Name', 'Times New Roman');
      await put(h(font), 'NameFarEast', '宋体');
      if (r === 1) {
        const shading = await at(h(cell), 'Shading');
        await put(h(shading), 'BackgroundPatternColor', wdColorYellow);
      }
    }
  }

  const first = await nth(h(paragraphs), 1);
  const firstRange = await at(h(first), 'Range');
  const bookmarks = await at(h(doc), 'Bookmarks');
  const sample = await run(h(bookmarks), 'Add', ['Sample', { handle: h(firstRange) }]);
  const find = await at(h(selection), 'Find');
  const found = await run(h(find), 'Execute', ['Hello world 1', true, false, false, false, false, true, wdFindContinue, false, 'Hi world 1', wdReplaceOne]);
  const comments = await at(h(doc), 'Comments');
  const noteRange = await at(h(await nth(h(paragraphs), 1)), 'Range');
  await run(h(comments), 'Add', [{ handle: h(noteRange) }, 'Sample note']);

  const report = {
    paragraphs: n(await at(h(paragraphs), 'Count')),
    title: n(await at(h(await at(h(await nth(h(paragraphs), 1)), 'Range')), 'Text')),
    documentTitle: n(await at(h(titleProp), 'Value')),
    words: n(await run(h(doc), 'ComputeStatistics', [wdStatisticWords])),
    indent: n(await at(h(body), 'LeftIndent')),
    rows: n(await at(h(rows), 'Count')),
    columns: n(await at(h(await at(h(table), 'Columns')), 'Count')),
    name: n(await at(h(await at(h(await run(h(table), 'Cell', [1, 1])), 'Range')), 'Text')),
    value: n(await at(h(await at(h(await run(h(table), 'Cell', [2, 2])), 'Range')), 'Text')),
    bookmark: n(await at(h(sample), 'Name')),
    bookmarkText: n(await at(h(await at(h(await nth(h(bookmarks), 'Sample')), 'Range')), 'Text')),
    comments: n(await at(h(comments), 'Count')),
    found: n(found),
  };
  console.log(report);
  return report;
}
