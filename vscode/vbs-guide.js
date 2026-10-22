/* What an agent needs before it writes VB6 for the open document.
 * Distilled from docs/quire_dom_api_specifications.md, docs/ledger_dom_api_specifications.md
 * and docs/lectern_dom_api_specifications.md. When a row in those inventories changes status,
 * update the matching section here in the same change. This is not the inventory. */
'use strict';

const INSTRUCTIONS = [
  'VibeOffice runs VB6 on the document in the active editor tab. It is not VBScript, and it is not a Microsoft Word, Excel, or PowerPoint VBA project. vbaProject.bin is kept and is not run.',
  'Call vibeoffice_dom_guide before writing a script, then vibeoffice_run_vbs. A module needs Public Sub Main unless you pass entry. Print values with Debug.Print. One Immediate statement uses mode "line" and starts with ? or Debug.Print when you need the value back.',
  'Collections are 1-based. True is -1 and False is 0. Omit optional arguments from the right. MsgBox and InputBox open in the page. There is no UserForm, CreateObject, or Declare.',
  'Error 438 with code G-STORY, G-TEXT, G-ENGINE, or G-HOST is a gap. Do not retry that member. E-APPROVAL means the user denied the call. E-DENIED means the editor always refuses it. HasHostAbility reports what this host implements. A RequestApproval call from this tool asks the user before it runs.',
].join(' ');

const SHARED = [
  'Language. Write a VB6 standard module, not a VBScript file, not JavaScript, and not a VBA project. vbaProject.bin is kept and is not run. The names below are already in the project. Do not call CreateObject("Word.Application"), CreateObject("Excel.Application"), or CreateObject("PowerPoint.Application").',
  'MsgBox and InputBox open in the page and wait. vbOK is 1, vbCancel is 2, vbYes is 6, vbNo is 7. InputBox returns an empty string on Cancel. There is no UserForm, control, or Declare.',
  'Collections start at 1. Paragraphs(1) and Paragraphs.Item(1) are the same call. True is -1 and False is 0.',
  'A member that is not in the inventory fails. Error number 438 and code G-STORY (Quire story index), G-TEXT (Lectern character index), G-ENGINE (no document operation), or G-HOST (this session does not have that host ability) means stop. E-APPROVAL means the user denied a RequestApproval call. E-DENIED means AlwaysDeny. Do not retry the same call.',
  'On this host HasHostAbility is true for Path, Save, SaveAs, Open, Print, FixedFormat, Hyperlink, Picture, FileDialog, Dir, and Shell. Path, Save, Print, and FixedFormat also need the open file to have a path. FullName and Path are that path and its folder. Save writes the open file and asks when an agent runs it. SaveAs and Open take a path and may be outside the workspace, which asks even for an AlwaysAllow call. PrintOut writes a PDF beside the file. ExportAsFixedFormat writes a PDF to the given path and rejects XPS before asking. FollowHyperlink opens http, https, or mailto. AddPicture embeds a picture. GetOpenFilename and FileDialog use the editor\'s file dialog, and cancel returns an empty string. Dir lists the engine\'s virtual names first, then the real folder. Shell and HostShell exist and always return E-DENIED. Application.Dialogs.Item(id).Show opens an in-app window in Quire and Ledger. A file-open id uses the editor\'s dialog.',
  'Edits go through the document and its undo. Debug.Print is how a script returns text. A Function return or a line that starts with ? returns a string, a number, a Boolean, or an array of those. A VB object is not returned.',
  'mode "module" compiles source as the standard module Script. With no entry it runs Public Sub Main. entry names a public procedure on Script, and args is an array of strings, numbers, and Booleans. mode "line" is one Immediate statement, as in the DOM console.',
].join('\n');

const APPS = {
  quire: [
    'Quire (the Word-shaped document in this tab). Entry points are ActiveDocument, Documents, and Selection. Constants use Word numbers (wdAlignParagraphCenter is 1, wdUnderlineSingle is 1). Creator is G-ENGINE.',
    'Read and write with anchors, not a character index. ActiveDocument.Paragraphs(1).Range.Text reads or replaces that paragraph and keeps the paragraph. Selection.Text, Selection.TypeText, Selection.TypeParagraph, InsertBefore, and InsertAfter edit the caret. Selection.Font.Size, Selection.Font.Bold, Selection.Font.Name, and Selection.Font.Color set the font on the selection. Paragraphs(1).Range.Font does the same for that paragraph. Paragraphs(1).Alignment takes wdAlignParagraphLeft 0, Center 1, Right 2, or Justify 3.',
    'Tables: ActiveDocument.Tables(1).Cell(1, 1).Range.Text. Find: Selection.Find.Execute with Word\'s argument order, not an options object. Bookmarks.Add takes a name and an anchor. Sentences, Words, and Characters.Count work. Characters(n) does not.',
    'G-STORY, do not call: Document.Range(Start, End), Range.Start, Range.End, Move, MoveStart, MoveEnd, GoTo, Expand, Selection.MoveLeft, MoveRight, MoveUp, MoveDown, HomeKey, EndKey, and Bookmark.Start or Bookmark.End. ActiveDocument.Range(0, 5) is this gap. A Range is a paragraph, a selection, a cell, a bookmark, or a find hit. Assigning Range.Text does not join the next paragraph the way Word does when the string has no paragraph mark.',
    'G-ENGINE: content controls, custom XML, form fields, mail merge, equations, building blocks, variables, chart objects, SmartArt nodes, OLE, Protect, and UndoClear. G-HOST when the host ability is absent: Windows, the bare name Dialogs, and PrintPreview. InsertFile is AlwaysAllow for txt, html, and docx. HasHostAbility("InsertFile") is true here.',
    'Example. Public Sub Main: Selection.Font.Size = 14: Selection.TypeText "A sentence.": Debug.Print ActiveDocument.Paragraphs(1).Range.Font.Size',
  ].join('\n'),
  ledger: [
    'Ledger (the Excel-shaped workbook in this tab). Entry points are ActiveWorkbook, ThisWorkbook (this workbook), Workbooks, ActiveSheet, ActiveCell, Selection, and WorksheetFunction. Constants use Excel numbers. Creator is G-ENGINE. There is no G-STORY and no G-TEXT: a cell is a row and a column.',
    'Read and write ActiveSheet.Range("A1").Value, Value2, Formula, FormulaR1C1, and NumberFormat. One cell is a scalar, never a 1 by 1 array. A block is a rectangle. In VB that rectangle is a 1-based array. A range larger than 1,000,000 cells rejects Value and Formula with G-ENGINE. A date is an Excel serial, not a VBA Date. An error such as #N/A is a string. An empty cell is Empty. A string Value that starts with = is stored as a formula.',
    'Font.Name, Font.Size, Font.Bold, Font.Italic, Font.Color, Interior.Color, and Borders work. Clear, Merge, Insert, Delete, Sort (three keys), Find, and Replace work. WorksheetFunction.Sum, If, and VLookup, and other names in the formula catalog, are callable. A dotted name is written with an underscore (BINOM.DIST is BINOM_DIST). Application.Calculate, Undo, and Redo work.',
    'G-ENGINE: semiautomatic calculation, SortFields, data bars, colour scales, icon sets, Chart.SeriesCollection, Axes, SetSourceData, an unknown chart type, ListObject.Refresh, PivotTable.Refresh, PivotFields, Validation.Modify, query tables, connections, slicers, sparklines, OLE, and threaded comments. A chart sheet has no Range. Editing a macro sheet or a dialog sheet is G-ENGINE.',
    'G-HOST: Application.Run, SendKeys, Wait, OnTime, windows, and add-ins. Worksheet.SaveAs is SheetSaveAs and writes csv, txt, or one xlsx sheet. HasHostAbility("SheetSaveAs") is true here. Hyperlink.Follow stays G-HOST. Use FollowHyperlink.',
    'Example. Public Sub Main: ActiveSheet.Range("A1").Value = "Title": ActiveSheet.Range("A1").Font.Size = 14: Debug.Print ActiveSheet.Range("A1").Value',
  ].join('\n'),
  lectern: [
    'Lectern (the PowerPoint-shaped presentation in this tab). Entry points are ActivePresentation, Presentations, and ActiveWindow.Selection. Constants use PowerPoint numbers (ppLayoutText is 2, ppLayoutBlank is 12). Creator is G-ENGINE. Positions are points.',
    'Read and write a shape\'s whole text: ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text. Paragraphs inside that string are joined by a carriage return. Font.Name, Font.Size, Font.Bold, Font.Italic, and Font.Color, and ParagraphFormat.Alignment, apply to that text. Slides.Add index, layout inserts a slide. Shapes.AddShape, AddTextbox (horizontal only), and AddTable work for the presets the inventory lists. Left, Top, Width, Height, and Rotation can be set. Characters.Count works.',
    'G-TEXT, do not call: TextRange.Start, Length, Characters(index), Words, Sentences, Lines, Runs, Find, Replace, and the Bound members. Characters(1, 5) is G-TEXT even when Characters.Count is 5.',
    'G-ENGINE: animation TimeLine, Chart and ChartData, SmartArt and Diagram, OLE, media, NotesPage (notes are a string, not a shape page), comments, tags, AddChart, AddSmartArt, an unmapped layout, a non-solid fill, Line.DashStyle, and shapes on the master. G-HOST: View, panes, and NewWindow. Export writes PNG or JPEG of every slide, including hidden ones. HasHostAbility("Export") is true here. There is no Dialogs index. Hyperlink.Follow stays G-HOST. Use FollowHyperlink.',
    'Example. Public Sub Main: ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "Hello": ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Font.Size = 32: Debug.Print ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text',
  ].join('\n'),
};

function forApp(app) {
  const body = APPS[app];
  if (!body) return '';
  return SHARED + '\n\n' + body;
}

module.exports = { INSTRUCTIONS, forApp, APPS };
