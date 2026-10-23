# Quire Word object model

JavaScript equivalents of the Microsoft Word VBA object model, for the open Quire document. Names, nesting, and parameter order follow the Word VBA reference (Microsoft Learn, Word VBA). This file is the inventory. The rows are provided by `public/quire/js/dom.js`. A member marked Not implemented throws, with `Error.code` set to `G-STORY`, `G-ENGINE`, or `G-HOST`.

Quire is not Microsoft Word. `Creator` is not reported as Word's creator code.

## Rules

- Collections are 1-based. `Paragraphs(1)` and `Paragraphs.Item(1)` are the same call. `Count`, `First`, and `Last` are properties.
- Members use the VBA spelling. Parameters stay in VBA order. Optional parameters may be omitted from the right. Callers do not pass an options object in their place.
- A property may be read and assigned: `Range.Text = "Hello"`.
- Constants use Word's numeric values (`wdAlignParagraphCenter` is `1`). They are defined beside the API.
- A returned `Range` is an anchor to content Quire already addresses (a paragraph, a selection, a cell, a bookmark, a find hit). It is not a story character index. `Start` and `End` are not part of that anchor.
- Edits go through the existing document operations (`js/ops.js`, the editor, tables, fields, notes, review, find, drawing, and the document session). Undo, redraw, and save stay on those paths.
- The script host is Help ▸ Quire DOM Console (`public/quire/js/dom-console.js`). A snippet runs in the page, and `console.log` prints in that window. The object model itself is this inventory.
- Calls from another language go through the DOM IR (`js/dom-ir.js`, [quire_dom_ir_specifications.md](quire_dom_ir_specifications.md)). The IR forwards this inventory. It does not add members and it does not lay out the document.

Shared members `Application` and `Parent` exist on every object below and are **Implemented** (navigation only). They are not repeated on each row. `Creator` is **Not implemented** on every object (`G-ENGINE`: Quire does not report Word's creator code).

## Status and gaps

| Status | Meaning |
|---|---|
| Implemented | The member is provided by `public/quire/js/dom.js`. A behavioral difference is written in the Gap column. |
| Not implemented | The member throws. The Gap column gives the code: G-STORY, G-ENGINE, or G-HOST. |

| Code | Why it is not wrapped |
|---|---|
| G-STORY | The member takes or returns a Word story character index (`Long`). Word's index counts a paragraph mark, and inside a table it also counts a cell mark and an end-of-row mark. Quire stores paragraphs and tables as blocks. `js/find.js` builds a search string (each paragraph's text plus a CR, table cells in reading order) and maps it back to Quire positions. That string is not Word's index, so a `Long` copied from VBA addresses different content as soon as a table, a cell mark, or a drawing is involved. |
| G-ENGINE | The document has no operation that does this. A wrapper would have nothing to call. |
| G-HOST | This session does not have that host ability. On the VS Code host the same member may run. The tiers are in [docs/dom_api_security.md](dom_api_security.md). |

**Update rule.** Any later change that adds a member, or that closes a gap, updates this table in the same change. A row that still says Not implemented is not a working member. A new member that is not listed here is not added in code; add the row first, with its status and gap. The same change updates the task-pane reference in `public/quire/js/vb-dom-ref.js`. That page is what a script sees. A member listed only here, or only there, is out of date.

## Story index (not implemented)

These members are **Not implemented**, gap **G-STORY**. They are listed once. Objects below point here instead of repeating them.

| Member | VBA signature |
|---|---|
| Document.Range | `Range(Start, End)` |
| Range.Start, Range.End | properties, `Long` |
| Range.Move | `Move(Unit, Count)` |
| Range.MoveStart, Range.MoveEnd | `MoveStart(Unit, Count)`, `MoveEnd(Unit, Count)` |
| Range.MoveWhile, Range.MoveUntil | `MoveWhile(Cset, Count)`, `MoveUntil(Cset, Count)` |
| Range.MoveStartWhile, Range.MoveStartUntil, Range.MoveEndWhile, Range.MoveEndUntil | same parameter names as VBA |
| Range.StartOf, Range.EndOf | `StartOf(Unit, Extend)`, `EndOf(Unit, Extend)` |
| Range.Expand | `Expand(Unit)` |
| Range.Information | `Information(Type)` |
| Range.GoTo, Range.GoToNext, Range.GoToPrevious | `GoTo(What, Which, Count, Name)`, `GoToNext(What)`, `GoToPrevious(What)` |
| Range.Next, Range.Previous | `Next(Unit, Count)`, `Previous(Unit, Count)` |
| Selection.Start, Selection.End | properties, `Long` |
| Selection.MoveLeft, Selection.MoveRight, Selection.MoveUp, Selection.MoveDown | `MoveLeft(Unit, Count, Extend)` and the same shape for Right, Up, Down |
| Selection.HomeKey, Selection.EndKey | `HomeKey(Unit, Extend)`, `EndKey(Unit, Extend)` |
| Selection.GoTo, Selection.GoToNext, Selection.GoToPrevious | same signatures as Range |
| Characters | `Characters(Index)`, `Characters.Item(Index)` |
| Bookmark.Start, Bookmark.End, Field.Code as a story range, Comment.Scope as a story range, Hyperlink.Range.Start, Footnote.Reference.Start | any `Long` story offset on these objects |

`Document.Content`, `Paragraph.Range`, `Cell.Range`, `Selection.Range`, and the other anchored ranges below do not grow a `Start` or `End`.

`Characters.Count` is separate and is **Implemented** (it counts text Quire already measures). `Characters(Index)` is not.

## Objects with no document operation

Every member of these objects is **Not implemented**. A later implementation adds a row of its own and stops inheriting this note.

| Object | Gap |
|---|---|
| ContentControls, ContentControl | G-ENGINE. Content controls are kept in the package when present. There is no editor for adding, deleting, or rewriting them. |
| XMLNodes, XMLNode, CustomXMLParts, CustomXMLPart | G-ENGINE. No custom-XML edit API. |
| FormFields, FormField, CheckBox, DropDown, TextInput | G-ENGINE. No legacy form-field editor. |
| Frames, Frame | G-ENGINE. Frame properties are stored on a paragraph. There is no Frames collection. |
| Subdocuments | G-ENGINE. No master-document editor. |
| MailMerge, MailMergeFields, MailMergeDataSource | G-ENGINE. No mail-merge data source. The mail-merge UI is not an object model. |
| OMaths, OMath | G-ENGINE. Equations are edited in the equation dialog. There is no Office Math object model. |
| BuildingBlocks, BuildingBlock | G-ENGINE. |
| Variables, Variable | G-ENGINE. No document-variables collection. |
| Windows, Window, Panes, Pane, View | G-HOST. Quire's view state is not the Word window model. |
| Dialogs, Dialog, FileDialog, System, Tasks, CommandBars, KeyBindings, Languages | G-HOST. |
| Chart, ChartData (Excel chart object model) | G-ENGINE. A chart can be inserted (see Shapes). The Excel chart object is not in Quire. |
| Diagram, DiagramNode (IntelliArt / SmartArt object model) | G-ENGINE. Inserting a diagram opens a dialog. There is no diagram-node API. |
| OLEFormat, OLEControl | G-ENGINE. |
| Tables of contents, indexes, and tables of figures as their own objects (`TablesOfContents`, `Indexes`, `TablesOfFigures`) | G-ENGINE as objects. Building them is a field (`Fields`), which is listed below. |
| AutoCorrect, Options, DefaultWebOptions | G-HOST. Application options, not the open document. |

## Application

Entry points only. Other `Application` members are in the host list above.

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| ActiveDocument | Property `Document` | Implemented | The open document in the current session. |
| Documents | Property `Documents` | Implemented | |
| Selection | Property `Selection` | Implemented | Empty when there is no caret. |
| HasHostAbility | `HasHostAbility(Ability)` | Implemented | True (`-1`) when this host implements the ability. Unknown names are False. See [docs/dom_api_security.md](dom_api_security.md). |
| MsgBox | `MsgBox(Prompt, Buttons, Title)` | Implemented | An in-app message. OK, Cancel, Yes, and No return 1, 2, 6, and 7. AlwaysAllow. |
| InputBox | `InputBox(Prompt, Title, Default)` | Implemented | An in-app prompt. Cancel returns an empty string. AlwaysAllow. |
| Dialogs | Property `Dialogs` | Implemented | `Dialogs.Item(id).Show` opens the matching in-app window, or a VS Code file dialog for open and save. An unmapped id is `G-HOST`. The bare name `Dialogs` remains a `G-HOST` gap object. |
| FileDialog | `FileDialog(Type)` | Implemented | VS Code's file or folder dialog. `Show` returns `-1` when a path was chosen. The bare name `FileDialog` remains a `G-HOST` gap object. |
| GetOpenFilename, GetSaveAsFilename | `GetOpenFilename(FileFilter, FilterIndex, Title, ButtonText, MultiSelect)`, `GetSaveAsFilename(InitialFilename, FileFilter, FilterIndex, Title, ButtonText)` | Implemented | The same dialogs. Cancel returns `""`. `MultiSelect` returns a string array. `GetSaveAsFilename` accepts the names and does not apply them to the dialog. |
| HostDir | `HostDir(PathName)` | Implemented | `Dir` in a script. Virtual files first, then the real folder (`RequestApproval`). |
| HostShell | `HostShell(PathName)` | Implemented | `Shell` in a script. AlwaysDeny (`E-DENIED`) on the VS Code host. `G-HOST` elsewhere. |

## Documents

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Count | Property | Implemented | |
| Item | `Item(Index)` | Implemented | |
| Add | `Add(Template, NewTemplate, DocumentType, Visible)` | Implemented | Only the call with no arguments, which creates a blank document. `Template`, `NewTemplate`, `DocumentType`, and `Visible` are G-ENGINE / G-HOST. |
| Open | `Open(FileName, ...)` | Implemented on the VS Code host | RequestApproval. Opens that file in its own tab and does not return it. Any argument after `FileName` is an error. Elsewhere `G-HOST`. |
| Close | `Close(SaveChanges, OriginalFormat, RouteDocument)` | Not implemented | G-HOST at collection scope. Close the `Document` instead. |
| Save | `Save(NoPrompt, OriginalFormat)` | Not implemented | G-HOST at collection scope. Save the `Document` instead. |

## Document

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Name | Property, read | Implemented | File name without a directory. |
| FullName | Property, read | Implemented on the VS Code host | AlwaysAllow. The open file's path. Elsewhere `G-HOST`. |
| Path | Property, read | Implemented on the VS Code host | AlwaysAllow. The folder that contains it. Elsewhere `G-HOST`. |
| Saved | Property | Implemented | `False` when the document is dirty. Assigning `True` clears the dirty flag and does not write the file. |
| Content | Property `Range` | Implemented | Anchor over the main story. No `Start` / `End` (G-STORY). |
| Paragraphs | Property | Implemented | |
| Sentences | Property | Implemented | Sentence breaks follow `O.sentenceExtent`. No story `Long`. |
| Words | Property | Implemented | Word breaks follow `O.wordExtent`. No story `Long`. |
| Characters | Property | Not implemented | G-STORY, except `Characters.Count` under Range. |
| Tables | Property | Implemented | |
| Sections | Property | Implemented | |
| StoryRanges | `StoryRanges(Index)` | Implemented | `Index` is a `WdStoryType` for a story Quire stores: main, footnotes, endnotes, comments, text boxes, and the header and footer stories that exist on the document. The result is an anchor range. Odd/even/first headers that the document does not have are absent, not empty Word stories. |
| Styles | Property | Implemented | |
| Bookmarks | Property | Implemented | |
| Fields | Property | Implemented | |
| Hyperlinks | Property | Implemented | |
| Comments | Property | Implemented | |
| Footnotes, Endnotes | Property | Implemented | |
| InlineShapes, Shapes | Property | Implemented | |
| Lists, ListParagraphs | Property | Implemented | Numbering Quire already computes. |
| Revisions | Property | Implemented | |
| BuiltInDocumentProperties | `BuiltInDocumentProperties(Index)` | Implemented | String core properties Quire stores: Title, Subject, Author, Keywords, Comments, Category, Manager, Company. Other built-in property ids are G-ENGINE. |
| CustomDocumentProperties | `CustomDocumentProperties(Index)` or by name | Implemented | String values in the custom-property map. Typed values other than string are G-ENGINE. |
| Undo | `Undo(Times)` | Implemented | |
| Redo | `Redo(Times)` | Implemented | |
| UndoClear | `UndoClear()` | Not implemented | G-ENGINE. There is no operation that drops the undo stack alone. |
| Range | `Range(Start, End)` | Not implemented | G-STORY. |
| Select | `Select()` | Implemented | Selects the main story (same as select-all). |
| Activate | `Activate()` | Implemented | Switches the session to this document. |
| Close | `Close(SaveChanges, OriginalFormat, RouteDocument)` | Implemented | `SaveChanges` accepts `wdDoNotSaveChanges`, `wdSaveChanges`, and `wdPromptToSaveChanges`. `OriginalFormat` and `RouteDocument` are G-HOST. |
| Save | `Save()` | Implemented | Writes the current file when it already has a destination. On the VS Code host an agent script is RequestApproval. A file that has never been saved stays `G-HOST`. |
| SaveAs, SaveAs2 | `SaveAs(FileName, FileFormat, ...)` | Implemented on the VS Code host | RequestApproval. Writes that path and opens it in another tab. A password or an unsupported `FileFormat` is an error and does not write. Elsewhere `G-HOST`. |
| ComputeStatistics | `ComputeStatistics(Statistic, IncludeFootnotesAndEndnotes)` | Implemented | `wdStatisticWords`, `wdStatisticCharacters`, `wdStatisticCharactersWithSpaces`, `wdStatisticParagraphs`, `wdStatisticPages`. `wdStatisticLines` and `wdStatisticLines` of a non-laid-out story are G-ENGINE. Pages use the layout already computed for the open document. |
| PrintOut | `PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)` | Implemented on the VS Code host | AlwaysAllow. A PDF of the whole document beside the file, named by dropping the extension and adding `.pdf`. Every argument is accepted and ignored. |
| PrintPreview | | Not implemented | G-HOST. |
| Protect, Unprotect | `Protect(Type, NoReset, Password, UseIRM, EnforceStyleLock)`, `Unprotect(Password)` | Not implemented | G-ENGINE. Permission in the UI is not the Word protection types. |
| Compare | `Compare(Name, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles, RemovePersonalInformation, RemoveDateAndTime)` | Not implemented | G-HOST. Compare-with is a file picker. The names bind, then the call returns G-HOST. |
| FollowHyperlink | `FollowHyperlink(Address, ...)` | Implemented on the VS Code host | AlwaysAllow. `http`, `https`, and `mailto` only. Any later argument is an error. VS Code asks before an external link. |
| GoTo | `GoTo(What, Which, Count, Name)` | Not implemented | G-STORY when `What` is a character, line, or page offset. Bookmark and comment by name are on those collections. |
| ExportAsFixedFormat | `ExportAsFixedFormat(OutputFileFormat, OutputFileName, ...)` | Implemented on the VS Code host | RequestApproval. PDF, or an omitted format, to that path. XPS (`wdExportFormatXPS`) errors before the confirmation and before any write. |

## Selection

The caret or the highlighted span in the active document.

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Text | Property | Implemented | Read returns the selected text. Assigning replaces the selection. A collapsed selection inserts. |
| TypeText | `TypeText(Text)` | Implemented | |
| TypeParagraph | `TypeParagraph()` | Implemented | |
| TypeBackspace | `TypeBackspace()` | Implemented | |
| Delete | `Delete(Unit, Count)` | Implemented | Only the call with no arguments (deletes the selection, or one character when collapsed, matching the Delete key). `Unit` and `Count` are G-STORY. |
| InsertBefore, InsertAfter | `InsertBefore(Text)`, `InsertAfter(Text)` | Implemented | |
| Font | Property | Implemented | |
| ParagraphFormat | Property | Implemented | |
| Style | Property | Implemented | Paragraph style on the selected paragraphs. A character style is G-ENGINE on this property (use `Font` / run formatting). |
| Range | Property `Range` | Implemented | Anchor. No `Start` / `End`. |
| Document | Property | Implemented | The active document. |
| Bold, Italic, Underline | Properties | Implemented | The same shortcuts as `Range`. |
| Paragraphs, Sentences, Words, Tables | Properties | Implemented | The same collections as `Range`. `Tables.Add` inserts at the given range, or at the selection when `Range` is omitted. |
| Bookmarks, Fields, Comments, Hyperlinks, InlineShapes | Properties | Implemented | The same collections as `Range`, including `Add` and `AddPicture`. |
| ListFormat, Case, LanguageID, LanguageIDFarEast | | Implemented | The same members as `Range`. |
| FormattedText | Property | Not implemented | G-ENGINE, the same as `Range`. |
| InsertParagraph, InsertParagraphBefore, InsertParagraphAfter | methods | Implemented | The same calls as `Range`. |
| Collapse | `Collapse(Direction)` | Implemented | `wdCollapseStart` or `wdCollapseEnd`. Moves the anchor to a caret. |
| Select | `Select()` | Implemented | |
| WholeStory | `WholeStory()` | Implemented | |
| Copy, Cut, Paste | `Copy()`, `Cut()`, `Paste()` | Implemented | Quire's own clipboard. `PasteAndFormat(Type)`, `PasteSpecial(IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel)`, and `PasteExcelTable(LinkedToExcel, WordFormatting, RTF)` are G-ENGINE. |
| Find | Property `Find` | Implemented | |
| Type | Property, read | Implemented | `wdSelectionNormal`, `wdSelectionIP`, `wdNoSelection`. Other selection types (frame, shape row) are reported as `wdSelectionNormal` when an object is selected (gap: no separate shape-selection type). |
| MoveLeft, MoveRight, MoveUp, MoveDown, HomeKey, EndKey, Start, End, GoTo | | Not implemented | G-STORY. |
| Information | `Information(Type)` | Not implemented | G-STORY. |
| InsertBreak | `InsertBreak(Type)` | Implemented | Page, column, and text-wrapping breaks that the break dialog already inserts. Section-break types the model can store are included. Types it cannot store are G-ENGINE. |
| InsertFile | `InsertFile(FileName, Range, ConfirmConversions, Link, Attachment)` | Implemented on the VS Code host | AlwaysAllow for `.txt`, `.htm`, `.html`, `.docx`, `.docm`, and `.dotx`. A path outside the workspace asks first. `Range`, `ConfirmConversions`, `Link`, and `Attachment` are errors. |
| InsertSymbol | `InsertSymbol(CharacterNumber, Font, Unicode, Bias)` | Implemented | Inserts one character in the given font, the same way Insert Symbol does. |
| InsertDateTime | `InsertDateTime(DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType)` | Implemented | `InsertAsField` True inserts a date field. Calendar and full-width flags are G-ENGINE. |
| InsertCaption | `InsertCaption(Label, Title, TitleAutoText, Position, ExcludeLabel)` | Not implemented | G-ENGINE. The caption dialog is separate from this call. |
| InsertCells, InsertRows, InsertColumns, and the matching Delete methods | the VBA signatures | Implemented | Forward to the table operations below when the selection is in a table. Outside a table they throw. |
| ConvertToTable | `ConvertToTable(Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior)` | Implemented | Separator and the resulting table. The `Apply*` autoformat flags and `Format` are G-ENGINE. |

## Range

Anchored span. Story-index members are in the story-index section.

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Text | Property | Implemented | Get returns the span's text. Set replaces that span. On a whole-paragraph range, the paragraph is kept: the new text replaces its runs. Word joins with the next paragraph when the assigned string has no trailing paragraph mark and the range included that mark. That join is not done. Setting `Text`, a paragraph's `Style`, or its `Alignment` does not move the Selection. |
| InsertBefore | `InsertBefore(Text)` | Implemented | |
| InsertAfter | `InsertAfter(Text)` | Implemented | |
| InsertParagraph | `InsertParagraph()` | Implemented | |
| InsertParagraphBefore | `InsertParagraphBefore()` | Implemented | |
| InsertParagraphAfter | `InsertParagraphAfter()` | Implemented | |
| Delete | `Delete()` | Implemented | No `Unit` / `Count` (G-STORY). |
| Select | `Select()` | Implemented | |
| Collapse | `Collapse(Direction)` | Implemented | Collapses this range. The Selection stays where it is. |
| Duplicate | `Duplicate()` | Implemented | A second anchor to the same span. |
| Copy, Cut, Paste | `Copy()`, `Cut()`, `Paste()` | Implemented | Same clipboard limit as Selection. |
| Font | Property | Implemented | |
| ParagraphFormat | Property | Implemented | |
| Style | Property | Implemented | Paragraph style. |
| Bold, Italic, Underline | Properties | Implemented | Shortcuts to `Font`. Underline uses `WdUnderline`. |
| Paragraphs | Property | Implemented | Paragraphs this anchor touches. |
| Sentences, Words | Property | Implemented | Inside this anchor, using Quire's breaks. |
| Tables | Property | Implemented | Tables this anchor touches. `Add` inserts at the given range, or at this anchor when `Range` is omitted. |
| Find | Property | Implemented | Limited to this anchor's story. |
| Characters | `Count` only | Implemented | `Characters(Index)` is G-STORY. |
| Case | `Case` property (`WdCharacterCase`) | Implemented | The case changes `O.changeCase` already applies. |
| ListFormat | Property | Implemented | |
| Hyperlinks, Bookmarks, Fields, Comments, InlineShapes | Properties | Implemented | Those that lie in the anchor. |
| FormattedText | Property | Not implemented | G-ENGINE. No rich-text value type separate from the document. |
| LanguageID | Property | Implemented | The run language Quire stores (`lang`). `LanguageIDFarEast` maps to `langEA`. |

## Paragraphs and Paragraph

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Paragraphs.Count, Item, First, Last | `Item(Index)` | Implemented | Reading order in the anchor's story, including paragraphs inside tables. |
| Paragraph.Range | Property | Implemented | That paragraph's text. No story `Long`. Text-set gap is the Range.Text gap. |
| Paragraph.Style | Property | Implemented | |
| Paragraph.Format | Property `ParagraphFormat` | Implemented | |
| Paragraph.LeftIndent, RightIndent, FirstLineIndent | Properties, points | Implemented | |
| Paragraph.SpaceBefore, SpaceAfter, LineSpacing | Properties | Implemented | |
| Paragraph.Alignment | Property `WdParagraphAlignment` | Implemented | |
| Paragraph.OutlineLevel | Property `WdOutlineLevel` | Implemented | |
| Paragraph.KeepTogether, KeepWithNext, PageBreakBefore, WidowControl | Properties | Implemented | |
| Paragraph.TabStops | Property | Implemented | |
| Paragraph.Next, Paragraph.Previous | `Next(Count)`, `Previous(Count)` | Implemented | Neighbouring paragraphs in reading order. Not story units. |
| Paragraph.IndentCharWidth, Space1, Space15, Space2 | methods | Implemented | Same effects as the indent and line-spacing commands. |
| Paragraph.OutlineDemote, OutlinePromote, OutlineDemoteToBody | | Not implemented | G-ENGINE. Outline view's promote/demote is not a separate stored operation from list level. List level is `ListFormat`. |
| Paragraph.DropCap | | Not implemented | G-ENGINE as an object. Drop-cap data is stored; there is no DropCap object to wrap. |

## Sentences and Words

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Count, Item, First, Last | `Item(Index)` | Implemented | Breaks from `O.sentenceExtent` and `O.wordExtent`. |
| Item.Text and the returned Range's font and delete | | Implemented | The item is an anchor range. No `Start` / `End`. |

## Tables, rows, columns, cells

Wrapped from `js/tables.js`. A `Cell.Range` is the cell's paragraphs, not a story `Long`.

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Tables.Count, Item | `Item(Index)` | Implemented | Document order. |
| Tables.Add | `Add(Range, NumRows, NumColumns, DefaultTableBehavior, AutoFitBehavior)` | Implemented | Inserts at the anchor. `DefaultTableBehavior` and `AutoFitBehavior` are G-ENGINE; the table uses Quire's current insert. When that paragraph is in a list, every cell paragraph joins the same list and the numbers continue through the table. |
| Table.Cell | `Cell(Row, Column)` | Implemented | 1-based. A merged cell is returned for each grid square it covers. |
| Table.Rows, Columns, Range | Properties | Implemented | `Range` is the whole table's cell contents. |
| Table.Style | Property | Implemented | Table styles Quire can apply. |
| Table.Select, Delete | `Select()`, `Delete()` | Implemented | |
| Table.Split | `Split(BeforeRow)` | Implemented | |
| Rows.Add, Rows.Item, Rows.Count | `Add(BeforeRow)` | Implemented | |
| Rows.Delete, Row.Delete, Row.Select, Row.Range, Row.Cells | | Implemented | |
| Row.HeadingFormat | Property | Implemented | Header-row flag. |
| Columns.Add, Columns.Item, Columns.Count, Column.Delete, Column.Select, Column.Width, Column.Cells | `Add(BeforeColumn)`, `Width` in points | Implemented | |
| Cells.Count, Item, Merge, Split, DistributeWidth, DistributeHeight | `Merge()`, `Split(NumRows, NumColumns, MergeBeforeSplit)` | Implemented | |
| Cell.Range, Cell.RowIndex, Cell.ColumnIndex | | Implemented | |
| Cell.Width, Cell.Height | Properties, points | Implemented | |
| Shading.BackgroundPatternColor, Shading.ForegroundPatternColor, Shading.Texture | | Implemented | Values Quire's cell and paragraph shading already store. |
| Borders | `Borders(Index)` for the edges the border command sets | Implemented | Line style, width, and color for edges Quire stores. `Borders.Apply` page-border modes are G-ENGINE. |
| Table.AutoFitBehavior | `AutoFitBehavior(Behavior)` | Implemented | The autofit modes `T.autofit` already has. |
| Table.Sort | `Sort(...)` | Implemented | The sort keys `T.sortTable` already takes. Extra VBA sort flags with no table-sort option are G-ENGINE. |
| Table.ConvertToText | `ConvertToText(Separator, NestedTables)` | Implemented | Nested-table handling is G-ENGINE. |
| Table.ApplyStyleHeadingRows and the other look flags | Properties | Implemented | The look flags stored on the table. |
| Table.Uniform, Table.Spacing | | Not implemented | G-ENGINE. |

## Sections, page setup, headers, footers

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Sections.Count, Item | `Item(Index)` | Implemented | |
| Section.Range | Property | Implemented | Paragraphs in that section. No story `Long`. |
| PageSetup.PageWidth, PageHeight, Orientation | Properties | Implemented | Points. Orientation is `wdOrientPortrait` or `wdOrientLandscape`. |
| PageSetup.TopMargin, BottomMargin, LeftMargin, RightMargin, Gutter, HeaderDistance, FooterDistance | Properties, points | Implemented | |
| PageSetup.TextColumns | Property | Implemented | Count, spacing, and equal-width. Preset line-between is included when stored. |
| PageSetup.PaperSize | Property | Not implemented | G-ENGINE as a `WdPaperSize` code. Width and height are set directly. |
| Headers, Footers | `Headers(Index)`, `Footers(Index)` | Implemented | `Index` is `wdHeaderFooterPrimary`, `wdHeaderFooterFirstPage`, or `wdHeaderFooterEvenPages`, when that story exists. |
| HeaderFooter.Range | Property | Implemented | Text of that header or footer story. |
| HeaderFooter.Exists | Property, read | Implemented | |
| HeaderFooter.LinkToPrevious | Property | Not implemented | G-ENGINE. No link-to-previous flag is edited. |
| Section.ProtectedForForms | | Not implemented | G-ENGINE. |

## Styles and Style

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Styles.Count, Item | `Item(Index)` by number or name | Implemented | |
| Styles.Add | `Add(Name, Type)` | Implemented | Paragraph and character styles. |
| Style.NameLocal | Property | Implemented | Quire has one name, not a localized alias. |
| Style.Type | Property, read | Implemented | |
| Style.Font, Style.ParagraphFormat | Properties | Implemented | |
| Style.Delete | `Delete()` | Implemented | Built-in styles that Quire will not delete throw, as the style pane does. |
| Style.InUse | Property, read | Implemented | |
| Style.BaseStyle, Style.NextParagraphStyle | Properties | Implemented | |
| Style.LinkStyle, Style.Visibility, Style.UnhideWhenUsed, Style.QuickStyle, Style.Priority, Style.Locked | | Not implemented | G-ENGINE. |

## Font

Properties Quire stores on a run. Assignment formats the anchor.

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Name, NameAscii, NameFarEast, NameBi | Properties | Implemented | `Name` writes the Latin face. `NameFarEast` writes the East Asian face. `NameBi` writes the complex-script face when stored. |
| Size, SizeBi | Properties, points | Implemented | |
| Bold, Italic, BoldBi, ItalicBi | Properties | Implemented | |
| Underline | Property `WdUnderline` | Implemented | Styles Quire stores. Color of the underline is `UnderlineColor`. |
| StrikeThrough, DoubleStrikeThrough | Properties | Implemented | |
| AllCaps, SmallCaps, Hidden | Properties | Implemented | `Hidden` is the vanish flag. |
| Superscript, Subscript | Properties | Implemented | |
| Color | Property | Implemented | Word BGR `Long`, plus `wdColorAutomatic`. |
| HighlightColorIndex | Property `WdColorIndex` | Implemented | The highlight names Quire stores. |
| Spacing, Position, Scaling, Kerning | Properties | Implemented | Spacing and position in points. Scaling is the character-width percent. |
| Outline, Shadow, Emboss, Engrave | Properties | Implemented | Stored flags. They are not a separate drawing effect. |
| EmphasisMark | Property | Implemented | Values Quire stores on `em`. |
| Borders, Shading | Properties | Implemented | The run border and shading already stored. |
| Animation | Property | Not implemented | G-ENGINE. |
| DiacriticColor, StylisticSet, NumberSpacing, NumberForm, Ligatures, ContextualAlternates | | Not implemented | G-ENGINE. |

## ParagraphFormat

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Alignment | Property `WdParagraphAlignment` | Implemented | `left`, `center`, `right`, `both`, `distribute`. |
| LeftIndent, RightIndent, FirstLineIndent, CharacterUnitLeftIndent, CharacterUnitRightIndent, CharacterUnitFirstLineIndent | Properties | Implemented | Point indents. Character-unit indents are G-ENGINE (Quire stores points). |
| SpaceBefore, SpaceAfter, LineSpacing, LineSpacingRule | Properties | Implemented | Rule is `wdLineSpaceSingle`, `wdLineSpace1pt5`, `wdLineSpaceDouble`, `wdLineSpaceAtLeast`, `wdLineSpaceExactly`, `wdLineSpaceMultiple`, matching the spacing Quire stores. |
| SpaceBeforeAuto, SpaceAfterAuto | Properties | Implemented | |
| KeepTogether, KeepWithNext, PageBreakBefore, WidowControl | Properties | Implemented | |
| OutlineLevel | Property | Implemented | |
| Hyphenation | Property | Implemented | The suppress-auto-hyphens flag, inverted to match VBA's Hyphenation meaning on that paragraph. |
| TabStops | Property | Implemented | |
| Borders, Shading | Properties | Implemented | Paragraph borders and shading already stored. |
| ReadingOrder | Property | Implemented | The bidi flag. |
| MirrorIndents, AutoAdjustRightIndent, DisableLineHeightGrid, FarEastLineBreakControl, WordWrap, HangingPunctuation, HalfWidthPunctuationOnTopOfLine, AddSpaceBetweenFarEastAndAlpha, AddSpaceBetweenFarEastAndDigit | | Not implemented | G-ENGINE unless the flag is the one already stored (`noSnap` is `DisableLineHeightGrid` and is Implemented). |
| TabStops.Add | `Add(Position, Alignment, Leader)` | Implemented | Position in points. Alignment and leader values Quire stores. |
| TabStops.Item, Count, ClearAll | `Item(Index)`, `ClearAll()` | Implemented | |
| TabStop.Position, Alignment, Leader, Clear | | Implemented | |

## ListFormat and Lists

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| ListFormat.ListType | Property, read | Implemented | |
| ListFormat.ListLevelNumber | Property | Implemented | 1-based in VBA. Quire stores 0-based levels; the wrapper converts. |
| ListFormat.ApplyBulletDefault, ApplyNumberDefault | methods | Implemented | The default bullet and number lists. |
| ListFormat.ApplyListTemplate | `ApplyListTemplate(ListTemplate, ContinuePreviousList, ApplyTo, DefaultListBehavior)` | Not implemented | G-ENGINE. There is no `ListTemplate` gallery object. Use the bullet and number defaults, or the list level. |
| ListFormat.RemoveNumbers | `RemoveNumbers(NumberType)` | Implemented | |
| ListFormat.ListIndent, ListOutdent | methods | Implemented | |
| ListFormat.CanContinuePreviousList | | Not implemented | G-ENGINE. |
| Lists.Count, Item | | Implemented | Lists `D.computeLists` already distinguishes. |
| List.Range, List.ListParagraphs, List.CountNumberedItems | `CountNumberedItems(NumberType, Level)` | Implemented | `CountNumberedItems` is not on the live list object, so the call is `missing-live`. |
| ListTemplates, ListTemplate, ListLevels, ListLevel | | Not implemented | G-ENGINE as a gallery. Level format is stored inside numbering definitions and is not this object. |

## Find and Replacement

`Find` is the object on a `Range` or `Selection`. Arguments follow VBA order. Arguments the finder already honors are implemented; the rest are accepted and ignored only when noted.

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Execute | `Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl)` | Implemented | `MatchKashida`, `MatchDiacritics`, `MatchAlefHamza`, and `MatchControl` are G-ENGINE (no such match). `Forward` False searches upward. `Wrap` uses `wdFindStop` and `wdFindContinue`. `wdFindAsk` is G-HOST (no prompt). `Replace` is `wdReplaceNone`, `wdReplaceOne`, or `wdReplaceAll`. |
| Found | Property, read | Implemented | |
| Text, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, Replacement | Properties | Implemented | Setting them and then calling `Execute` with no arguments uses those properties. |
| Replacement.Text | Property | Implemented | |
| Replacement.Font, Replacement.ParagraphFormat | | Not implemented | G-ENGINE. Replace-with formatting is only the format criteria `js/find.js` already applies, not a full Font object on Replacement. |
| ClearFormatting | `ClearFormatting()` | Implemented | |
| HitHighlight, ClearHitHighlight | `HitHighlight(FindText, HighlightColor, TextColor, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl, IgnoreSpace, IgnorePunct, HighlightColorIndex)`, `ClearHitHighlight()` | Not implemented | G-ENGINE. Highlight-all is a view decoration, not these methods. |
| MatchPrefix, MatchSuffix, MatchPhrase, MatchByte, IgnoreSpace, IgnorePunct | | Not implemented | G-ENGINE. |

## Bookmarks

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Bookmarks.Count, Item, Exists | `Item(Index)`, `Exists(Name)` | Implemented | |
| Bookmarks.Add | `Add(Name, Range)` | Implemented | `Range` is an anchor. Omitted range uses the selection. |
| Bookmark.Name | Property | Implemented | |
| Bookmark.Range | Property | Implemented | Anchor whose `Text` is the bookmarked text. No `Start` / `End`. |
| Bookmark.Delete | `Delete()` | Implemented | |
| Bookmark.Select | `Select()` | Implemented | |
| Bookmark.Empty | Property, read | Implemented | |
| Bookmark.Start, Bookmark.End, Bookmark.Copy | | Not implemented | G-STORY for the offsets. `Copy(Name)` is G-ENGINE. |

## Fields

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Fields.Count, Item | `Item(Index)` | Implemented | |
| Fields.Add | `Add(Range, Type, Text, PreserveFormatting)` | Implemented | `Type` is a `WdFieldType` Quire can evaluate or store. `Text` is the instruction. Unknown types are stored and not evaluated (gap: they are not updated). |
| Fields.Update | `Update()` | Implemented | Updates fields Quire knows how to evaluate. |
| Fields.ToggleShowCodes | | Not implemented | G-ENGINE. |
| Field.Result | Property `Range` | Implemented | The result text. |
| Field.Code | Property `Range` | Implemented | The instruction text as an anchor, not a story `Long`. |
| Field.Update | `Update()` | Implemented | |
| Field.Unlink | `Unlink()` | Implemented | |
| Field.Delete | `Delete()` | Implemented | |
| Field.Locked | Property | Implemented | |
| Field.Type | Property, read | Implemented | |
| Field.Select | `Select()` | Implemented | |
| Field.Copy | | Not implemented | G-ENGINE. |
| Field.DoClick | | Not implemented | G-HOST. |

## Hyperlinks

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Hyperlinks.Count, Item | | Implemented | Links stored on runs. |
| Hyperlinks.Add | `Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay, Target)` | Implemented | `Anchor` is an anchor range. `Target` (frame name) is G-ENGINE. |
| Hyperlink.Address, SubAddress, TextToDisplay, ScreenTip | Properties | Implemented | |
| Hyperlink.Range | Property | Implemented | The linked text. No story `Long`. |
| Hyperlink.Delete | `Delete()` | Implemented | Removes the link, keeps the text. |
| Hyperlink.Follow | `Follow(NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo)` | Not implemented | G-HOST. |
| Hyperlink.Type, Hyperlink.Target | | Not implemented | G-ENGINE. |

## Comments

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Comments.Count, Item | | Implemented | |
| Comments.Add | `Add(Range, Text)` | Implemented | |
| Comment.Range | Property | Implemented | The comment body. |
| Comment.Scope | Property | Implemented | The anchored document span. No story `Long`. |
| Comment.Author, Comment.Initial | Properties | Implemented | |
| Comment.Delete | `Delete()` | Implemented | |
| Comment.Edit | | Not implemented | G-HOST. Editing the body is `Comment.Range.Text`. There is no comment pane object. |
| Comments.AddReply and reply threads beyond the reply Quire stores | `AddReply(Comment, Text)` | Implemented | The call is `missing-live`. Deeper Word comment-thread properties are G-ENGINE. |

## Footnotes and endnotes

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Footnotes.Count, Item, Endnotes.Count, Item | | Implemented | |
| Footnotes.Add, Endnotes.Add | `Add(Range, Reference, Text)` | Implemented | |
| Footnote.Range, Endnote.Range | Property | Implemented | The note story. |
| Footnote.Index | Property, read | Implemented | |
| Footnote.Reference | Property | Not implemented | G-STORY. The reference mark's story offset is not published. The note is still reachable through `Item`. |
| Footnotes.Convert, Endnotes.Convert | | Implemented | The conversion `N.convertAll` already performs. |

## InlineShapes and Shapes

Insert and the properties Quire already edits on a picture or a shape. The Excel chart object and the diagram-node object stay in the no-engine list.

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| InlineShapes.Count, Item | | Implemented | Inline pictures and inline shapes. |
| InlineShapes.AddPicture | `AddPicture(FileName, LinkToFile, SaveWithDocument, Range)` | Implemented on the VS Code host | RequestApproval. Embeds PNG, JPEG, EMF, or WMF. `LinkToFile` True is an error, and so is a range. |
| InlineShape.Width, Height, ScaleWidth, ScaleHeight | Properties | Implemented | Points and percent, as stored. |
| InlineShape.Type | Property, read | Implemented | Picture, shape, chart, or text box, mapped to the nearest `WdInlineShapeType`. |
| InlineShape.Select, Delete | | Implemented | |
| InlineShape.Range | Property | Implemented | The inline anchor. No story `Long`. |
| InlineShape.AlternativeText, Title | Properties | Implemented | |
| InlineShape.Hyperlink | Property | Implemented | When the shape stores a link. |
| InlineShape.PictureFormat | crop and brightness/contrast only when those values are stored | Implemented | The rest of `PictureFormat` is G-ENGINE. |
| Shapes.Count, Item | | Implemented | Floating shapes. |
| Shape.Left, Top, Width, Height | Properties | Implemented | |
| Shape.WrapFormat.Type | Property | Implemented | Wrap modes `DR.setWrap` already sets. |
| Shape.TextFrame.TextRange | Property | Implemented | Text of a text box, as an anchor range. |
| Shape.Select, Delete | | Implemented | |
| Shape.Fill.ForeColor, Shape.Line.ForeColor, Shape.Line.Weight | | Implemented | When the shape stores a solid fill and a line. Gradients, patterns, and arrows are G-ENGINE. |
| Shapes.AddShape, AddTextbox | `AddShape(Type, Left, Top, Width, Height, Anchor)`, `AddTextbox(Orientation, Left, Top, Width, Height, Anchor)` | Implemented | Types `DR.newShape` can build. Other auto-shape type codes are G-ENGINE. |
| Shapes.AddChart | `AddChart(Type, Left, Top, Width, Height, Anchor)` | Not implemented | G-ENGINE. `A.insertChart` opens the chart inserter; it is not this call. |
| Shape.OLEFormat, Shape.Chart, Shape.GroupItems, Shape.Nodes | | Not implemented | G-ENGINE. |

## Revisions

| Member | VBA signature | Status | Gap |
|---|---|---|---|
| Revisions.Count, Item | | Implemented | |
| Revisions.AcceptAll, RejectAll | | Implemented | |
| Revision.Accept, Reject | | Implemented | |
| Revision.Type | Property, read | Implemented | Insert or delete, mapped to `WdRevisionType`. Formatting revisions that Quire does not list are absent. |
| Revision.Author, Date | Properties, read | Implemented | |
| Revision.Range | Property | Implemented | The revised span's text. No story `Long`. |
| Document.TrackRevisions | Property | Implemented | |

## Enumerations

Defined as numeric constants with Word's values. Only constants named by an **Implemented** row are required. Further `wd*` names are not defined until a row needs them.

| Group | Constants |
|---|---|
| Paragraph alignment | `wdAlignParagraphLeft` 0, `wdAlignParagraphCenter` 1, `wdAlignParagraphRight` 2, `wdAlignParagraphJustify` 3, `wdAlignParagraphDistribute` 4 |
| Collapse | `wdCollapseEnd` 0, `wdCollapseStart` 1 |
| Save on close | `wdDoNotSaveChanges` 0, `wdSaveChanges` -1, `wdPromptToSaveChanges` -2 |
| Find wrap and replace | `wdFindStop` 0, `wdFindContinue` 1, `wdReplaceNone` 0, `wdReplaceOne` 1, `wdReplaceAll` 2 |
| Underline | `wdUnderlineNone` 0, `wdUnderlineSingle` 1, `wdUnderlineWords` 2, `wdUnderlineDouble` 3, `wdUnderlineDotted` 4, `wdUnderlineThick` 6, `wdUnderlineDash` 7, `wdUnderlineDotDash` 9, `wdUnderlineDotDotDash` 10, `wdUnderlineWavy` 11 |
| Line spacing | `wdLineSpaceSingle` 0, `wdLineSpace1pt5` 1, `wdLineSpaceDouble` 2, `wdLineSpaceAtLeast` 3, `wdLineSpaceExactly` 4, `wdLineSpaceMultiple` 5 |
| Story | `wdMainTextStory` 1, `wdFootnotesStory` 2, `wdEndnotesStory` 3, `wdCommentsStory` 4, `wdTextFrameStory` 5, `wdEvenPagesHeaderStory` 6, `wdPrimaryHeaderStory` 7, `wdEvenPagesFooterStory` 8, `wdPrimaryFooterStory` 9, `wdFirstPageHeaderStory` 10, `wdFirstPageFooterStory` 11 |
| Header/footer index | `wdHeaderFooterPrimary` 1, `wdHeaderFooterFirstPage` 2, `wdHeaderFooterEvenPages` 3 |
| Orientation | `wdOrientPortrait` 0, `wdOrientLandscape` 1 |
| Statistics | `wdStatisticWords` 0, `wdStatisticLines` 1, `wdStatisticPages` 2, `wdStatisticCharacters` 3, `wdStatisticParagraphs` 4, `wdStatisticCharactersWithSpaces` 5 |
| Selection type | `wdNoSelection` 0, `wdSelectionIP` 1, `wdSelectionNormal` 2 |
| Color | `wdColorAutomatic` -16777216, and the named `wdColor*` values of the palette Quire's color control already lists, as Word BGR longs |

`WdFieldType` values are Word's numbers for the field types `js/fields.js` evaluates. Types it only stores still use Word's number on `Field.Type` when the instruction was read from a file; `Fields.Add` rejects a type it cannot store.

## Worked calls

These are the kind of calls step 2 must run against the open document. They use anchors, not story longs.

```javascript
ActiveDocument.Paragraphs(1).Range.Text
ActiveDocument.Paragraphs(1).Range.Text = "Hello"
Selection.TypeText(" ")
Selection.Font.Bold = true
ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphCenter
ActiveDocument.Tables(1).Cell(1, 1).Range.Text
ActiveDocument.Bookmarks.Add("Here", Selection.Range)
Selection.Find.Execute("Hello", true, false, false, false, false, true, wdFindContinue, false, "Hi", wdReplaceOne)
```

`ActiveDocument.Range(0, 5)` is not in this set. It stays **Not implemented** (G-STORY).
