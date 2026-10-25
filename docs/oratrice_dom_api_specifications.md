# Oratrice PowerPoint object model

The object model on the open presentation. `public/oratrice/js/dom.js` implements this inventory. The VB host calls it through [oratrice_dom_ir_specifications.md](oratrice_dom_ir_specifications.md). Milestones are [oratrice_dom_ir_porting_milestones.md](oratrice_dom_ir_porting_milestones.md).

A row is implemented when `dom.js` performs it. A gap is callable and rejects with `err.code`. The codes are `G-TEXT` (a character index inside a shape), `G-ENGINE` (no document operation) and `G-HOST` (this session does not have that host ability). Host abilities and the three audit tiers are in [docs/dom_api_security.md](dom_api_security.md).

## Rules

- The entry is `Application.ActivePresentation` and `Presentations`. `Selection` is `ActiveWindow.Selection`.
- Collections are 1-based. VBA parameter order. Omit arguments from the right. There is no options object.
- VBA `True` and `False` are `-1` and `0`. A tri-state write treats `-1` and `true` as on, and `0` and `false` as off.
- Colour is `ColorFormat`. `RGB` is an Office BGR long. `ObjectThemeColor` is a Oratrice slot (`dk1`, `lt1`, `accent1`, …). `SchemeColor` is one of the eight 2003 slots. Tint, shade and brightness writes are `G-ENGINE`.
- `TextFrame2` is the same object as `TextFrame`. Members that exist only on `TextFrame2` are `G-ENGINE` or absent.
- Edits go through the presentation model, the text operations and undo. The object model does not layout the slide and does not add members that are not in this inventory.
- The task-pane reference is `public/oratrice/js/vb-dom-ref.js`. A change to a member or a gap updates that file in the same change. That page is what a script sees. A member listed only here, or only there, is out of date.
- `Creator` is `G-ENGINE` on every object. Oratrice is not PowerPoint.
- Shape `Left`, `Top`, `Width` and `Height` are points.
- A text range is an anchor `{p, o}` inside one shape. `Text` joins paragraphs with a carriage return. `Characters.Count` counts stored characters and does not include that carriage return. `Characters(Start, Length)`, `Words`, `Sentences`, `Lines`, `Runs`, `Find`, `Replace`, `Start`, `Length` and the `Bound*` members are `G-TEXT`.
- Document macros are VB6 source in the custom XML part `https://acane.net/melinaoffice/macros`. `vbaProject.bin` is not listed and is not run. The object model does not expose either store.

## Status and gaps

Implemented in `public/oratrice/js/dom.js`. A held shape wrapper re-reads the shape by id, so undo does not leave `Left` on the snapshot that undo replaced. A text range still closes over the text body it was given; read it again after undo.

`ActiveWindow` is thin. `Selection` and `Presentation` are implemented. `View`, `ViewType`, `Caption`, `ActivePane`, `Panes` and `Split` are `G-HOST`.

`SlideShowWindow` has the navigation calls (`Next`, `Previous`, `GotoSlide`, `Exit`). Pointer, state and accelerators are `G-HOST`.

## Text index (not implemented)

PowerPoint's `TextRange.Start` and `TextRange.Length` are a character index inside one shape, counting the paragraph mark. Oratrice stores a paragraph index and an offset. Those members, and any member that needs the index (`Characters(start, length)`, `Words`, `Sentences`, `Lines`, `Runs`, `Find`, `Replace`, `BoundLeft`, `BoundTop`, `BoundWidth`, `BoundHeight`), reject with `G-TEXT`.

`Characters` is still a collection. `Count` is implemented. `Item` is `G-TEXT`.

## Objects with no document operation

These names are registered. Any member is that object's gap. They are not reduced wrappers.

| Object | Code | Why |
|---|---|---|
| `TimeLine`, `Sequence`, `Effect` | `G-ENGINE` | A slide's animation records are not a VBA timeline |
| `Chart`, `ChartData` | `G-ENGINE` | A chart shape is not a `Chart` object |
| `Diagram`, `DiagramNode`, `SmartArt` | `G-ENGINE` | IntelliArt nodes are not this model |
| `OLEFormat`, `MediaFormat`, `CustomerData` | `G-ENGINE` | No operation on those parts |
| `NotesPage` | `G-ENGINE` | Notes are `slide.notes`, a string, not a shape page |
| `Comments`, `Comment` | `G-ENGINE` | Comment parts are kept; there is no comment editor |
| `Tags` | `G-ENGINE` | No tag bag |
| `PrintOptions` | `G-ENGINE` | Printing is the print dialog |
| `Panes`, `Pane`, `View`, `CommandBars`, `FileDialog`, `Dialogs` | `G-HOST` | Windows, panes and the printer |
| `ActionSettings`, `NamedSlideShows`, `SectionProperties`, `ColorSchemes`, `AnimationSettings` | `G-ENGINE` | No matching operation |

`Slide.NotesPage`, `Slide.Comments`, `Slide.TimeLine`, `Slide.Tags` and `Slide.CustomerData` are the same gap on the slide. `Shape.Chart`, `Shape.OLEFormat`, `Shape.MediaFormat`, `Shape.ActionSettings`, `Shape.AnimationSettings` and `Shape.Nodes` are the same gap on the shape.

## Application

| Member | | |
|---|---|---|
| `ActivePresentation` | get | The open presentation |
| `Presentations` | get | The collection |
| `ActiveWindow` | get | The document window |
| `SlideShowWindows` | get | The running show, or empty |

## Presentations

| Member | | |
|---|---|---|
| `Count`, `Item`, `First`, `Last` | | 1-based. `Presentations(1)` is `Item` |
| `Add` | call | No arguments. Returns the new presentation and makes it active. `WithWindow` is `G-HOST` |
| `Open` | call | RequestApproval on the VS Code host. Opens that file in its own tab. Elsewhere `G-HOST` |

Closing the last presentation opens a blank one. The editor always has a presentation.

## Presentation

| Member | | |
|---|---|---|
| `Name` | get | The file name, or `Presentation1` |
| `FullName`, `Path` | get | AlwaysAllow on the VS Code host. Elsewhere `G-HOST` |
| `Saved` | read/write | `True` when the deck is not dirty. A new unedited deck reads `True` |
| `Slides` | get | |
| `PageSetup` | get | |
| `SlideShowSettings` | get | |
| `SlideShowWindows` | get | |
| `Designs` | get | The built-in design list |
| `SlideMaster` | get | Name and headers/footers. Shapes, layouts, design and background are `G-ENGINE` |
| `BuiltInDocumentProperties` | get | Title, Subject, Author, Keywords, Comments, Category, Company |
| `CustomDocumentProperties` | get | String values. `Add` with `LinkToContent` or a type other than 4 is `G-ENGINE` |
| `Undo`, `Redo` | call | Return `-1` when a step ran |
| `Save` | call | Calls the app save when a file is already saved. Otherwise, in the browser, `G-HOST`. With no save function, clears the dirty flag |
| `Close` | call | Drops this presentation. The last one is replaced by a blank deck |
| `SaveAs`, `SaveAs2` | call | RequestApproval on the VS Code host. The deck's presentation format, or PDF |
| `Export` | call | RequestApproval. PNG or JPEG of every slide, including hidden slides, as `Slide1.png` in the folder |
| `ExportAsFixedFormat` | call | RequestApproval. PDF. XPS errors before the confirmation. The path is the first argument |
| `PrintOut` | call | `PrintOut(From, To, PrintToFile, Copies, Collate)`. AlwaysAllow. A PDF of the slides beside the file. Every argument is accepted and ignored |
| `FollowHyperlink` | call | AlwaysAllow. `http`, `https`, or `mailto` |
| `NewWindow` | call | `G-HOST` |
| `HasHostAbility`, `MsgBox`, `InputBox`, `FileDialog`, `HostDir`, `HostShell`, `GetOpenFilename`, `GetSaveAsFilename` | call | See [docs/dom_api_security.md](dom_api_security.md). Oratrice has no `Dialogs` index. `GetSaveAsFilename(InitialFilename, FileFilter, FilterIndex, Title, ButtonText)` accepts the names and does not apply them |
| `CreateVideo`, `Merge`, `PublishSlides` | call | `G-ENGINE`. `CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality)`, `Merge(Path)`, `PublishSlides(SlideLibraryUrl, Overwrite, UseSlideOrder)` |

## DocumentWindow and Selection

| Member | | |
|---|---|---|
| `Selection` | get | |
| `Presentation` | get | The active presentation |
| `View`, `ViewType`, `Caption`, `ActivePane`, `Panes`, `Split` | get | `G-HOST` |
| `Selection.Type` | get | `ppSelectionNone` 0, `ppSelectionSlides` 1, `ppSelectionShapes` 2, `ppSelectionText` 3 |
| `ShapeRange` | get | `Left`, `Top`, `Delete`, and the shapes. `G-ENGINE` when nothing is selected |
| `SlideRange` | get | The selected slides, or the current slide |
| `TextRange` | get | The selected text anchor. `G-ENGINE` when the selection is not text |
| `Unselect`, `Delete`, `Copy`, `Cut`, `Paste` | call | Shape and text clipboard. A slide-only paste is not this call |
| `ChildShapeRange` | get | `G-ENGINE` |

## Slides and Slide

| Member | | |
|---|---|---|
| `Add(index, layout)` | call | Index is 1-based. An unmapped layout code is `G-ENGINE` |
| `Range` | call | |
| `Name` | read/write | |
| `SlideIndex`, `SlideNumber` | get | `SlideNumber` adds `FirstSlideNumber` |
| `SlideID` | get | `G-ENGINE`. Slide ids are strings |
| `Layout` | read/write | Unmapped stored layouts read `ppLayoutCustom` (32). An unmapped write is `G-ENGINE` |
| `Shapes` | get | |
| `Design` | read/write | A design name or key. An unknown name is `G-ENGINE` |
| `HeadersFooters` | get | The presentation's header and footer record |
| `SlideShowTransition` | get | |
| `Select`, `Delete`, `Duplicate`, `MoveTo` | call | |
| `Hyperlinks` | get | Links on the slide's text. `Add` writes onto those paragraphs |
| `NotesPage`, `Comments`, `TimeLine`, `Tags`, `CustomerData` | get | `G-ENGINE` |

## Shapes and Shape

`Shapes` on a slide can add. `GroupItems` cannot.

| Member | | |
|---|---|---|
| `Item`, `Count`, `Title`, `Placeholders` | | `Title` is the title placeholder |
| `AddShape(type, left, top, width, height)` | call | The types in Enumerations. Any other `AutoShapeType` is `G-ENGINE` |
| `AddTextbox(orientation, left, top, width, height)` | call | Horizontal only |
| `AddTable(rows, cols, left, top, width)` | call | |
| `AddPicture` | call | RequestApproval on the VS Code host. Embeds the picture. `LinkToFile` True is an error |
| `AddChart`, `AddSmartArt`, `AddDiagram`, `AddMediaObject`, `AddOLEObject`, `AddTextEffect` | call | `G-ENGINE`. `AddChart(Type, Left, Top, Width, Height)`, `AddSmartArt(Layout, Left, Top, Width, Height)`, `AddDiagram(Type, Left, Top, Width, Height)`, `AddMediaObject(FileName, Left, Top, Width, Height)`, `AddOLEObject(Left, Top, Width, Height, ClassName, FileName, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Link)`, `AddTextEffect(PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top)` |
| `Left`, `Top`, `Width`, `Height`, `Rotation` | read/write | Points. `Rotation` is degrees |
| `Name` | read/write | |
| `Id`, `Type`, `HasTextFrame`, `HasTable`, `ZOrderPosition` | get | `Type` is `msoAutoShape`, `msoTextBox`, `msoTable`, `msoPicture`, `msoPlaceholder`, `msoGroup`, `msoChart`, `msoTextEffect` or `msoLine` |
| `AutoShapeType` | read/write | The mapped presets only |
| `TextFrame`, `TextFrame2` | get | The same object |
| `Fill`, `Line` | get | Solid fill and a solid line. A gradient reads `msoFillGradient`. Writing a type other than solid is `G-ENGINE`. `Line.DashStyle` is `G-ENGINE` |
| `PlaceholderFormat` | get | `Type` and `ContainedType`. Vertical title and body are types 5 and 6. `G-ENGINE` when the shape is not a placeholder |
| `Table` | get | `G-ENGINE` when the shape is not a table |
| `GroupItems` | get | `G-ENGINE` when the shape is not a group |
| `Delete`, `Duplicate`, `Select`, `ZOrder` | call | `ZOrder` is `msoBringToFront` 0, `msoSendToBack` 1, `msoBringForward` 2, `msoSendBackward` 3. Z-order index 0 is the back, so `ZOrderPosition` is the index plus one |
| `Hyperlinks` | get | Links stored on the shape's text |
| `Chart`, `OLEFormat`, `MediaFormat`, `ActionSettings`, `AnimationSettings`, `Nodes`, `Tags` | get | `G-ENGINE` |

## TextFrame, TextRange, Paragraph, Font

| Member | | |
|---|---|---|
| `TextRange` | get | The whole body |
| `HasText` | get | |
| `WordWrap`, `AutoSize` | read/write | `ppAutoSizeNone` 0, `ppAutoSizeShapeToFitText` 1. `msoAutoSizeTextToFitShape` 2 is Oratrice autofit `norm` |
| `MarginLeft`, `MarginTop`, `MarginRight`, `MarginBottom` | read/write | Points |
| `VerticalAnchor` | read/write | `ppAnchorTop` 1, `ppAnchorMiddle` 3, `ppAnchorBottom` 4 |
| `Orientation` | get | `G-ENGINE` |
| `Text` | read/write | Paragraphs joined by a carriage return. A whole-body write replaces the body. A partial write deletes that span and inserts |
| `InsertBefore`, `InsertAfter`, `Delete`, `Select` | call | |
| `Font`, `ParagraphFormat` | get | |
| `IndentLevel` | read/write | 1-based. Oratrice stores 0-based, levels 1 through 9 |
| `Paragraphs` | get | |
| `Characters` | get | `Count` only |
| `InsertSymbol`, `InsertDateTime`, `InsertSlideNumber` | call | The date uses the `datetime1` format when the page has it |
| `Start`, `Length`, `BoundLeft`, `BoundTop`, `BoundWidth`, `BoundHeight` | get | `G-TEXT` |
| `Find`, `Replace` | call | `G-TEXT`. `Find(FindWhat, After, MatchCase, WholeWords)`, `Replace(FindWhat, ReplaceWhat, After, MatchCase, WholeWords)` |
| `Words`, `Sentences`, `Lines`, `Runs` | get | `G-TEXT` |
| `Paragraph.Alignment` | read/write | `ppAlignLeft` 1 through `ppAlignDistribute` 5 |
| `Font.Name`, `NameFarEast`, `Size`, `Bold`, `Italic`, `Underline` | read/write | `Name` writes the Latin face. Underline `msoTrue` is a single line. `NameAscii`, `NameComplexScript`, `SizeBi`, `BoldBi` and `ItalicBi` are `G-ENGINE` |
| `Font.Color` | get | `ColorFormat` |
| `ParagraphFormat.Alignment`, `SpaceBefore`, `SpaceAfter`, `SpaceWithin`, `LineRuleWithin` | read/write | Point spacing stores points. A percent spacing value reads as `G-ENGINE`. `LineRuleWithin` is `-1` for a multiple and `0` for points |
| `Bullet.Type`, `Bullet.Visible` | read/write | `ppBulletNone` 0, `ppBulletUnnumbered` 1, `ppBulletNumbered` 2 |

## Table, fill, line, colour, hyperlink

| Member | | |
|---|---|---|
| `Table.Cell(row, column)`, `Rows`, `Columns` | | `Rows.Add` and `Columns.Add` append or insert |
| `Row.Height`, `Column.Width`, `Cell.Shape` | | |
| `Table.Merge`, `Table.ApplyStyle` | call | `G-ENGINE`. `Merge(MergeTo)`, `ApplyStyle(StyleID, SaveFormatting)` |
| `Fill.Type` | read/write | 0 none, 1 solid. A gradient reads 3. Writing another type is `G-ENGINE` |
| `Fill.ForeColor`, `Fill.Transparency`, `Fill.Visible` | | Transparency is `1 - alpha` |
| `Line.Weight`, `Line.ForeColor`, `Line.Visible` | | `Visible` false sets the line to none |
| `Line.DashStyle` | get | `G-ENGINE` |
| `ColorFormat.RGB` | read/write | BGR long |
| `ObjectThemeColor`, `SchemeColor` | read/write | See Enumerations |
| `TintAndShade`, `Brightness` | get | `G-ENGINE` |
| `Hyperlinks.Add(anchor, address, subAddress, screenTip, textToDisplay)` | call | Writes the link onto the anchor's runs |
| `Address`, `SubAddress`, `ScreenTip`, `TextToDisplay` | read/write | |
| `Hyperlink.Follow` | call | `G-HOST` |
| `Hyperlink.Type` | get | `G-ENGINE` |

## Page setup, design, headers, transition, show

| Member | | |
|---|---|---|
| `SlideWidth`, `SlideHeight`, `FirstSlideNumber` | read/write | Points and the first number |
| `SlideSize`, `NotesOrientation`, `SlideOrientation` | get | `G-ENGINE` |
| `Designs`, `Design.Name` | | Built-in designs |
| `HeadersFooters.DateAndTime`, `Footer`, `SlideNumber` | get | `Visible` and `Text` |
| `DisplayOnTitleSlide` | read/write | The opposite of "not on title" |
| `HeaderFooter.Format`, `UseFormat` | get | `G-ENGINE` |
| `EntryEffect` | read/write | The effects in Enumerations. A stored effect outside that list reads `G-ENGINE` |
| `Speed` | read/write | `ppTransitionSpeedSlow` 1, medium 2, fast 3 |
| `AdvanceOnClick`, `AdvanceOnTime`, `AdvanceTime`, `Hidden` | read/write | `AdvanceTime` is seconds. The deck stores milliseconds |
| `Duration`, `SoundEffect` | get | `G-ENGINE` |
| `StartingSlide`, `EndingSlide` | read/write | 1-based. Both empty means the whole deck |
| `RangeType` | read/write | `ppShowAll` 1, `ppShowSlideRange` 2. `ppShowNamedSlideShow` 3 is `G-ENGINE` |
| `LoopUntilStopped`, `AdvanceMode`, `ShowType` | read/write | `ppShowTypeWindow` 2 is `G-HOST`. Kiosk is 3 |
| `Run` | call | Starts the headless show position. When the page has a running show, it starts that too |
| `NamedSlideShows`, `PointerColor` | get | `G-ENGINE` |
| `SlideShowWindows(1).View.Next`, `Previous`, `GotoSlide`, `Exit` | call | `Next` before `Run` is `G-ENGINE` |
| `CurrentShowPosition` | get | 1-based while the show is on, otherwise 0 |
| `State`, `PointerType`, `AcceleratorsEnabled` | get | `G-HOST` |

Hidden slides are skipped in the show order.

## Enumerations

The numbers live on `L.dom` and on the page, and the VB module injects them as `Public Const`.

Layouts: `ppLayoutTitle` 1, `ppLayoutText` 2, `ppLayoutTwoColumnText` 3, `ppLayoutTable` 4, `ppLayoutChart` 8, `ppLayoutTitleOnly` 11, `ppLayoutBlank` 12, then the content layouts 13–16, 19–22, 24, 25, 27, 29–31, and `ppLayoutCustom` 32.

Placeholders: title 1, body 2, center title 3, subtitle 4, vertical title 5, vertical body 6, object 7, chart 8, table 12, slide number 13, footer 15, date 16.

AutoShapes: rectangle 1, diamond 4, rounded rectangle 5, oval 9, hexagon 10, cross 11, can 13, cube 14, donut 18, block arc 20, heart 21, lightning bolt 22, sun 23, moon 24, arc 25.

Shape types: auto shape 1, chart 3, group 6, line 9, picture 13, placeholder 14, text effect 15, text box 17, table 19.

Theme colours: Dark1 1 through Background2 16, in the `msoThemeColor*` constants. Scheme: `ppBackground` 1 is `lt1`, `ppForeground` 2 is `dk1`, `ppShadow` 3 is `lt2`, `ppTitle` 4 is `dk2`, `ppFill` 5 is `accent1`, `ppAccent` 6 is `accent2`, `ppAccentAndHyperlink` 7 is `hlink`, `ppAccentAndFollowedHyperlink` 8 is `folHlink`.

Transitions: none 0, cut 257, cut through black 258, dissolve 1537, wipe left/up/right/down 2817–2820, fade smoothly 3849, push down/left/right/up 3852–3855.

Show: `ppShowAll` 1, `ppShowSlideRange` 2, `ppShowNamedSlideShow` 3. `ppShowTypeSpeaker` 1, `ppShowTypeWindow` 2, `ppShowTypeKiosk` 3. Advance manual 1, use slide timings 2.

## Worked calls

```vb
ActivePresentation.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "Hello"
ActivePresentation.Slides.Add 2, ppLayoutText
ActiveWindow.Selection.ShapeRange.Left = 72
```

`TextRange.Characters(Start:=1, Length:=5)` is `G-TEXT`. `Characters.Count` on `"Hello"` is 5.
