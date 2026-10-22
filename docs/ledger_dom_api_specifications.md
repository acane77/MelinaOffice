# Ledger Excel object model

The object model on the open workbook. `public/ledger/js/dom.js` implements this inventory. The VB host calls it through [ledger_dom_ir_specifications.md](ledger_dom_ir_specifications.md). Milestones are [ledger_dom_ir_porting_milestones.md](ledger_dom_ir_porting_milestones.md). The schedule is [ledger_vbs_support_schedule.md](ledger_vbs_support_schedule.md).

A row is implemented when `dom.js` performs it. A gap is callable and rejects with `err.code`. The codes are `G-ENGINE` (no document operation) and `G-HOST` (a path, a window, a printer, a modal, or a file picker). There is no `G-STORY` and no `G-TEXT`: a cell is a row and a column.

## Rules

- The entry is `Application`. `ThisWorkbook` in the console is the active workbook. A later document-macro phase would make `ThisWorkbook` the workbook that stores the source.
- Collections are 1-based. The workbook model is 0-based. Conversion stays inside the wrapper.
- VBA parameter order. Omit arguments from the right. There is no options object. VBA `True` and `False` are `-1` and `0`.
- A single cell `Value`, `Value2` or `Formula` is a scalar. A block is a rectangle. One cell is never a 1×1 array. In JavaScript a block is a 0-based array of rows. In VB it is a 1-based `VBArray`. A range larger than 1,000,000 cells rejects `Value` and `Formula` with `G-ENGINE`.
- A date is an Excel serial number, not a VBA `Date`. `Value` and `Value2` both return the serial. An error is a string such as `#N/A`, not a thrown gap. An empty cell is `null` in JavaScript and `Empty` in VB. A string `Value` that starts with `=` is stored as a formula. A formula written onto a block shifts relative references from the top-left cell.
- Colour is an Office BGR long. `ThemeColor` uses the Excel theme enum: 1 `dk1` is theme index 1, 2 `lt1` is 0, 3 `dk2` is 3, 4 `lt2` is 2, 5–10 accent 1–6 are 4–9, 11 hyperlink is 10, 12 followed hyperlink is 11.
- `WorksheetFunction` is not a catalog of 519 rows. A name in Ledger's function catalog is callable. A dotted name (`BINOM.DIST`) is also `BINOM_DIST`. A name that is not in the catalog is absent. Tests sample `Sum`, `If` and `VLookup`.
- `ScreenUpdating` and `DisplayAlerts` are host flags. `Calculation` maps automatic and manual onto the workbook. Semiautomatic is `G-ENGINE`. `ReferenceStyle` maps `xlA1` and `xlR1C1`.
- Classic `FormatConditions` (`cellIs` and `expression`) are implemented. Data bars, colour scales and icon sets stay display-only (`G-ENGINE`). `SortFields` is `G-ENGINE`. `Sort` is Excel 2003 three-key sort. `AutoFilter` stores `{ref, cols:[{values:[string]}]}`.
- Copy and `PasteSpecial` use an in-app snapshot, not the operating-system clipboard. `Save` works when the workbook already has a destination. `SaveAs`, `Open`, `PrintOut`, `FullName` and `Path` are `G-HOST`. Closing the last workbook opens a blank one. With `DisplayAlerts` and a dirty book, `Close` with no `SaveChanges` argument is `G-HOST`. `SaveChanges` true is also `G-HOST`.
- `Creator` is `G-ENGINE` on every object. Ledger is not Excel.
- Edits go through the workbook operations and undo. The object model does not invent members that are not in this inventory.
- The task-pane reference is `public/ledger/js/vb-dom-ref.js`. A change to a member or a gap updates that file in the same change. That page is what a script sees. A member listed only here, or only there, is out of date.
- `vbaProject.bin` is kept, not listed and not run. Macros written here are VB6 source in the custom XML part `urn:vibeoffice:macros`, or in this browser for All templates and documents. The object model does not expose either store. `Application.Run` is `G-HOST`.

## Status and gaps

Implemented in `public/ledger/js/dom.js`. Chart, pivot, slicer, sparkline and query-table surfaces are thin or whole-object gaps. `PivotTables` lists tables already on the workbook. `Refresh` and `PivotFields` are `G-ENGINE`. Macro sheets and dialog sheets appear in `Sheets`. Editing their cells is `G-ENGINE`. A chart sheet has no `Range`.

## Objects with no document operation

These names are registered. Any member is that object's gap. `Application` and `Parent` still navigate. `Creator` and every other member return the gap.

| Object | Code | Why |
|---|---|---|
| `QueryTables`, `QueryTable`, `Connections`, `WorkbookConnection`, `XmlMaps`, `XmlMap` | `G-ENGINE` | Connections and maps are kept, not executed |
| `Slicers`, `SlicerCaches`, `Slicer`, `SparklineGroups`, `SparklineGroup` | `G-ENGINE` | Display and package parts, no editor |
| `OLEObjects`, `OLEObject` | `G-ENGINE` | Embedded objects are kept |
| `CommentsThreaded`, `CommentThreaded`, `Phonetics`, `Phonetic` | `G-ENGINE` | Not this note model |
| `Model`, `Queries` | `G-ENGINE` | Power Pivot and Power Query are not run |
| `Windows`, `Window`, `AddIns`, `AddIn`, `Dialogs`, `Dialog`, `FileDialog`, `CommandBars` | `G-HOST` | Windows, add-ins and file pickers |

## Application

| Member | | |
|---|---|---|
| `ActiveWorkbook`, `ThisWorkbook` | get | The console workbook |
| `Workbooks` | get | Open workbooks |
| `ActiveSheet` | get | The active worksheet |
| `ActiveCell`, `Selection` | get | The selection, when the grid's workbook matches; otherwise the sheet's stored selection |
| `Names` | get | Workbook names |
| `WorksheetFunction` | get | Catalog functions |
| `ScreenUpdating`, `DisplayAlerts` | rw | Host flags |
| `Calculation` | rw | Automatic or manual. Semiautomatic is `G-ENGINE` |
| `ReferenceStyle` | rw | `xlA1` or `xlR1C1` |
| `Calculate`, `Undo`, `Redo` | call | Recalculate, or one undo or redo step (`-1` if a step ran, else `0`) |
| `Windows`, `AddIns`, `Dialogs`, `FileDialog`, `CommandBars` | get | `G-HOST` |
| `Run`, `SendKeys`, `Wait`, `OnTime`, `GetOpenFilename`, `GetSaveAsFilename` | call | `G-HOST` |
| `EnableEvents` | get | `G-ENGINE` |

## Workbooks, Workbook

| Member | | |
|---|---|---|
| `Workbooks.Item`, `Add` | call | By name or index. `Add` opens a blank book |
| `Workbooks.Open`, `Close` | call | `G-HOST` |
| `Name` | get | The workbook name |
| `FullName`, `Path` | get | `G-HOST` |
| `Saved` | rw | False after an edit until a successful save |
| `Worksheets`, `Sheets`, `Charts`, `ActiveSheet`, `Names`, `Styles` | get | `Charts` is the chart sheets inside `Sheets` |
| `BuiltInDocumentProperties`, `CustomDocumentProperties` | get | Built-in map: title, subject, creator, keywords, description, category, manager, company |
| `Activate` | call | Makes the book current |
| `Save` | call | Clears dirty when a destination exists. Otherwise `G-HOST` |
| `Close` | call | See the rules above |
| `Protect`, `Unprotect` | call | Workbook structure. The password is hashed when the dialog helper is loaded |
| `SaveAs`, `SaveAs2`, `PrintOut`, `ExportAsFixedFormat`, `FollowHyperlink`, `SendMail`, `Windows` | | `G-HOST` |

`DocumentProperties.Add` with `LinkToContent`, or a type other than a string, is `G-ENGINE`.

## Names, Styles

| Member | | |
|---|---|---|
| `Names.Item`, `Add` | call | `Add(Name, RefersTo)` |
| `Name.Name`, `RefersTo` | rw | |
| `Name.RefersToR1C1` | get | |
| `Name.Delete` | call | |
| `Styles.Item` | call | An existing cell style |
| `Style.Name` | get | |

## Sheets, Worksheets, Worksheet

| Member | | |
|---|---|---|
| `Item`, `Add` | call | `Add(Before, After, Count)` inserts at the Before sheet, or after After |
| `Name` | rw | |
| `Index` | get | 1-based |
| `Visible` | rw | `xlSheetVisible` `-1`, `xlSheetHidden` `0`, `xlSheetVeryHidden` `2` |
| `UsedRange`, `Range`, `Cells`, `Rows`, `Columns` | | `Cells` with no arguments is the used sheet. `Range` takes an address or two corners |
| `Comments`, `Hyperlinks`, `Names`, `ListObjects`, `ChartObjects`, `Shapes` | get | |
| `PivotTables`, `PageSetup`, `HPageBreaks`, `VPageBreaks`, `Outline`, `AutoFilter`, `Scenarios` | get | |
| `ProtectionMode` | get | |
| `Activate`, `Select`, `Delete`, `Move`, `Calculate`, `Protect`, `Unprotect` | call | |
| `Copy` | call | A shallow sheet copy |
| `QueryTables`, `XmlMaps`, `OLEObjects`, `CommentsThreaded` | get | `G-ENGINE` |
| `PrintOut`, `ExportAsFixedFormat`, `SaveAs` | call | `G-HOST` |

## Range

| Member | | |
|---|---|---|
| `Count`, `Row`, `Column`, `Address`, `Text`, `MergeCells` | get | `Text` is the number format applied to the value |
| `Rows`, `Columns`, `Worksheet`, `Font`, `Interior`, `Borders` | get | |
| `Formula`, `FormulaR1C1`, `Value`, `Value2`, `NumberFormat` | rw | |
| `ColumnWidth`, `RowHeight`, `Hidden` | rw | `Hidden` applies to a whole row or a whole column. Width `0` hides a column. A hidden row reports height `0` |
| `Item`, `Cells`, `Resize`, `Offset`, `End`, `SpecialCells` | call | `SpecialCells` covers constants, formulas, blanks, the last cell, visible cells and comments, plus the value-type bits. Other types are `G-ENGINE`. More than 64 hits returns the first cell |
| `EntireRow`, `EntireColumn`, `CurrentRegion` | get | |
| `Clear`, `ClearContents`, `ClearFormats`, `ClearComments`, `Merge`, `UnMerge` | call | |
| `Select`, `Activate`, `Copy`, `PasteSpecial`, `Insert`, `Delete`, `Sort`, `AutoFilter` | call | |
| `Find` | call | `Nothing` when there is no hit |
| `Replace` | call | `0` or `-1` |
| `AddComment`, `Comment`, `Hyperlinks`, `Validation`, `FormatConditions`, `Areas` | | |
| `SortFields` | get | `G-ENGINE` |

## Font, Interior, Borders

| Member | | |
|---|---|---|
| `Font.Name`, `Size`, `Bold`, `Italic`, `Underline`, `Color`, `ColorIndex`, `ThemeColor`, `TintAndShade` | rw | |
| `Interior.Color`, `Pattern`, `ColorIndex`, `ThemeColor`, `TintAndShade`, `PatternColor` | rw | |
| `Borders.Item`, `LineStyle` | | An edge is `xlEdgeLeft` 7, top 8, bottom 9, right 10, inside vertical 11, inside horizontal 12 |
| `Border.LineStyle`, `Weight`, `Color` | rw | `xlNone` clears. Continuous, dash, dot, double, medium and thick map onto the stored line |

## Comments, hyperlinks, validation, conditional formats

| Member | | |
|---|---|---|
| `Comments.Item`, `Add` | call | A note `{author, text, visible}` |
| `Comment.Text`, `Author`, `Delete` | | |
| `Hyperlinks.Item`, `Add`, `Delete` | | |
| `Hyperlink.Address`, `SubAddress`, `Range`, `TextToDisplay`, `Delete` | | |
| `Hyperlink.Follow` | call | `G-HOST` |
| `Validation.Add`, `Delete` | call | |
| `Validation.Modify` | call | `G-ENGINE` |
| `FormatConditions.Item`, `Add`, `Delete` | | Classic cell-is and expression rules |
| `FormatCondition.Type`, `Formula1`, `Delete` | | |

## Lists, charts, shapes

| Member | | |
|---|---|---|
| `ListObjects.Item`, `Add` | call | `Add` uses the destination's first row as column names and ignores the source type |
| `ListObject.Name`, `Range`, `ListColumns`, `Delete`, `Resize` | | |
| `ListObject.Refresh` | call | `G-ENGINE` |
| `ListColumn.Name` | rw | |
| `ChartObjects.Item`, `Add` | call | `Add` creates a column chart drawing |
| `ChartObject.Name`, `Left`, `Top`, `Width`, `Height`, `Chart`, `Delete` | | Positions are points |
| `Chart.ChartType`, `Name` | | Known types: column clustered 51, bar clustered 57, line 4, pie 5, scatter `-4169`, area 1, doughnut `-4120`. Any other type is `G-ENGINE` |
| `Chart.SeriesCollection`, `Axes`, `SetSourceData` | | `G-ENGINE` |
| `Shapes.Item`, `AddTextbox` | | `AddTextbox` ignores orientation |
| `Shapes.AddPicture` | call | `G-HOST` |
| `Shape.Name`, `Left`, `Top`, `Width`, `Height`, `TextFrame`, `Delete` | | |
| `Shape.OLEFormat` | get | `G-ENGINE` |
| `TextFrame.TextRange`, `TextRange.Text` | | |

## Page setup, breaks, outline, filter, scenarios, pivots

| Member | | |
|---|---|---|
| `PageSetup` orientation, zoom, fit, margins, print area, centering | rw | Margins are inches. `PrintOut` is `G-HOST` |
| `HPageBreaks` / `VPageBreaks` `Item`, `Add` | | `Add` stores the 1-based row or column |
| `PageBreak.Location`, `Delete` | | |
| `Outline.SummaryRow`, `ShowLevels` | | |
| `AutoFilter.Range`, `ShowAllData` | | |
| `Scenarios.Item`, `Add` | | |
| `Scenario.Name`, `Comment`, `Show`, `Delete` | | |
| `Scenario.ChangeScenario` | call | `G-ENGINE` |
| `PivotTables.Item` | call | Existing tables only |
| `PivotTable.Name` | get | |
| `PivotTable.Refresh`, `PivotFields` | | `G-ENGINE` |

## Constants

`xlSheetVisible` `-1`, `xlSheetHidden` `0`, `xlSheetVeryHidden` `2`. `xlCalculationAutomatic` `-4105`, `xlCalculationManual` `-4135`, `xlCalculationSemiautomatic` `2`. `xlA1` `1`, `xlR1C1` `-4150`. `xlUp` `-4162`, `xlDown` `-4121`, `xlToLeft` `-4159`, `xlToRight` `-4161`. `xlWhole` `1`, `xlPart` `2`. `xlByRows` `1`, `xlByColumns` `2`. `xlAscending` `1`, `xlDescending` `2`, `xlGuess` `0`, `xlYes` `1`, `xlNo` `2`.

`xlCellTypeConstants` `2`, `xlCellTypeFormulas` `-4123`, `xlCellTypeBlanks` `4`, `xlCellTypeLastCell` `11`, `xlCellTypeVisible` `12`, `xlCellTypeComments` `-4144`. `xlNumbers` `1`, `xlTextValues` `2`, `xlLogical` `4`, `xlErrors` `16`.

`xlPasteAll` `-4104`, `xlPasteValues` `-4163`, `xlPasteFormulas` `-4123`, `xlPasteFormats` `-4122`, `xlPasteComments` `-4144`. `xlShiftDown` `-4121`, `xlShiftToRight` `-4161`, `xlShiftUp` `-4162`, `xlShiftToLeft` `-4159`.

`xlContinuous` `1`, `xlDash` `-4115`, `xlDot` `-4118`, `xlDouble` `-4119`, `xlNone` `-4142`, `xlLineStyleNone` `-4142`. `xlHairline` `1`, `xlThin` `2`, `xlMedium` `-4138`, `xlThick` `4`. Edge indexes are under Borders. `xlUnderlineStyleNone` `-4142`, `xlUnderlineStyleSingle` `2`. `xlPatternNone` `-4142`, `xlPatternSolid` `1`. `xlPortrait` `1`, `xlLandscape` `2`.

`xlCellValue` `1`, `xlExpression` `2`. `xlEqual` `3`, `xlNotEqual` `4`, `xlGreater` `5`, `xlLess` `6`, `xlGreaterEqual` `7`, `xlLessEqual` `8`. `xlBetween` `1`, `xlNotBetween` `2`. `xlValidateList` `3`, `xlValidateWholeNumber` `1`, `xlValidateDecimal` `2`, `xlValidateDate` `4`, `xlValidateTime` `5`, `xlValidateTextLength` `6`, `xlValidateCustom` `7`.

Chart types are under Chart. Theme colours are under Rules. `xlColorIndexAutomatic` `-4105`, `xlColorIndexNone` `-4142`. `msoTrue` `-1`, `msoFalse` `0`, `msoTextOrientationHorizontal` `1`.

## Worked calls

```text
ActiveSheet.Range("A1").Value = "Item"
ActiveSheet.Range("B2").Value = 12.5
ActiveSheet.Range("B4").Formula = "=SUM(B2:B3)"
block = ActiveSheet.Range("B2:B3").Value
ActiveSheet.Range("A1:B1").Font.Bold = True
ActiveSheet.Range("B2:B4").NumberFormat = "0.00"
ActiveWorkbook.Names.Add "Total", "=Sheet1!$B$4"
ActiveSheet.Range("A6").Value = WorksheetFunction.Sum(ActiveSheet.Range("B2:B3"))
```

`ActiveWorkbook.FullName` is `G-HOST`. `ActiveWorkbook.Creator` is `G-ENGINE`. The same text is `tests/ledger/vb-playground.bas`.
