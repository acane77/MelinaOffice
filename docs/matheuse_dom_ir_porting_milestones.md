# Matheuse DOM IR porting milestones

What the VB host can call after each milestone. The inventory is [matheuse_dom_api_specifications.md](matheuse_dom_api_specifications.md). The protocol is [matheuse_dom_ir_specifications.md](matheuse_dom_ir_specifications.md). The schedule is [matheuse_vbs_support_schedule.md](matheuse_vbs_support_schedule.md).

**Status:** M1 through M5 are implemented in `public/matheuse/js/dom.js` and `public/matheuse/js/dom-ir.js`. `tests/matheuse/dom.test.js` and `tests/matheuse/vb-playground.test.js` cover the worked calls, a block array through VB, and the two gap codes.

## What a port is

Each milestone adds members to the same session. A member from a later milestone is absent until that milestone lands. A gap member is present and returns its code. The executor does not grow the inventory.

## What VB sees

The injected module is `MatheuseDOM`. It publishes `Application`, `ActiveWorkbook`, `ThisWorkbook`, `Workbooks`, `ActiveSheet`, `ActiveCell`, `Selection`, `WorksheetFunction`, and the `xl*` and `mso*` constants in the inventory. `ThisWorkbook` is the active workbook.

`evalVB6` and the DOM Console share one Immediate VM. `runVB6` compiles the module `Script` plus `MatheuseDOM` on a second VM, calls `Sub Main` or a named public procedure, and stops that VM. A missing procedure is `VBError` 35. The two VMs do not run together. An IR error is `VBError` 438 with `err.code`.

## M1. Identity

`Application`, `Workbooks`, `Workbook`, `Sheets`, `Worksheets` and `Worksheet`: name, index, visibility, activate, add, delete, move and copy. `Range` is handed out (`Range`, `Cells`, `Rows`, `Columns`, `Item`) and can be selected. Values are later.

## M2. Values and structure

`Value`, `Value2`, `Formula`, `FormulaR1C1` and block arrays. Number format, font, fill and borders. Row height, column width and `Hidden`. `Clear*`, merge, insert and delete. `Text` reads the formatted value.

## M3. Names and the sheet tools

Defined names. Styles by name. Comments, hyperlinks, `Find` and `Replace`. Validation. AutoFilter. Excel 2003 `Sort`. Classic conditional formats. `WorksheetFunction` for names in the function catalog.

## M4. Lists, drawings and page setup

`ListObjects`. Thin `ChartObjects` and `Shapes` (position, name, text, chart type). Page setup, page breaks, scenarios, protection and outline. `PrintOut` stays `G-HOST`. `Shapes.AddPicture` stays `G-HOST`.

## M5. Pivots and the remaining gaps

`PivotTables` lists tables that are already on the workbook. `Refresh` and `PivotFields` are `G-ENGINE`. Query tables, connections, slicers, sparklines, OLE, threaded comments, phonetics, the data model, windows, add-ins, dialogs and command bars are the gap objects in the inventory. Any other member is their gap code.
