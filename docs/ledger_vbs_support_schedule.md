# Ledger VB support schedule

**Status:** P1, P2, P3 M1–M5, P4M1, P4M2, P5 and document macros are implemented. P6 has not been started. P1 is `public/ledger/js/dom-ir.js` and [ledger_dom_ir_specifications.md](ledger_dom_ir_specifications.md). P2 is `public/ledger/js/vb-host.js` and the headless entry `third_party/VB6/src/runtime/engine-entry.js`. P3 is [ledger_dom_ir_porting_milestones.md](ledger_dom_ir_porting_milestones.md). P4M1 is `evalVB6`. P4M2 is `runVB6`. P5 is Help ▸ Ledger DOM Console.

**Goal.** Run VB6 source against the open workbook. Document calls go through a JavaScript DOM IR into the object model in `public/ledger/js/dom.js`. A later VBS Form uses the same VB6 engine; its window is a host slot left empty.

The language is the VB6 engine in `third_party/VB6` (MIT). Call shape matches the Excel VBA surface in [ledger_dom_api_specifications.md](ledger_dom_api_specifications.md). This is not a separate VBScript interpreter, and it is not a WebAssembly port of that engine.

## Decisions

- **Engine.** `VirtualMachine` from the fork, constructed directly. User VB is compiled to that VM's instructions. It is not passed to JavaScript `eval`. The built engine bundle is gitignored (`public/ledger/vendor/vb6/`). The submodule stays the source. `tools/vb6/build-engine.mjs` writes the Quire, Lectern and Ledger bundles. Quire and Lectern paths stay where they are.
- **Embed.** Ledger does not mount `ApplicationHost`. Ledger already has windows and a grid.
- **IR.** A handle protocol. The executor is a Ledger copy of the Lectern executor, plus the `array` value and the `both` member kind. It is not moved into `public/common/`.
- **Gaps.** Only `G-ENGINE` and `G-HOST`. There is no story index and no character index. `Creator` is `G-ENGINE` on every object. Objects listed with no members return that object's gap. The Excel reference is not mined for a longer member list.
- **Arrays.** A block `Value` or `Formula` is kind `array`. JavaScript stays 0-based. VB uses a 1-based `VBArray`. The engine exports `VBArray`. One cell stays a scalar.
- **One host, two VMs.** `evalVB6` and the DOM Console share one Immediate VM. `runVB6` compiles one standard module onto a second VM on that host and the same IR, then stops that VM. The Immediate VM and its locals stay. The two VMs do not run at once. Nested `eval` or `run` while a script is inside an IR call rejects with "A VB script is running".
- **Visual Basic IDE.** View ▸ Visual Basic IDE (`public/common/vb-ide.js`) debugs one standard module, Module1, on a VM created by `debugBegin`. Start Run does not install breakpoints. The window calls the virtual machine; it does not mount VB6 Studio, and it is not P6. Step Into in Macros stays disabled. Completion (`public/common/vb-complete.js`) lists names the caret's procedure can see, and the procedures and modules this virtual machine actually runs. Keywords are listed anywhere except after a dot. A string, a comment or a date literal does not.
- **Document macros.** A third VM on that host, separate from Immediate and from `runVB6`. Tools ▸ Macros reads and writes VB6 source. Current Document is the custom XML part `urn:vibeoffice:macros` on the workbook. All templates and documents is `vibeoffice.ledger.allMacros` in this browser, not in the file. The modules stay loaded so they can call each other. Run Script from the console is still `runVB6` and does not see them. `vbaProject.bin` is not listed or run. `Application.Run` is `G-HOST`.
- **Two surfaces.** Document objects are Automation adapters whose `invoke` emits IR. Form controls would be the engine's `__control` objects. Controls are not IR handles. No form is created.

## Layers

```text
DOM Console / evalVB6
        → VirtualMachine (third_party/VB6)
        → LedgerVbHost
              automation  → DOM IR handles → public/ledger/js/dom.js → workbook operations
              createForm  → empty until P6
```

`Application`, `ActiveWorkbook`, `ThisWorkbook`, `Workbooks`, `ActiveSheet`, `ActiveCell`, `Selection`, `WorksheetFunction` and the `xl*` / `mso*` constants are injected as the standard module `LedgerDOM` on every Ledger VB project. In the console, `ThisWorkbook` is the active workbook.

## Host slots

| Slot | Now | P6 |
|---|---|---|
| `automation` | Workbook factory. `invoke` only emits IR | Same factory |
| `print`, `error` | Console and `evalVB6`. `Debug.Print` uses the print sink | Same sink |
| `createForm` | Absent | A Ledger pane |
| `beginModal`, `msgBox`, `inputBox`, `openFile`, `saveFile` | Absent | Ledger dialogs, and only an explicit user choice for files |
| `loadProject` | Absent | Compiles a form project |

The engine `Clipboard` is not connected. Copy and paste of cells stay on `Range` through the IR. The VM's own `Err`, `Debug` and `Clipboard` stay in place.

## Sessions

- **Immediate.** `prepareImmediateContext`. No `Sub Main`. The debugger evaluation cap applies. The same frame is reused. A `Sub` typed in the console does not become a project procedure. A line ending in `_` continues. An unfinished `Sub`, `Function`, `If` or `With` stays open.
- **Running.** `runVB6` calls `Sub Main` or a named public procedure on a second VM, then stops it. A missing procedure is `VBError` 35. The instruction budget is the engine default.

## Phases

### P1. JavaScript DOM IR

`public/ledger/js/dom-ir.js`. `get`, `set`, `call` and `item`. Results are a scalar, a handle, nothing, empty, an array, or an error. Implemented members call `dom.js`. Gap members return the inventory's code.

### P2. VB6 engine and the form placeholder

Headless bundle `public/ledger/vendor/vb6/vb6-engine.js`, gitignored. `LedgerVbHost`. Form slots stay null. No `ApplicationHost`.

### P3. M1–M5

The milestones in [ledger_dom_ir_porting_milestones.md](ledger_dom_ir_porting_milestones.md), all in `public/ledger/js/dom.js`.

### P4. evalVB6 and runVB6

`window.evalVB6` and `window.runVB6`, also `L.evalVB6` and `L.runVB6`. `?` and `Print` return a value. Any other Immediate line returns undefined.

### P5. DOM Console

Help ▸ Ledger DOM Console, and the same command on the Paper 2016 ribbon File menu. The script editor defaults to Visual Basic 6. Run Script runs `runVB6`. Debug.Print is written to the Output box when the statement runs, so the line is visible while a message box is still open. An empty buffer opens on the sample in `tests/ledger/vb-playground.bas`. The page is generated by `tools/ledger/build/make.py`.

### Document macros

`public/ledger/js/macros.js` and the shared window in `public/common/macros.js`. The command first shows Notice on Macros in Ledger. OK opens the window. Current Document and All templates and documents are the two lists. `vbaProject.bin` stays unread.

### P6. Forms

Not started. `docs/ledger_vbs_form.md` is not written. No form window, no controls, no `MsgBox`.

## Files

| File | Role |
|---|---|
| `public/ledger/js/dom.js` | The object model |
| `public/ledger/js/dom-ir.js` | The handle protocol |
| `public/ledger/js/vb-host.js` | `LedgerVbHost`, `evalVB6`, `runVB6` |
| `public/ledger/js/dom-console.js` | Help ▸ Ledger DOM Console |
| `public/ledger/vendor/vb6/vb6-engine.js` | Gitignored bundle |
| `tests/ledger/vb-playground.bas` | The sample `Sub Main` |
| `tests/ledger/dom.test.js`, `tests/ledger/vb-playground.test.js` | The object model, the IR and the sample |

## Out of scope

Moving the executor into `public/common/`. WebAssembly. Mounting VB6 Studio. Rewriting workbook operations to fake unlisted members. Running `vbaProject.bin`. Translating UI strings. Quire and Lectern hosts and bundles stay where they are.
