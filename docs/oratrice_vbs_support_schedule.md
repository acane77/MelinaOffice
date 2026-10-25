# Oratrice VB support schedule

**Status:** P1, P2, P3 M1–M5, P4M1, P4M2, P5, document macros and P6 are implemented. P1 is `public/oratrice/js/dom-ir.js` and [oratrice_dom_ir_specifications.md](oratrice_dom_ir_specifications.md). P2 is `public/oratrice/js/vb-host.js` and the headless bundle from `node tools/vb6/build-engine.mjs`. P3 is [oratrice_dom_ir_porting_milestones.md](oratrice_dom_ir_porting_milestones.md). P4M1 is `evalVB6`. P4M2 is `runVB6`. P5 is the DOM Console. P6 is [ecrivaine_vbs_form.md](ecrivaine_vbs_form.md): `public/common/vb-forms.js`, drawn with `ui.dialog`.

**Goal.** Run VB6 source against the open Oratrice presentation. Calls go through a JavaScript DOM IR into `public/oratrice/js/dom.js`. A later form would use the same VB6 engine. Its window is a host slot left empty. This schedule does not add a form document.

The language is the VB6 engine in `third_party/VB6` (the same engine Ecrivaine uses). Call shape matches the PowerPoint surface in [oratrice_dom_api_specifications.md](oratrice_dom_api_specifications.md). This is not a separate VBScript interpreter, and it is not a WebAssembly port of that engine.

## Decisions

- **Engine.** `VirtualMachine` from `third_party/VB6`, constructed directly. User VB is compiled to that VM's instructions. It is not passed to JavaScript `eval`. The built bundle is gitignored at `public/oratrice/vendor/vb6/`. `build-engine.mjs` also writes Ecrivaine's bundle and still returns that path. The two files are copies. Ecrivaine's host is not moved.
- **Embed.** Oratrice does not mount `ApplicationHost`. That host builds its own stage. Oratrice already has a presentation surface.
- **IR.** A handle protocol. The executor calls `dom.js` for members that module implements, and returns the specification's gap for the rest. It does not lay out slides or invent a character index. The executor is a copy of Ecrivaine's, with Oratrice's catalog. It is not shared code in `public/common/`.
- **Gaps.** Objects listed without per-member rows are registered by object name. Any member returns that object's gap code (`G-ENGINE` or `G-HOST`).
- **One host, three VMs.** `evalVB6` and the DOM Console share one Immediate-mode VM. `runVB6` compiles one standard module onto a second VM on that host and the same IR, then stops that VM. The Immediate VM and its locals stay.
- **Visual Basic IDE.** View ▸ Visual Basic IDE (`public/common/vb-ide.js`) debugs one standard module, Module1, on a VM created by `debugBegin`. Start Run does not install breakpoints. The window calls the virtual machine; it does not mount `ApplicationHost`, and it is not P6. Step Into in Macros stays disabled. Completion (`public/common/vb-complete.js`) lists names the caret's procedure can see, and the procedures and modules this virtual machine actually runs. Keywords are listed anywhere except after a dot. A string, a comment or a date literal does not.
- **Document macros.** A third VM on that host. Tools ▸ Macro ▸ Macros, and View ▸ Macros on the Paper 2016 ribbon, read and write VB6 source. Current Document is the custom XML part `https://acane.net/melinaoffice/macros` on `ppt/presentation.xml`. All templates and documents is `melinaoffice.oratrice.allMacros` in this browser. The project gives each macro module another name: this engine treats a module's name as an object, and that object would hide a procedure of the same name, so the other macros call the Sub by the macro's name. The modules stay loaded so they can call each other. Run Script from the console is still `runVB6` and does not see them. Opened from Macros, that window shows the macro name read-only and can save the source; its Run compiles the buffer with the other saved macros on a machine that is then stopped. A notice titled “Notice on Macros in Oratrice” comes first; OK opens the window, and the title's close button does not. “Never show this message again” is `oratrice.macroNotice`. Step Into stays disabled. Office `vbaProject.bin` is not listed or run. A presentation with no macros does not load the engine. A macro that fails to compile does not stop the presentation opening. Saving as HTML, PNG, PDF or text warns that the macros cannot be stored.
- **Two surfaces.** Presentation objects are Automation adapters whose `invoke` emits IR. Form controls are the engine's `__control` objects, created by `host.createForm` when a form project is loaded. Controls are not IR handles.

## Layers

```text
DOM Console / evalVB6 / document macros
        → VirtualMachine (third_party/VB6)
        → OratriceVbHost
              automation  → DOM IR handles → public/oratrice/js/dom.js → presentation operations
              createForm  → public/common/vb-forms.js, drawn with ui.dialog
```

`Application`, `ActivePresentation`, `Presentations`, `ActiveWindow` and the `pp*` / `mso*` constants are injected as the standard module `OratriceDOM` on every Oratrice VB project. The properties read the current presentation through one Automation object.

## Host slots

| Slot | P2 | Later form |
|---|---|---|
| `automation` | Presentation factory. `invoke` only emits IR | Same factory |
| `print`, `error` | Console and `evalVB6`. `Debug.Print` uses `print` | Same sink |
| `createForm` | Null until `loadProject` | A Melina Office dialog (`public/common/vb-forms.js`) |
| `beginModal`, `msgBox`, `inputBox` | Null until `loadProject` | Suite dialogs. Other VB forms are marked inert |
| `openFile`, `saveFile` | Null | Still null |
| `loadProject` | A method on the host. The other form slots stay null until it runs | Compiles a form project and calls `start()` |

P2 does not connect the engine clipboard to Oratrice's clipboard. Copy and paste of shapes and text stay on `Selection` through the IR.

## Sessions

`L.vb.open()` builds a host and an Immediate VM. `L.vb.session()` returns the shared host, starting it if needed. `evalVB6(text)` runs one Immediate line on that host. A line that starts with `?` or `Print` returns the value. Any other line returns `undefined`. `runVB6(source, entry, args)` compiles `source` as a standard module. With no entry it calls `Sub Main`. A missing procedure is `VBError` number 35.

Closing the host stops the VM and releases the IR. The presentation stays. The next `evalVB6` starts a new host, so Immediate locals do not survive `close`.

A script and an Immediate line do not run at the same time. A nested `eval` or `run` while a script is inside an IR call rejects with "A VB script is running".

## Phases

### P1. JavaScript DOM IR

Implemented. `tests/oratrice/dom-ir.test.js` gets `ActivePresentation`, writes the title text, reads `Characters.Count`, and checks `G-TEXT`, `G-HOST`, `G-ENGINE`, `readonly`, `absent` and `released`.

### P2. VB6 engine and the form placeholder

Implemented. `node tools/vb6/build-engine.mjs` writes the bundle. `tests/oratrice/vb-host.test.js` checks the bundle has no `ApplicationHost` or `BrowserForm`, the form slots are null, `formInstances` is empty, and one Immediate assignment writes the title.

### P3. VB port of the DOM specification

Implemented. M1 through M5, as [oratrice_dom_ir_porting_milestones.md](oratrice_dom_ir_porting_milestones.md).

### P4M1. `evalVB6`

Implemented. `tests/oratrice/vb-eval.test.js`. One Immediate session. A local survives a later line. `FullName` is `G-HOST`. After `close`, a new session does not see the old local.

### P4M2. `runVB6`

Implemented. `tests/oratrice/vb-run.test.js`. `Twice(21)` is 42 and does not change the title. `Main` writes `"10"`. The Immediate local survives. A missing `Main` is number 35. `FullName` inside `Main` is `G-HOST`.

### P5. DOM Console

Implemented. Help ▸ Oratrice DOM Console, and the same command on the File menu of the Paper 2016 ribbon. The window is modeless. The language starts as Visual Basic 6 and is remembered in `melinaoffice.oratrice.domConsole.lang`. JavaScript still runs. Run Script stores its buffer in `melinaoffice.oratrice.vbScript`. The default sample sets the title placeholder and points at the PowerPoint VBA object model. After a line or a script, the slide is drawn again. Debug.Print is written to the Output box when the statement runs, so the line is visible while a message box is still open.

`tests/oratrice/vb-milestones-sample.js` is pasted into that window with the language set to JavaScript, then `await runVbMilestones()`. It is not a Node test. Node runs the `vb-m*.test.js` files.

### Document macros

Implemented. `public/oratrice/js/macros.js` stores Current Document in `https://acane.net/melinaoffice/macros` and All templates and documents in `melinaoffice.oratrice.allMacros`. `tests/oratrice/macros.test.mjs` round-trips the part on `ppt/presentation.xml`, undoes Save Macro and Delete Macro, and keeps the browser list off Ecrivaine's key. `tests/oratrice/vb-macros.test.js` runs two macros that call each other, a draft that does not replace the saved project, and a module variable that survives a second run.

### P6. VBS Form

Implemented, shared with Ecrivaine and Matheuse. [ecrivaine_vbs_form.md](ecrivaine_vbs_form.md). View ▸ Run Form opens a `.frm`. The window is `ui.dialog`. `BrowserForm` is not loaded. `openFile` and `saveFile` stay null. A change to the presentation still goes through the DOM IR.

## Files

| File | |
|---|---|
| `public/oratrice/js/dom.js` | The object model |
| `public/oratrice/js/dom-ir.js` | Catalog and executor |
| `public/oratrice/js/vb-host.js` | Immediate VM, `runVB6`, and the document-macro VM |
| `public/oratrice/js/macros.js` | Current Document and All templates and documents |
| `public/oratrice/js/dom-console.js` | The console window |
| `public/oratrice/index.html` | Scripts and the console styles |
| `public/oratrice/js/commands.js`, `app.js`, `ribbon.js`, `dialogs.js` | Tools ▸ Macro, the View ribbon group, Alt+F8, and the notice |
| `tests/oratrice/dom.test.js`, `dom-ir.test.js` | The object model and the IR without VB |
| `tests/oratrice/vb-m1.test.js` … `vb-m5.test.js`, `vb-host.test.js`, `vb-eval.test.js`, `vb-run.test.js`, `vb-macros.test.js` | The VB checks |
| `tests/oratrice/macros.test.mjs` | The custom XML part, undo, and the browser list |
| `tests/oratrice/vb-milestones-sample.js` | The pasteable console script |
| `tools/vb6/build-engine.mjs` | Writes both gitignored bundles |

## Out of scope

`vbaProject.bin`, a WebAssembly port, `ApplicationHost`, Matheuse, and rewriting `dom.js` behind a second API. Ecrivaine's host and Ecrivaine's bundle path stay where they are.
