# Lectern VB support schedule

**Status:** P1, P2, P3 M1–M5, P4M1, P4M2 and P5 are implemented. Document macros are implemented. P6 has not been started. P1 is `public/lectern/js/dom-ir.js` and [lectern_dom_ir_specifications.md](lectern_dom_ir_specifications.md). P2 is `public/lectern/js/vb-host.js` and the headless bundle from `node tools/vb6/build-engine.mjs`. P3 is [lectern_dom_ir_porting_milestones.md](lectern_dom_ir_porting_milestones.md). P4M1 is `evalVB6`. P4M2 is `runVB6`. P5 is the DOM Console.

**Goal.** Run VB6 source against the open Lectern presentation. Calls go through a JavaScript DOM IR into `public/lectern/js/dom.js`. A later form would use the same VB6 engine. Its window is a host slot left empty. This schedule does not add a form document.

The language is the VB6 engine in `third_party/VB6` (the same engine Quire uses). Call shape matches the PowerPoint surface in [lectern_dom_api_specifications.md](lectern_dom_api_specifications.md). This is not a separate VBScript interpreter, and it is not a WebAssembly port of that engine.

## Decisions

- **Engine.** `VirtualMachine` from `third_party/VB6`, constructed directly. User VB is compiled to that VM's instructions. It is not passed to JavaScript `eval`. The built bundle is gitignored at `public/lectern/vendor/vb6/`. `build-engine.mjs` also writes Quire's bundle and still returns that path. The two files are copies. Quire's host is not moved.
- **Embed.** Lectern does not mount `ApplicationHost`. That host builds its own stage. Lectern already has a presentation surface.
- **IR.** A handle protocol. The executor calls `dom.js` for members that module implements, and returns the specification's gap for the rest. It does not lay out slides or invent a character index. The executor is a copy of Quire's, with Lectern's catalog. It is not shared code in `public/common/`.
- **Gaps.** Objects listed without per-member rows are registered by object name. Any member returns that object's gap code (`G-ENGINE` or `G-HOST`).
- **One host, three VMs.** `evalVB6` and the DOM Console share one Immediate-mode VM. `runVB6` compiles one standard module onto a second VM on that host and the same IR, then stops that VM. The Immediate VM and its locals stay.
- **Document macros.** A third VM on that host. Tools ▸ Macro ▸ Macros, and View ▸ Macros on the Paper 2016 ribbon, read and write VB6 source. Current Document is the custom XML part `urn:vibeoffice:macros` on `ppt/presentation.xml`. All templates and documents is `vibeoffice.lectern.allMacros` in this browser. The project gives each macro module another name: this engine treats a module's name as an object, and that object would hide a procedure of the same name, so the other macros call the Sub by the macro's name. The modules stay loaded so they can call each other. Run Script from the console is still `runVB6` and does not see them. Opened from Macros, that window shows the macro name read-only and can save the source; its Run compiles the buffer with the other saved macros on a machine that is then stopped. A notice titled “Notice on Macros in Lectern” comes first; OK opens the window, and the title's close button does not. “Never show this message again” is `lectern.macroNotice`. Step Into stays disabled. Office `vbaProject.bin` is not listed or run. A presentation with no macros does not load the engine. A macro that fails to compile does not stop the presentation opening. Saving as HTML, PNG, PDF or text warns that the macros cannot be stored.
- **Two surfaces.** Presentation objects are Automation adapters whose `invoke` emits IR. Form controls would be the engine's `__control` objects, created only by `host.createForm`. That slot stays null.

## Layers

```text
DOM Console / evalVB6 / document macros
        → VirtualMachine (third_party/VB6)
        → LecternVbHost
              automation  → DOM IR handles → public/lectern/js/dom.js → presentation operations
              createForm  → empty until a form plan; not started
```

`Application`, `ActivePresentation`, `Presentations`, `ActiveWindow` and the `pp*` / `mso*` constants are injected as the standard module `LecternDOM` on every Lectern VB project. The properties read the current presentation through one Automation object.

## Host slots

| Slot | P2 | Later form |
|---|---|---|
| `automation` | Presentation factory. `invoke` only emits IR | Same factory |
| `print`, `error` | Console and `evalVB6`. `Debug.Print` uses `print` | Same sink |
| `createForm` | Null | Not started |
| `beginModal`, `msgBox`, `inputBox` | Null | Not started |
| `openFile`, `saveFile` | Null | Not started |
| `loadProject` | Null | Not started |

P2 does not connect the engine clipboard to Lectern's clipboard. Copy and paste of shapes and text stay on `Selection` through the IR.

## Sessions

`L.vb.open()` builds a host and an Immediate VM. `L.vb.session()` returns the shared host, starting it if needed. `evalVB6(text)` runs one Immediate line on that host. A line that starts with `?` or `Print` returns the value. Any other line returns `undefined`. `runVB6(source, entry, args)` compiles `source` as a standard module. With no entry it calls `Sub Main`. A missing procedure is `VBError` number 35.

Closing the host stops the VM and releases the IR. The presentation stays. The next `evalVB6` starts a new host, so Immediate locals do not survive `close`.

A script and an Immediate line do not run at the same time. A nested `eval` or `run` while a script is inside an IR call rejects with "A VB script is running".

## Phases

### P1. JavaScript DOM IR

Implemented. `tests/lectern/dom-ir.test.js` gets `ActivePresentation`, writes the title text, reads `Characters.Count`, and checks `G-TEXT`, `G-HOST`, `G-ENGINE`, `readonly`, `absent` and `released`.

### P2. VB6 engine and the form placeholder

Implemented. `node tools/vb6/build-engine.mjs` writes the bundle. `tests/lectern/vb-host.test.js` checks the bundle has no `ApplicationHost` or `BrowserForm`, the form slots are null, `formInstances` is empty, and one Immediate assignment writes the title.

### P3. VB port of the DOM specification

Implemented. M1 through M5, as [lectern_dom_ir_porting_milestones.md](lectern_dom_ir_porting_milestones.md).

### P4M1. `evalVB6`

Implemented. `tests/lectern/vb-eval.test.js`. One Immediate session. A local survives a later line. `FullName` is `G-HOST`. After `close`, a new session does not see the old local.

### P4M2. `runVB6`

Implemented. `tests/lectern/vb-run.test.js`. `Twice(21)` is 42 and does not change the title. `Main` writes `"10"`. The Immediate local survives. A missing `Main` is number 35. `FullName` inside `Main` is `G-HOST`.

### P5. DOM Console

Implemented. Help ▸ Lectern DOM Console, and the same command on the File menu of the Paper 2016 ribbon. The window is modeless. The language starts as Visual Basic 6 and is remembered in `vibeoffice.lectern.domConsole.lang`. JavaScript still runs. Run Script stores its buffer in `vibeoffice.lectern.vbScript`. The default sample sets the title placeholder and points at the PowerPoint VBA object model. After a line or a script, the slide is drawn again. Debug.Print is written to the Output box when the statement runs, so the line is visible while a message box is still open.

`tests/lectern/vb-milestones-sample.js` is pasted into that window with the language set to JavaScript, then `await runVbMilestones()`. It is not a Node test. Node runs the `vb-m*.test.js` files.

### Document macros

Implemented. `public/lectern/js/macros.js` stores Current Document in `urn:vibeoffice:macros` and All templates and documents in `vibeoffice.lectern.allMacros`. `tests/lectern/macros.test.mjs` round-trips the part on `ppt/presentation.xml`, undoes Save Macro and Delete Macro, and keeps the browser list off Quire's key. `tests/lectern/vb-macros.test.js` runs two macros that call each other, a draft that does not replace the saved project, and a module variable that survives a second run.

### P6. VBS Form

Not started. No `docs/lectern_vbs_form.md`. The form slots stay null. Do not fill them from this schedule.

## Files

| File | |
|---|---|
| `public/lectern/js/dom.js` | The object model |
| `public/lectern/js/dom-ir.js` | Catalog and executor |
| `public/lectern/js/vb-host.js` | Immediate VM, `runVB6`, and the document-macro VM |
| `public/lectern/js/macros.js` | Current Document and All templates and documents |
| `public/lectern/js/dom-console.js` | The console window |
| `public/lectern/index.html` | Scripts and the console styles |
| `public/lectern/js/commands.js`, `app.js`, `ribbon.js`, `dialogs.js` | Tools ▸ Macro, the View ribbon group, Alt+F8, and the notice |
| `tests/lectern/dom.test.js`, `dom-ir.test.js` | The object model and the IR without VB |
| `tests/lectern/vb-m1.test.js` … `vb-m5.test.js`, `vb-host.test.js`, `vb-eval.test.js`, `vb-run.test.js`, `vb-macros.test.js` | The VB checks |
| `tests/lectern/macros.test.mjs` | The custom XML part, undo, and the browser list |
| `tests/lectern/vb-milestones-sample.js` | The pasteable console script |
| `tools/vb6/build-engine.mjs` | Writes both gitignored bundles |

## Out of scope

`vbaProject.bin`, a WebAssembly port, `ApplicationHost`, Ledger, and rewriting `dom.js` behind a second API. Quire's host and Quire's bundle path stay where they are.
