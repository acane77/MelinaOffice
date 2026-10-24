# Quire VB support schedule

**Status:** P1, P2, P3 M1–M5, P4M1, P4M2, P5 and P6 are implemented. P1 is `public/quire/js/dom-ir.js` and [quire_dom_ir_specifications.md](quire_dom_ir_specifications.md). P2 is `public/quire/js/vb-host.js`, the headless entry `third_party/VB6/src/runtime/engine-entry.js`, and `make`. P3 is [quire_dom_ir_porting_milestones.md](quire_dom_ir_porting_milestones.md). P4M1 is `evalVB6`. P4M2 is `runVB6`. P5 is the DOM Console. P6 is [quire_vbs_form.md](quire_vbs_form.md): `public/common/vb-forms.js` on the same host, drawn with `ui.dialog`.

**Goal.** Run VB6 source against the open Quire document. Document calls go through a JavaScript DOM IR into the object model in `public/quire/js/dom.js`. A later VBS Form uses the same VB6 engine; its window is a host slot left empty until the last phase.

The language is the VB6 engine in `third_party/VB6` (submodule `git@github.com:acane77/VB6.git`, MIT). Call shape matches the Word VBA surface in [quire_dom_api_specifications.md](quire_dom_api_specifications.md). This is not a separate VBScript interpreter, and it is not a WebAssembly port of that engine.

## Decisions

- **Engine.** `VirtualMachine` from the fork (`src/runtime/vm.js`), constructed directly. User VB is compiled to that VM's instructions. It is not passed to JavaScript `eval`. The fork has no VB6 WebAssembly runtime: its WASM file is SQLite, and `packages/vb6-compute` emits WGSL for a typed subset that rejects objects. Those are not the document engine. The built engine bundle is gitignored; the submodule stays the source.
- **Embed.** Quire does not mount `ApplicationHost` (`src/runtime/host.js`). That host builds its own stage, theme, MDI and WebGPU or Canvas session. Quire already has windows and a document surface.
- **IR.** A handle protocol. The IR executor calls `dom.js` for members that module implements, and returns the specification's gap for the rest. It does not lay out pages or invent story indexes.
- **Gaps.** Objects that the specification lists without per-member rows (content controls, windows, dialogs, mail merge, equations, and the others in that section) are registered by object name. Any member returns that object's gap code (`G-ENGINE` or `G-HOST`). The Word reference is not mined for a longer member list.
- **One host, two VMs.** `evalVB6` and the DOM Console share the host and one Immediate-mode VM. `runVB6` compiles one standard module onto a second VM on that host and the same IR, then stops that VM. The Immediate VM and its locals stay. Each VM has its own Automation object table. Loading a form project (P6) replaces the project VM and keeps the host. The document lives in `dom.js`, so it survives the replacement. Immediate-frame locals do not survive a host replacement.
- **Visual Basic IDE.** View ▸ Visual Basic IDE (`public/common/vb-ide.js`) debugs one standard module, Module1, on a VM created by `debugBegin`. Start Run does not install breakpoints. The window calls the virtual machine; it does not mount the VB6 Studio IDE, and it is not P6. Step Into in Macros stays disabled. Completion (`public/common/vb-complete.js`) lists names the caret's procedure can see, and the procedures and modules this virtual machine actually runs. Keywords are listed anywhere except after a dot. A string, a comment or a date literal does not.
- **Document macros.** A third VM on that host, separate from Immediate and from `runVB6`. Tools ▸ Macros reads and writes VB6 source in the document's custom XML part `urn:vibeoffice:macros` (one standard module and one public Sub of the macro's name). The project gives that module another name: this engine treats a module's name as an object, and that object would hide a procedure of the same name, so the other macros call the Sub by the macro's name. The modules stay loaded so they can call each other. Run Script from the console is still `runVB6` and does not see them. Opened from Macros, that window shows the macro name read-only and can save the source; its Run compiles the buffer with the other saved macros on a machine that is then stopped. Step Into stays disabled. Office `vbaProject.bin` is not listed or run. A document with no macros does not load the engine. A macro that fails to compile does not stop the document opening. Opening Macros in Quire shows the notice first, unless it was dismissed with Never show this message again. Ledger's Macros command shows that workbooks keep their macros and Ledger does not run them. Lectern has no Macros command.
- **Two surfaces.** Document objects are Automation adapters whose `invoke` emits IR. Form controls are the engine's `__control` objects, created only by `host.createForm`. Controls are not IR handles and do not paint the document.

## Layers

```text
DOM Console / evalVB6 / a form event
        → VirtualMachine (third_party/VB6)
        → QuireVbHost
              automation  → DOM IR handles → public/quire/js/dom.js → document operations
              createForm  → public/common/vb-forms.js, drawn with ui.dialog
```

`Application`, `ActiveDocument`, `Documents`, `Selection` and the `wd*` constants from the specification are injected as a standard module on every Quire VB project, including a later form project. The properties read the current document through one Automation object. They are not frame locals and they are not VB controls.

## Host slots

The host object is created in P2 and kept. P6 fills the form slots on that same object.

| Slot | P2 | P6 |
|---|---|---|
| `automation` | Document factory. `invoke` only emits IR. | Same factory. A button and the console edit one document. |
| `print`, `error` | Console and `evalVB6`. | Same sink. `Debug.Print` uses it. |
| `createForm` | Absent. `attachForm` then does nothing. | A VibeOffice dialog. Returned controls become form fields. |
| `beginModal` | Absent. | Other VB forms are marked inert. |
| `msgBox`, `inputBox` | Absent. | Suite dialogs (`ui.msg`, `ui.prompt`). |
| `openFile`, `saveFile` | Absent. | Still null. |
| Virtual filesystem, Win32 registry | The engine's private in-memory storage. | Still private. Document bytes are not stored there. |
| `loadProject` | Reserved, not implemented. | Compiles the form project, builds a running VM on this host, and calls `start()`. |

P2 does not set `dataFetch`, and it does not connect the engine `Clipboard` to Quire's document clipboard. Copy and paste of document content stay on `Selection` / `Range` through the IR.

The VM always constructs its own `Err`, `App`, `Screen`, `Forms`, `Debug`, `Clipboard`, Win32 bridge and `DataContext`. P2 leaves those in place and does not strip `vm.js` to a second interpreter.

## Sessions

The VM has two modes and they exclude each other:

- **Immediate.** `prepareImmediateContext`. No `Sub Main`, no startup form. Events are not delivered until `configureImmediateEvents`. `start()` is rejected. `vm.immediate` is the console and `evalVB6`. The debugger evaluation cap applies (the engine default is 100,000 instructions and 5 seconds). The same frame is reused, so a later line sees an immediate local such as `x = 1`. A `Sub` or `Function` typed in the console does not become a project procedure.
- **Running.** `start()`. Startup form or `Sub Main`, event queue, VM instruction budget (the engine default is 5,000,000). `vm.immediate` works when the stack is empty or the VM is paused. Form events use this budget, not the console cap.

P2 uses Immediate only. P4M2 runs one standard module on a second VM and then stops it; it does not call `prepareImmediateContext` and it does not replace the Immediate VM. The two VMs do not run at once. P6's `loadProject` switches to Running on a new VM with the same host.

## Phases

Work proceeds in order. A later phase starts only when the previous phase's deliverable is accepted. Inside P3, each milestone is a separate step. P6 writes its own plan and waits for acceptance before form code.

### P1. JavaScript DOM IR

**Deliverables.** [quire_dom_ir_specifications.md](quire_dom_ir_specifications.md). A JavaScript IR API that sits on `public/quire/js/dom.js`.

**Scope.** The full inventory in the DOM specification, including members `dom.js` does not implement. The catalog is that specification: status and gap codes are forwarded, not decided again.

A call is `get`, `set`, `call` or `item`, with an object handle, a member name and arguments in VBA order. The result is a scalar, a new handle, or an error carrying `G-STORY`, `G-ENGINE` or `G-HOST`. Omitted optional arguments stay Missing. Collections stay 1-based. `Paragraphs(1)` is the default member. `Range.Text = "Hello"` is a property write. VBA `True` / `False` cross the boundary as `-1` / `0`; the executor passes what the existing member already accepts.

Implemented members call the live `dom.js` object. A gap member is still present and returns the specification's code. The executor does not reimplement layout, undo or file dialogs.

**Done when.** Every specification member has an IR operation. The worked calls in the specification, run through the IR, match a direct `dom.js` call. Gap operations return the documented code. `ActiveDocument.Range(0, 5)` stays `G-STORY`.

### P2. VB6 engine and the form placeholder

**Deliverables.** A headless engine entry in the fork (compiler, `VirtualMachine`, values, `AutomationRegistry`, and the modules `vm.js` already imports). A Quire build of that entry. `QuireVbHost` with the slots in the table above. The injected standard module, with the document properties wired as far as the IR allows and the `wd*` constants from the specification.

**Not in this phase.** `ApplicationHost`, controls, `BrowserForm`, the IDE, themes, the WebGPU application renderer, OCX, and `vb6-compute`. No form window. `createForm`, `beginModal`, `msgBox`, `inputBox`, `openFile`, `saveFile` and `loadProject` stay absent or reserved. The engine bundle is not committed.

**Build.** The fork's bundler emits one browser script, `public/quire/vendor/vb6/vb6-engine.js` (`globalThis.VB6Engine`). That directory is gitignored. `make` rebuilds it from the submodule (`tools/vb6/build-engine.sh`). Nothing is fetched from a third-party server. Constructing the host inserts that file with a classic script tag. Quire does not request it at page load, so a missing build does not fail every page. P6 may add `vb6-forms.js` in that same directory; P2 does not build it.

**Done when.** A Quire page can construct the host and an Immediate-mode VM, evaluate one VB statement that reads or writes the document through the IR, and shut the VM down without a form surface or an `ApplicationHost` stage. The host object exposes the empty form slots as named absences, so P6 fills them instead of replacing the host.

### P3. VB port of the DOM specification

**Deliverable first.** [quire_dom_ir_porting_milestones.md](quire_dom_ir_porting_milestones.md), then implementation one milestone at a time.

The port is VB that runs on the P2 engine and emits IR. It does not skip a member because the document engine has no operation for it. Each milestone ports the whole specification surface of its objects, gaps included.

| Milestone | Surface |
|---|---|
| M1 | IR values, handles and errors as seen from VB; the `wd*` constants; `Application`, `Documents`, `Document`; `Application`, `Parent` and `Creator` on every object |
| M2 | `Selection`, `Range`; `Paragraphs`, `Sentences`, `Words`, `Characters`; story-index members, returning `G-STORY` |
| M3 | `Font`, `ParagraphFormat`, `TabStops`, `Styles`; `ListFormat`, `Lists`; `Sections`, page setup, headers and footers |
| M4 | Tables, borders, shading; `Find`; bookmarks, fields, hyperlinks; comments, footnotes, endnotes |
| M5 | Shapes and revisions; the specification's gap-only objects, by object name and gap code |

**Done when.** The milestones document is accepted, and the accepted milestones' members are callable from VB through the IR. A milestone that is not accepted is not started.

### P4M1. `evalVB6`

**Deliverable.** `evalVB6` on the page, sharing the P2 host and Immediate VM with the console.

An expression returns the VB value. A statement returns Empty. Compile errors and gap errors reject. Calls from the developer tools and from the DOM Console are the same `vm.immediate`. This phase does not compile `Sub`, `Function`, `For`, `Do`, `While`, `If`, `Select Case` or `With`, and it does not start a form.

**Done when.** A snippet typed in the developer console performs a document edit that the IR already supports, and a following call sees the Immediate locals and the injected document properties.

### P4M2. `runVB6`

**Deliverable.** `runVB6` on the page, on the same host and the same DOM IR as `evalVB6`. The engine stays the gitignored bundle from `make`. No form, no control, no class module, no `Declare`, no Win32 call, no file dialog and no IDE. `createForm`, `beginModal`, `msgBox`, `inputBox`, `openFile`, `saveFile` and `loadProject` stay null. `settings.xaml` and `settings.anchoring` stay unset.

The source is one standard module named `Script`, compiled with the injected `QuireDOM` module. `Script` is not named `QuireDOM`. Public constants and public property gets on `QuireDOM` are visible in `Script`, including `ActiveDocument`, `Selection` and the `wd*` constants. A public property get on another standard module is resolved by the VM the same way a public function is.

`await runVB6(code)` compiles the module and runs `Sub Main`. It returns Empty. `await runVB6(code, "Twice", [21])` does not run `Main`; it calls that public procedure and returns its value to JavaScript. No `Main` and no procedure name is the engine's startup error (number 35). A missing or private procedure is the same number. Compile errors, runtime errors and `G-STORY`, `G-ENGINE` and `G-HOST` reject the promise, with `err.code` set for a gap.

The script VM is built with `debuggerEnabled` false and the engine's running budget (5,000,000 instructions). It is not the Immediate debugger cap. While it is running, another `runVB6` or `evalVB6` rejects. When the script finishes, that VM stops. The Immediate VM stays, so an Immediate local such as `scratch` is still there.

**Done when.** A module with a loop, a condition and a function edits the open document, a named public function returns its value without running `Main`, and an Immediate local set before the script still reads back after it.

### P5. DOM Console

**Deliverable.** Help ▸ Quire DOM Console. [quire.md](quire.md) describes it. The bar under the title has a Language box: Visual Basic 6 or JavaScript. Visual Basic 6 is the default, and the choice is kept. Choosing a language clears the terminal and prints that language's welcome.

JavaScript is the console that was already there. Its welcome adds a third tip: JavaScript runs here as well, but the F12 console is the better place to run a JavaScript script. That tip is translated. The other new console strings go through `L.t()` and the English catalog.

Visual Basic 6 types into the same terminal and runs with `evalVB6`. A statement returns Empty and is not printed. `?expr` and `Print` print the VB value (`True`, not the JavaScript `true`). A line ending in `_` continues. An unfinished `Sub`, `Function`, `If` or `With` stays open until it ends. Immediate locals carry to the next line. The console does not compile a procedure into the running program.

In Visual Basic 6 the bar also has Run Script. It opens a modeless window titled Run Visual Basic Script: a CodeMirror editor (MIT, `public/common/vendor/codemirror/`, VB mode), a read-only Output box, and Run and Close. Run calls `runVB6` on the whole buffer, so the script needs `Sub Main`. Output shows what the script prints and any error. Debug.Print is written there when the statement runs, so the line is visible while a message box is still open. A later Run replaces that output. The buffer and the latest output are kept in `localStorage` across a refresh. A buffer that was never edited, or that was empty when the window last closed, opens on a welcome `Sub Main`. The comments welcome the user, say to write Visual Basic and click Run, and link to the Word VBA object model. The procedure inserts one sentence at the cursor. A macro opened in the same window keeps its own source.

**Done when.** A VB statement in the console edits the open document and a following `?` reads it back, JavaScript still runs when chosen, Run Script runs a module with a loop, and the script text is still there after a refresh.

### P6. VBS Form

Implemented. [quire_vbs_form.md](quire_vbs_form.md). The same `QuireVbHost`. `loadProject` compiles the project, starts a running VM and keeps the Immediate VM. `createForm`, `beginModal`, `msgBox` and `inputBox` are installed for that run. `openFile` and `saveFile` stay null. The window is `ui.dialog` in `public/common/vb-forms.js`. `BrowserForm` is not loaded and is not folded into the engine script. Ledger and Lectern use that same file. A control event that edits the document calls the injected module, which emits IR. The form does not draw the document. View ▸ Run Form opens a `.frm`.

## Files this plan expects to add or touch

| Piece | Where | Phase |
|---|---|---|
| IR specification | `docs/quire_dom_ir_specifications.md` | P1 |
| IR API | Quire script beside `public/quire/js/dom.js` | P1 |
| Engine entry | `third_party/VB6/src/runtime/engine-entry.js` | P2 |
| Engine bundle | `public/quire/vendor/vb6/vb6-engine.js`, gitignored | P2 |
| Form runtime | `public/common/vb-forms.js`, [quire_vbs_form.md](quire_vbs_form.md) | P6 |
| Rebuild tool | `Makefile`, `tools/vb6/build-engine.sh` | P2 |
| Host and injected module | `public/quire/js/vb-host.js` | P2 |
| Porting milestones | `docs/quire_dom_ir_porting_milestones.md` | P3 |
| VB port | with the host, one milestone at a time | P3 |
| `evalVB6` | page global on the P2 Immediate session | P4M1 |
| `runVB6` | page global on the same host, second VM | P4M2 |
| Console | `public/quire/js/dom-console.js`, `docs/quire.md`, console tests | P5 |
| Form plan and host slots | `docs/quire_vbs_form.md`, `public/common/vb-forms.js` | P6 |

`docs/quire.md` is updated when the console or the form changes what a person can do. The DOM specification stays the inventory: an IR or VB member that the specification does not list is not added, and a gap closed in `dom.js` updates that specification in the same change.

## Out of scope

A WebAssembly build of the VM. Mounting the VB6 Studio IDE or `ApplicationHost` inside Quire. Rewriting `dom.js` behind the IR. Treating form controls as document ranges. Teaching the console to define project procedures. Translating UI strings. Ledger and Lectern.
