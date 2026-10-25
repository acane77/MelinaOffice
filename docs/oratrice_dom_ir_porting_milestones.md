# Oratrice DOM IR porting milestones

**Status:** M1 through M5 are implemented (`tests/oratrice/vb-m1.test.js` through `tests/oratrice/vb-m5.test.js`).

Gap-only objects are public fields of the injected module, filled from `session.named`. A gap type accepts any member name so the call reaches the IR. `Application` and `Parent` still navigate.

The inventory is [oratrice_dom_api_specifications.md](oratrice_dom_api_specifications.md). The call protocol is [oratrice_dom_ir_specifications.md](oratrice_dom_ir_specifications.md). The runtime is the host in `public/oratrice/js/vb-host.js`.

## What a port is

A VB statement on the Immediate VM names a member. The Automation adapter emits one IR operation. `dom.js` performs the member, or the IR returns that member's gap. The port does not add a second object model, and it does not generate a VB procedure per member. Each live wrapper is its own Automation object.

A milestone ports every inventory member of its objects, including gaps. A gap is callable and rejects. It is not omitted and it is not given new behaviour. A name the inventory does not list is not added. The port does not lay out slides or invent a character index.

## What VB sees

| IR result | From Immediate VB |
|---|---|
| Scalar | The value `dom.js` returned. Numbers, strings and booleans are not rewritten. VBA `True` and `False` are `-1` and `0` when that is what the live member accepts or returns. |
| Handle | An object. A later `.Member` is another IR operation. |
| Nothing | VB `Nothing`. |
| Empty | The call returns no value. |
| Error | The call rejects. The error is the engine's `VBError`, number 438, and `code` is the IR code (`G-TEXT`, `G-ENGINE`, `G-HOST`, `missing`, `missing-live`, `absent`, `readonly`). |

A get-only member has no Let. Assigning it fails in the VM and does not write. Trailing omitted arguments stay omitted. `Slides(1)` is the collection's default member. Collections stay 1-based.

`Application`, `Parent` and `Creator` are on every object from M1 on. `Creator` is `G-ENGINE`. `Parent` is the live parent: a presentation's parent is the `Presentations` collection, and `Application.Parent` is `Application`.

The injected module is `OratriceDOM`. Its properties are `Application`, `ActivePresentation`, `Presentations` and `ActiveWindow`. Constants are the `pp*` and `mso*` numbers from `dom.js`.

## Milestones

### M1. Application, Presentations, Presentation

Implemented. The test opens one Immediate host and checks:

- The `pp*` and `mso*` constants, against `L.dom`.
- `Application.ActivePresentation`, `Presentations`, `ActiveWindow.Selection` as an object, and `ActiveWindow.Presentation`.
- `Presentations.Count`, `Item`, `First` and `Last`. `Add` with a window argument and `Open` are `G-HOST`.
- `Name` is get-only. `Saved`, the built-in Title property, a custom property, `PageSetup.SlideWidth`, `Undo` and `Redo`.
- `Creator` is `G-ENGINE`. `FullName`, `Path`, `SaveAs`, `PrintOut` and `ActiveWindow.View` are `G-HOST`.

`Slides` is only an object here. Slide members are M2.

### M2. Slides and shapes

Implemented.

- `Slides.Add`, `Name`, `MoveTo`, `Duplicate`.
- `AddTextbox` and `AddShape` for `msoShapeRectangle`. `Left` on the shape and on `ShapeRange` after `Select`. `Delete` through the selection.
- `Selection.Type` is `ppSelectionShapes`.

### M3. Text, font, paragraph format

Implemented.

- `Shapes.Title.TextFrame.TextRange.Text`, `Font.Name`, `Font.Size`, `Font.Bold`, `ParagraphFormat.Alignment`, `InsertAfter`.
- `Characters.Count`. `Bullet.Type`. The title placeholder's `PlaceholderFormat.Type`.
- `Characters(1)`, `Start`, `Length` and `Find` are `G-TEXT`.

### M4. Tables, fill, links, transition, show

Implemented.

- `AddTable` and `Cell(1, 1)` text. Oval `Fill.ForeColor.RGB` and `Line.Weight`.
- `Hyperlinks.Add` on the title range. Footer text on `SlideMaster.HeadersFooters`. Setting `Slide.Design` by name.
- `SlideShowTransition.EntryEffect` and `Speed` for fade smoothly.
- `SlideShowSettings.Run`, `SlideShowWindows(1).View.Next`, `GotoSlide` and `Exit`.
- `AddPicture` is `G-HOST`. `AddChart` and `Slide.TimeLine` are `G-ENGINE`.

### M5. Gap-only objects

Implemented. `TimeLine`, `Sequence`, `Effect`, `Chart`, `Diagram`, `OLEFormat`, `MediaFormat`, `CustomerData`, `NotesPage`, `Comments`, `Tags`, `PrintOptions`, `Panes` and `CommandBars` are named objects. `Parent` is `Application`. A member on a `G-ENGINE` object returns `G-ENGINE`. `CommandBars` and `Panes` return `G-HOST`. `Slide.Comments`, `SlideMaster.Shapes` and `ActiveWindow.Caption` reject with the codes in the inventory.

`SlideShowWindow` has no extra members in this milestone. Pointer, state and accelerators stay `G-HOST` on `SlideShowView`.
