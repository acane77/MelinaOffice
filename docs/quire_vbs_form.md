# VB6 forms

**Status.** The first round runs a form. It does not include a designer.

A form is a VB6 project on the same host as `evalVB6`. The virtual machine still owns events, `Form_Load`, `Show` and `MsgBox`. The window is a VibeOffice dialog (`ui.dialog`), in the current look. `ApplicationHost`, `BrowserForm`, the VB6 Studio IDE, the designer, themes and the WebGPU renderer are not loaded. The engine bundle stays the headless script from `make`.

## What runs

View ▸ Run Form (and the same command in the View ribbon's Macros group) asks for files. A `.frm` is required. `examples/vb/calculator/Calculator.frm` is a four-function calculator on that path. Its Insert button writes the displayed value through the document module (`InsertText`: the caret in Quire, the active cell in Ledger, a text box on the first slide in Lectern). `examples/vb/food-energy/FoodEnergy.frm` turns energy per 100 g (kJ) and a weight (g) into the portion's kJ and kCal (1 kCal = 4.184 kJ) and inserts either figure the same way. `.bas` modules and a `.vbp` selected with it are part of the project. A `.frx` picture is left out and recorded on the host as `formLosses`, with `notify: false`. Nothing is shown for that.

`loadProject` also accepts the engine's project object: `schema: 1`, modules with `kind: 'form'`, and `form: { type, properties, controls, menus }`. `settings.anchoring` and `settings.xaml` are not set.

The call stops a form already running, keeps the Immediate virtual machine, and starts a new one with the running budget (5,000,000 instructions). The app's DOM module (`QuireDOM`, `LedgerDOM` or `LecternDOM`) is added. `startup` is the `.vbp` startup, or the first form. While that machine's stack is empty, or the machine is paused, the DOM Console's Visual Basic lines use its `immediate`. Closing the last loaded form stops the machine. While it is running, the VB Form toolbar appears on its own, as the Picture toolbar does for a picture, and the ribbon shows VB Form Tools. Both carry Stop. A script, a macro or the Visual Basic IDE asks "A VB form is running, do you want to stop it?" and continues only when the form machine stops.

`createForm`, `beginModal`, `msgBox` and `inputBox` stay null until `loadProject`. `openFile` and `saveFile` stay null.

## Window and controls

`createForm` returns an object whose `controlMap` the virtual machine stores on the form. A control is a JavaScript object with `__control` and `type`, so property values keep the VB types the virtual machine already applies. Events are `vm.dispatch(instance, Name_Event, args)`. A control array passes `Index` first. Closing the window calls `requestUnload`, so `Form_QueryUnload` and `Form_Unload` can cancel it. When the unload leaves no form loaded, the form machine stops. A modeless form leaves the document editable. `Show` with a modal argument uses a modal dialog, and `beginModal` marks the other VB forms inert.

Drawn and live: Form, Label, TextBox, CommandButton, CheckBox, OptionButton, ListBox, ComboBox, Frame, Timer, HScrollBar, VScrollBar. Positions are twips, 15 to a pixel, inside the dialog. System colours (the negative OLE values) use the dialog face, the dialog text and the input background of the current look. A menu is a property object; a top-level menu is a button that dispatches `Name_Click`.

Other control types are objects with geometry and `Tag`. They are not drawn. A member the object does not have is the virtual machine's error 438.

`MsgBox` and `InputBox` use `ui.msg` and `ui.prompt`. The button values are OK 1, Cancel 2, Abort 3, Retry 4, Ignore 5, Yes 6, No 7. InputBox returns an empty string on Cancel. The document module's own `MsgBox` and `InputBox`, which macros call, are not copied into the form machine, so a form's call is the virtual machine's and goes through those host slots.

## Out of this round

The designer, MDI, anchoring, data binding, OCX controls, and PictureBox drawing (`Line`, `Circle`, `PSet`). Document edits from a form still go through the DOM module, not through a control.
