# Quire DOM IR

The call protocol between a language runtime and the Word object model in `public/quire/js/dom.js`. The inventory is [quire_dom_api_specifications.md](quire_dom_api_specifications.md). This file defines values, operations and errors. The member list is that inventory, stored as the catalog in `public/quire/js/dom-ir.js`.

The IR does not lay out pages, invent story indexes, or replace `dom.js`. A member the inventory marks Implemented is invoked on the live object. A member the inventory marks Not implemented is an IR operation that returns that row's gap code. An object the inventory lists with no per-member rows is registered by name; any member returns that object's gap.

## Session

`L.domIr.open()` returns a session. `session.root` is `{ handle, type: 'Application' }`, the live `Application` object. `session.exec(request)` performs one operation and returns a result object. When the live member returns a promise, `exec` returns a promise of that result, so `await session.exec(...)` waits for it. `session.release()` drops every handle. `L.domIr.Missing` is the omitted-argument value.

Handles are integers, unique for one session and one live object. The same live object yields the same handle. A new wrapper from `dom.js` is a new handle.

## Values

| Kind | Result |
|---|---|
| Scalar | `{ ok: true, kind: 'scalar', value }` Numbers, strings and booleans pass through. VBA `True` and `False` are the numbers `-1` and `0`. The executor does not convert them; `dom.js` already accepts both those numbers and JavaScript booleans. |
| Handle | `{ ok: true, kind: 'handle', handle, type }` `type` is the inventory object name. |
| Nothing | `{ ok: true, kind: 'nothing' }` The live member returned `null`. |
| Empty | `{ ok: true, kind: 'empty' }` A method that does not return a value. |
| Error | `{ ok: false, code, member, message }` |

An argument is a scalar, `L.domIr.Missing`, or `{ handle }`. Trailing `Missing` values are dropped, so the live call is made with a shorter argument list. A `Missing` that still has a value after it is passed as `undefined`. A shorter argument list omits parameters from the right. Callers do not pass an options object.

Constants stay the `wd*` numbers defined by `dom.js`. The IR does not re-encode them.

## Operations

`session.exec` takes `{ op, handle, member, args, value }`.

| op | Meaning |
|---|---|
| `get` | Read `member`. |
| `set` | Assign `value` to `member`. A read-only member returns `readonly` and does not write. |
| `call` | Call `member` with `args` in VBA order. A property that is a getter may be called with no arguments; that reads it. Arguments on a getter return `missing-live`. |
| `item` | 1-based index. `args[0]` is the index. This is `Paragraphs(1)` and `Item(1)`. The executor calls the live `Item`. |

`get` of `Parent` returns the live parent. Its type is the type already recorded for that object, or `Object` when the parent has not been seen. `Object` has only `Application`, `Parent` and `Creator`.

## Errors

| Code | When |
|---|---|
| `G-STORY`, `G-ENGINE`, `G-HOST` | The inventory's gap, returned before the live member runs, or thrown by the live member on a partial gap (for example `Documents.Add` with a template). `G-HOST` means this session does not have that host ability. |
| `E-APPROVAL` | The user denied a RequestApproval call. |
| `E-DENIED` | The member is AlwaysDeny. Shell is this code on the VS Code host. |
| `missing` | The live collection has no such index. |
| `missing-live` | The inventory lists the member and the live object does not have it. The IR does not invent the behaviour. |
| `absent` | The name is not in the inventory for that object. Prototype members are not visible. |
| `readonly` | `set` on a member that is not writable. |
| `handle` | The handle is unknown or the session was released. |
| `released` | `exec` after `release`. |

`ActiveDocument.Range(0, 5)` is `call` on `Range` and returns `G-STORY`. `Creator` returns `G-ENGINE` on every object.

## Objects with no members of their own

These names are `session.named(name)`. The result is a handle of that type. `Application` and `Parent` still navigate. `Creator` is `G-ENGINE`. Every other member returns the gap from the inventory: `G-ENGINE` or `G-HOST`. The Word reference is not used to add members.

`ContentControls`, `ContentControl`, `XMLNodes`, `XMLNode`, `CustomXMLParts`, `CustomXMLPart`, `FormFields`, `FormField`, `CheckBox`, `DropDown`, `TextInput`, `Frames`, `Frame`, `Subdocuments`, `MailMerge`, `MailMergeFields`, `MailMergeDataSource`, `OMaths`, `OMath`, `BuildingBlocks`, `BuildingBlock`, `Variables`, `Variable`, `Chart`, `ChartData`, `Diagram`, `DiagramNode`, `OLEFormat`, `OLEControl`, `TablesOfContents`, `Indexes`, `TablesOfFigures` are `G-ENGINE`.

`Windows`, `Window`, `Panes`, `Pane`, `View`, `Dialogs`, `Dialog`, `FileDialog`, `System`, `Tasks`, `CommandBars`, `KeyBindings`, `Languages`, `AutoCorrect`, `Options`, `DefaultWebOptions` are `G-HOST` when used as those gap objects. `Application.Dialogs` and `Application.FileDialog` are live members. See [docs/dom_api_security.md](dom_api_security.md).

A catalog member may end with `AlwaysAllow`, `RequestApproval`, or `AlwaysDeny`, and then the ability name. `dialog` means the file dialog is the approval, so the IR does not ask as well. `ExportAsFixedFormat` checks for XPS in the live member before that tier runs. Without `L.hostAudit`, a RequestApproval member behaves as it did before the tier: the live member returns `G-HOST` when the host ability is absent.

## Catalog

A call line, including a gap call, may end with a parameter list. `[Name]` is optional and a bare `Name` is required. The names are the VBA names in positional order, and the host publishes them, so `NumColumns:=2` binds to that slot. A listed call with no parameters has no list and still accepts sixteen optional `ArgN` parameters. A collection `Item` with no list is `Index`. A type marked `gap` has no per-member list, so an unlisted member on that type still uses `ArgN`.

Every collection in the catalog accepts `item` and exposes `Count`, `First` and `Last`. `First` and `Last` follow the inventory rule. Where `dom.js` does not provide them, the result is `missing-live`.

One live spelling differs from the inventory name: `Comment.Initial` reads and writes the live property `Initials`.

`Font.Borders` and `Font.Shading` are Implemented in the inventory. The live getters still throw `G-ENGINE`. The IR calls them and returns that error. It does not supply a border model of its own.

`Table.ApplyStyleHeadingRows and the other look flags` uses the flag names `dom.js` already stores: `ApplyStyleLastRow`, `ApplyStyleFirstColumn`, `ApplyStyleLastColumn`, `ApplyStyleRowBands`, `ApplyStyleColumnBands`. A border edge is `Border.LineStyle`, the property the live edge stores. `PictureFormat` is the object the inventory names; crop and brightness are not given member names there, so they are not catalog entries.

## Worked calls

These are the inventory's worked calls as IR operations. `wdAlignParagraphCenter` is `1`, `wdFindContinue` is `1`, `wdReplaceOne` is `1`.

```text
get Application.ActiveDocument
get Document.Paragraphs
item Paragraphs 1
get Paragraph.Range
get Range.Text
set Range.Text "Hello"
get Application.Selection
call Selection.TypeText [" "]
get Selection.Font
set Font.Bold -1
set Paragraph.Alignment 1
get Document.Tables
call Tables.Add [range handle, 2, 2]
item Tables 1
call Table.Cell [1, 1]
get Bookmarks
call Bookmarks.Add ["Here", range handle]
get Selection.Find
call Find.Execute ["Hello", -1, 0, 0, 0, 0, -1, 1, 0, "Hi", 1]
call Document.Range [0, 5]   → G-STORY
```
