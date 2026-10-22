# Lectern DOM IR

The handle protocol between the VB host and `public/lectern/js/dom.js`. The inventory is [lectern_dom_api_specifications.md](lectern_dom_api_specifications.md). The executor and the catalog are `public/lectern/js/dom-ir.js`.

The IR does not layout slides and does not add members. A member the catalog does not list is `absent`, except `Application`, `Parent` and `Creator`, which every object has. `Creator` is `G-ENGINE`.

## Session

`L.domIr.open()` returns a session.

| | |
|---|---|
| `session.root` | `{ handle, type: 'Application' }` for `L.dom.Application` |
| `session.exec(req)` | One operation. Returns `{ ok: true, ... }` or `{ ok: false, code, member, message }` |
| `session.named(name)` | A gap-only object from the catalog (`TimeLine`, `Chart`, `CommandBars`, …). An unknown name is `absent` |
| `session.release()` | Later `exec` returns `released` |

Handles are integers starting at 1. The same live object is the same handle for the life of the session. A wrapper created by a later get is a new object and a new handle.

## Values

| `kind` | |
|---|---|
| `scalar` | `value` is the number, string or boolean `dom.js` returned. The IR does not rewrite VBA `True` |
| `handle` | `handle` and `type`. The caller uses the handle on the next operation |
| `nothing` | The live member returned null |
| `empty` | A call with no result |

A catalog return type of `empty` still reports a scalar when the live call returned a number. `Undo` does that.

`L.domIr.Missing` is an omitted argument. Trailing `Missing` values are dropped. A hole in the middle is passed as `undefined`.

An argument `{ handle }` is unwrapped to the live object. An unknown handle fails the operation with `handle`.

## Operations

```js
{ op: 'get', handle, member }
{ op: 'set', handle, member, value }
{ op: 'call', handle, member, args }
{ op: 'item', handle, args }
```

`item` is `Item` with one index. Collections use it as the default member, so `Slides(1)` is `item`.

`get` of a call-only member is `absent`. `set` of a get-only member is `readonly`. A live object that does not have the member is `missing-live`. A thrown `dom.js` gap is returned as that code (`G-TEXT`, `G-ENGINE`, `G-HOST`, `missing`).

## Errors

| `code` | |
|---|---|
| `G-TEXT` | The member needs a character index inside the shape |
| `G-ENGINE` | No document operation, including `Creator` and every member of a `G-ENGINE` gap object |
| `G-HOST` | A path, a window, a printer, or a modal, including every member of a `G-HOST` gap object |
| `missing` | The index is outside the collection |
| `missing-live` | The catalog lists the member and this live object does not have it |
| `absent` | The catalog has no such member |
| `readonly` | The member is get-only |
| `handle` | The handle is not in this session |
| `released` | `release` has been called |

## Objects with no members of their own

A catalog line `# TimeLine gap G-ENGINE` registers the name. `session.named('TimeLine')` returns a handle of that type. `Application` and `Parent` navigate. Any other member returns `G-ENGINE`. A `# Panes gap G-HOST` object returns `G-HOST`. The host publishes each of these names as `Public Name As Object` on the injected module.

## Catalog

`L.domIr.types` is parsed from the catalog string in `dom-ir.js`.

```text
# TypeName
# TypeName collection
# TypeName gap G-ENGINE
Member get [ReturnType]
Member rw
Member call [ReturnType]
Member gap G-TEXT get|call|rw
```

A collection gains `Count`, `First`, `Last` and `Item` when the catalog does not list them. `Item`'s return type is the collection's element type.

`L.domIr.member(type, name)` is what the host's Automation metadata reads. The executor is the only caller of the live object.

## Worked calls

```js
const s = L.domIr.open();
const pres = s.exec({ op: 'get', handle: s.root.handle, member: 'ActivePresentation' });
const slides = s.exec({ op: 'get', handle: pres.handle, member: 'Slides' });
const slide = s.exec({ op: 'item', handle: slides.handle, args: [1] });
const shapes = s.exec({ op: 'get', handle: slide.handle, member: 'Shapes' });
const title = s.exec({ op: 'get', handle: shapes.handle, member: 'Title' });
const frame = s.exec({ op: 'get', handle: title.handle, member: 'TextFrame' });
const range = s.exec({ op: 'get', handle: frame.handle, member: 'TextRange' });
s.exec({ op: 'set', handle: range.handle, member: 'Text', value: 'Hello' });
```

`Characters(1)` is `item` on `Characters` and returns `G-TEXT`. `FullName` returns `G-HOST`. `session.named('TimeLine')` then a get of `MainSequence` returns `G-ENGINE`.
