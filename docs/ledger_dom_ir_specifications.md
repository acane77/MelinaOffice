# Ledger DOM IR

The handle protocol between a caller and `public/ledger/js/dom.js`. The inventory is [ledger_dom_api_specifications.md](ledger_dom_api_specifications.md). The executor is `public/ledger/js/dom-ir.js`. It is a Ledger copy. It is not shared in `public/common/`.

## Session

`L.domIr.open()` returns a session. `exec` runs one operation. `named` publishes a gap object by type name. `release` drops every handle. A released handle returns `released`.

The root handle is `Application`. A returned object is a new handle. The caller does not see the live JavaScript object.

## Values

| Kind | Meaning |
|---|---|
| `scalar` | Number, string, or boolean. VBA `True` and `False` are `-1` and `0` |
| `handle` | An object in this session |
| `nothing` | VBA `Nothing`. `Find` with no hit |
| `empty` | VBA `Empty`. An empty cell, or a call that returns no value |
| `array` | A rectangle. JavaScript is a 0-based array of rows. VB is a 1-based `VBArray`. One cell stays a scalar |
| `missing` | An omitted optional argument |

A range larger than 1,000,000 cells rejects `Value` and `Formula` with `G-ENGINE`.

## Operations

| `op` | Fields | |
|---|---|---|
| `get` | `handle`, `member` | Read a property. A `both` member with no arguments calls the function |
| `set` | `handle`, `member`, `value` | Write a property |
| `call` | `handle`, `member`, `args` | Call a method. A property get that carries arguments is sent as `call`, so `Cells(1, 1)` and `Range("A1")` work |
| `item` | `handle`, `args` | The default member, one index. `Item` with two arguments is `call` |

`WorksheetFunction` is not listed member by member. Any name in the function catalog is `call` and returns a scalar. Any other name is absent. `exec` calls `invoke`.

## Errors

| Code | When |
|---|---|
| `G-ENGINE` | The inventory says there is no document operation, including `Creator` and semiautomatic calculation |
| `G-HOST` | The inventory says the member needs a path, a window, a printer, a modal or a file picker |
| `missing` | The member is in the catalog and the live object rejected the arguments |
| `missing-live` | The catalog says the member exists and the live object does not have it |
| `absent` | The name is not in the catalog, and it is not a catalog function |
| `readonly` | A write to a get-only member |
| `handle` | The handle is not in this session |
| `released` | The session was released |

The VB host turns these into `VBError` 438 with `err.code` set to the IR code.

## Objects with no members

A type line `# Name gap G-ENGINE` or `# Name gap G-HOST` registers that name. `named` returns a handle. `Application` and `Parent` navigate. Any other member returns that gap.

## Catalog

The grammar matches Lectern, plus `both` and `array`.

```text
# TypeName
# TypeName collection
# TypeName gap G-ENGINE
Member get|rw|call|both|gap [return type] [G-ENGINE|G-HOST]
```

`both` is a function: a get with no arguments calls it, and a call passes the arguments. `Cells` is `both`. `Range` is `call`. Collections receive `Count`, `Item`, `First` and `Last` when the line does not list them. `Range` and the collections use `Item` as the default member. `WorksheetFunction` and the gap types accept any member name at the host.

## Worked calls

Open a session. Get `ActiveSheet`. Call `Range` with `"A1"`. Set `Value` to `"Hi"`. Get `Value` and receive `{ok:true, kind:"scalar", value:"Hi"}`. Call `Range` with `"A1:B1"`, set `Value` to `[["Hi", 2]]`, and get `{kind:"array", value:[["Hi", 2]]}`. Get `ActiveWorkbook.FullName` and receive `G-HOST`. Get `NotAMember` and receive `absent`. After `release`, the root returns `released`.

`tests/ledger/dom.test.js` runs these calls.
