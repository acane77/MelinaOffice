# Quire DOM IR porting milestones

**Status:** M1 through M5 are implemented (`tests/quire/vb-m1.test.js` through `tests/quire/vb-m5.test.js`).

Gap-only objects are public fields of the injected module, filled from `session.named`. A gap type accepts any member name so the call reaches the IR. `Application` and `Parent` still navigate.

The inventory is [quire_dom_api_specifications.md](quire_dom_api_specifications.md). The call protocol is [quire_dom_ir_specifications.md](quire_dom_ir_specifications.md). The runtime is the P2 host in `public/quire/js/vb-host.js`.

## What a port is

A VB statement on the Immediate VM names a member. The Automation adapter emits one IR operation. `dom.js` performs the member, or the IR returns that member's gap. The port does not add a second object model, and it does not generate a VB procedure per member. A second list would drift from the catalog. Each live wrapper is its own Automation object. One session holds 8192 of them, because a walk allocates a new wrapper on each access.

A milestone ports every inventory member of its objects, including gaps. A gap is callable and rejects. It is not omitted and it is not given new behaviour. A name the inventory does not list is not added. The port does not lay out pages or invent story indexes.

## What VB sees

| IR result | From Immediate VB |
|---|---|
| Scalar | The value `dom.js` returned. Numbers, strings and booleans are not rewritten. VBA `True` and `False` are `-1` and `0` when that is what the live member accepts or returns. |
| Handle | An object. A later `.Member` is another IR operation. |
| Nothing | VB `Nothing`. |
| Empty | The call returns no value. |
| Error | The call rejects. The error is the engine's `VBError`, number 438, and `code` is the IR code (`G-STORY`, `G-ENGINE`, `G-HOST`, `missing`, `missing-live`, `absent`, `readonly`). |

A get-only member has no Let. Assigning it fails in the VM and does not write. Trailing omitted arguments stay omitted. `Paragraphs(1)` is the collection's default member. Collections stay 1-based.

`Application`, `Parent` and `Creator` are on every object from M1 on. `Creator` is `G-ENGINE`. `Parent` is the live parent: a document's parent is the `Documents` collection, and `Application.Parent` is `Application`. Later milestones test those three on the objects they add. They do not reopen M1's list.

## Milestones

Each milestone is a separate step. The next one starts only after this document's row for it is accepted. The test for a milestone calls the inventory members of that row from VB.

### M1. Application, Documents, Document

Implemented. The test opens one Immediate host and checks:

- The `wd*` constants injected from `dom.js`, against those numbers.
- `Application`: `ActiveDocument`, `Documents`, `Selection` (the Selection object only; its members are M2), `Application`, `Parent`, `Creator`.
- `Documents`: `Count`, `Item`, `First`, `Last`, `Add` with no arguments, `Add` with a template (`G-ENGINE`), `Open`, `Close`, `Save` (`G-HOST`), and the three shared members.
- `Document`: every row in the inventory's Document table. Reads and writes that the live member already performs (`Name`, `Saved`, `TrackRevisions`, `ComputeStatistics` for a statistic it implements, `BuiltInDocumentProperties` and `CustomDocumentProperties` through `DocumentProperty.Name` and `DocumentProperty.Value`, `StoryRanges` for a stored story, `Undo`, `Redo`, `Select`, `Activate`, `Save`, `Close` with `wdDoNotSaveChanges`). Collection properties are read as objects; their own members wait for the milestone that owns that collection.
- Gaps on `Document`: `FullName`, `Path`, `SaveAs`, `SaveAs2`, `PrintOut`, `PrintPreview`, `Compare`, `FollowHyperlink`, `ExportAsFixedFormat` (`G-HOST`); `UndoClear`, `Protect`, `Unprotect`, `ComputeStatistics(wdStatisticLines)` (`G-ENGINE`); `Characters`, `Range`, `GoTo` (`G-STORY`); `Close` with `wdPromptToSaveChanges` or a second argument (`G-HOST`).

`DocumentProperty` is in this milestone because it is the value `BuiltInDocumentProperties` and `CustomDocumentProperties` return. It is not a separate object family.

### M2. Selection, Range, paragraphs, sentences, words, characters

Implemented. The test writes a paragraph, then calls Selection, Range, Paragraphs, Sentences, Words and Characters from VB. Story-index members return `G-STORY`. `Characters(Index)` returns `G-STORY`. `Characters.Count` matches the live count. `Sentences.First` and `Words.Last` are `missing-live`: the live collections have `Count` and `Item` only. `Paragraph.Next` with no argument reads the neighbouring paragraph. `Paragraph.Next(2)` is `missing-live` because the live member is a getter and does not take a count.

### M3. Font, paragraph format, styles, lists, sections

Implemented. The test sets run font, paragraph format, a tab stop, a new paragraph style, numbering, page setup and a header or footer from VB, and reads them back from the live document. `Font.Borders`, `Font.Shading` and the other font gaps return `G-ENGINE`. `ParagraphFormat.Borders` and `ParagraphFormat.Shading` are `missing-live`: the inventory lists them and the live format object does not have them. `ListFormat.ListType`, `List.Range` and `List.CountNumberedItems` are `missing-live` for the same reason. `ListTemplates` and the other gallery members return `G-ENGINE`. The port does not add those members to `dom.js`.

### M4. Tables, find, bookmarks, fields, links, notes

Implemented. The test inserts a table, a find-and-replace, a bookmark, a hyperlink, a comment, a footnote, an endnote and a field from VB, and reads them back from the live document. `Bookmark.Start`, `Bookmark.End`, `Field.Code.Start`, `Hyperlink.Range.Start` and `Footnote.Reference` return `G-STORY`. `Find` with `wdFindAsk` returns `G-HOST`. `Find.ClearFormatting` is `missing-live`: the inventory lists it and the live find object does not have it. `Comment.Scope` and `Comments.AddReply` are `missing-live` (the live reply method is `Reply`). `Comment.Initial` reads `Initials`; that live property has no setter. `Footnotes.Convert` is `missing-live`. `Field.Locked`, `Field.Unlink` and `Field.Select` are `missing-live`. `Rows.Delete` is `missing-live`; `Rows.Add` and `Columns.Add` run. The port does not add those members to `dom.js`.

### M5. Shapes, revisions, gap-only objects

Implemented. The test reads shape and revision collections from VB. `InlineShapes.AddPicture` returns `G-HOST`. `Shapes.AddChart` returns `G-ENGINE`. `Shapes.AddShape` returns `G-ENGINE` when the drawing module is not loaded; when it is, width, height and fill write through and `Shape.Left` and `Shape.WrapFormat` are `missing-live`. `Revision.Author`, `Revision.Date` and `Revision.Range` are `missing-live` when a revision exists. Every object in "Objects with no document operation" is a public object on the injected module. `Application` and `Parent` still navigate. `Creator` returns `G-ENGINE`. Any other member returns that object's `G-ENGINE` or `G-HOST`. The Word reference is not mined for further members.
