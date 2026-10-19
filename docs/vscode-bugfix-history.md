# VS Code save-reporting fixes

The bytes of a save and the editor tab are different. The page can write the workspace file and show "Saved" while VS Code still shows the dirty dot, or a notification that the save failed. This file records those bugs and the fix, so a later change does not put the two paths back.

## 2026-10-19 — Dirty tab, "Save was cancelled", "Editor was closed"

Seen in VS Code and in code-server. Not on every save. Once it started, further saves in that window kept failing the same way until the window, or code-server, was restarted. The file on disk had the saved bytes in each case.

### What the user saw

1. Quire (and the same shape in the other apps) wrote the file and showed "Saved" with the file name. The editor tab kept the dirty dot.
2. Ledger showed `Failed to save 'Book1.xlsx': Save was cancelled.` The workbook, including one just created with New File, was already written. Quire and Lectern did not show this notification.
3. While the webview was busy, a save showed `Failed to save '…': The editor was closed.` The bytes were already written.

### Why

Ctrl+S was two saves.

The page handles Ctrl+S, writes the file, then posts `{ type: 'dirty', dirty: false }`. The extension turned that into `workbench.action.files.save` with `skipSave`, hoping VS Code would clear the dot. VS Code's own Save calls `saveCustomDocument`, which posts `do-save` and sets `savingFromHost`. While that flag was set, the page's `dirty: false` was ignored, so the clear never ran. A `dirty: true` that arrived after `save-result` pushed a new edit. VS Code drops only the edits it had when the save started, so the new one left the dot on. A slow webview makes the overlap last, which is why it then kept happening.

`do-save` did `await L.ui.exec('save')` and treated every result other than `'saved'` or `true` as `Save was cancelled.` Ledger's command wrapper called `A().save()` and did not return the promise, so the extension got `undefined` at once and reported cancellation while `A().save()` was still writing. Quire and Lectern return the promise, so they did not show that notification. Book1.xlsx was simply the first Ctrl+S after New File, which is when VS Code's save runs.

`OfficeDocument.dispose()` rejected every in-flight request with `The editor was closed.` A busy webview can be disposed after `writeFile` has returned and before `save-result` is delivered. `savingFromHost` then stayed true until something rejected that request, and later `dirty: false` messages kept being ignored.

### What changed

An in-place save of the open file is one save, `saveCustomDocument`.

- The page's Save command asks the extension (`request-save`) instead of writing on its own. `vscode.workspace.save` runs `saveCustomDocument` for that file. A request that arrives while a save is running waits for that save.
- `do-save` is the only writer. If the app already built the blob (Save, or Save before Close), that blob is written and the app is not asked to build it again. Otherwise `do-save` calls the real save command. A second `do-save` waits for the first.
- Ledger's command wrapper returns the promise from `A().save()`, so `do-save` sees `true` or `false`.
- `dirty: false` does not start a save. The tab clears because `saveCustomDocument` resolved. A dirty notice during the save is held; if the document is still dirty when the save ends, one new edit is fired so a keystroke during the save is not dropped and a notice caused by the save itself does not leave the dot on.
- After the in-place bytes are on disk, disposing the document or the webview resolves that save. Disposing it before the write rejects the save once and clears the in-flight state, so the next save can run. Close still clears the dot without writing, and that skip cannot stick to the next save.

A copy (PDF, CSV, PNG, or another name or folder) is still written beside the file and does not clear the open file's dot.

### Where

| | |
|---|---|
| `vscode/protocol.js` | `createSaveSession`: one save, join, dirty-during-save, dispose after write |
| `vscode/extension.js` | `saveCustomDocument`, `request-save`, no `dirty: false` save |
| `public/common/suite.js` | `request-save`, `do-save`, the staged blob |
| `public/common/core.js` | `L.saveFile` calls `saveThroughEditor` for the open file |
| `public/ledger/js/commands.js` | the save command returns its promise |

`node --test tests/vscode/extension.test.mjs` covers the session, the Ledger return, and that the old `savingFromHost` / `workbench.action.files.save` path is gone. The editor tab itself is checked in VS Code (Run and Debug, **VibeOffice extension**), not in a browser.
