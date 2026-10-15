# VibeOffice in VS Code and code-server

The same Quire, Ledger and Lectern pages run inside a VS Code webview, including [code-server](https://coder.com/docs/code-server). Opening a workspace file (a double-click on `.docx`, `.xlsx` or `.pptx`, or **Open in VibeOffice** in the Explorer) loads that file in the app that owns it. **Save** (Ctrl+S, or File ▸ Save in the app) writes the bytes back to that file. A different name, a PDF, or a CSV is written in the same folder and opened as well. The document does not leave the machine: the extension host reads and writes the workspace, and the webview only talks to the extension.

| | |
|---|---|
| `package.json` | the extension: custom editors, the Open in VibeOffice command |
| `vscode/extension.js` | activation, the editors, save and backup |
| `vscode/page.js` | the webview document: a `<base>` so the app's relative links resolve, and the host script |
| `vscode/protocol.js` | which extension opens in which app, and how a save is split into messages |
| `public/` | the pages, unchanged; the webview loads them from here |

`VO.host.kind` is `'vscode'`. The page side of that host is `public/common/suite.js` (`takeDocument`, `writeFile`); `L.saveFile` in `public/common/core.js` uses it. The rules shared with the desktop app are in [docs/desktop.md](desktop.md) (Host interface).

`.docx`, `.xlsx`, `.pptx` and the other Office Open XML types open in VibeOffice by default. Text, HTML, CSV, Markdown and RTF stay with the text editor; **Open With** offers Quire or Ledger for them.

## Install in code-server

The extension runs in the Node extension host (`extensionKind: workspace`), which is what code-server uses. The pages have to sit next to `package.json` (`public/quire/index.html` and the rest).

On the code-server machine, from this repository (Linux and macOS; the extensions directory is `~/.local/share/code-server/extensions` on Linux and `~/Library/Application Support/code-server/extensions` on macOS):

```
mkdir -p ~/.local/share/code-server/extensions
ln -sfn "$PWD" ~/.local/share/code-server/extensions/vibeoffice.vibeoffice-0.1.1
```

On Windows the same directory is `%APPDATA%\code-server\extensions`. Restart code-server. If it was started with `--extensions-dir`, link into that directory instead. The folder name is `publisher.name-version` from `package.json`.

Or pack a VSIX and install that (the archive contains `package.json`, `vscode/` and `public/`):

```
python3 tools/vscode_vsix.py
code-server --install-extension tmp/vibeoffice-0.1.1.vsix
```

In desktop VS Code the same folder is an extension: Run and Debug, **VibeOffice extension** (`.vscode/launch.json`), which starts an Extension Development Host.

## Opening and saving

Click a `.docx`, `.xlsx` or `.pptx` in the Explorer. The editor tab is the app (Quire, Ledger or Lectern) with that file open, in place of the sample document. Edits mark the tab dirty. Ctrl+S asks the app to save and writes the result over the workspace file. Workspace files open from the Explorer. File ▸ Open inside the app is the page's file picker; in a webview that picker often cannot see the workspace, so use the Explorer for those files.

File ▸ Save As, PDF, PNG and CSV use the name the app chose. When that name is the open file, the open file is written. When it is another name, the new file is written beside it (with ` (1)` before the extension if the name is already there) and opened in its own tab. The tab you saved from still shows unsaved changes for the original file.

File ▸ Exit closes the editor tab. Links to a web page or a mail address open outside the editor. The service worker is not registered.

Help ▸ Language chooses the UI language (Auto, English, 日本語, 한국어, Français, Deutsch, 简体中文, 繁體中文), the same list as the Start Center's Settings. The extension remembers it (`vibeoffice.lang` in its global state) and loads the page again with that choice, which is what makes the editor show it: assigning the same HTML does not. If the file has unsaved changes, the choice is kept and used the next time an editor opens. With the ribbon, the list is at the end of File, beside Theme. It is not on the menu in the browser or the desktop app.

After you change a page under `public/`, close the editor tab and open the file again. The extension reads `index.html` when the tab opens.

`node --test tests/vscode/extension.test.mjs` covers the file types, the webview page and the save messages.
