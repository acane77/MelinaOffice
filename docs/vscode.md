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
ext=$(python3 -c "import json; p=json.load(open('package.json')); print(p['publisher']+'.'+p['name']+'-'+p['version'])")
mkdir -p ~/.local/share/code-server/extensions
ln -sfn "$PWD" ~/.local/share/code-server/extensions/$ext
```

On Windows the same directory is `%APPDATA%\code-server\extensions`. Restart code-server. If it was started with `--extensions-dir`, link into that directory instead. The folder name is `publisher.name-version` from `package.json`. The version lives only there.

Or pack a VSIX and install that (the archive contains `package.json`, `vscode/` and `public/`). The script reads the version from `package.json` and writes `tmp/vibeoffice-v<version>-<commit>.vsix`, where `<commit>` is the short id of HEAD, then prints that path:

```
python3 tools/vscode_vsix.py
code-server --install-extension <the path the script prints>
```

In desktop VS Code the same folder is an extension: Run and Debug, **VibeOffice extension** (`.vscode/launch.json`), which starts an Extension Development Host.

## Checking a change

A feature or a fix for the extension is checked by the user in VS Code (Run and Debug, **VibeOffice extension**), not by starting Chromium. The webview is not the browser: do not start Chromium, `tools/shot.mjs` or CDP for it. `node --test tests/vscode/extension.test.mjs` is the automated check; it does not open a browser. On WSL2, do not look for Chromium on Windows or on the WSL2 host, and do not run commands outside the workspace.

## Opening and saving

Click a `.docx`, `.xlsx` or `.pptx` in the Explorer. The editor tab is the app (Quire, Ledger or Lectern) with that file open, in place of a blank document. As soon as the page's dialogs exist, and until that file has finished opening, a dialog stays up: title "Starting Quire" (or Ledger, or Lectern). The frame is the ordinary dialog. Inside, on that app's colour (Quire `#2D5999`, Ledger `#137E43`, Lectern `#B8492B`), are the app's name, "Opening the file. Please wait...", and Cancel. It has no close button, and Escape does not dismiss it (`ui.dialog` with `canClose: false`). Cancel closes this editor tab. The dialog closes itself when the file is open. Edits mark the tab dirty. Save (the app's Save, Ctrl+S, or VS Code's Save) is one write of the open file: the page asks the extension, and `saveCustomDocument` tells the page to save. The tab clears when that save finishes. A copy (PDF, CSV, another name) is written beside the file and leaves the open file's unsaved changes on the tab. See [docs/vscode-bugfix-history.md](vscode-bugfix-history.md).

Ctrl+A is the page's own Select All. The webview host also binds that key and then selects the whole page; `suite.js` stops the key before that listener and ignores the follow-up, so the selection stays. Edit ▸ Select All does not press the key.

**New File…** (Create: New File, and the New File button in the Explorer) lists Word Document (.docx), Excel Workbook (.xlsx) and PowerPoint Presentation (.pptx). Choosing one creates the next free `Document1.docx`, `Book1.xlsx` or `Presentation1.pptx` in the first workspace folder and opens it in its editor. The page builds the same blank document as File ▸ New and writes it into that file. With no folder open, VS Code asks where to save it first. Typing a name such as `Report.docx` in that box is VS Code's own New File: it creates an empty file and opens it here, and an empty `.docx`, `.xlsx` or `.pptx` is filled the same way.

File ▸ New asks for a file with VS Code's save dialog, in the folder of the file this tab has open. Quire suggests `Document1.docx`, Ledger `Book1.xlsx` and Lectern `Presentation1.pptx`, using the next free number. The page then builds a new document of that type and the extension opens it in its own tab. The tab you started from stays on its file. The new file's extension is the app's (`docx`, `xlsx` or `pptx`). A name that ends with something else keeps that ending and gains the app's extension, so `新建文档.txt` is saved as `新建文档.txt.docx`. Choosing a file that is already open in VibeOffice switches to that tab and leaves it unchanged. A template chosen in the app is written the same way. A web page, an e-mail message, the welcome document and Lectern's design-template pane still open in the current tab.

File ▸ Open uses VS Code's open dialog (on code-server, the files on that machine) and opens each chosen file in its own tab. Ledger can select more than one. The page's own file picker stays in use for a file dropped on the page, and in Quire for From existing document.

Recent Files in Getting Started and the File menu remember the workspace path of a file this editor opened or saved. Choosing one opens that file in its own tab, and Save writes it back there. Unsaved changes kept for that entry are opened into the new tab, still bound to the same file. An entry with no path asks for a file name, as File ▸ New does, and writes the copy Recent Files kept. A file that is already open is shown instead of being replaced.

File ▸ Close, when this tab has only that one document, closes the editor tab. With several documents open inside the one tab, it closes the current one.

File ▸ Save As asks for a file name, a folder and a type. Location starts as the folder of the file this tab has open. Browse opens VS Code's folder dialog (on code-server, the folders on that machine). When the name and the folder are the open file, that file is written. Otherwise the new file is written in the chosen folder (with ` (1)` before the extension if that name is already there) and opened in its own tab. The tab you saved from still shows unsaved changes for the original file. PDF, PNG and CSV from the other commands use the name the app chose and are written beside the open file.

File ▸ Exit closes the editor tab. Links to a web page or a mail address open outside the editor. The service worker is not registered.

Help ▸ Language chooses the UI language (Auto, English, 日本語, 한국어, Français, Deutsch, 简体中文, 繁體中文), the same list as the Start Center's Settings. The extension remembers it (`vibeoffice.lang` in its global state) and loads the page again with that choice, which is what makes the editor show it: assigning the same HTML does not. If the file has unsaved changes, the choice is kept and used the next time an editor opens. With the ribbon, the list is at the end of File, beside Theme. It is not on the menu in the browser or the desktop app.

After you change a page under `public/`, close the editor tab and open the file again. The extension reads `index.html` when the tab opens.

`node --test tests/vscode/extension.test.mjs` covers the file types, the webview page and the save messages.
