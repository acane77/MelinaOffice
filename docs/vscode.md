# VibeOffice in VS Code and code-server

The same Quire, Ledger and Lectern pages run inside a VS Code webview, including [code-server](https://coder.com/docs/code-server). Opening a workspace file (a double-click on `.docx`, `.xlsx` or `.pptx`, or **Open in VibeOffice** in the Explorer) loads that file in the app that owns it. **Save** (Ctrl+S, or File ▸ Save in the app) writes the bytes back to that file. A different name, a PDF, or a CSV is written in the same folder and opened as well. The document does not leave the machine: the extension host reads and writes the workspace, and the webview only talks to the extension.

| | |
|---|---|
| `package.json` | the extension: custom editors, the Open in VibeOffice command, the MCP provider id |
| `vscode/extension.js` | activation, the editors, save and backup, and registering the MCP server |
| `vscode/mcp.js` | the MCP process: `vibeoffice_dom_guide` and `vibeoffice_run_vbs` |
| `vscode/mcp-bridge.js` | the localhost bridge from that process back to the extension host |
| `vscode/documents.js` | which open tab a script may run on |
| `vscode/vbs-guide.js` | the object-model note the agent reads, distilled from the three DOM inventories |
| `vscode/page.js` | the webview document: a `<base>` so the app's relative links resolve, and the host script |
| `vscode/protocol.js` | which extension opens in which app, and how a save is split into messages |
| `public/` | the pages; `public/common/suite.js` runs a script the extension posts into the open tab |

`VO.host.kind` is `'vscode'`. The page side of that host is `public/common/suite.js` (`takeDocument`, `writeFile`); `L.saveFile` in `public/common/core.js` uses it. The rules shared with the desktop app are in [docs/desktop.md](desktop.md) (Host interface).

When no look is stored, this host opens in **Golden Gate (Follow VS Code Color Theme)**. The Themes list adds that look, one look, not a light one and a dark one. A look the user has already chosen stays. A browser does not list it, and a stored choice of it opens as Golden Gate outside this host. The window background is the editor's (`editor.background`). The ribbon tab bar is the title bar (`titleBar.activeBackground`), which is where the menu bar is drawn; VS Code has no separate menu-bar background. Buttons, fields, menus and lists use the workbench colours (`button`, `input`, `menu`, `list`, `focusBorder`). Liquid glass follows the theme kind: light for `vscode-light` and `vscode-high-contrast-light`, dark for `vscode-dark` and `vscode-high-contrast`. Switching the VS Code theme updates those colours. The selected look stays Golden Gate (Follow VS Code Color Theme). The webview already exposes the colours as `--vscode-*` and the kind as a class on `<body>`.

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

Or pack a VSIX and install that (the archive contains `package.json`, `vscode/` and `public/`). `tools/package_vsix.sh` writes `public/version.json` (version from `package.json`, the short commit id, the current branch, and today's date) and then packs `tmp/vibeoffice-v<version>-<commit>.vsix`. Help ▸ About reads that file and shows `Version: v<version> (<commit>), Build Date: <date>`, with half-width colons and no translation.

```
tools/package_vsix.sh
code-server --install-extension <the path the script prints>
```

In desktop VS Code the same folder is an extension: Run and Debug, **VibeOffice extension** (`.vscode/launch.json`), which starts an Extension Development Host.

## Checking a change

A feature or a fix for the extension is checked by the user in VS Code (Run and Debug, **VibeOffice extension**), not by starting Chromium. The webview is not the browser: do not start Chromium, `tools/shot.mjs` or CDP for it. `node --test tests/vscode/extension.test.mjs` and `node --test tests/vscode/mcp.test.mjs` are the automated checks; they do not open a browser. On WSL2, do not look for Chromium on Windows or on the WSL2 host, and do not run commands outside the workspace.

## Running a script from the agent

The extension registers one MCP server, named VibeOffice, when it activates. Activation is still the first Quire, Ledger, or Lectern tab, so the tools appear after a document is open. VS Code uses `vscode.lm.registerMcpServerDefinitionProvider` when that API exists. Cursor uses `vscode.cursor.mcp.registerServer` when that API exists. Nothing is written to an MCP configuration file. A host with neither API opens and saves files as before. The engine requirement stays `^1.75.0`.

The server is a stdio process (`vscode/mcp.js`) started with the editor's executable and `ELECTRON_RUN_AS_NODE=1`. It calls back to a `127.0.0.1` bridge that only the extension knows the token for. The bridge does not open files, save, or read the workspace. It posts `run-vbs` into the webview of the active editor tab, and that page runs the same `runVB6` / `evalVB6` entry the DOM console uses. A hidden VibeOffice tab is not scripted. If the active tab is not one of the three apps, the tool names the open VibeOffice tabs and asks for a click on the right one. If none is open, both tools say to open the file first.

`vibeoffice_dom_guide` is read-only. It returns the guide for the app that owns the active tab (Quire, Ledger, or Lectern), including the gaps versus Microsoft VBA: `G-STORY` in Quire, `G-TEXT` in Lectern, and `G-ENGINE` / `G-HOST` in all three, and the host abilities that this editor implements. The text is distilled in `vscode/vbs-guide.js` from `docs/quire_dom_api_specifications.md`, `docs/ledger_dom_api_specifications.md`, `docs/lectern_dom_api_specifications.md`, and [docs/dom_api_security.md](dom_api_security.md). A change to a member's status in those inventories updates the guide in the same change.

`vibeoffice_run_vbs` runs the script. `mode` `module` is a standard module with `Public Sub Main`, or the public procedure named by `entry`. `mode` `line` is one Immediate statement. Values come back from `Debug.Print`, or from a scalar a `Function` or a `?` line returns. The edit uses the document's own operations, so the tab is marked dirty and Undo in the app removes it. One script runs at a time on that document. A script that does not finish within a minute returns an error. That minute pauses while an approval window or a VS Code file dialog is open. A RequestApproval call from this tool asks Allow, Always allow, or Deny. Always allow is remembered for that ability. Dismissing the window, or leaving it for two minutes, is Deny (`E-APPROVAL`). AlwaysDeny, including Shell, is `E-DENIED`. The page must be loaded again after `public/` changes: close the tab and open the file.

To try it, run **VibeOffice extension**, open a `.docx`, `.xlsx`, or `.pptx` on the left, and ask the agent on the right to change that document (for example to set a font size or to write a paragraph). The agent should call the guide, then the script, on that tab only.

`tests/vscode/host-ability.vbs` calls every host ability on the open saved document. The same file runs in Quire, Ledger, and Lectern: it branches on `HasHostAbility`. Paste it into the DOM console (Run Script), or pass it to `vibeoffice_run_vbs`. A file call opens a dialog first and uses the path you pick; Cancel skips that call. `PrintOut` writes a PDF beside the open file. Shell and XPS are expected to fail (`E-DENIED`, and XPS before any write).

## Opening and saving

Click a `.docx`, `.xlsx` or `.pptx` in the Explorer. The editor tab is the app (Quire, Ledger or Lectern) with that file open, in place of a blank document. As soon as the page's dialogs exist, and until that file has finished opening, a dialog stays up: title "Starting Quire" (or Ledger, or Lectern). The frame is the ordinary dialog. Inside, on that app's colour (Quire `#2D5999`, Ledger `#137E43`, Lectern `#B8492B`), are the app's name, "Opening the file. Please wait...", and Cancel. It has no close button, and Escape does not dismiss it (`ui.dialog` with `canClose: false`). Cancel closes this editor tab. So does refusing the password on a protected file: Cancel or the close button on the Password window closes the tab, and so does OK or the close button on the window that reports a wrong password or an encryption method the app cannot read. The blank document is not left on screen. The opening dialog closes itself when the file is open. Edits mark the tab dirty. Save (the app's Save, Ctrl+S, or VS Code's Save) is one write of the open file: the page asks the extension, and `saveCustomDocument` tells the page to save. The tab clears when that save finishes. A copy (PDF, CSV, another name) is written beside the file and leaves the open file's unsaved changes on the tab. See [docs/vscode-bugfix-history.md](vscode-bugfix-history.md).

Ctrl+A is the page's own Select All. The webview host also binds that key and then selects the whole page; `suite.js` stops the key before that listener and ignores the follow-up, so the selection stays. Edit ▸ Select All does not press the key.

**New File…** (Create: New File, and the New File button in the Explorer) lists Word Document (.docx), Excel Workbook (.xlsx) and PowerPoint Presentation (.pptx). Choosing one creates the next free `Document1.docx`, `Book1.xlsx` or `Presentation1.pptx` in the first workspace folder and opens it in its editor. The page builds the same blank document as File ▸ New and writes it into that file. With no folder open, VS Code asks where to save it first. Typing a name such as `Report.docx` in that box is VS Code's own New File: it creates an empty file and opens it here, and an empty `.docx`, `.xlsx` or `.pptx` is filled the same way.

File ▸ New asks for a file with VS Code's save dialog, in the folder of the file this tab has open. Quire suggests `Document1.docx`, Ledger `Book1.xlsx` and Lectern `Presentation1.pptx`, using the next free number. The page then builds a new document of that type and the extension opens it in its own tab. The tab you started from stays on its file. The new file's extension is the app's (`docx`, `xlsx` or `pptx`). A name that ends with something else keeps that ending and gains the app's extension, so `新建文档.txt` is saved as `新建文档.txt.docx`. Choosing a file that is already open in VibeOffice switches to that tab and leaves it unchanged. A template chosen in the app is written the same way. A web page, an e-mail message, the welcome document and Lectern's design-template pane still open in the current tab.

File ▸ Open uses VS Code's open dialog (on code-server, the files on that machine) and opens each chosen file in its own tab. Ledger can select more than one. The page's own file picker stays in use for a file dropped on the page, and in Quire for From existing document.

Recent Files in Getting Started and the File menu remember the workspace path of a file this editor opened or saved. Choosing one opens that file in its own tab, and Save writes it back there. Unsaved changes kept for that entry are opened into the new tab, still bound to the same file. An entry with no path asks for a file name, as File ▸ New does, and writes the copy Recent Files kept. A file that is already open is shown instead of being replaced.

File ▸ Close, when this tab has only that one document, closes the editor tab. With several documents open inside the one tab, it closes the current one.

File ▸ Save As asks for a file name, a folder and a type. Location starts as the folder of the file this tab has open. Browse opens VS Code's folder dialog (on code-server, the folders on that machine). When the name and the folder are the open file, that file is written. Otherwise the new file is written in the chosen folder (with ` (1)` before the extension if that name is already there) and opened in its own tab. The tab you saved from still shows unsaved changes for the original file. PDF, PNG and CSV from the other commands use the name the app chose and are written beside the open file.

File ▸ Enable Auto-Save sits directly under Save As, on the menu bar and on the ribbon's File menu. It is off until you turn it on, and a check mark shows that it is on. About a second after you stop editing, the open file is written, the same write as Save, and the tab clears. The extension remembers the choice (`vibeoffice.autoSave` in its global state), so the next time VibeOffice starts it is still on or off, in every Quire, Ledger and Lectern tab. The browser and the desktop app do not show the command. An edit still marks the tab dirty either way, and Save (Ctrl+S) still writes immediately.

File ▸ Exit closes the editor tab. Links to a web page or a mail address open outside the editor. The service worker is not registered.

Help ▸ Language chooses the UI language (Auto, English, 日本語, 한국어, Français, Deutsch, 简体中文, 繁體中文), the same list as the Start Center's language menu. The extension remembers it (`vibeoffice.lang` in its global state) and loads the page again with that choice, which is what makes the editor show it: assigning the same HTML does not. If the file has unsaved changes, the choice is kept and used the next time an editor opens. With the ribbon, the list is at the end of File, beside Theme. It is not on the menu in the browser or the desktop app.

After you change a page under `public/`, close the editor tab and open the file again. The extension reads `index.html` when the tab opens.

`node --test tests/vscode/extension.test.mjs` covers the file types, the webview page and the save messages.
