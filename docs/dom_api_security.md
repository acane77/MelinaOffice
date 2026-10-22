# Host abilities and audit

A Visual Basic member is either engine support or host support. Engine support edits the open document and works in the browser, the desktop app, and VS Code. Host support needs the machine the editor is running on: a path, a file, a folder, or a process. `Application.HasHostAbility(Ability)` reports whether this host implements that ability. It does not report whether the next call will be allowed. An unknown name is False. The comparison ignores case.

On the VS Code and code-server host (`VO.host.kind` is `vscode`) the abilities are Path, Save, SaveAs, Open, Print, FixedFormat, Hyperlink, Picture, FileDialog, Dir, and Shell. Quire also has InsertFile. Ledger also has SheetSaveAs. Lectern also has Export. Path, Save, Print, and FixedFormat are true only while the open file has a path. Everywhere else every ability is false, and a host member fails with `G-HOST`. The desktop app still uses handles, not paths ([docs/desktop.md](desktop.md)). `FullName` is the path on the extension host. On code-server that is the server's path.

## Three tiers

The tier is on the member. Who started the run decides whether a confirmation appears. An agent-written macro that the user later runs is a user run.

| Tier | Agent script (`vibeoffice_run_vbs`) | User run (DOM console, Run Script, a macro the user starts) |
|---|---|---|
| AlwaysAllow | No VibeOffice confirmation. VS Code may still ask, for example before an external link. | The same. |
| RequestApproval | Allow, Always allow, or Deny. Always allow is remembered for that ability in the workspace and the tier stays RequestApproval. Dismissing the window, or leaving it for two minutes, is Deny. | No extra confirmation. Starting the script is the allow. A file picker or an input box still appears, because that window is the call. |
| AlwaysDeny | Refused. Code `E-DENIED`. There is no Always allow. | The same. |

`E-APPROVAL` means the user denied a RequestApproval call. `E-DENIED` means the member is AlwaysDeny. Neither code is `G-HOST`. `G-HOST` means this session does not have that host ability.

Any read or write whose path is outside the workspace is RequestApproval for that call, even when the member's own tier is AlwaysAllow. The call is offered. If VS Code then blocks it (workspace trust, a remote file system, an external link), the error is VS Code's.

A file or folder window is VS Code's open or save dialog. That window is the approval: the script does not also ask Allow or Deny. Message boxes, input boxes, and the Office windows the app already has open inside the page. The script waits until the window closes.

The minute allowed for one agent script does not run while an approval window or a VS Code file dialog is open. The approval window itself gives up after two minutes.

Shell is AlwaysDeny. The function exists, `HasHostAbility("Shell")` is true on this host, and the call returns `E-DENIED` without starting a process. Declare stays unsupported. It is not an AlwaysDeny stub, and there is no Win32.

## Members

| Ability | Members | Tier | What the call does |
|---|---|---|---|
| Path | `FullName`, `Path` | AlwaysAllow | The open file's path, and the folder that contains it. |
| Save | `Save` | RequestApproval | Writes the open file, including a close that saves. A document that has never been saved stays `G-HOST`. |
| SaveAs | `SaveAs`, `SaveAs2` | RequestApproval | Writes the chosen path, which may be outside the workspace, and opens it in another tab. The tab you started from keeps its unsaved edits. A password or an unsupported format names that parameter and does not write. |
| Open | `Documents.Open`, `Workbooks.Open`, `Presentations.Open` | RequestApproval | Opens the path in its own tab. Extra arguments, including a password or read-only, are an error. The call does not return the other document. |
| Print | `PrintOut` on the document, the workbook, a worksheet, and page setup | AlwaysAllow | A PDF of the whole document, or of that sheet, beside the open file. The name is the file name without its extension, plus `.pdf`. Copies, the printer, and a page range are ignored. Print preview stays `G-HOST`. |
| FixedFormat | `ExportAsFixedFormat` | RequestApproval | PDF, or an omitted format, to the given path. The same name as PrintOut when the path is omitted. XPS fails before any confirmation and before any write. Page ranges are ignored. The file is not opened afterwards. A Ledger worksheet exports that sheet. A workbook exports the visible sheets. |
| Hyperlink | `FollowHyperlink` | AlwaysAllow | Opens an `http`, `https`, or `mailto` address. Any other scheme is an error. `NewWindow`, `AddHistory`, `ExtraInfo`, `Method`, `HeaderInfo`, and `SubAddress` are errors when passed. VS Code asks before an external link. |
| Picture | `AddPicture` | RequestApproval | Reads the file and embeds a PNG, JPEG, EMF, or WMF. Ledger converts SVG to PNG. `LinkToFile` True is an error. A position or a range argument is an error. |
| InsertFile | Quire `Range.InsertFile` and `Selection.InsertFile` | AlwaysAllow | Inserts `.txt`, `.htm`, `.html`, `.docx`, `.docm`, or `.dotx`. A bookmark range, `ConfirmConversions`, `Link`, or `Attachment` is an error. A path outside the workspace asks first. Ledger and Lectern have no InsertFile. |
| Export | Lectern `Presentation.Export` | RequestApproval | PNG or JPEG of every slide, including hidden slides, as `Slide1.png` in the given folder. `ScaleWidth` and `ScaleHeight` are pixels. One of them follows the slide's shape. Both omitted uses the slide width times 2.6667. GIF, EMF, TIFF, and BMP are errors. |
| SheetSaveAs | Ledger `Worksheet.SaveAs` | RequestApproval | That sheet as CSV, text, or a one-sheet workbook. Another format, a password, or a code page is an error. |
| FileDialog | `GetOpenFilename`, `GetSaveAsFilename`, `FileDialog`, and the file-open and file-save Office ids | RequestApproval | VS Code's dialog. Cancel returns an empty string. `MultiSelect` returns an array of paths. The dialog is the approval. |
| Dir | `Dir`, `HostDir` | RequestApproval | The engine's private virtual files first. Those do not hold the document. A miss lists the real folder, and that step asks. A further `Dir` with no argument continues the same list. A hit in the virtual files does not ask. |
| Shell | `Shell`, `HostShell` | AlwaysDeny | The call is refused. No process is started. |

`Application.Dialogs.Item(id).Show` (Quire and Ledger) suspends until the matching in-app window closes. The ids that open a file use the VS Code dialog instead. An id the app does not have stays `G-HOST`. The call returns `-1` when the window closes; the button that closed it is not recovered. Lectern has no `Dialogs` index. `MsgBox` returns 1, 2, 6, or 7 for OK, Cancel, Yes, and No. `InputBox` returns an empty string on Cancel.

Windows, panes, the view, and `NewWindow` stay `G-HOST`.

The catalog line records the tier and the ability after the member, for example `FullName get AlwaysAllow Path` and `Save call RequestApproval Save`. `dialog` on a line means the file dialog itself is the approval. `ExportAsFixedFormat` checks for XPS before the tier is applied.
