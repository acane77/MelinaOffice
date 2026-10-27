# Ecrivaine — Web Edition

A Word 2003–style word processor written in plain HTML and vanilla JavaScript (no build step, no
libraries). It reads and writes Office Open XML (`.docx`, `.docm`, `.dotx`, `.dotm`) files.
In a ribbon look (Golden Gate, the default, or Ribbon 2010; Help ▸ Theme, or Tools ▸ Options ▸ General ▸ Theme) a
ribbon with Word 2016's tabs (Home, Insert, Design, Layout, References, Mailings, Review, View and
the tool tabs) takes the place of the menus and toolbars.

## Using it

Open https://acane.net/melinaoffice/ecrivaine/. It starts with a blank document (`Document1`). The first time, the task pane stays closed on every device; opening it keeps that choice ([suite.md](suite.md#the-ribbon)). The app is `public/ecrivaine/`; every path below is relative to it.
Everything runs in the browser: there is no back end, no account and no network service behind any
feature. The proofing dictionaries are fetched at run time from `../common/dict/` (shared with Matheuse),
so the site is served rather than opened from disk (a `file://` page cannot fetch them; editing still
works, but spelling and the thesaurus stay off). To run it locally, see [testing.md](testing.md).
The metric-compatible fonts (see *Rendering notes*) are served with the site from `public/fonts/`.
The status bar ends in a zoom slider (10–500%, with 100% in the middle). The percent can be edited, and the slider moves the zoom as it is dragged. A pinch over the page zooms it; two-finger scrolling moves the page. A finger held within 10px of where it landed opens the right-click menu at the press, even when text is selected. A slide past that shows the formatting bar one line above the caret. A pinch on the toolbars or the status bar does not zoom the window. Typing keeps the page still while the caret is in the top nine tenths of the window. Once the caret's line reaches the bottom tenth, the page moves up by that line and the caret stays on that line. A line that has left the top of the window is brought back to the top edge.
In Print Layout, double-click the gray gap between two pages to hide the white space there: the top and bottom margins (and the header and footer in them) drop away, and a little of each margin stays so the pages meet on a black line. Double-click that edge to show the margins again. Tools ▸ Options ▸ View ▸ White space between pages is the same switch. Reading Layout, Print Preview and printing keep the whole page.
Selecting text with the mouse or the keyboard shows a small formatting bar just above and to the right of the pointer (in the Windows 7 look, a glass window with the controls drawn on the glass): font, size, grow and shrink, clear formatting, highlight, font colour, bold, italic, underline, strikethrough, subscript, superscript, bullets, numbering and the style list. It starts partly see-through, becomes solid while the pointer is on it or on a menu it opened, and fades out with the distance from where it opened, gone past 50 pixels. Once it has faded out it stays hidden until the next selection. A caret, typing, Escape or scrolling the page also puts it away.
Help ▸ Ecrivaine DOM Console (File, in the Paper 2016 look) opens a terminal on this page. What you run there uses the open document and stays on the machine. The bar under the title is Copy, Paste, Clear, DOM API References, and Language. Language is Visual Basic 6 or JavaScript; Visual Basic 6 is selected until you choose the other, and the choice is kept. Choosing a language clears the terminal and prints that language's welcome. Closing the window is the title's close button. Paste puts the clipboard into the terminal; when that text does not arrive, a box takes the paste and OK puts it in. DOM API References opens the Visual Basic DOM API Reference on the task pane. That page lists the Visual Basic this app runs, how it differs from Microsoft VBA and VBScript, and the members on the open document. It stays open beside the console.

Visual Basic 6 runs each finished entry with the Immediate session. Enter runs a finished line. Shift+Enter, a line ending in `_`, or an unfinished `Sub`, `Function`, `If` or `With`, keeps the line open. Assign with `x = 1`. The name stays, and a later line can use it. `?x` prints the value. A whole script is not compiled here. Run Script, on the bar only in Visual Basic 6, opens a modeless window: an editor, a read-only Output box, Run and Close. Run compiles the buffer and runs `Sub Main`. A compile error opens a message and does not run. Save, including Save when closing or switching macros, does the same and does not write the macro. Each Run replaces the Output box with that run. Debug.Print is written into that box when the statement runs, so the line is visible while a message box is still open. The editor text and that output remain after a refresh. When that editor has never held a script, or the script was empty when the window last closed, it opens on a welcome `Sub Main`. The comments welcome the user, say to write Visual Basic and click Run, and link to the Word VBA object model. The procedure inserts one sentence at the cursor. A macro opened in the same window keeps its own source. That window does not call macros stored in the document.

Tools ▸ Macro ▸ Macros (Alt+F8; View ▸ Macros in the Paper 2016 look) first shows a notice titled Notice on Macros in Ecrivaine, unless Never show this message again was checked the last time OK was clicked. The box is unchecked when the notice opens. The notice says macros created by Microsoft Word are not executed in Ecrivaine, that Ecrivaine saves the document text rather than a Word VBA project, that macros written here and macros written in Word do not load in the other, and that these macros are for editing with Visual Basic and for agents rather than for Word compatibility. OK opens the Macros window. The title's close button leaves it closed. The window lists macros. Macros in is Current Document or All templates and documents. Current Document is selected when the window opens. Its macros stay in the open document, and Run, Edit, Delete and Create behave as before. All templates and documents keeps a separate list in this browser, the same list for every document opened here. Choosing it lists, runs, edits, deletes and creates only that list. Step Into and Organizer stay disabled. Run closes the window and runs the saved macro. Edit and Create close it and open the Run Script window with that macro's name, which cannot be edited there, and a Save button. A name that is not one identifier as the Visual Basic lexer reads it — a letter, underscore or code point from U+0080 through U+FFFF, then those or digits, with one optional type suffix, or a bracketed name — is refused at once. The message says the procedure name is not a legal Visual Basic identifier, and dismissing it closes that window. Run Script opened on its own has no macro name to check. Save writes the source back to where that macro is kept. Run in that window tries the buffer together with the other macros in the same list, without saving. Closing it with unsent changes asks whether to save. There is no macro recording.

View ▸ Melina Office Visual Basic (and the same command on the classic menu) opens a modeless window of that title. It edits one standard module named Module1 and debugs it against the open document. File ▸ New Blank Project starts from Sub Main, with the comment Insert your code here and a blank line (with 4 spaces) before End Sub, and clears the breakpoints and Output. File ▸ Open Project reads a JSON file that keeps the source, the breakpoints (a line, a condition and whether it is enabled), the watches, the expression and the last result. An older project whose breakpoints are only line numbers still opens. File ▸ Save Project (Ctrl+S) stores the source, the breakpoints and the Output text in this browser; the next time the window opens, that is what it loads. Nothing is stored until Save Project. File ▸ Save Project As writes the JSON file. Open Visual Basic File puts a source file into the editor and leaves the breakpoints. Clicking the line-number gutter sets or clears a breakpoint on that line. Double-clicking a line number opens a small panel with that line number, its condition and whether the breakpoint is enabled; OK writes them back, and Cancel leaves them. An empty condition always stops. Start Debug (F5) runs Sub Main and stops on the enabled breakpoints. A condition is evaluated there: False, 0 and an empty value continue without stopping, and a condition that fails to evaluate stops and writes the error to Output. Start Run runs the same module and does not install breakpoints. While the program is paused, Continue, Step Into and Step Over are the virtual machine's statement steps. The Breakpoints tab lists each breakpoint with its line, condition and the source at that line. A call-stack row shows the line and the procedure name, as in Line 71| Bind, and jumps to that line when it is in the editor. Object values are the paused frame: a scalar is shown as its value, and an object as [Object]. Expression evaluates in the selected frame, and only while paused. Each enabled watch is evaluated at that same pause. A row whose value changed since the previous pause is shown on the stronger highlight; the first value, and a row that was not evaluated this pause, are not. Debug.Print is written to Output. Holding the pointer over a name shows that name's signature or type. While Ctrl is held, that name is drawn as a link. Ctrl+click jumps to its definition in this module. Holding the pointer over an error's wavy line shows that error. A status bar along the bottom reads Ready while the program is idle, Running while it runs, Triggered Breakpoint when a breakpoint stops it, Triggered Error Interrupt when a run-time error stops it, and Break when a step stops it. Dragging an edge or a corner changes the window's size. The next time the window opens, it has the width, height and place it had when it last closed. The bar between the editor and the lower area changes their heights, and a later resize of the window keeps that proportion. That editor height is kept, and the next time the window opens it is the same. Closing the window stops the session. This window is not the document-macro editor, so Step Into in Macros stays disabled, and it does not design a form. View ▸ Run Form opens a `.frm`, and any `.bas`, `.frx` or `.vbp` selected with it. The window is a Melina Office dialog. Events run on the VB6 engine. A change to the document still goes through the DOM IR ([ecrivaine_vbs_form.md](ecrivaine_vbs_form.md)).

A macro is one standard module whose public Sub has the macro's name. Other procedures in that source are not listed; they can be called from that module. The modules are compiled together when the document is current, so one macro can call another by name. Document macros are kept in a custom XML part (`https://acane.net/melinaoffice/macros`) of a Word document (`.docx`, `.docm`, `.dotx`, `.dotm`). Word leaves that part in place. All templates and documents is stored in the browser (`melinaoffice.ecrivaine.allMacros`), not in the file. Normal.dotm is not a location in this window. A new document is built in memory, and `Normal.dotm` is only the template name written when the document does not name one. Attaching a template stores that file's name and, when asked, copies its styles; the file and its macros are not kept. Opening a `.dotm` makes that file the current document, so its macros are listed there. A VBA project in the file is a different thing: it is kept and is not listed or run. Saving as HTML, text, RTF or Markdown warns that those files cannot hold the macros. PDF export does not change the document file. Matheuse's Tools ▸ Macro ▸ Macros shows Notice on Macros in Matheuse, then the same window. A VBA project in an `.xlsm` file is kept and is not listed or run. Oratrice has the same window for a presentation.

JavaScript is the other language. Enter runs a finished line. Shift+Enter, or a line that still has an open bracket or string, keeps the line open. A function written at the top level stays, so the next line can call it. The welcome says that if the right-click menu cannot paste, use the Paste button; that at the top level you write `x = 1` rather than `let`, `var` or `const`, which do not carry to the next line; and that JavaScript runs here as well, but the F12 console is the better place to run a JavaScript script.

## Proofing tools (offline)

| Feature | How it works |
| --- | --- |
| Spelling (as you type and Tools ▸ Spelling and Grammar) | Word lists for English (U.S.) and English (U.K.) in `common/dict/en_US.words` and `common/dict/en_GB.words`: SCOWL's own lists at size 60 (the ones its Hunspell en_US and en_GB-ise dictionaries are made from; `tools/dict.py` rebuilds the British one), ranked by word frequency so suggestions come out in a sensible order. Unknown words get red wavy underlines (CSS Custom Highlight API); right-click for suggestions, Ignore All and Add to Dictionary. Underlines are off when a document opens (and the dictionary is not loaded): click the spelling icon in the status bar to check that document as you type, click again to stop. |
| Grammar | A rule checker for the mistakes Word 2003 flagged most often: repeated words, a/an, capitalisation at sentence start, spacing around punctuation, commonly confused words, simple subject–verb agreement. Green wavy underlines. |
| Thesaurus and Research pane | `common/dict/en.thes`, built from the WordNet-based thesaurus LibreOffice ships (`th_en_US_v2`, MyThes format), with meanings grouped by part of speech, related words, antonyms and a Back history. "Search This Document" finds the word in the open document. |
| Custom dictionary | Tools ▸ Options ▸ Spelling & Grammar; words you add are stored in the browser. |
| Readability statistics | Tools ▸ Options ▸ Spelling & Grammar ▸ Readability Statistics (counts, Flesch Reading Ease, Flesch–Kincaid Grade Level, passive sentences). |

Licences for the word lists are in `public/common/dict/LICENSES.txt`.

## Rendering notes

- Layout follows Word's own rules where they differ from CSS: space before/after collapses to the
  larger of the two (unless the document asks for HTML-style spacing), space before is dropped at the
  top of a page reached by overflow, by "page break before" or (Word 2013+ documents) by a section
  break, a line with "multiple" line spacing only needs its single-spaced height to fit at the bottom
  of a page, and Word 2013+ justified text may squeeze spaces by up to 20% to fit a word on a line.
  Adjacent paragraphs with the same borders and the same left and right indents share one box: the
  rule is drawn only under the last of them, so the blank lines in front of a horizontal line stay
  blank lines.
- More of Word's own habits: a right or centred tab stop placed beyond the right indent lets the page
  number of a table of contents run out to it; superscripts and footnote references do not push their
  line down; table rows are as tall as Word makes them even though the browser draws hairline borders a
  whole pixel wide; East Asian documents with a line grid snap their lines to it (below); a "next column"
  section break in a one-column layout starts a new page; "page break before" on the first paragraph
  of a table breaks before the table; Word 2013+ documents drop space before after a manual page break;
  Word 2010+ documents keep the paragraph mark of a page break that ends a section's last paragraph on
  the page of the break (older ones carry it over to a page of its own); an empty paragraph holding only
  a section break makes no page of its own when the next section starts on a new page anyway: pushed off
  a full page, or past a page break unless it has "page break before" of its own, it stays at the foot
  of that page. Of 153 Word-saved corpus files with such a paragraph, 3 more now show Word's stored page
  count (113), and none moves away from it; for issue 20, of the 57 Word 2007/2010 files with a page
  break and a section break or a break at their end, 4 more show it (42 in all), one more comes closer,
  and none moves away. A centred table wider than its column overhangs both margins evenly, a
  right-aligned one the left margin. A top-and-bottom wrapped object clears its vertical span in every
  column it crosses, the way LibreOffice widens that obstacle to the column's print area; an object that
  stays inside one column leaves the others where they are.
- East Asian text is drawn with the run's East Asian font (`w:eastAsia`, or the theme's for the
  document's East Asian language) and marked with its East Asian language, and takes that font's line
  height where it is the taller ([i18n.md](i18n.md), Fonts). Format ▸ Font sets the two separately:
  Asian text font is `w:eastAsia`, Font is `w:ascii` and `w:hAnsi`. "(Use Asian text font)" writes the
  Asian family into the Latin slot as well. The formatting toolbar's Font box writes by family: an East
  Asian family goes to both slots (it draws Latin as well as Chinese), a Latin family goes to the Latin
  font only and leaves the East Asian face in place. Typing in a Font box, on the toolbar or in Format ▸ Font,
  completes to a listed family, a recently used one before the others. When the two
  names differ, East Asian characters are drawn with the East Asian font and the other characters with
  the Latin font. In a Chinese document every symbol of East Asian Width Ambiguous (the ellipsis,
  the middle dot, the multiplication sign, the degree sign, the dashes, the quotation marks, the
  roman numerals and the rest) is drawn with the East Asian font, including beside digits. Accented
  Latin, Greek and Cyrillic keep their own font. Other general punctuation is drawn with the font of
  the characters beside it, and a run of only that punctuation uses the East Asian font. Where East Asian text (ideographs, kana or bopomofo) touches a Latin letter or a digit, Ecrivaine leaves a quarter of that East Asian character's size between them, as Word does (`w:autoSpaceDE` for letters, `w:autoSpaceDN` for digits; both on unless the paragraph turns one off). The quarter is the East Asian size, including when the two characters are in different runs. Nothing is added beside punctuation, Hangul or a space that is already there, and a paragraph snapped to a character grid does not add it. Distributed alignment spreads the spare space between fullwidth characters. A
  justified or distributed line that is only a little wider than the measure narrows fullwidth
  punctuation (down to half an em, and only by as much as it overflows) so the line stays one line;
  left alignment does not, and wraps. With "balance single-byte and double-byte characters",
  an ASCII space becomes half a fullwidth em when it touches another space or sits at the end of its
  run: the two spaces between a heading number and its title, and an underlined blank on an East Asian
  cover. A single space between Latin words keeps the font's own width. A new document's Latin and East Asian fonts are both 等线 (a Simplified Chinese
  document keeps that face at 10.5pt). The page follows the UI's language: A4, Office's margins, and
  Japanese or Chinese line grids (i18n.md, Defaults). A new document is laid out as if created in
  Microsoft Word 2013 and later (Tools ▸ Options ▸ Compatibility). An opened file keeps the compatibility
  it was saved with.
- Each line is as tall as what it holds, as Word sizes lines (`js/portions.js`, after Apache OpenOffice
  Writer's line portions): a line with text by its text, a picture alone in its line exactly the
  picture's height (Ecrivaine added the text's descent below it, 3 to 7pt), a large first word no longer
  makes every line of its paragraph that tall, and blanks and tabs don't raise a line. A line holding
  nothing that counts takes the paragraph mark's font, or that of the line break that ends it; with
  exact spacing every line is the value, whatever sizes its text has; a Symbol or Wingdings bullet's
  line is as tall as that font makes it, though it is drawn in Arial. A Word-saved 19-page walkthrough
  that made 20 pages (two pictures pushed a page break onto a page of its own) now breaks where Word's
  PDF does on every page, and two lines that follow each other there are as far apart as in Word
  within 0.51pt (`tests/ecrivaine/lines.test.mjs`); 15 synthetic cases match LibreOffice within 0.2pt. Of
  2,289 Word-saved corpus files, 2,136 show Word's stored page count (2,127 before); the 10 whose count
  changed all moved toward Word's, none away (`tools/ecrivaine-pages.mjs`).
  Not covered yet: a picture in a line with exact spacing (Word clips it to the line), a picture alone
  with multiple spacing or on a grid, a line of only a small picture (Ecrivaine follows LibreOffice: the
  mark's line, the picture at its foot), and a large paragraph mark after small text (matched to
  LibreOffice only).
- On a document line grid (East Asian documents) a line is the grid pitch times its "multiple"
  spacing, and at least the whole pitches its text needs: 1.5 lines on the usual 15.6pt grid are
  23.4pt, and 14pt text at 1.7 lines takes two pitches, 31.2pt (issue #11). Word centres the text in
  its line, and at the foot of a page only the space below the text may run into the margin. Against
  Word's PDF of 90 cases (`tools/linegrid.py`: 10.5 to 28pt text, 0.8 to 3 lines, 15.6pt and 18pt
  grids, English and Chinese), all 90 line heights match, and pages 1 to 5 and the 18pt grid's pages
  break where Word's do. Of the 135 Word-saved corpus files with a grid, 126 show Word's stored page
  count, as before. A table cell does not take that grid unless the document sets
  `adjustLineHeightInTable`: the cell keeps the paragraph's own line spacing, so a font taller than
  one pitch does not gain a whole extra pitch of empty space above and below the text (LibreOffice
  `CalcRealHeight`, the same default as a DOCX that omits the flag). Body text still snaps.
- Page and section totals are filled in once pagination is complete, and tabbed text around them
  is measured again. Pages are laid out again only from the first page where a new total changes a
  paragraph's height (a line wraps, a footer grows), so a long document with "Page X of Y" opens in
  one layout pass (181 pages: 330 ms, was 600 ms with a second full pass). This handles changes such
  as 9 to 10 or 99 to 100 pages, including when deleting pages reduces the count again.
- Formatting symbols are drawn over the text without changing run shaping, line wrapping or
  paragraph heights. The symbols are omitted from PDF and print output. Showing hidden text can
  still reveal additional content and affect layout. Marks follow rotated text boxes, zoom,
  right-to-left paragraphs and vertical table text.
- Encoded carriage returns (`&#13;`) inside DOCX text runs render as line breaks; CRLF pairs render
  as one break, including pairs split by formatting, bookmarks or inline content controls.
  Their original text characters and editing offsets survive save and reopen. Forced lines
  are counted during pagination even when a centered or right-aligned line starts farther right.
  This reproduces the line-break symptom in issue #1 with a synthetic fixture; the reporter's
  original document has not been available to confirm that it uses the same encoding.
- Equations (Office Math, `m:oMath`) are rendered as MathML: fractions, scripts, radicals, delimiters,
  n-ary operators, accents, matrices and equation arrays. An equation that is not edited is saved
  unchanged. An equation is one character: a click selects the whole of it, and the caret moves to the
  character before it or the character after it. Double-click, or Edit Equation on its menu,
  opens a LaTeX editor with a live preview. The field is filled from the equation (linear LaTeX as
  stored, otherwise LaTeX derived from the OMML). Update writes that LaTeX back as OMML and stays
  disabled while the LaTeX has a syntax error. The same menu has Change to Inline on a display
  equation and Change to Display on an inline one: the OMML is wrapped or unwrapped, and what was
  not edited is kept. Justification (Left, Right, Centered, Centered as Group) is on that menu for a
  display equation whose paragraph is justified, and is grey otherwise. The choice is saved as `m:jc`
  on the math paragraph; when it is omitted, Word's Centered as Group is used. Left, Right and Centered
  place the equation that way. Centered as Group centers one equation, centers a group of equations
  that stay aligned with each other, and puts an equation number on the right margin. A display
  equation that shares the paragraph with other text breaks to the next line: under Justify the text
  stays at the left and the equation is centered (Centered as Group) or placed by the justification
  that was chosen, at the size of an inline equation. Insert ▸ Equation (and Insert ▸ Object ▸ Equation,
  and Equation in the ribbon's Symbols group) opens the same editor; Insert adds the equation at the
  caret, as a display equation when the paragraph is empty and as an inline equation when it already
  has text. Insert ▸ Object ▸ Equation (Legacy) still inserts Cambria Math italics.
  An equation sits in the line, as Word does, so a table cell does not gain a blank line above
  and below it. An inline
  equation that is the only thing in its paragraph is saved with a space after it, so Word keeps it
  inline. An equation wider than its column or table cell
  breaks before a binary operator (+, =, and the like), as Word does. A line break beside an
  equation is saved as its own paragraph: Word would otherwise join the equations into one and drop the break.
- Content controls bound to custom XML or to the document properties show the bound values, as Word
  does when it opens a file. Linked text boxes continue their story from one box to the next.
- SmartArt is drawn from the picture Word stores with it; files saved without that picture (Word 2007
  and some generators) get a simple diagram built from the SmartArt data.
- Fonts are drawn with the fonts installed on your computer. When a document uses a font you don't
  have, Ecrivaine substitutes a metric-compatible one where one exists (Carlito for Calibri, Caladea for
  Cambria and Gelasio for Georgia, served with the site; Liberation or Tinos/Arimo/Cousine for Times New
  Roman/Arial/Courier New where installed; TeX Gyre
  fonts for Palatino, Bookman, Century Schoolbook and Century Gothic), so line breaks and page counts
  stay close to Word's.

## File formats

| Format | Open | Save |
| --- | --- | --- |
| Word Document `.docx` / `.docm` | yes | yes |
| Strict Open XML `.docx` | yes | standard Transitional OOXML (recorded, not shown) |
| Word Template `.dotx` / `.dotm` | yes | yes |
| Web Page `.htm` / `.html` | yes | yes (single file, images inlined) |
| Rich Text Format `.rtf` | yes | yes |
| Plain Text `.txt` | yes | yes |
| Markdown `.md` / `.markdown` (GitHub flavour) | yes | yes |
| Markdown with pictures `.zip` (a `.md` and its picture files) | yes | yes |
| PDF | — | yes (File ▸ Print, or Save As ▸ PDF) |
| Password-protected `.docx` | yes (Word 2007 "standard" and Word 2010+ "agile" encryption) | yes: Tools ▸ Options ▸ Security ▸ Password to open (AES-256, as Word 2013+) |

Encryption and decryption happen in the browser (WebCrypto for hashing, so the page must be served over
https or from localhost). A document opened with its password keeps that password when saved.

## Markdown

Ecrivaine reads and writes the Markdown GitHub uses: CommonMark plus tables, task lists, strikethrough,
extended autolinks, footnotes and alerts (`> [!NOTE]`).

| In the file | In Ecrivaine | Written back as |
| --- | --- | --- |
| `#` … `######` headings | Heading 1–6 | `#` headings (Title as `#`, Subtitle as `##`) |
| `**bold**`, `*italic*`, `~~strike~~`, `` `code` `` | Bold, italic, strikethrough, the HTML Code character style | The same; where a delimiter could not open or close, `<strong>`, `<em>`, `<del>` |
| Links, autolinks, `[text](#heading)` | Hyperlinks (titles as tips); `#heading` follows GitHub's heading anchors | The same |
| `-` / `1.` lists, nested, tight or loose | Bullets and numbering per list level (List Paragraph) | `-` and `1.`; tight unless the items have space between them |
| `- [ ]` / `- [x]` task lists | Check boxes that tick on click | `[ ]` / `[x]` |
| `>` quotes, nested; `> [!NOTE]` and the other four alerts | Quote style with a bar per level; alerts with their title and colour | `>` and `[!KIND]` |
| Fenced and indented code | HTML Preformatted paragraphs, the language kept | Fenced code with the language |
| Pipe tables with alignment | Tables (Markdown Table style: header row, banded rows) | Pipe tables; merged cells or several paragraphs in a cell as an HTML `<table>` |
| `---` | A paragraph with a bottom border | `---` |
| `[^label]` footnotes | Footnotes | `[^label]` and its definition |
| Pictures | Pictures from `data:` URLs, or from the `.zip`; others as a box with the alt text, the address kept | `![alt](path)`; a resized picture as `<img width>`; new pictures go into the `.zip` as `images/…` |
| `$…$` math, emoji `:shortcodes:`, Mermaid and other diagram code | Kept as written (math in the Markdown Source character style, diagrams as code) | Unchanged |
| Inline HTML GitHub allows (`<sub>`, `<sup>`, `<ins>`, `<kbd>`, `<br>`, `<img>`, `<a href>`, `<a name>`) | Subscript, superscript, underline, HTML Keyboard style, line break, picture, link, bookmark | The same tags |
| HTML blocks, YAML front matter, link reference definitions, comments | Blocks through the HTML reader; front matter, definitions and comments are kept but not shown | Unchanged unless edited |

Saving keeps the file as it was where it was not edited: each block of the file remembers its lines,
and a block whose content is unchanged is written back byte for byte (line endings, a byte-order mark,
`*` or `_`, setext headings, reference links and line wrapping included). Edited and new blocks are
written in one style. What Markdown cannot express (fonts, colours, alignment, page setup, headers and
footers) is not saved; a `.md` file cannot hold pictures, so saving one that has new pictures offers
the `.zip` form instead. A `.zip` is saved with every other file it held, unchanged. Unsaved versions
of a Markdown document (autosave, Document Recovery) are kept as Markdown.

Checks (tools in docs/testing.md):

| Check | Result |
| --- | --- |
| Reader against the CommonMark 0.31.2 spec examples (rendered as the reference implementation renders) | 652 of 652 |
| Reader against the GFM spec's extension examples (tables, task lists, strikethrough, autolinks, disallowed raw HTML) | 22 of 22 |
| Spec examples opened, every block rewritten, saved and read again, same HTML | 559 of 652; 507 of the 544 without raw HTML (the rest: nested emphasis of the same kind, `*a *b**`, which a document cannot hold; empty links; a quote that starts inside a list item's first line) |
| 24 real files (READMEs of React, Node.js, VS Code, Rust, Kubernetes, TensorFlow, Vue, Deno, Astro, Express, Go, TypeScript, cmark-gfm, markdown-it, awesome, public-apis, free-programming-books, the Airbnb style guide, GitHub's own formatting docs) saved unchanged | 24 of 24 byte for byte |
| The same files with one paragraph edited | Only the edited block's lines change in all 24 |
| The same files with every block rewritten, same HTML | 7 of 24; all 5 without raw HTML (the others turn logos, `<picture>` and centred `<p align>` blocks into formatting) |
| Opening time | 300 kB (public-apis): 80 ms; the spec examples twenty times over (300 kB): 70 ms to parse and render |

## What a save keeps, converts and drops

Ecrivaine keeps package dependencies, control boundaries, text/section properties and opaque object XML beside
its editable model. Save and drafts retain DOCX/DOCM/DOTX/DOTM and encryption. The Compatibility
Checker shows, before a user download, only the losses that cost the user content or a working
feature (see docs/suite.md, What the user is told). Drafts retain notices silently, and cancelling a conversion does not leave its notice on a later preserving save.

The current pinned **2,902-file** run uses frozen writer `984eaf3`, including section and numbering
preservation. **2,880 save/reopen and draft-recovery attempts complete**, with **22 failures** in each
scenario. Text edit/save/undo/redo completes in **2,281**, with **599 exclusions and 22 failures**.
Protected, malformed and mislabeled inputs remain explicit failures.

Across **19,028 emitted/attempted states**, package/SDK comparison reports **18,056 OK, 373 failed
and 599 excluded**, with no added diagnostics. Save/reopen alone has **2,828 OK and 74 failed**;
unvalidatable originals remain failures. Comment, formatting, compatibility-attribute, drawing-ID,
break-recovery and structure-loss notice corrections are included in this full run. Two initial text
timeouts pass focused retries; the original failed attempts remain recorded alongside those results.
Retained content properties normalize known invalid order, redundant identical properties and
unambiguous color spelling, with one recorded notice per affected source part.
Styles, conflicting values and unknown Word children remain intact. A corpus scan found 74 candidate
files; all save/reopen, and all 71 with validatable originals add no package/SDK errors. The three
unvalidatable originals remain failures. Independent previous-writer comparisons find only the
declared formatting repairs and generated timestamps; other content and opaque bytes are unchanged.
All six regression saves open in LibreOffice with unchanged page counts. Diagnostic matching follows
style IDs and unchanged property content when positions move, while retaining occurrence counts.
Newly generated drawings reserve both document drawing IDs and part-local shape IDs, so converting
one object cannot collide with the identity of an unrelated retained group.
An Office-found comment repair is also fixed: pre-release comment extensions retain their original
namespace and unchanged bytes, and a comment is written with one reference. Removing a duplicate
reference records a loss entry. The pinned scan finds one saved file with excess
anchors and 55 automatic extension-namespace promotions; the corrected point-comment save and a
current-format reply control both pass Word. The current full run includes these corrections.

Rewritten parts now hoist fragment namespaces and `mc:Ignorable` to the root. On the same 400-file
size sample, main XML shrinks from **22,183,474 to 4,770,570 bytes**, against **5,373,347 bytes** in
the originals. Independent expanded-XML comparisons preserve all content in **1,845 changed parts**;
there are no new package issues or inner-`mc:Ignorable` convention signatures. The compact Word
sample passes Word, SDK and LibreOffice. Opaque parts retain their bytes. This is an isolated
serialization measurement on existing saves, with fresh app saves for Office, not a new full corpus run.

The feature inventory compares **2,864 readable pairs**, with **38 failed attempts**. It retains
**776,207 of 776,869 counted words**; **2,811 documents** keep every counted word. Data-bound values,
malformed input, unreferenced stories and unsupported markup remain exceptions.

| Status | Feature | Current measured result / limit |
|---|---|---|
| Kept, with approved Strict conversions | Package metadata, custom XML and templates | 965/965 custom-XML graphs and property dependencies (960 exact, 5 namespace conversions); 353/353 typed custom-property bags, 102/102 template links (one namespace conversion), 17/17 VBA and 1/1 label relationships |
| Kept, with source exceptions | Glossary, embedded fonts and identity metadata | 193/195 glossary relationships (two missing original dependencies), 20/20 referenced fonts, 95/95 people and 15/15 comment-identity relationships; 3 orphan font parts are not reattached |
| Kept, with a remaining exception | Content controls | 2,675/2,676 wrappers; inline, block, row and cell controls retain properties and scalar bindings |
| Kept | OLE, SmartArt, charts and embedded documents | 516/516 OLE/object inventory nodes, 39/39 SmartArt data parts, 133/133 chart parts and 4/4 altChunk blocks; geometry edits retain payloads |
| Kept | Text effects, pictures in page backgrounds and drawing properties while untouched | 205/205 Word 2010 effect entries, 62/62 backgrounds, 1,054/1,054 text boxes and 1,803/1,803 DrawingML shape entries |
| Kept, with a reported source exception | Section properties, their alternatives and header/footer dependencies | 3,404/3,405 section entries and 175/175 multi-column sections; 48 focused states cover property edits, commands, copies, history and drafts without added package/SDK diagnostics. Malformed `tdf108849` has two final section definitions; one is omitted with a notice |
| Kept; Office samples accepted | Picture bullets, including unused definitions, image dependencies and imported identities | 81/81 definitions and 106/106 references in the full inventory. The 36-input focused sweep retains image bytes and identities through edits, copies, history and drafts; `tdf149089` keeps its original two pages |
| Kept; Office samples accepted | Numbering identities, definition/level metadata, legacy settings and format alternatives | 9,202/9,202 definitions, 8,923/8,923 list identities, 8,910/8,910 template entries, 232/232 legacy settings and 103/103 cleanup entries. List commands, history, drafts and independent copies pass 72 assertions across 33 distinct saved states |
| Kept, with range/story exceptions | Comments, notes, bookmarks and permissions | 1,404/1,421 comments, 674/684 footnotes, 57/57 endnotes, 4,870/4,965 bookmarks and 50/54 permission ranges |
| Kept, with remaining losses | Revisions, fields, pictures and tables | 1,385/1,427 insertions/deletions, 1,087/1,103 formatting revisions, 2,590/2,711 fields, 1,239/1,242 pictures and 1,605/1,606 tables |
| Converted when edited | Unsupported contents inside an opaque frame; typed/complex controls | Edited or ungrouped converted content uses the model representation; incompatible typed-control edits unbind and report the conversion |
| Normalized, with a notice | Invalid paragraph/run formatting that edits can duplicate | Six regression files / 36 emitted states pass; 74-file candidate scan checked, with 34 saves changing formatting; conflicting or unknown source properties remain intact |
| Recovered, with a notice | Missing/incomplete table grids, conflicting default styles and misplaced page/line breaks | The writer reports reconstructed column widths, normalized defaults and breaks moved into valid paragraphs/runs. Removing recovered content clears its save notice; undo/drafts restore it. Grid/default recovery can change pagination; retaining misplaced breaks restores the affected four-page document |
| Mostly converted, with a notice | Smart tags and inline custom-XML markup | 18/519 original wrappers remain; their text is represented as ordinary content |
| Dropped, with a notice | Extra document bodies in malformed input | `MultipleBodyBug` contains three bodies; the first remains and the other two are omitted. Their content loss is explicitly reported before Save |
| Kept; Office accepted | Per-property drawing edits, row/cell alternatives and watermarks | 62 focused checks and 48 save/edit/history/draft states pass; full corpus completed; 14 original batch-2 files pass Office and the comment repair has a confirmed correction |

Package counts verify bytes, types, rIds and targets; unreferenced parts have separate loss entries.
The prepared-save audit reports missing wrappers, fields, move/revision markup, controls, bookmarks,
permissions and notes/comments. Editing an imported chart records its formatting conversion.
UTF-16 XML parts use the shared decoder. The final full run includes these reporting changes.
The raw carry audit reports 2,795 exact passes and 107 findings: 38 failed driver/package reads,
67 approved Strict conversions and two incomplete glossary graphs. An independent namespace-aware
comparison confirms that the converted graphs retain their other content and targets. All 2,653
web-settings relationships survive (2,586 exact and 67 converted).
The edited inventory compares 2,270 readable pairs and retains 2,505/2,505 control wrappers,
809/809 text boxes and 1,046/1,046 DrawingML shape entries. Its 632 missing/unreadable copies include
the driver's 599 deliberately excluded edits; they are not counted as successful comparisons.
The user accepted the bound-text, formatted-effects, resized-SmartArt, pasted-OLE, macro and namespace
samples in Office batch 1. All 15 batch 2 files open in LibreOffice with unchanged counts. Strict
conversion passes **67 documents / 268 emitted states** against original package/SDK diagnostics and
keeps its notice through draft recovery; its Office review is included in batch 2.
Current full reports and exceptions: `~/corpora/results/ecrivaine-final-7c25bd7/`. Earlier focused
measurements and failed attempts remain in the report directories linked by its README. Section,
numbering and picture-bullet corrections are included in the full run. Missing table grids and
conflicting default styles have save notices; these recoveries can still change pagination.
Office batch 5 files 01–05 now cover picture bullets, custom formats, edited/copied style-linked
lists and section settings. All five add no package/SDK diagnostics and retain their LibreOffice
page counts; their previous-writer differences are reviewed. The user accepted all five in Word.
A fresh LibreOffice check of the exact accepted files also passes; hashes and results are in
`~/corpora/results/office-accepted-2026-10-09/`.
The pinned inventory finds numbering parts in 810 inputs. The writer keeps their source level counts,
including absent defaults, rather than serializing all nine levels used internally for layout.
Independent list copies now also receive fresh eight-digit `nsid` identities. Two command fixtures
pass 44 assertions and 23 package/SDK comparisons, including repeated imports, undo/redo and drafts.
Multi-column index insertion keeps the original section start before the index and makes the index
and remainder continuous; two fixtures pass 18 assertions and 14 package/SDK comparisons through
undo/redo and draft recovery. All four selected saves open in LibreOffice with unchanged counts.
These corrections passed Word in Office batch 7 files 01–02 and are included in the full refresh. Focused evidence:
`~/corpora/results/lossless-evening-review-2026-10-09/`.
The render review also found seven misplaced breaks in `tdf108714`, which the reader previously
skipped. Recovering them in valid paragraphs/runs retains all four pages on save and text edit,
with a recorded notice. The pinned scan finds one affected readable input and
records 33 archive/XML read failures separately. Seven command states and seven corpus
save/edit/history/draft states add no package/SDK diagnostics; 25 browser assertions and all 27
Ecrivaine unit tests pass. All seven emitted corpus states retain four LibreOffice pages.
Reports: `~/corpora/results/ecrivaine-break-recovery-2026-10-09/`.
Office batch 6 file 05 checks that correction in Word; its package/SDK and four-page LibreOffice
comparisons pass, and the user accepted it in Word.

The full LibreOffice pass renders **2,869/2,880 unedited saves**. Every unsuccessful saved render
also fails or times out in its original. With the break correction included, **12** readable
unedited pairs change page count. Across all **19,028** states, the paired render/count results are
**18,000 OK, 429 failed and 599 excluded**. Missing outputs, unrenderable originals, saved conversion
failures and count differences remain separate in the reports; these are not lossless-save percentages.
The current run and explicit retry results are retained in `~/corpora/results/lossless-final-984eaf3/`.
The twelve remaining pairs are listed in its `consolidated/render-exceptions.json`; known source defects and
unresolved reflow are distinguished. Malformed extra-body/final-section notices cover two readable
corpus inputs: six save/edit/draft attempts and 14 package/SDK states pass, as do all 29 Ecrivaine unit
tests. All emitted package content matches the frozen saves except generated metadata timestamps,
so the existing layout evidence applies. The scope scan retains 35 failed archive/main-XML reads
separately. Evidence: `~/corpora/results/ecrivaine-structure-notices-2026-10-09/`.

The issue #1 regressions pass all **4 browser checks**, each of which fails on the previous
commit. The full save/reopen recheck completes **2,880/2,902 inputs**, with the same **22 failures**
reproduced on the previous commit. Nine independently authored page-field documents save with
byte-identical package contents under both writers and add no package/SDK diagnostics.
LibreOffice opens **8 original/saved pairs** with matching page counts; the remaining original
fails to load. The generated carriage-return fixtures also pass package/SDK checks and retain
their separate lines in LibreOffice. Reports: `~/corpora/results/issue-1-2026-10-11/`.

Review of that fix adds **5 edge-case regression groups**, all failing on the first fix and
passing after correction: protected fill destinations, split CRLF pairs through formatting and
save/reopen, forced-line pagination, mark placement in rotated, RTL and vertical text at
three zoom levels, and omission of all formatting symbols from print/PDF rendering.
The original four groups still pass, as do **79 Matheuse and 29 Ecrivaine unit
tests**. Three split-CRLF packages pass package/SDK validation without diagnostics and retain
their line count in LibreOffice. Reports: `~/corpora/results/issue-1-review-2026-10-11/`.

The convention scan has 193 signatures, unchanged from the preceding full writer. The earlier
scoped MC rules and Strict list-alignment entries retain their reviewed source properties.
No new package/SDK diagnostic accompanies them. Office batches 1–2, including the
comment repair and all namespace follow-ups, are accepted. Batch 5's section/numbering samples
are also accepted.
The Office checklist is `~/Downloads/lossless-check/office-batch-2/CHECKLIST.md`; file 12 records its
original repair and links the accepted corrected save. Both samples in
`~/Downloads/lossless-check/office-batch-2-comment-fix/` pass Word, package/SDK checks and LibreOffice,
with unchanged page counts.

### Content controls: edit behavior

| Edit | Save behavior |
|---|---|
| Text inside plain/rich controls | Keep the wrapper; update scalar custom-XML XPath targets, including attributes |
| Date/list value command or checkbox click | Keep typed properties and update the bound value |
| Ordinary typing inside a date/list/checkbox control | Convert to unbound text and report it |
| Formatting | Keep control properties; content locks prevent forbidden changes |
| Enter in a whole-paragraph control | Promote boundaries to a block control with the same identity |
| Enter in a partial-paragraph control | Split into unbound controls and report the conversion |
| Join across a multi-paragraph control boundary | Convert the affected wrapper to ordinary content and report it |
| Delete one boundary | Balance the remaining content and report conversion; deletion/content locks roll back the whole forbidden transaction |
| Copy/paste | Remap complete control/bookmark identities; carry XML stores and glossary dependencies; foreign copies receive separate store identities |
| Undo/redo and drafts | Restore boundaries, binding updates and pending conversions |

The context menu offers **Content Control Value…** for dates, lists and checkboxes. Complex XML/Flat
OPC bindings remain intact while untouched; editing them unbinds and reports the conversion. Copying
a built-in document-property binding into another document also unbinds it and reports it.

### Text properties, ranges and backgrounds: edit behavior

| Edit | Save behavior |
|---|---|
| Type or change an unrelated property | Keep the original property XML, compatibility alternatives and dependencies |
| Change font, colour, shadow, spacing or another modeled property | Replace that property; retain unrelated extensions and attributes in schema order |
| Clear all formatting | Remove direct formatting, including retained extensions |
| Copy a bookmark/permission range | Remap both endpoints; incomplete copies omit the markers |
| Delete one range endpoint | Omit the incomplete pair on save and report it |
| Set a page colour | Replace the original background; other edits keep its fill/picture |
| Change a page size, margin, numbering or grid setting | Change its owned attributes; keep other section settings, revision metadata and compatibility branches |
| Change page-number format or starting number | Keep the existing chapter heading level and separator unless chapter numbering is explicitly disabled |
| Insert a section, apply page setup/columns forward, build a multi-column index or merge letters | Allocate independent retained revision IDs for new sections; keep identities when reusing an existing boundary |
| Edit text or list alignment/spacing around a picture bullet | Keep its picture definition and image bytes; undo and drafts restore its metadata |
| Replace a picture bullet with another symbol/number format | Convert that level to the chosen symbol, remove its picture reference and report the conversion; unused original definitions remain available |
| Insert a file containing picture bullets | Carry its referenced definitions and image parts, allocating identities that do not collide with the current document |
| Change a list's alignment, starting value or number format | Replace its owned property; retain IDs, template metadata, legacy attributes and other compatibility alternatives. Unknown/custom formats remain selected when only alignment is changed |
| Create a new list from an existing definition or insert another document | Allocate independent list identities and remap retained references; keep every original definition's identity |
| Edit a list delegated to a numbering style | Give that list direct levels, retain the former style and its other users, and report the conversion; undo restores the link |
| Copy malformed list-symbol highlighting | Remove redundant `none` highlighting or convert its color to supported shading, with a compatibility notice; untouched originals remain intact |
| Edit body text or save without changing section settings | Keep original section XML, including absent defaults and header/footer references |
| Undo/redo or drafts | Restore property fragments and range metadata |

### Opaque objects and compatibility alternatives: edit behavior

| Edit | Save behavior |
|---|---|
| Edit surrounding text | Keep the original DrawingML/VML, OLE, ActiveX, SmartArt or altChunk fragment and dependencies |
| Move, resize, rotate or flip one object | Update its anchor/transform in both compatibility branches |
| Resize a SmartArt preview group | Keep diagram data with scaled children or a changed outer extent |
| Edit or ungroup converted contents; replace a picture | Write the model representation and report conversion |
| Change picture crop, border, brightness/contrast, grayscale or opacity | Replace that property; retain recolouring, effects, blip extensions and sibling properties |
| Change shape fill, line, shadow or text-box layout | Merge the owning properties in schema order; a shadow edit keeps glow, reflection, soft edges and 3-D siblings |
| Change DrawingML wrapping or anchor reference mode | Merge anchor properties; changing inline/floating status retains the original graphic and identities |
| Change a complex VML fill, VML geometry/text layout, or a shadow in an effect graph | Convert the affected frame and report it |
| Duplicate or paste into another document | Carry preview and payload bytes; remap definitions and references together, including SmartArt's implicit drawing link |
| Delete all or part of a multi-member compatibility wrapper | Never resurrect deleted content; convert surviving members and report it when the wrapper is incomplete |
| Undo and drafts | Restore frame/dependency identities |

Embedded documents and objects without a usable preview display a placeholder. Run, inline, block,
row and cell alternatives are captured before normalization and emitted once. Empty table choices
remain attached to adjacent members, retaining their fallback without adding a visible row or cell.
Moving one of several drawings owned by a single wrapper currently converts that wrapper.

Original text/picture watermarks stay in their own header/footer locations with their XML and
relationships. Editing the watermark replaces those originals; removing it omits them. Undo restores
the originals, and drafts retain the current state. Eight focused cases cover drawing properties,
row/cell alternatives (including empty choices) and watermarks: **62 checks and 48 emitted states**
pass, with no added package/SDK diagnostics. The two watermark inputs retain all six original VML
watermark elements byte for byte before editing. The relevant Office batch-2 samples are accepted.

### Package and file identity: edit behavior

| Edit | Save behavior |
|---|---|
| Change settings, custom properties or font use | Merge owned properties; retain unrelated settings, property types/pids and referenced embedding metadata |
| Save as a macro-free variant | Report removed VBA before download; cancellation keeps the original filename/type and acknowledgements |
| Recover an encrypted or macro-enabled draft | Keep its main-part variant, password and pending loss entries |
| Save an invalid original dependency or orphaned part | Report unavailable/unreferenced content instead of silently claiming it was retained |

## Compatibility testing

The reproducible OOXML run uses the pinned LibreOffice, Apache POI and Open XML SDK inputs from
`tools/corpora.sh`. Ecrivaine runs in Chromium through CDP; independent package checks, SDK diagnostics
and LibreOffice compare the saved file with its original. That Chromium is for the pages in a browser. A VS Code extension change does not start it; the user checks the extension in VS Code. On WSL2, do not look for Chromium on Windows or on the WSL2 host, and do not run commands outside the workspace (AGENTS.md). Page counts, Word's saved page-break markers
and reference PDFs help diagnose layout. Missing fonts, tight wrapping around irregular pictures,
vertical text and complex data bindings remain known differences. RC4-encrypted legacy binary files
are outside the supported OOXML encryption path.

Markdown checks remain separate: **24/24 unchanged documents save byte for byte**, edits stay inside
the edited block in **24/24**, and CommonMark/GFM specifications pass **652/652 and 22/22** cases.
Commands and comparison policies are in [testing.md](testing.md); earlier broad-corpus layout and
compatibility results are archived with the run history rather than mixed into current save numbers.

## Source layout

| File | What it does |
| --- | --- |
| `index.html` | Window chrome and page CSS, script order |
| `../common/core.js`, `../common/zip.js` | Utilities, colour maths, storage, ZIP reader/writer, PDF writer, downloads |
| `../common/xml.js`, `../common/opc.js`, `../common/opc-order.js` | Shared OOXML package graph, identity reservations, AlternateContent capture and schema-order merges |
| `../common/sha.js`, `../common/crypto.js` | Password-protected documents: compound-file reader/writer, AES, ECMA-376 agile and standard encryption |
| `js/dmodel.js`, `js/ops.js` | Document model (paragraphs, runs, tables, sections, styles, numbering), undo history, editing operations |
| `js/dom.js` | Word object model on the open document ([ecrivaine_dom_api_specifications.md](ecrivaine_dom_api_specifications.md)) |
| `js/dom-ir.js` | Handle protocol over that object model ([ecrivaine_dom_ir_specifications.md](ecrivaine_dom_ir_specifications.md)) |
| `js/vb-host.js` | VB6 host. Document properties, `wd*` constants and the inventory's gap-only objects go through the IR. `evalVB6` is the Immediate page global on the shared session. `runVB6` compiles one standard module onto a second VM on that session. Document macros use a third VM that stays loaded. The Visual Basic IDE uses `debugBegin` / `debugEnd` for one more VM that stays paused until Stop ([ecrivaine_vbs_support_schedule.md](ecrivaine_vbs_support_schedule.md)). View ▸ Run Form runs a `.frm` on a Melina Office dialog ([ecrivaine_vbs_form.md](ecrivaine_vbs_form.md)). The engine script is gitignored; `make` writes `vendor/vb6/vb6-engine.js`, and the host loads it when a session starts |
| `../common/vb-ide.js`, `../common/vb-complete.js` | View ▸ Melina Office Visual Basic: the editor, conditional breakpoints, watches, Output, Expression, call stack and object values. An edit is checked after a short pause: a wavy underline, and an Errors tab after Output (line, column, number, source, message); a click on a Module1 row moves to that line. Start Debug and Start Run show that list and do not run when the compile fails. Completion lists the names the caret's procedure can see, and the procedures and modules this virtual machine actually runs. Every completion name opens the same box, with its own icon. The reference note is shown when there is one. A variable, constant or member shows Type when it has one. A procedure shows Signature only when it has arguments, and Returns only when it has a return type. A member that needs a host ability this session does not have is gray, and the box says why. Keywords are listed anywhere except after a dot. A string, a comment or a date literal does not |
| `js/macros.js`, `../common/macros.js` | Document macros and the All templates and documents list: the Tools ▸ Macros window, the custom XML part, and loading those modules onto the macro VM |
| `js/dom-console.js`, `../common/vendor/xterm/`, `../common/vendor/codemirror/` | Help ▸ Ecrivaine DOM Console: a terminal for Visual Basic 6 (`evalVB6`) or JavaScript, and Run Script (`runVB6`). From Macros, the same window loads and saves one document macro. DOM API References opens the task pane |
| `../common/vb-dom-ref.js`, `js/vb-dom-ref.js` | Visual Basic DOM API Reference task pane. The member list follows [ecrivaine_dom_api_specifications.md](ecrivaine_dom_api_specifications.md); a change to a member updates both |
| `js/render.js`, `js/layout.js` | Paragraph/run rendering, pagination, headers/footers, footnotes, columns, floats |
| `js/portions.js` | How tall each line is, from what it holds (after Apache OpenOffice Writer's line portions) |
| `js/editor.js` | contentEditable bridge: selection mapping, typing, IME, clipboard, keyboard |
| `js/preserve.js` | Package ownership/merges; controls and XML bindings; property/range preservation; opaque frames, compatibility alternatives and clipboard dependencies |
| `js/preserve-drawing.js` | Property-level DrawingML/VML merges, including compatibility choices, picture effects and anchor edits |
| `js/docx-read.js`, `js/docx-write.js`, `../common/dml.js` | `.docx` import/export, DrawingML/VML shapes, charts, themes |
| `js/rtf.js`, `js/htmlio.js` | RTF and HTML import/export |
| `js/markdown.js`, `js/mdio.js` | Markdown reader (CommonMark + GitHub extensions, with source lines; also loads in Node) and Markdown/`.zip` import and export |
| `js/fields.js` | Field engine (PAGE, DATE, REF, SEQ, TOC, INDEX, IF, formulas, …), footnotes and endnotes |
| `js/review.js` | Track changes, comments, balloons, Reviewing pane, compare documents |
| `js/tables.js` | Tables and Borders toolbar, table styles, draw/erase, merge/split, sort, formulas |
| `js/drawing.js`, `js/drawtool.js`, `../common/geometry.js`, `../common/charts.js`, `../common/metafile.js`, `../common/clipart.js` | Pictures, AutoShapes, LettersArt (WordArt), diagrams, charts, WMF/EMF, clip art |
| `js/autocorrect.js`, `../common/spell.js`, `js/spell.js` | AutoCorrect, AutoFormat as you type, AutoText; the shared spelling engine and Ecrivaine's Spelling and Grammar (offline) |
| `js/omml.js` | Office Math (OMML) to MathML, and OMML to LaTeX for the equation editor |
| `js/latexprev.js` | LaTeX preview, and the OMML written when an equation is updated |
| `../common/dict/` | Spelling word lists, thesaurus, licences (shared with Matheuse) |
| `js/outline.js`, `js/rulers.js`, `js/panes.js` | Outline view, rulers, task panes (Research/Thesaurus, Clip Art, Styles…; while open, Picture or Drawing Tools for a selected object, docs/suite.md), Document Map, thumbnails, Print Preview, Reading Layout |
| `js/mailmerge.js`, `js/templates.js` | Mail Merge wizard and toolbar, document templates |
| `../common/ui.js`, `../common/search.js`, `../common/icons.js`, `../common/minibar.js`, `../common/looks.css`, `../common/ui.css` | Menus, toolbars, dialogs, colour pickers, Search commands, original 16×16 icon set, the formatting bar after a text selection, the looks (shared) |
| `js/commands.js`, `js/dialogs*.js`, `js/find.js`, `js/app.js` | Command registry, all dialogs, Find and Replace, application controller |
| `js/ribbon.js`, `../common/ribbon.js` | The Paper 2016 look's ribbon: Word's tabs and groups over Ecrivaine's commands ([suite.md](suite.md#the-ribbon)) |
| `../common/scout/` | Nikki, the assistant: Help ▸ Show the Assistant ([suite.md](suite.md#the-assistant-nikki)) |

Ecrivaine is an independent program. Microsoft, Word and Office are trademarks of Microsoft Corporation;
Ecrivaine is not affiliated with Microsoft. More in the README (Legal).
