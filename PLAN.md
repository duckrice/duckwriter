# DuckWriter plan

A free word processor for Windows that opens and saves Word documents (.docx),
for you and your friends. Built like Duckument: Tauri (a small Windows app
using the built-in Edge engine), with the editing done in TypeScript.

The rule that runs through every stage: **a document someone sends you must
survive being opened and saved in DuckWriter.** Whatever DuckWriter can't edit
yet (a chart, a content control, a comment) is kept exactly as it was and
written back unchanged.

## Done

### Foundation
- Page view at the document's own page size and margins; typing, selecting,
  copy and paste, undo and redo.
- Styles from the document (Normal, Title, Heading 1–3, Quote…) with a style
  picker; font, size, bold, italic, underline, strike-through, superscript,
  subscript, colour, highlight, clear formatting.
- Alignment, indents, line spacing, space before/after paragraphs.
- Bulleted and numbered lists with levels, numbered the way Word does.
- New, Open, Save, Save As for .docx; asks before losing unsaved changes.
- Word count and zoom.

### Scientific writing
- **Equations**: typed in a LaTeX-like shorthand (`x^2`, `\frac{a}{b}`,
  `\sum_{i=1}^{n}`, `\alpha`…) with a live preview and buttons for the common
  shapes; Alt+= as in Word. Saved as real Word equations; Word's equations are
  shown and can be edited.
- **Symbols**: Greek letters, maths signs, arrows and units.
- **Real pages**: the text is laid out on pages, with widow and orphan
  control, headings kept with their text, and page breaks.
- **Headers and footers** with page numbers and page counts; a different
  first page. Double-click to edit, as in Word.
- **Footnotes** at the bottom of their page, numbered in order (Ctrl+Alt+F).
- **Captions** (Figure 1, Table 1, Equation 1) that number themselves,
  **cross-references** ("see Figure 2") that stay right, **numbered headings**
  (1, 1.1, 1.1.1) and a **table of contents** that updates itself.
- **Pictures**: in line with the text or floating (square, top and bottom, in
  front, behind); insert, paste or drop; resize and move.
- **Snap to grid**: pictures snap to a half-centimetre grid and to the margins
  when moved or resized (Alt to place freely); a **line grid** puts every line
  of text on the same evenly spaced lines on every page (Word's document grid).

### Everyday work
- **Keyboard shortcuts** for everything, Word's where possible, working on
  Norwegian and other keyboards; F1 lists and searches them.
- **Auto-save** to the document's file, and **recovery copies** of unsaved
  work, all on this computer; work left by a crash is offered back.
- **Tables**: Word tables open as real tables (merged cells, widths,
  shading, header rows, table styles); insert, add and delete rows and
  columns, merge and split cells, lines and shading, Tab between cells,
  resize columns by dragging; long tables continue on the next page.
- **Print** (Ctrl+P) and **Save as PDF**, page for page as on screen.
- **Find and replace** (Ctrl+F, Ctrl+H, F3): every match highlighted, match
  case, whole words, replace one or all.
- **Spelling** with Windows' own spell checker, in every language in Windows'
  language settings at once (Norwegian and English together): a red line
  under unknown words, suggestions on right-click or F7, Ignore all, Add to
  dictionary, Ctrl+F7 to switch it off.
- **Tracked changes**: Word's tracked changes are shown (new text underlined,
  deleted text struck through, a colour per reviewer) and kept when saving;
  accept or reject one at a time or all; **Track changes** (Ctrl+Shift+E)
  records your own edits, under your name, for Word to read.
- **Comments** beside the page, level with their text: add (Ctrl+Alt+M),
  edit, reply, resolve, delete; Word's replies and resolved comments are
  read and written.
- **Links** (Ctrl+K) to web pages, e-mail addresses and headings; Ctrl+click
  follows them, in the browser.
- A long table's **header rows are repeated** at the top of each page.
- Formatting changes are tracked too, and can be rejected back to what they were.
- **Recent documents** (Alt+F) and opening a document by dropping it on the window.
- **Page setup** (Alt+P): paper size, upright or sideways, margins in centimetres.
- **Headings beside the page** (Alt+Shift+N), to go to with a click or the keyboard.
- **Copy and paste** keeps all formatting inside DuckWriter; from Word and the
  web, headings, lists, links, bold, italic and superscripts come through.
  Ctrl+Alt+V pastes plain text.
- **Format painter** (the brush, or Ctrl+Shift+C and Ctrl+Shift+V).
- **AutoFormat as you type**: addresses become links; "1. " or "- " starts a list.
- **Word count** of the selection in the status bar; Ctrl+Shift+G for characters
  and paragraphs.
- **Pictures in headers and footers** (a logo in a letterhead) are shown and kept.

### Round 3
- **Citations and a reference list**: sources as Word keeps them; APA,
  Harvard, Chicago, IEEE and Vancouver; BibTeX import and DOI lookup.
- **Writing together** (see COLLABORATION.md): on the local network, through
  a relay server, or through a shared folder; encrypted end to end.
- **Markup views**: all markup, the final text, or the original (Alt+Shift+M).
- **Go to page** (Ctrl+G), **page numbers** from the header button, **change
  case** (Alt+Shift+K).
- **Line numbers** in the margin (Alt+Shift+L), for manuscripts.
- **Lists of figures and tables**, and **numbered equations** (Alt+Shift+E)
  that can be referred to as (1).
- **Changing a style** for all its paragraphs (Alt+Shift+Y), written for Word.
- **Earlier versions** kept on this computer at each save (Alt+Shift+V).
- **Templates**: scientific article, report, letter, thesis chapter (Alt+Shift+T).
- **Each document in its own window**; pictures and links in headers.
- **Installer and auto-update**, ready to publish.

### Round 4
- **Ribbon tabs**: Home, Insert, References, Review, Layout (Alt+1 to Alt+5).
- **Maths shorthand** in the text: \alpha → α, -> → →, <= → ≤.
- **Focus mode** (F11) and a **dark page** for night writing (Alt+Shift+D).
- **Tables**: sort rows (Alt+Shift+O), convert text to a table.
- **Markdown, HTML and text**: open them as Word documents, save copies as them.
- **Endnotes** (Ctrl+Alt+D), numbered i, ii, iii.
- **Track changes** covers paragraph formatting and table rows too.

### Round 5
- **Compare** two documents, or this one with an earlier version: the
  differences as tracked changes.
- **Check before sending**: comments, tracked changes, hidden author names,
  pictures without a description.
- **Picture descriptions** (alt text), read aloud by screen readers.
- **Own templates**: save a document as a template, start new ones from it.
- **AutoCorrect** for common typos, and **fit zoom** (page width, whole page).
- **Long documents**: 250 pages open in under 2 s, typing stays at about
  20 ms a key (it used to hang); scripts/perf.mjs measures it.

### Round 6
- **LaTeX**: save a copy as .tex, with the sources as .bib, for journals
  that want LaTeX.
- **RIS** sources (EndNote, Mendeley, Zotero), and the sources out as BibTeX.
- **Moving sections** in the headings pane: drag a heading, or
  Alt+Shift+Up/Down; promote and demote with Alt+Shift+Left/Right.
- **Footnotes to endnotes** and back, all at once.
- **A word goal**, with words written today, in the status bar.
- **Non-breaking space**
  (Ctrl+Shift+Space) and **non-breaking hyphen** (Ctrl+Shift+-); paste as
  plain text was already there (Ctrl+Alt+V).

### Round 7
- **Document properties**: title, subject, author, keywords; the save date
  and "last saved by" kept up to date; Check before sending lists the names
  in them.
- **CSL-JSON** sources (Zotero, Mendeley), and **searching Zotero** while it
  runs, to add sources from it.
- **Readability** in the word count: sentences, words per sentence, and LIX
  (which works for Norwegian and English alike).
- **Read aloud** with Windows' voices, from the cursor or the selection,
  for proofreading. (Automatic hyphenation was on the list, but the browser
  engine has no hyphenation dictionaries on Windows.)

### Round 8: sturdiness
- **A package check** on saved files (it found copied pictures sharing an id,
  now fixed): every relationship points to a part
  that is there, every part has a content type, the XML is well formed,
  and ids (comments, changes, bookmarks, pictures) are unique, so Word never
  calls a saved file damaged. Run over every sample and over edited files.
- **Speed with many changes and comments**: 450 comments and 2,700 tracked
  changes in 124 pages, 22 ms a key (was 205).
- **Every button has a name** for screen readers (checked in the app test).

## Next

- Text wrapping on both sides of a picture (the browser engine can't do this).
- Publishing: the GitHub repository, the first release.
