# DuckWriter plan

A free word processor for Windows that opens and saves Word documents (.docx),
for you and your friends. Built like Duckument: Tauri (a small Windows app
using the built-in Edge engine), with the editing done in TypeScript.

The rule that runs through every stage: **a document someone sends you must
survive being opened and saved in DuckWriter.** Whatever DuckWriter can't edit
yet (a table, a picture, a header, a field) is kept exactly as it was and
written back unchanged.

## Stage 1: Foundation (done)

- Page view: the page at its real size and margins from the document.
- Typing, selecting, copy and paste, undo and redo.
- Styles from the document (Normal, Title, Heading 1–3, Quote…), shown the way
  the document defines them, with a style picker.
- Font, size, bigger/smaller, bold, italic, underline, strike-through,
  superscript, subscript, text colour, highlight, clear formatting.
- Alignment, indents, line spacing, space before/after paragraphs.
- Bulleted and numbered lists with levels (Tab / Shift+Tab), numbered the way
  Word numbers them.
- Page breaks.
- New, Open, Save, Save As for .docx; asks before losing unsaved changes;
  opens a document double-clicked in Explorer.
- Word count and zoom.
- Kept but not yet editable: tables, pictures, fields, content controls,
  bookmarks, comments, headers and footers, footnotes.

## Stage 2: Tables and pictures

- Insert and edit tables: rows, columns, merging, borders, widths.
- Show and keep pictures; insert, resize and remove them.
- Links: insert and edit (Ctrl+K).

## Stage 3: Real pages

- The text split into pages as it will print.
- Page size, orientation and margins.
- Headers, footers and page numbers.

## Stage 4: Writing tools

- Find and replace.
- Spell check (Windows' own dictionaries).
- Word count details, zoom to page width.

## Stage 5: Print and PDF

- Print, and save as PDF.

## Stage 6: Everyday comfort and release

- Several documents in tabs, recent files.
- Auto-update and an installer, published on GitHub like Duckument.
