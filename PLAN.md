# DuckWriter plan

A free word processor for Windows that opens and saves Word documents (.docx),
for you and your friends. Built like Duckument: Tauri (a small Windows app
using the built-in Edge engine), with the editing done in TypeScript.

The rule that runs through every stage: **a document someone sends you must
survive being opened and saved in DuckWriter.** Whatever DuckWriter can't edit
yet (a table, a chart, a content control, a comment) is kept exactly as it was
and written back unchanged.

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

## Next

- Tables: insert and edit rows, columns, merging, borders, widths.
- Links: insert and edit (Ctrl+K).
- Text wrapping on both sides of a picture in the middle of the page.
- Find and replace; spell check (Windows' own dictionaries).
- Print and save as PDF.
- Several documents in tabs, recent files.
- Auto-update and an installer, published on GitHub like Duckument.
