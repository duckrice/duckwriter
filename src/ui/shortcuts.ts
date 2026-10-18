import { showInspection } from "./inspect";
import { insertEndnote } from "./endnotes";
import { moveSection, shiftSection } from "../editor/outline";
import { modifyStyle } from "./styles";
import { insertBibliography, insertCitation, manageSources } from "./sources";
import { copyFormatting, pasteFormatting, pastePlain } from "./painter";
import { editLink } from "./links";
// Every keyboard shortcut in one list: the keys handled, and the help window
// (F1) that lists them, both come from here, so they always agree.

import { redo, undo } from "prosemirror-history";
import { selectAll } from "prosemirror-commands";
import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { h } from "../dom";
import type { DocContext } from "../docx/model";
import {
  clearFormatting,
  insertPageBreak,
  setAlign,
  setLineSpacing,
  setParagraphs,
  setStyle,
  stepIndent,
  stepSize,
  stepSizeBy,
  toggleFormat,
  toggleVertAlign,
  formattingAt,
  changeCase,
} from "../editor/commands";
import { changeListLevel, toggleList } from "../editor/lists";
import { matches } from "../keys";
import { insertEquation, insertNumberedEquation } from "./equation";
import { choosePicture } from "./pictures";
import { captionDialog, insertToc, toggleHeadingNumbers } from "./references";
import { TABLE_COMMANDS } from "./tables";

export interface ShortcutHost {
  readonly view: EditorView;
  readonly body: EditorView;
  readonly ctx: DocContext;
  newDocument(): void;
  open(): void;
  save(saveAs?: boolean): void;
  print(): void;
  savePdf(): void;
  insertFootnote(): void;
  editHeader(): void;
  zoomBy(direction: 1 | -1 | 0): void;
  setLineGrid(on: boolean): void;
  toggleSnap(): void;
  toggleAutoSave(): void;
  find(replace: boolean): void;
  findNext(direction: 1 | -1): void;
  nextSpellingMistake(): void;
  toggleSpelling(): void;
  newComment(): void;
  cycleMarkup(): void;
  cycleZoomFit(): void;
  compareWithFile(): void;
  exportAs(format: "markdown" | "html" | "text" | "latex"): void;
  sortRows(): void;
  toggleFocus(): void;
  toggleDarkPage(): void;
  versions(): void;
  styleChanged(id: string): void;
  cycleLineNumbers(): void;
  goToPage(): void;
  share(): void;
  pageSetup(): void;
  wordCount(): void;
  toggleNavigation(): void;
  toggleTracking(): void;
  stepChange(direction: 1 | -1): void;
  decideChange(accept: boolean): void;
  decideAllChanges(accept: boolean): void;
  /** The toolbar's boxes and menus. */
  focusBox(box: "font" | "size" | "style"): void;
  openMenu(menu: "symbol" | "references" | "layout" | "textColor" | "highlight" | "table" | "open" | "newDoc"): void;
  /** Shows a tab of the toolbar, from 1. */
  selectTab(n: number): void;
  toggleTableLines(): void;
  applyTextColor(): void;
  applyHighlight(): void;
  showShortcuts(): void;
}

export interface Shortcut {
  group: string;
  label: string;
  /** One or more ways to press it, as shown: "Ctrl+B". */
  keys: string[];
  /** Left out for keys handled elsewhere (typing, Windows itself); they are only listed. */
  run?: (host: ShortcutHost) => void;
  /** Also works while typing in a box, such as the font size. */
  everywhere?: boolean;
}

const cmd = (make: (ctx: DocContext) => Command) => (host: ShortcutHost) => {
  const { view } = host;
  make(host.ctx)(view.state, view.dispatch, view);
  view.focus();
};
const plain = (command: Command) => cmd(() => command);
const typeChar = (char: string): Command => (state, dispatch) => {
  dispatch?.(state.tr.insertText(char).scrollIntoView());
  return true;
};

/** Heading 1–3, if the document has that style. */
const heading = (level: number) =>
  cmd((ctx) => (state, dispatch) => (ctx.styles.byId.has(`Heading${level}`) ? setStyle(ctx, `Heading${level}`)(state, dispatch) : false));

export const SHORTCUTS: Shortcut[] = [
  // ---- documents
  { group: "Documents", label: "New document", keys: ["Ctrl+N"], run: (x) => x.newDocument(), everywhere: true },
  { group: "Documents", label: "Open", keys: ["Ctrl+O"], run: (x) => x.open(), everywhere: true },
  { group: "Documents", label: "Recent documents", keys: ["Alt+F"], run: (x) => x.openMenu("open"), everywhere: true },
  ...["Home", "Insert", "References", "Review", "Layout"].map((tab, i) => ({ group: "Documents", label: `The toolbar's ${tab} tab`, keys: [`Alt+${i + 1}`], run: (x: ShortcutHost) => x.selectTab(i + 1), everywhere: true })),
  { group: "Documents", label: "New from a template (article, report, letter, thesis chapter)", keys: ["Alt+Shift+T"], run: (x) => x.openMenu("newDoc"), everywhere: true },
  { group: "Documents", label: "Page setup: paper size, orientation, margins", keys: ["Alt+P"], run: (x) => x.pageSetup() },
  { group: "Documents", label: "Word count (of the selection too)", keys: ["Ctrl+Shift+G"], run: (x) => x.wordCount(), everywhere: true },
  { group: "Documents", label: "Go to page", keys: ["Ctrl+G"], run: (x) => x.goToPage() },
  { group: "Documents", label: "Save a copy as a web page (HTML)", keys: ["Ctrl+Alt+Shift+W"], run: (x) => x.exportAs("html"), everywhere: true },
  { group: "Documents", label: "Save a copy as Markdown", keys: ["Ctrl+Alt+Shift+K"], run: (x) => x.exportAs("markdown"), everywhere: true },
  { group: "Documents", label: "Save a copy as LaTeX (.tex, with the sources as .bib)", keys: ["Ctrl+Alt+Shift+X"], run: (x) => x.exportAs("latex"), everywhere: true },
  { group: "Documents", label: "Focus mode: only the page, full screen (Esc to leave)", keys: ["F11"], run: (x) => x.toggleFocus(), everywhere: true },
  { group: "Documents", label: "Dark page, for writing at night", keys: ["Alt+Shift+D"], run: (x) => x.toggleDarkPage(), everywhere: true },
  { group: "Documents", label: "Earlier versions of this document (kept on this computer)", keys: ["Alt+Shift+V"], run: (x) => x.versions(), everywhere: true },
  { group: "Documents", label: "Line numbers: through the document, from 1 on each page, or none", keys: ["Alt+Shift+L"], run: (x) => x.cycleLineNumbers() },
  { group: "Documents", label: "Headings beside the page (the navigation pane); arrows and Enter go to one", keys: ["Alt+Shift+N"], run: (x) => x.toggleNavigation(), everywhere: true },
  { group: "Documents", label: "Open a document by dropping it on the window", keys: ["Drag and drop"] },
  { group: "Documents", label: "Save", keys: ["Ctrl+S"], run: (x) => x.save(), everywhere: true },
  { group: "Documents", label: "Save as", keys: ["F12"], run: (x) => x.save(true), everywhere: true },
  { group: "Documents", label: "Print", keys: ["Ctrl+P"], run: (x) => x.print(), everywhere: true },
  { group: "Documents", label: "Save as PDF", keys: ["Ctrl+Alt+Shift+P"], run: (x) => void x.savePdf(), everywhere: true },
  { group: "Documents", label: "Auto-save, on or off", keys: ["Ctrl+Alt+Shift+S"], run: (x) => x.toggleAutoSave(), everywhere: true },
  { group: "Documents", label: "Keyboard shortcuts (this list)", keys: ["F1"], run: (x) => x.showShortcuts(), everywhere: true },

  // ---- editing
  { group: "Editing", label: "Undo", keys: ["Ctrl+Z"], run: plain(undo) },
  { group: "Editing", label: "Redo", keys: ["Ctrl+Y", "Ctrl+Shift+Z"], run: plain(redo) },
  { group: "Editing", label: "Select everything", keys: ["Ctrl+A"], run: plain(selectAll) },
  { group: "Editing", label: "Cut, copy, paste", keys: ["Ctrl+X", "Ctrl+C", "Ctrl+V"] },
  {
    group: "Editing",
    label: "Paste as plain text (without formatting)",
    keys: ["Ctrl+Alt+V"],
    run: (x) =>
      void pastePlain(x.view).then((ok) => {
        if (!ok) x.view.dom.dispatchEvent(new CustomEvent("duckwriter-toast", { bubbles: true, detail: "Couldn't read the clipboard; use Ctrl+V" }));
      }),
  },
  { group: "Text", label: "Copy the formatting (format painter)", keys: ["Ctrl+Shift+C"], run: (x) => copyFormatting(x.view) },
  { group: "Text", label: "Change case: lower, UPPER, Title", keys: ["Alt+Shift+K"], run: plain(changeCase()) },
  { group: "Text", label: "Put the copied formatting on the selection", keys: ["Ctrl+Shift+V"], run: (x) => pasteFormatting(x.view) },
  { group: "Editing", label: "Find", keys: ["Ctrl+F"], run: (x) => x.find(false), everywhere: true },
  { group: "Editing", label: "Find and replace", keys: ["Ctrl+H"], run: (x) => x.find(true), everywhere: true },
  { group: "Editing", label: "Next match", keys: ["F3"], run: (x) => x.findNext(1), everywhere: true },
  { group: "Editing", label: "Previous match", keys: ["Shift+F3"], run: (x) => x.findNext(-1), everywhere: true },
  { group: "Editing", label: "Next spelling mistake, with suggestions", keys: ["F7", "Alt+F7"], run: (x) => x.nextSpellingMistake() },
  { group: "Editing", label: "Spell checking on or off", keys: ["Ctrl+F7"], run: (x) => x.toggleSpelling(), everywhere: true },
  { group: "Editing", label: "New line in the same paragraph", keys: ["Shift+Enter"] },
  { group: "Editing", label: "Page break", keys: ["Ctrl+Enter"], run: plain(insertPageBreak()) },
  { group: "Headings and outline", label: "Move this section (the heading and what is under it) up", keys: ["Alt+Shift+ArrowUp"], run: cmd((ctx) => moveSection(ctx, -1)) },
  { group: "Headings and outline", label: "Move this section down", keys: ["Alt+Shift+ArrowDown"], run: cmd((ctx) => moveSection(ctx, 1)) },
  { group: "Headings and outline", label: "Take this section's headings a level up (Heading 2 → 1)", keys: ["Alt+Shift+ArrowLeft"], run: cmd((ctx) => shiftSection(ctx, -1)) },
  { group: "Headings and outline", label: "Take this section's headings a level down (Heading 1 → 2)", keys: ["Alt+Shift+ArrowRight"], run: cmd((ctx) => shiftSection(ctx, 1)) },
  { group: "Editing", label: "Non-breaking space (keeps “10 kg” on one line)", keys: ["Ctrl+Shift+Space"], run: plain(typeChar(String.fromCharCode(0xa0))) },
  { group: "Editing", label: "Non-breaking hyphen", keys: ["Ctrl+Shift+-", "Ctrl+_"], run: plain(typeChar(String.fromCharCode(0x2011))) },

  // ---- text
  { group: "Text", label: "Bold", keys: ["Ctrl+B"], run: cmd((c) => toggleFormat(c, "bold")) },
  { group: "Text", label: "Italic", keys: ["Ctrl+I"], run: cmd((c) => toggleFormat(c, "italic")) },
  { group: "Text", label: "Underline", keys: ["Ctrl+U"], run: cmd((c) => toggleFormat(c, "underline")) },
  { group: "Text", label: "Strikethrough", keys: ["Ctrl+Shift+X"], run: cmd((c) => toggleFormat(c, "strike")) },
  { group: "Text", label: "Superscript (x²)", keys: ["Ctrl++"], run: plain(toggleVertAlign("superscript")) },
  { group: "Text", label: "Subscript (x₂)", keys: ["Ctrl+="], run: plain(toggleVertAlign("subscript")) },
  { group: "Text", label: "Font size up, to the next size", keys: ["Ctrl+>"], run: cmd((c) => stepSize(c, 1)) },
  { group: "Text", label: "Font size down, to the next size", keys: ["Ctrl+<"], run: cmd((c) => stepSize(c, -1)) },
  { group: "Text", label: "Font size up by one point", keys: ["Ctrl+]"], run: cmd((c) => stepSizeBy(c, 1)) },
  { group: "Text", label: "Font size down by one point", keys: ["Ctrl+["], run: cmd((c) => stepSizeBy(c, -1)) },
  { group: "Text", label: "Font size: type a size (Enter to use it)", keys: ["Ctrl+Shift+P"], run: (x) => x.focusBox("size") },
  { group: "Text", label: "Font: type a font name (Enter to use it)", keys: ["Ctrl+Shift+F"], run: (x) => x.focusBox("font") },
  { group: "Text", label: "Text colour (the last one used)", keys: ["Ctrl+Shift+O"], run: (x) => x.applyTextColor() },
  { group: "Text", label: "Choose a text colour", keys: ["Ctrl+Alt+Shift+O"], run: (x) => x.openMenu("textColor") },
  { group: "Text", label: "Highlight (the last colour used)", keys: ["Ctrl+Alt+H"], run: (x) => x.applyHighlight() },
  { group: "Text", label: "Choose a highlight colour", keys: ["Ctrl+Alt+Shift+H"], run: (x) => x.openMenu("highlight") },
  { group: "Text", label: "Clear formatting", keys: ["Ctrl+Space"], run: cmd((c) => clearFormatting(c)) },

  // ---- paragraphs
  { group: "Paragraphs", label: "Choose a style (type it, Enter)", keys: ["Ctrl+Shift+S"], run: (x) => x.focusBox("style") },
  { group: "Paragraphs", label: "Normal text", keys: ["Ctrl+Shift+N"], run: cmd((c) => setStyle(c, null)) },
  { group: "Paragraphs", label: "Change the paragraph's style (all paragraphs in it change)", keys: ["Alt+Shift+Y"], run: (x) => modifyStyle(x.view, x.ctx, (id) => x.styleChanged(id)) },
  { group: "Paragraphs", label: "Heading 1", keys: ["Ctrl+Alt+1"], run: heading(1) },
  { group: "Paragraphs", label: "Heading 2", keys: ["Ctrl+Alt+2"], run: heading(2) },
  { group: "Paragraphs", label: "Heading 3", keys: ["Ctrl+Alt+3"], run: heading(3) },
  { group: "Paragraphs", label: "Align left", keys: ["Ctrl+L"], run: plain(setAlign("left")) },
  { group: "Paragraphs", label: "Centre", keys: ["Ctrl+E"], run: plain(setAlign("center")) },
  { group: "Paragraphs", label: "Align right", keys: ["Ctrl+R"], run: plain(setAlign("right")) },
  { group: "Paragraphs", label: "Justify", keys: ["Ctrl+J"], run: plain(setAlign("justify")) },
  { group: "Paragraphs", label: "Bullets", keys: ["Ctrl+Shift+L", "Ctrl+Shift+8"], run: cmd((c) => toggleList(c, "bullet")) },
  { group: "Paragraphs", label: "Numbering", keys: ["Ctrl+Shift+7"], run: cmd((c) => toggleList(c, "number")) },
  {
    group: "Paragraphs",
    label: "Indent more (a list item: a level in)",
    keys: ["Ctrl+M"],
    run: cmd((c) => (state, dispatch, view) => changeListLevel(c, 1)(state, dispatch) || stepIndent(c, 1)(state, dispatch, view)),
  },
  {
    group: "Paragraphs",
    label: "Indent less (a list item: a level out)",
    keys: ["Ctrl+Shift+M"],
    run: cmd((c) => (state, dispatch, view) => changeListLevel(c, -1)(state, dispatch) || stepIndent(c, -1)(state, dispatch, view)),
  },
  { group: "Paragraphs", label: "List item a level in / out; in a table, the next / previous cell (in the last cell, Tab adds a row)", keys: ["Tab", "Shift+Tab"] },
  { group: "Paragraphs", label: "Single line spacing", keys: ["Ctrl+1"], run: plain(setLineSpacing(1)) },
  { group: "Paragraphs", label: "1.5 line spacing", keys: ["Ctrl+5"], run: plain(setLineSpacing(1.5)) },
  { group: "Paragraphs", label: "Double line spacing", keys: ["Ctrl+2"], run: plain(setLineSpacing(2)) },
  {
    group: "Paragraphs",
    label: "Space before the paragraph, on or off",
    keys: ["Ctrl+0"],
    run: cmd((c) => (state, dispatch) => {
      const before = formattingAt(c, state).para.spaceBefore ?? 0;
      return setParagraphs(() => ({ spaceBefore: before > 0 ? 0 : 12 }))(state, dispatch);
    }),
  },

  // ---- inserting
  { group: "Insert", label: "Equation", keys: ["Alt+="], run: (x) => insertEquation(x.view) },
  { group: "Insert", label: "Endnote (at the end of the document)", keys: ["Ctrl+Alt+D"], run: (x) => insertEndnote(x.body, x.ctx) },
  { group: "Insert", label: "Numbered equation, (1) on the right", keys: ["Alt+Shift+E"], run: (x) => insertNumberedEquation(x.body, x.ctx) },
  { group: "Insert", label: "Link, or change the link at the cursor", keys: ["Ctrl+K"], run: (x) => editLink(x.view, x.ctx) },
  { group: "Insert", label: "Cite a source", keys: ["Alt+Shift+C"], run: (x) => insertCitation(x.body, x.ctx) },
  { group: "Insert", label: "Reference list (bibliography)", keys: ["Alt+Shift+B"], run: (x) => insertBibliography(x.body, x.ctx) },
  { group: "Insert", label: "Sources and citation style (BibTeX, DOI)", keys: ["Alt+Shift+S"], run: (x) => manageSources(x.body, x.ctx) },
  { group: "Insert", label: "Follow a link", keys: ["Ctrl+click"] },
  { group: "Review", label: "Write together: share this document, or join one", keys: ["Alt+Shift+W"], run: (x) => x.share(), everywhere: true },
  { group: "Review", label: "New comment", keys: ["Ctrl+Alt+M"], run: (x) => x.newComment() },
  { group: "Review", label: "Show all markup, the final text, or the original", keys: ["Alt+Shift+M"], run: (x) => x.cycleMarkup() },
  { group: "Review", label: "Check before sending (open changes, comments, undescribed pictures, broken references)", keys: ["Alt+Shift+I"], run: (x) => showInspection(x.body, x.ctx), everywhere: true },
  { group: "Review", label: "Compare with another version (the differences as tracked changes)", keys: ["Ctrl+Alt+Shift+C"], run: (x) => void x.compareWithFile(), everywhere: true },
  { group: "Review", label: "Track changes, on or off", keys: ["Ctrl+Shift+E"], run: (x) => x.toggleTracking(), everywhere: true },
  { group: "Review", label: "Next change or comment", keys: ["Alt+PageDown"], run: (x) => x.stepChange(1) },
  { group: "Review", label: "Previous change or comment", keys: ["Alt+PageUp"], run: (x) => x.stepChange(-1) },
  { group: "Review", label: "Accept the change and go to the next", keys: ["Ctrl+Alt+Shift+A"], run: (x) => x.decideChange(true) },
  { group: "Review", label: "Reject the change and go to the next", keys: ["Ctrl+Alt+Shift+D"], run: (x) => x.decideChange(false) },
  { group: "Insert", label: "Symbol (Greek letters, maths signs…)", keys: ["Ctrl+Shift+Y"], run: (x) => x.openMenu("symbol") },
  { group: "Insert", label: "Picture", keys: ["Ctrl+Alt+P"], run: (x) => void choosePicture(x.body, x.ctx) },
  { group: "Insert", label: "Footnote", keys: ["Ctrl+Alt+F"], run: (x) => x.insertFootnote() },
  { group: "Insert", label: "Caption (Figure 1, Table 1…)", keys: ["Ctrl+Alt+C"], run: (x) => captionDialog(x.body, x.ctx) },
  { group: "Insert", label: "Cross-reference, and the other references", keys: ["Ctrl+Alt+R"], run: (x) => x.openMenu("references") },
  { group: "Insert", label: "Table of contents", keys: ["Ctrl+Alt+T"], run: (x) => insertToc(x.body, x.ctx) },
  { group: "Insert", label: "Numbered headings, on or off", keys: ["Ctrl+Alt+N"], run: (x) => toggleHeadingNumbers(x.body, x.ctx) },
  { group: "Insert", label: "Edit the header and footer", keys: ["Ctrl+Alt+O"], run: (x) => x.editHeader() },

  // ---- tables
  { group: "Tables", label: "Insert a table (choose the size with the arrow keys)", keys: ["Ctrl+Alt+Shift+T"], run: (x) => x.openMenu("table") },
  { group: "Tables", label: "Insert a row below", keys: ["Ctrl+Alt+Shift+ArrowDown"], run: plain(TABLE_COMMANDS.rowBelow) },
  { group: "Tables", label: "Insert a row above", keys: ["Ctrl+Alt+Shift+ArrowUp"], run: plain(TABLE_COMMANDS.rowAbove) },
  { group: "Tables", label: "Insert a column to the right", keys: ["Ctrl+Alt+Shift+ArrowRight"], run: plain(TABLE_COMMANDS.columnRight) },
  { group: "Tables", label: "Insert a column to the left", keys: ["Ctrl+Alt+Shift+ArrowLeft"], run: plain(TABLE_COMMANDS.columnLeft) },
  { group: "Tables", label: "Delete the row", keys: ["Ctrl+Alt+Shift+Backspace"], run: plain(TABLE_COMMANDS.deleteRow) },
  { group: "Tables", label: "Delete the column", keys: ["Ctrl+Alt+Shift+Delete"], run: plain(TABLE_COMMANDS.deleteColumn) },
  { group: "Tables", label: "Merge the selected cells (select by dragging, or Shift+arrows)", keys: ["Ctrl+Alt+Shift+M"], run: plain(TABLE_COMMANDS.merge) },
  { group: "Tables", label: "Split a merged cell", keys: ["Ctrl+Alt+Shift+N"], run: plain(TABLE_COMMANDS.split) },
  { group: "Tables", label: "Header row, on or off", keys: ["Ctrl+Alt+Shift+R"], run: plain(TABLE_COMMANDS.headerRow) },
  { group: "Tables", label: "Lines around the cells, on or off", keys: ["Ctrl+Alt+Shift+L"], run: (x) => x.toggleTableLines() },
  { group: "Tables", label: "Sort the rows by the cursor's column (again: the other way)", keys: ["Alt+Shift+O"], run: (x) => x.sortRows() },

  // ---- equations, headers, notes
  { group: "Equations, headers and notes", label: "Finish the equation", keys: ["Enter"] },
  { group: "Equations, headers and notes", label: "Leave the equation, header, footer, note or comment", keys: ["Esc"] },

  // ---- view and layout
  { group: "View and layout", label: "Zoom in", keys: ["Ctrl+Alt++", "Ctrl+Alt+="], run: (x) => x.zoomBy(1), everywhere: true },
  { group: "View and layout", label: "Zoom out", keys: ["Ctrl+Alt+-"], run: (x) => x.zoomBy(-1), everywhere: true },
  { group: "View and layout", label: "Zoom to 100%", keys: ["Ctrl+Alt+0"], run: (x) => x.zoomBy(0), everywhere: true },
  { group: "View and layout", label: "Zoom to the page width, the whole page, or 100%", keys: ["Alt+Shift+Z"], run: (x) => x.cycleZoomFit(), everywhere: true },
  { group: "View and layout", label: "Snap pictures to the grid, on or off", keys: ["Ctrl+Alt+G"], run: (x) => x.toggleSnap() },
  { group: "View and layout", label: "Line grid, on or off", keys: ["Ctrl+Alt+L"], run: (x) => x.setLineGrid(x.ctx.page.lineGrid === null) },
  { group: "View and layout", label: "Move or resize a picture without snapping", keys: ["Alt (hold while dragging)"] },
];

/** The shortcut for a key press, if there is one. */
export function shortcutFor(event: KeyboardEvent): Shortcut | null {
  for (const s of SHORTCUTS) if (s.run && s.keys.some((k) => matches(k, event))) return s;
  return null;
}

/** Handles shortcuts before the editor sees the key. */
export function handleShortcuts(host: ShortcutHost) {
  window.addEventListener(
    "keydown",
    (event) => {
      if (event.defaultPrevented || document.querySelector("dialog:modal")) return;
      const shortcut = shortcutFor(event);
      if (!shortcut) return;
      const target = event.target as HTMLElement | null;
      const inBox = !!target && ["INPUT", "TEXTAREA", "SELECT"].includes(target.tagName);
      if (inBox && !shortcut.everywhere) return;
      // In the headings pane they work on the heading chosen there (navigation.ts).
      if (shortcut.group === "Headings and outline" && target?.closest?.(".nav-list")) return;
      event.preventDefault();
      event.stopPropagation();
      shortcut.run!(host);
    },
    true,
  );
}

/** The help window listing every shortcut, with a box to search them. */
export function showShortcutHelp() {
  document.querySelector(".shortcut-help")?.remove();
  const search = h("input", { class: "tb-input shortcut-search", placeholder: "Search, e.g. “size” or “heading”", "aria-label": "Search shortcuts" });
  const list = h("div", { class: "shortcut-list" });
  const draw = () => {
    const q = search.value.trim().toLowerCase();
    list.replaceChildren();
    const groups = [...new Set(SHORTCUTS.map((s) => s.group))];
    for (const group of groups) {
      const rows = SHORTCUTS.filter((s) => s.group === group && (!q || s.label.toLowerCase().includes(q) || s.keys.join(" ").toLowerCase().includes(q) || group.toLowerCase().includes(q)));
      if (!rows.length) continue;
      list.append(
        h("h3", {}, group),
        ...rows.map((s) =>
          h(
            "div",
            { class: "shortcut-row" },
            h("span", { class: "shortcut-label" }, s.label),
            h("span", { class: "shortcut-keys" }, ...s.keys.flatMap((k, i) => [...(i ? [h("span", { class: "shortcut-or" }, "or")] : []), keysView(k)])),
          ),
        ),
      );
    }
    if (!list.children.length) list.append(h("p", { class: "shortcut-none" }, "No shortcut matches."));
  };
  const dialog = h(
    "dialog",
    { class: "dialog shortcut-help" },
    h("div", { class: "shortcut-head" }, h("h2", {}, "Keyboard shortcuts"), h("button", { type: "button", class: "btn subtle small", onclick: () => dialog.close() }, "Close")),
    search,
    list,
  );
  search.addEventListener("input", draw);
  dialog.addEventListener("close", () => dialog.remove());
  document.body.append(dialog);
  draw();
  dialog.showModal();
  search.focus();
}

/** A shortcut drawn as keys: [Ctrl] + [Shift] + [L]. */
function keysView(spec: string): HTMLElement {
  const parts = spec.endsWith("++") ? [...spec.slice(0, -2).split("+").filter(Boolean), "+"] : spec.split("+");
  return h("span", { class: "keys" }, ...parts.flatMap((p, i) => [...(i ? ["+"] : []), h("kbd", {}, p)]));
}
