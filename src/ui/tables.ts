// Making and changing tables: a grid to pick the size of a new table, and the
// table commands for the toolbar and the keyboard.

import { Fragment } from "prosemirror-model";
import { TextSelection, type Command, type EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  CellSelection,
  TableMap,
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  findTable,
  isInTable,
  mergeCells,
  setCellAttr,
  splitCell,
  toggleHeaderRow,
} from "prosemirror-tables";
import { h } from "../dom";
import { ensureStyles } from "../docx/builtins";
import type { DocContext } from "../docx/model";
import { schema } from "../editor/schema";

/** A new table of rows × columns, as wide as the text, in Word's Table Grid style. */
export function insertTable(view: EditorView, ctx: DocContext, rows: number, cols: number) {
  ensureStyles(ctx, "TableGrid");
  const p = ctx.page;
  const width = (p.width - p.marginLeft - p.marginRight) * (4 / 3);
  const colwidth = Math.floor(width / cols);
  const row = () => schema.nodes.table_row.create(null, Array.from({ length: cols }, () => schema.nodes.table_cell.create({ colwidth: [colwidth] }, schema.nodes.paragraph.create())));
  const table = schema.nodes.table.create({ style: "TableGrid" }, Array.from({ length: rows }, row));
  const tr = view.state.tr.replaceSelectionWith(table);
  // As in Word, there is always a paragraph after a table to carry on typing in.
  const after = tr.mapping.map(view.state.selection.to);
  if (!tr.doc.resolve(after).nodeAfter && after >= tr.doc.content.size - 1) tr.insert(tr.doc.content.size, schema.nodes.paragraph.create());
  // The cursor goes in the first cell.
  const start = tr.mapping.map(view.state.selection.from);
  let firstCell = -1;
  tr.doc.nodesBetween(Math.max(0, start - table.nodeSize - 2), Math.min(tr.doc.content.size, start + table.nodeSize), (node, pos) => {
    if (firstCell < 0 && node.type === schema.nodes.table_cell) firstCell = pos;
    return firstCell < 0;
  });
  if (firstCell >= 0) tr.setSelection(TextSelection.near(tr.doc.resolve(firstCell + 2)));
  view.dispatch(tr.scrollIntoView());
  view.focus();
}

/** The grid to choose a new table's size: move over it and click. */
export function tableGrid(onPick: (rows: number, cols: number) => void): HTMLElement {
  const label = h("div", { class: "table-grid-label" }, "Insert table");
  const grid = h("div", { class: "table-grid", role: "grid" });
  const cells: HTMLElement[] = [];
  const mark = (r: number, c: number) => {
    for (const cell of cells) cell.classList.toggle("on", Number(cell.dataset.r) <= r && Number(cell.dataset.c) <= c);
    label.textContent = `${r} × ${c} table`;
  };
  for (let r = 1; r <= 8; r++) {
    for (let c = 1; c <= 10; c++) {
      const cell = h("button", { type: "button", class: "table-grid-cell", title: `${r} × ${c}`, "data-r": String(r), "data-c": String(c) });
      cell.addEventListener("mouseenter", () => mark(r, c));
      cell.addEventListener("focus", () => mark(r, c));
      cell.addEventListener("click", () => onPick(r, c));
      cells.push(cell);
      grid.append(cell);
    }
  }
  return h("div", { class: "menu table-menu" }, label, grid);
}

/** Lines on the table's cells: on, off. */
export function setTableLines(on: boolean): Command {
  return (state, dispatch) => {
    if (!isInTable(state)) return false;
    const table = findTable(state.selection.$from);
    if (!table) return false;
    if (dispatch) dispatch(state.tr.setNodeMarkup(table.pos, undefined, { ...table.node.attrs, borders: on, lines: null }));
    return true;
  };
}

export const TABLE_COMMANDS = {
  rowAbove: addRowBefore,
  rowBelow: addRowAfter,
  columnLeft: addColumnBefore,
  columnRight: addColumnAfter,
  deleteRow,
  deleteColumn,
  deleteTable,
  merge: mergeCells,
  split: splitCell,
  headerRow: toggleHeaderRow,
  shade: (color: string | null) => setCellAttr("background", color),
  valign: (where: "top" | "middle" | "bottom") => setCellAttr("valign", where),
  selectTable: selectCells("table"),
  selectRow: selectCells("row"),
  selectColumn: selectCells("column"),
};

/** Selects the whole table, or the row or column the cursor is in, as Word's Select menu. */
function selectCells(what: "table" | "row" | "column"): Command {
  return (state, dispatch) => {
    const selection = cellsAt(state, what, state.selection.$from);
    if (!selection) return false;
    if (dispatch) dispatch(state.tr.setSelection(selection));
    return true;
  };
}

/** The whole table (or row, or column) around `$at`, as a cell selection. */
export function cellsAt(state: EditorState, what: "table" | "row" | "column", $at: import("prosemirror-model").ResolvedPos): CellSelection | null {
  const found = findTable($at);
  if (!found) return null;
  const map = TableMap.get(found.node);
  const start = found.pos + 1;
  const $cell = (index: number) => state.doc.resolve(start + map.map[index]);
  if (what === "table") return new CellSelection($cell(0), $cell(map.map.length - 1));
  const sel = state.selection;
  const rect = map.findCell((sel instanceof CellSelection ? sel.$headCell.pos : cellStart(state)) - start);
  return what === "row"
    ? CellSelection.rowSelection($cell(rect.top * map.width), $cell(rect.top * map.width + map.width - 1))
    : CellSelection.colSelection($cell(rect.left), $cell((map.height - 1) * map.width + rect.left));
}

/** Where the cell the cursor is in starts. */
function cellStart(state: EditorState): number {
  const $from = state.selection.$from;
  for (let d = $from.depth; d > 0; d--) if (/^table_(cell|header)$/.test($from.node(d).type.name)) return $from.before(d);
  return $from.pos;
}

/**
 * Backspace on selected cells, as in Word: the whole table selected removes the table,
 * whole rows or columns remove them; other cells are emptied (as Delete does).
 */
export const backspaceCells: Command = (state, dispatch) => {
  const sel = state.selection;
  if (!(sel instanceof CellSelection)) return false;
  if (sel.isRowSelection() && sel.isColSelection()) return deleteTable(state, dispatch);
  if (sel.isRowSelection()) return deleteRow(state, dispatch);
  if (sel.isColSelection()) return deleteColumn(state, dispatch);
  return false;
};

/** The lines of a table, side by side, as Word's Borders menu sets them. Size is in eighths of a point. */
export interface TableLines {
  top: boolean;
  bottom: boolean;
  left: boolean;
  right: boolean;
  insideH: boolean;
  insideV: boolean;
  style: "single" | "double" | "dashed" | "dotted";
  size: number;
  color: string;
}

export const DEFAULT_LINES: TableLines = { top: true, bottom: true, left: true, right: true, insideH: true, insideV: true, style: "single", size: 4, color: "auto" };

/** The table's lines as they are now (all on, in thin black, when the table only says "lines" or follows its style). */
export function linesOf(state: EditorState): TableLines | null {
  const table = findTable(state.selection.$from);
  if (!table) return null;
  const { lines, borders } = table.node.attrs;
  if (lines) return lines as TableLines;
  const on = borders !== false;
  return { ...DEFAULT_LINES, top: on, bottom: on, left: on, right: on, insideH: on, insideV: on };
}

/** Changes some of the table's lines: which sides, and their style, width and colour. */
export function setLines(change: Partial<TableLines>): Command {
  return (state, dispatch) => {
    const table = findTable(state.selection.$from);
    const current = linesOf(state);
    if (!table || !current) return false;
    if (dispatch) dispatch(state.tr.setNodeMarkup(table.pos, undefined, { ...table.node.attrs, borders: null, lines: { ...current, ...change } }));
    return true;
  };
}

export { isInTable };

/**
 * Sorts the table's rows by the column the cursor is in: A to Z (numbers as
 * numbers), or Z to A when it already is (Alt+Shift+O). Header rows stay on top.
 * A table with merged cells can't be sorted.
 */
export function sortTable(): Command {
  return (state, dispatch) => {
    const found = findTable(state.selection.$from);
    if (!found) return false;
    const { node: table, pos } = found;
    const $from = state.selection.$from;
    let cellDepth = $from.depth;
    while (cellDepth > 0 && !/^table_(cell|header)$/.test($from.node(cellDepth).type.name)) cellDepth--;
    if (!cellDepth) return false;
    const column = $from.index(cellDepth - 1);
    const rows: import("prosemirror-model").Node[] = [];
    let merged = false;
    table.forEach((row) => {
      rows.push(row);
      row.forEach((cell) => {
        if (cell.attrs.colspan > 1 || cell.attrs.rowspan > 1) merged = true;
      });
    });
    if (merged) return false;
    const headers = rows.findIndex((r) => r.firstChild?.type !== schema.nodes.table_header);
    const top = rows.slice(0, headers < 0 ? rows.length : headers);
    const body = rows.slice(top.length);
    if (body.length < 2) return false;
    const key = (row: import("prosemirror-model").Node) => (column < row.childCount ? row.child(column).textContent.trim() : "");
    const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });
    const compare = (a: import("prosemirror-model").Node, b: import("prosemirror-model").Node) => {
      const x = key(a);
      const y = key(b);
      const nx = Number(x.replace(",", "."));
      const ny = Number(y.replace(",", "."));
      if (x && y && !Number.isNaN(nx) && !Number.isNaN(ny)) return nx - ny;
      return collator.compare(x, y);
    };
    const ascending = [...body].sort(compare);
    const already = ascending.every((r, i) => compare(r, body[i]) === 0);
    const sorted = already ? [...body].sort((a, b) => compare(b, a)) : ascending;
    if (dispatch) {
      const next = table.copy(Fragment.from([...top, ...sorted]));
      const tr = state.tr.replaceWith(pos, pos + table.nodeSize, next);
      tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(state.selection.from, tr.doc.content.size))));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/** The selected paragraphs as a table: a row each, cells split at tabs (or, without tabs, at commas). */
export function textToTable(ctx: DocContext): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;
    if ($from.depth < 1 || isInTable(state)) return false;
    const start = $from.before(1);
    const end = $to.after(1);
    const paras: import("prosemirror-model").Node[] = [];
    state.doc.nodesBetween(start, end, (node, _pos, parent) => {
      if (parent === state.doc && node.type === schema.nodes.paragraph) paras.push(node);
      return false;
    });
    if (!paras.length) return false;
    const separator = paras.some((p) => p.textContent.includes("\t")) ? "\t" : ",";
    // Each paragraph cut into pieces at the separator, keeping their formatting.
    const pieces = (para: import("prosemirror-model").Node) => {
      const out: import("prosemirror-model").Node[] = [];
      let from = 0;
      let at = 0;
      para.forEach((child) => {
        if (child.isText) {
          const text = child.text!;
          for (let i = 0; i < text.length; i++) {
            if (text[i] === separator) {
              out.push(para.cut(from, at + i));
              from = at + i + 1;
            }
          }
        }
        at += child.nodeSize;
      });
      out.push(para.cut(from));
      // Spaces after a comma aren't part of the next cell.
      return out.map((p) => (separator === "," && p.firstChild?.isText ? p.cut(p.textContent.length - p.textContent.trimStart().length) : p));
    };
    const rows = paras.map(pieces);
    const columns = Math.max(...rows.map((r) => r.length));
    const page = ctx.page;
    const width = Math.floor(((page.width - page.marginLeft - page.marginRight) * (4 / 3)) / columns);
    ensureStyles(ctx, "TableGrid");
    const table = schema.nodes.table.create(
      { style: "TableGrid" },
      rows.map((r) =>
        schema.nodes.table_row.create(
          null,
          Array.from({ length: columns }, (_, i) => schema.nodes.table_cell.create({ colwidth: [width] }, r[i] ? schema.nodes.paragraph.create(null, r[i].content) : schema.nodes.paragraph.create())),
        ),
      ),
    );
    if (dispatch) {
      const tr = state.tr.replaceWith(start, end, table);
      if (!tr.doc.resolve(start + table.nodeSize).nodeAfter) tr.insert(start + table.nodeSize, schema.nodes.paragraph.create());
      tr.setSelection(TextSelection.near(tr.doc.resolve(start + 3)));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}
