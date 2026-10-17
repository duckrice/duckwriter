// Making and changing tables: a grid to pick the size of a new table, and the
// table commands for the toolbar and the keyboard.

import { TextSelection, type Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
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
    if (dispatch) dispatch(state.tr.setNodeMarkup(table.pos, undefined, { ...table.node.attrs, borders: on }));
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
};

export { isInTable };
