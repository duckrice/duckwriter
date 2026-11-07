// How a table shows: the table add-on's view (with column widths and
// resizing), plus the table's style and whether it has lines.

import type { Node } from "prosemirror-model";
import { TableView } from "prosemirror-tables";

export class DuckTableView extends TableView {
  constructor(node: Node, cellMinWidth: number) {
    super(node, cellMinWidth);
    // Word's handle at the table's corner: a click selects the whole table (editor.ts handles it).
    const handle = document.createElement("span");
    handle.className = "table-handle";
    handle.title = "Select the whole table (then copy, cut, or Backspace to remove it)";
    handle.contentEditable = "false";
    handle.textContent = "⊞";
    this.dom.prepend(handle);
    this.dress(node);
  }

  update(node: Node): boolean {
    if (!super.update(node)) return false;
    this.dress(node);
    return true;
  }

  private dress(node: Node) {
    const { borders, style } = node.attrs;
    this.table.classList.toggle("lines", borders === true);
    this.table.classList.toggle("no-lines", borders === false);
    if (style) this.table.dataset.style = style;
    else delete this.table.dataset.style;
    // Lines chosen side by side, as custom properties the stylesheet draws with.
    const lines = node.attrs.lines as { [side: string]: unknown; style: string; size: number; color: string } | null;
    this.table.classList.toggle("custom-lines", !!lines);
    for (const side of ["top", "bottom", "left", "right", "insideH", "insideV"]) {
      if (!lines) this.table.style.removeProperty(`--line-${side}`);
      else this.table.style.setProperty(`--line-${side}`, lines[side] ? lineCss(lines.style, lines.size, lines.color) : "none");
    }
  }
}

function lineCss(style: string, size: number, color: string): string {
  const css = { single: "solid", double: "double", dashed: "dashed", dotted: "dotted" }[style] ?? "solid";
  // A double line needs room for its two strokes.
  const width = Math.max(css === "double" ? 2.25 : 0.5, size / 8);
  return `${width}pt ${css} ${/^[0-9a-f]{6}$/i.test(color) ? `#${color}` : "#000000"}`;
}
