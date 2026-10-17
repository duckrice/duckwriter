// How a table shows: the table add-on's view (with column widths and
// resizing), plus the table's style and whether it has lines.

import type { Node } from "prosemirror-model";
import { TableView } from "prosemirror-tables";

export class DuckTableView extends TableView {
  constructor(node: Node, cellMinWidth: number) {
    super(node, cellMinWidth);
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
  }
}
