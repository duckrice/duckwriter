import { describe, expect, it } from "vitest";
import { EditorState, NodeSelection, TextSelection } from "prosemirror-state";
import { schema } from "../src/editor/schema";
import { captionFor } from "../src/ui/references";

const cell = (text: string) => schema.nodes.table_cell.create(null, schema.nodes.paragraph.create(null, schema.text(text)));

describe("captions follow what the cursor is on", () => {
  it("a table gets a Table caption above it", () => {
    const doc = schema.nodes.doc.create(null, [schema.nodes.table.create(null, schema.nodes.table_row.create(null, [cell("a")])), schema.nodes.paragraph.create()]);
    const state = EditorState.create({ doc, selection: TextSelection.create(doc, 4) });
    expect(captionFor(state)).toEqual({ label: "Table", above: true });
  });

  it("an equation gets an Equation caption below it, and plain text a Figure one", () => {
    const math = schema.nodes.math.create({ latex: "x^2", display: true });
    const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, math), schema.nodes.paragraph.create(null, schema.text("Text"))]);
    expect(captionFor(EditorState.create({ doc, selection: NodeSelection.create(doc, 1) }))).toEqual({ label: "Equation", above: false });
    expect(captionFor(EditorState.create({ doc, selection: TextSelection.create(doc, 5) }))).toEqual({ label: "Figure", above: false });
  });
});
