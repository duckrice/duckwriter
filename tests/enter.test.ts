import { describe, expect, it } from "vitest";
import { EditorState, TextSelection } from "prosemirror-state";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { ensureStyles } from "../src/docx/builtins";
import { schema } from "../src/editor/schema";
import { newParagraph } from "../src/editor/commands";

const cell = (text: string) => schema.nodes.table_cell.create(null, schema.nodes.paragraph.create(null, text ? schema.text(text) : null));
const table = () => schema.nodes.table.create(null, schema.nodes.table_row.create(null, [cell("a"), cell("b")]));

async function setup(doc: import("prosemirror-model").Node, at: (doc: import("prosemirror-model").Node) => number) {
  const { context } = await readDocx(await blankDocx());
  ensureStyles(context, "Heading1");
  let state = EditorState.create({ doc });
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, at(state.doc))));
  newParagraph(context)(state, (tr) => (state = state.apply(tr)));
  return state;
}

// Bug: after a table, Enter at the start of the heading below made an empty heading, not a line to type in.
describe("Enter around tables", () => {
  it("at the start of a heading after a table, puts an ordinary line between them", async () => {
    const heading = schema.nodes.paragraph.create({ style: "Heading1" }, schema.text("Results"));
    const doc = schema.nodes.doc.create(null, [table(), heading]);
    const state = await setup(doc, (d) => d.child(0).nodeSize + 1);
    expect(state.doc.childCount).toBe(3);
    expect(state.doc.child(1).type).toBe(schema.nodes.paragraph);
    expect(state.doc.child(1).attrs.style).toBeNull();
    expect(state.doc.child(1).content.size).toBe(0);
    expect(state.doc.child(2).attrs.style).toBe("Heading1");
  });

  it("at the start of a table's first cell, with nothing above it, puts a line above the table", async () => {
    const doc = schema.nodes.doc.create(null, [table(), schema.nodes.paragraph.create()]);
    const state = await setup(doc, () => 4);
    expect(state.doc.child(0).type).toBe(schema.nodes.paragraph);
    expect(state.doc.child(1).type).toBe(schema.nodes.table);
    expect(state.selection.$from.parent).toBe(state.doc.child(0));
    // The cell's text is untouched.
    expect(state.doc.child(1).textContent).toBe("ab");
  });
});
