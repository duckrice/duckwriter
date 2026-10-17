import { TextSelection } from "prosemirror-state";
import { keydownHandler } from "prosemirror-keymap";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { createState } from "../src/editor/editor";
import { schema } from "../src/editor/schema";

describe("Tab in a table", () => {
  it("in the last cell adds a row and moves into it", async () => {
    const { context } = await readDocx(await blankDocx());
    const cell = () => schema.nodes.table_cell.create(null, schema.nodes.paragraph.create());
    const doc = schema.nodes.doc.create(null, [schema.nodes.table.create(null, [schema.nodes.table_row.create(null, [cell()])]), schema.nodes.paragraph.create()]);
    let state = createState(doc, () => context);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));
    const tab = state.plugins.map((p) => p.props.handleKeyDown).filter(Boolean);
    const view = { state, dispatch: (tr: import("prosemirror-state").Transaction) => (view.state = view.state.apply(tr)) } as unknown as import("prosemirror-view").EditorView;
    const event = new KeyboardEvent("keydown", { key: "Tab" });
    tab.some((handler) => handler!.call(null as never, view, event));
    const rows = view.state.doc.firstChild!.childCount;
    expect(rows).toBe(2);
    expect(view.state.selection.$from.node(-2)).toBe(view.state.doc.firstChild!.child(1));
    void keydownHandler;
  });
});
