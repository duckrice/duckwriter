import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { EditorState, TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { schema } from "../src/editor/schema";
import { TABLE_COMMANDS, backspaceCells, linesOf, setLines } from "../src/ui/tables";

const cell = (text: string) => schema.nodes.table_cell.create(null, schema.nodes.paragraph.create(null, schema.text(text)));
const row = (...texts: string[]) => schema.nodes.table_row.create(null, texts.map(cell));

function inTable() {
  const doc = schema.nodes.doc.create(null, [
    schema.nodes.paragraph.create(null, schema.text("Before")),
    schema.nodes.table.create({ style: "TableGrid" }, [row("a1", "b1"), row("a2", "b2")]),
    schema.nodes.paragraph.create(null, schema.text("After")),
  ]);
  let state = EditorState.create({ doc });
  // The cursor in "a1".
  state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 12)));
  const run = (cmd: (s: EditorState, d: (tr: import("prosemirror-state").Transaction) => void) => boolean) => {
    const ok = cmd(state, (tr) => (state = state.apply(tr)));
    return ok;
  };
  return { get: () => state, run };
}

const blocks = (state: EditorState) => state.doc.content.content.map((n) => `${n.type.name}:${n.textContent}`);

// Bug: a table couldn't easily be removed; Backspace on its selected cells only emptied them.
describe("removing tables, rows and columns", () => {
  it("Select table, then Backspace, removes the table", () => {
    const t = inTable();
    expect(t.run(TABLE_COMMANDS.selectTable)).toBe(true);
    expect(t.get().selection).toBeInstanceOf(CellSelection);
    expect(t.run(backspaceCells)).toBe(true);
    expect(blocks(t.get())).toEqual(["paragraph:Before", "paragraph:After"]);
  });

  it("a selected row or column goes with Backspace", () => {
    const t = inTable();
    t.run(TABLE_COMMANDS.selectRow);
    t.run(backspaceCells);
    expect(blocks(t.get())[1]).toBe("table:a2b2");
    const c = inTable();
    c.run(TABLE_COMMANDS.selectColumn);
    c.run(backspaceCells);
    expect(blocks(c.get())[1]).toBe("table:b1b2");
  });
});

describe("the table's lines", () => {
  it("are set side by side, and saved so Word draws the same", async () => {
    const t = inTable();
    expect(linesOf(t.get())).toMatchObject({ top: true, insideV: true, style: "single", size: 4 });
    t.run(setLines({ insideH: false, insideV: false, style: "double", size: 12, color: "c00000" }));
    const table = t.get().doc.child(1);
    expect(table.attrs.lines).toMatchObject({ top: true, insideH: false, style: "double", size: 12, color: "c00000" });
    const { context } = await readDocx(await blankDocx());
    const zip = await JSZip.loadAsync(await writeDocx(t.get().doc, context));
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:top w:val="double" w:sz="12" w:space="0" w:color="c00000"/>');
    expect(xml).toContain('<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>');
  });
});
