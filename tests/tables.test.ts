import * as docx from "docx";
import JSZip from "jszip";
import type { Node } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { addColumnAfter, addRowAfter, mergeCells, CellSelection } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { createState } from "../src/editor/editor";
import { schema } from "../src/editor/schema";

const cell = (text: string, options: Partial<docx.ITableCellOptions> = {}) => new docx.TableCell({ children: [new docx.Paragraph(text)], ...options });

/** A results table: a header row, a cell spanning two columns, one spanning two rows, a shaded cell. */
async function withTable(): Promise<Uint8Array> {
  const d = new docx.Document({
    sections: [
      {
        children: [
          new docx.Paragraph("Results"),
          new docx.Table({
            columnWidths: [2000, 3000, 4000],
            rows: [
              new docx.TableRow({ tableHeader: true, children: [cell("Sample"), cell("Mass (g)"), cell("Notes")] }),
              new docx.TableRow({ children: [cell("A", { rowSpan: 2 }), cell("1.20"), cell("clear", { shading: { fill: "FFF2CC", type: docx.ShadingType.CLEAR, color: "auto" } })] }),
              new docx.TableRow({ children: [cell("1.25"), cell("cloudy")] }),
              new docx.TableRow({ children: [cell("Mean", { columnSpan: 2 }), cell("—")] }),
            ],
          }),
          new docx.Paragraph("After the table"),
        ],
      },
    ],
  });
  return new Uint8Array(await docx.Packer.toBuffer(d));
}

/** The table as text: rows of "text[colspan x rowspan]". */
function shape(table: Node): string[][] {
  const rows: string[][] = [];
  table.forEach((row) => {
    const cells: string[] = [];
    row.forEach((c) => cells.push(`${c.textContent}${c.attrs.colspan > 1 || c.attrs.rowspan > 1 ? `[${c.attrs.colspan}x${c.attrs.rowspan}]` : ""}`));
    rows.push(cells);
  });
  return rows;
}

const firstTable = (doc: Node) => {
  let found: Node | null = null;
  doc.descendants((n) => {
    if (!found && n.type === schema.nodes.table) found = n;
    return !found;
  });
  return found!;
};

describe("tables", () => {
  it("reads a table with merged cells, widths, shading and a header row", async () => {
    const { doc } = await readDocx(await withTable());
    const table = firstTable(doc);
    expect(shape(table)).toEqual([
      ["Sample", "Mass (g)", "Notes"],
      ["A[1x2]", "1.20", "clear"],
      ["1.25", "cloudy"],
      ["Mean[2x1]", "—"],
    ]);
    expect(table.firstChild!.firstChild!.type.name).toBe("table_header");
    expect(table.child(1).child(2).attrs.background).toBe("#FFF2CC");
    expect(table.child(1).child(1).attrs.colwidth).toEqual([200]);
    expect(table.attrs.grid).toEqual([2000, 3000, 4000]);
  });

  it("saves a table that wasn't changed exactly as it was", async () => {
    const original = await withTable();
    const { doc, context } = await readDocx(original);
    const xml = async (bytes: Uint8Array) => {
      const text = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
      return text.slice(text.indexOf("<w:tbl>"), text.indexOf("</w:tbl>") + 8);
    };
    expect(await xml(await writeDocx(doc, context))).toBe(await xml(original));
  });

  it("saves an edited table so it opens again the same, merged cells and all", async () => {
    const { doc, context } = await readDocx(await withTable());
    let state = createState(doc, () => context);
    // Put the cursor in "1.20" (beside the merged "A") and add a row below it, then a column after.
    let at = 0;
    state.doc.descendants((n, pos) => {
      if (n.isText && n.text === "1.20") at = pos;
    });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, at)));
    addRowAfter(state, (tr) => (state = state.apply(tr)));
    addColumnAfter(state, (tr) => (state = state.apply(tr)));
    const before = shape(firstTable(state.doc));
    expect(before[1][0]).toBe("A[1x3]");

    const saved = await writeDocx(state.doc, context);
    const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml.match(/<w:vMerge\/>/g)).toHaveLength(2);
    expect(xml).toContain('<w:tblHeader/>');
    const again = await readDocx(saved);
    expect(shape(firstTable(again.doc))).toEqual(before);
    let shaded: string | null = null;
    firstTable(again.doc).descendants((n) => {
      if (n.type === schema.nodes.table_cell && n.textContent === "clear") shaded = n.attrs.background;
    });
    expect(shaded).toBe("#FFF2CC");
  });

  it("merges cells into one spanning cell", async () => {
    const { doc, context } = await readDocx(await withTable());
    let state = createState(doc, () => context);
    const cells: number[] = [];
    state.doc.descendants((n, pos) => {
      if ((n.type === schema.nodes.table_cell || n.type === schema.nodes.table_header) && ["1.20", "clear"].includes(n.textContent)) cells.push(pos);
    });
    state = state.apply(state.tr.setSelection(CellSelection.create(state.doc, cells[0], cells[1])));
    mergeCells(state, (tr) => (state = state.apply(tr)));
    const saved = await writeDocx(state.doc, context);
    const merged = shape(firstTable((await readDocx(saved)).doc))[1];
    expect(merged).toHaveLength(2);
    expect(merged[1]).toMatch(/^1\.20\s*clear\[2x1\]$/);
  });

  it("fills new cells with an empty paragraph", () => {
    expect(schema.nodes.table_cell.createAndFill()!.firstChild!.type.name).toBe("paragraph");
    expect(schema.nodes.doc.createAndFill()!.firstChild!.type.name).toBe("paragraph");
  });

  it("adds a new table with Word's Table Grid style and lines", async () => {
    const { context } = await readDocx(await blankDocx());
    const row = () => schema.nodes.table_row.create(null, [0, 1].map(() => schema.nodes.table_cell.create({ colwidth: [150] }, schema.nodes.paragraph.create())));
    const table = schema.nodes.table.create({ style: "TableGrid" }, [row(), row()]);
    const saved = await writeDocx(schema.nodes.doc.create(null, [table, schema.nodes.paragraph.create()]), context);
    const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:tblStyle w:val="TableGrid"/>');
    expect(xml).toContain('<w:gridCol w:w="2250"/><w:gridCol w:w="2250"/>');
    expect(xml.match(/<w:tc>/g)).toHaveLength(4);
  });
});
