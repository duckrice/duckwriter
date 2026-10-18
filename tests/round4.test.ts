import { TextSelection, type EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { toHtml, toMarkdown, toText } from "../src/convert/export";
import { htmlToDoc, markdownToHtml } from "../src/convert/import";
import type { DocContext } from "../src/docx/model";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { mathShorthand } from "../src/editor/autoformat";
import { createState } from "../src/editor/editor";
import { schema } from "../src/editor/schema";
import { sortTable, textToTable } from "../src/ui/tables";

async function ctx(): Promise<DocContext> {
  return (await readDocx(await blankDocx())).context;
}

const para = (text: string) => schema.nodes.paragraph.create(null, text ? schema.text(text) : null);

describe("maths shorthand", () => {
  it("turns \\alpha and -> into symbols when the word is finished", async () => {
    const c = await ctx();
    const at = (text: string) => {
      let state = createState(schema.nodes.doc.create(null, [para(text)]), () => c);
      state = state.apply(state.tr.setSelection(TextSelection.atEnd(state.doc)));
      const tr = mathShorthand(state, " ");
      return tr ? state.apply(tr).doc.textContent : null;
    };
    expect(at("angle \\alpha")).toBe("angle α");
    expect(at("x \\leq")).toBe("x ≤");
    expect(at("A ->")).toBe("A →");
    expect(at("x <=")).toBe("x ≤");
    expect(at("5 +-")).toBe("5 ±");
    expect(at("\\notacommand")).toBeNull();
    expect(at("<!--")).toBeNull();
    expect(at("plain words")).toBeNull();
  });
});

describe("tables", () => {
  const table = (header: string[], rows: string[][]) => {
    const cell = (type: "table_cell" | "table_header", text: string) => schema.nodes[type].create(null, para(text));
    return schema.nodes.table.create(null, [schema.nodes.table_row.create(null, header.map((h) => cell("table_header", h))), ...rows.map((r) => schema.nodes.table_row.create(null, r.map((t) => cell("table_cell", t))))]);
  };
  const column = (state: EditorState, c: number) => {
    const out: string[] = [];
    state.doc.firstChild!.forEach((row) => out.push(row.child(c).textContent));
    return out;
  };

  it("sort by the cursor's column, numbers as numbers, the header staying on top; again for the other way", async () => {
    const c = await ctx();
    let state = createState(schema.nodes.doc.create(null, [table(["Name", "Mass"], [["b", "10"], ["a", "9"], ["c", "100"]]), para("")]), () => c);
    // In the second cell of the first body row.
    let pos = -1;
    state.doc.descendants((n, p) => {
      if (pos < 0 && n.isText && n.text === "10") pos = p;
    });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
    sortTable()(state, (tr) => (state = state.apply(tr)));
    expect(column(state, 1)).toEqual(["Mass", "9", "10", "100"]);
    expect(column(state, 0)).toEqual(["Name", "a", "b", "c"]);
    sortTable()(state, (tr) => (state = state.apply(tr)));
    expect(column(state, 1)).toEqual(["Mass", "100", "10", "9"]);
  });

  it("are made from lines with tabs or commas", async () => {
    const c = await ctx();
    let state = createState(schema.nodes.doc.create(null, [para("Run\tRate"), para("1\t0.5"), para("2\t0.7")]), () => c);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, state.doc.content.size - 1)));
    expect(textToTable(c)(state, (tr) => (state = state.apply(tr)))).toBe(true);
    expect(state.doc.firstChild!.type).toBe(schema.nodes.table);
    expect(column(state, 1)).toEqual(["Rate", "0.5", "0.7"]);
    let commas = createState(schema.nodes.doc.create(null, [para("a, b, c")]), () => c);
    commas = commas.apply(commas.tr.setSelection(TextSelection.create(commas.doc, 2)));
    textToTable(c)(commas, (tr) => (commas = commas.apply(tr)));
    expect(column(commas, 2)).toEqual(["c"]);
  });
});

describe("Markdown, HTML and text", () => {
  const md = `# Rates of things

## Methods

We measured **mass** and *time* at [the lab](https://duck.no), with $E = mc^2$.

- first
  - inner
- second

1. one
2. two

| Run | Rate |
| --- | --- |
| 1 | 0.5 |

$$\\frac{a}{b}$$
`;

  it("Markdown opens as a document: title, headings, lists, links, equations, tables", async () => {
    const c = await ctx();
    const doc = htmlToDoc(markdownToHtml(md), c);
    const kinds: string[] = [];
    doc.forEach((n) => kinds.push(n.type === schema.nodes.paragraph ? (n.attrs.style ?? (n.attrs.numId ? `list${n.attrs.ilvl}` : "p")) : n.type.name));
    expect(kinds.slice(0, 3)).toEqual(["Title", "Heading1", "p"]);
    expect(kinds.filter((k) => k.startsWith("list") || k === "ListParagraph").length).toBeGreaterThanOrEqual(5);
    expect(kinds).toContain("table");
    let maths = 0;
    let link = "";
    doc.descendants((n) => {
      if (n.type === schema.nodes.math) maths++;
      const l = schema.marks.link.isInSet(n.marks);
      if (l) link = l.attrs.href;
    });
    expect(maths).toBe(2);
    expect(link).toBe("https://duck.no");
  });

  it("and goes back out as Markdown, HTML with MathML, and text", async () => {
    const c = await ctx();
    const doc = htmlToDoc(markdownToHtml(md), c);
    const out = await toMarkdown(doc, c);
    expect(out).toContain("# Rates of things");
    expect(out).toContain("## Methods");
    expect(out).toContain("**mass**");
    expect(out).toContain("[the lab](https://duck.no)");
    expect(out).toContain("$E = mc^2$");
    expect(out).toMatch(/- first\n {3}- inner\n- second/);
    expect(out).toContain("| Run | Rate |");
    const html = await toHtml(doc, c, "Rates");
    expect(html).toContain("<h1>Rates of things</h1>");
    expect(html).toContain("<math");
    expect(html).toContain('<a href="https://duck.no">the lab</a>');
    expect(html).toMatch(/<ul><li>first<ul><li>inner<\/li><\/ul><\/li><li>second<\/li><\/ul>/);
    const text = toText(doc, c);
    expect(text).toContain("Run\tRate");
  });
});

describe("endnotes", () => {
  it("are saved in Word's endnotes part and read back, listed at the end", async () => {
    const JSZip = (await import("jszip")).default;
    const { writeDocx } = await import("../src/docx/write");
    const { doc: blank, context } = await readDocx(await blankDocx());
    void blank;
    context.endnotes.notes.set("1", schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text("A note at the end."))));
    const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, [schema.text("Text"), schema.nodes.endnote.create({ id: "1" })]), schema.nodes.endnotes.create()]);
    const saved = await writeDocx(doc, context);
    const zip = await JSZip.loadAsync(saved);
    const endnotes = await zip.file("word/endnotes.xml")!.async("string");
    expect(endnotes).toContain('<w:endnote w:id="1">');
    expect(endnotes).toContain("<w:endnoteRef/>");
    expect(endnotes).toContain("A note at the end.");
    expect(await zip.file("word/document.xml")!.async("string")).toContain('<w:endnoteReference w:id="1"/>');
    expect(await zip.file("word/settings.xml")!.async("string")).toContain("<w:endnotePr>");
    const again = await readDocx(saved);
    expect(again.context.endnotes.notes.get("1")?.textContent).toBe("A note at the end.");
    expect(again.doc.lastChild!.type).toBe(schema.nodes.endnotes);
  });
});
