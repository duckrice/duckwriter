import JSZip from "jszip";
import { TextSelection, type EditorState, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { describe, expect, it } from "vitest";
import { parseBibtex, makeTag, type Bibliography } from "../src/docx/bibliography";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { formatCitation, formatEntry } from "../src/editor/citations";
import { changeCase } from "../src/editor/commands";
import { createState } from "../src/editor/editor";
import { schema } from "../src/editor/schema";
import { TEMPLATES, templateDoc } from "../src/templates";
import { referenceTargets, updateTocs } from "../src/ui/references";

const part = async (bytes: Uint8Array, path: string) => (await JSZip.loadAsync(bytes)).file(path)!.async("string");

/** An editor without a window. */
function fakeView(state: EditorState) {
  const view = {
    state,
    dispatch(tr: Transaction) {
      view.state = view.state.apply(tr);
    },
  };
  return view as unknown as EditorView & { state: EditorState };
}

describe("citation styles", () => {
  const list = parseBibtex(`@article{a, author={Smith, Anna and Jones, Bob}, title={Rates of things}, journal={Journal of Ducks}, year=2020, volume={3}, number={2}, pages={1--10}, doi={10.1/x}}
@book{b, author={Haaland, Kari}, title={Fjordar}, publisher={Samlaget}, address={Oslo}, year=2019}
@misc{c, author={Lie, Ola}, title={Notes}, year=2021}`);
  for (const s of list) s.tag = makeTag(s, list);
  const bib = (style: Bibliography["style"]): Bibliography => ({ path: null, style, sources: list, version: 0, savedVersion: 0 });
  const text = (pieces: { text: string; italic?: boolean }[]) => pieces.map((p) => (p.italic ? `*${p.text}*` : p.text)).join("");

  it("Vancouver numbers in citing order, with ranges", () => {
    const order = ["Smi20", "Haa19", "Lie21"];
    expect(formatCitation({ tags: ["Lie21", "Smi20", "Haa19"], page: "" }, bib("Vancouver"), order)).toBe("(1–3)");
    expect(formatCitation({ tags: ["Lie21", "Smi20"], page: "" }, bib("Vancouver"), order)).toBe("(1, 3)");
    expect(text(formatEntry(list[0], "Vancouver", 1))).toBe("1.\tSmith A, Jones B. Rates of things. Journal of Ducks. 2020;3(2):1–10. doi:10.1/x");
    expect(text(formatEntry(list[1], "Vancouver", 2))).toBe("2.\tHaaland K. Fjordar. Oslo: Samlaget; 2019.");
  });

  it("Chicago author-date", () => {
    expect(formatCitation({ tags: ["Smi20"], page: "23" }, bib("Chicago"), [])).toBe("(Smith and Jones 2020, 23)");
    expect(text(formatEntry(list[0], "Chicago", 1))).toBe("Smith, Anna, and Bob Jones. 2020. “Rates of things.” *Journal of Ducks* 3 (2): 1–10. https://doi.org/10.1/x.");
    expect(text(formatEntry(list[1], "Chicago", 1))).toBe("Haaland, Kari. 2019. *Fjordar.* Oslo: Samlaget.");
  });
});

describe("lists of figures, numbered equations, line numbers", () => {
  it("lists the captions with their pages, and saves Word's table of figures", async () => {
    const { context } = await readDocx(await blankDocx());
    const seq = (label: string, n: number) => schema.nodes.field.create({ instr: ` SEQ ${label} \\* ARABIC `, result: String(n) });
    const caption = (label: string, n: number, text: string) => schema.nodes.paragraph.create({ style: "Caption" }, [schema.text(`${label} `), seq(label, n), schema.text(`: ${text}`)]);
    const doc = schema.nodes.doc.create(null, [schema.nodes.toc.create({ label: "Figure", title: "List of figures" }), caption("Figure", 1, "The setup"), caption("Table", 1, "Data"), caption("Figure", 2, "Results")]);
    const view = fakeView(createState(doc, () => context));
    updateTocs(view, context, { count: 2, notes: [], pageOfBlock: new Map([[view.state.doc.child(0).nodeSize, 0], [view.state.doc.child(0).nodeSize + view.state.doc.child(1).nodeSize + view.state.doc.child(2).nodeSize, 1]]) });
    const entries = view.state.doc.firstChild!.attrs.entries as { text: string; page: number }[];
    expect(entries.map((e) => [e.text, e.page])).toEqual([["Figure 1: The setup", 1], ["Figure 2: Results", 2]]);
    const saved = await writeDocx(view.state.doc, context);
    const xml = await part(saved, "word/document.xml");
    expect(xml).toContain(' TOC \\h \\z \\c "Figure" ');
    expect(xml).toContain('<w:pStyle w:val="TableofFigures"/>');
    const again = await readDocx(saved);
    expect(again.doc.firstChild!.attrs).toMatchObject({ label: "Figure", title: "List of figures" });
  });

  it("refers to a numbered equation as (1)", async () => {
    const { context } = await readDocx(await blankDocx());
    const number = [schema.text("("), schema.nodes.field.create({ instr: " SEQ Equation \\* ARABIC ", result: "1" }), schema.text(")")];
    const cell = (content: ReturnType<typeof schema.nodes.paragraph.create>) => schema.nodes.table_cell.create(null, content);
    const table = schema.nodes.table.create({ borders: false }, schema.nodes.table_row.create(null, [cell(schema.nodes.paragraph.create()), cell(schema.nodes.paragraph.create(null, schema.nodes.math.create({ latex: "E=mc^2" }))), cell(schema.nodes.paragraph.create({ align: "right" }, number))]));
    const doc = schema.nodes.doc.create(null, [table, schema.nodes.paragraph.create()]);
    const targets = referenceTargets(context, doc);
    expect(targets.map((t) => [t.kind, t.label])).toEqual([["caption", "Equation (1)"]]);
  });

  it("saves line numbers as Word's lnNumType", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    context.page = { ...context.page, lineNumbers: { countBy: 5, restart: "newPage" } };
    context.pageChanged = true;
    const saved = await writeDocx(doc, context);
    expect(await part(saved, "word/document.xml")).toContain('<w:lnNumType w:countBy="5" w:restart="newPage"/>');
    expect((await readDocx(saved)).context.page.lineNumbers).toEqual({ countBy: 5, restart: "newPage" });
  });

  it("saves a page border as Word's pgBorders, and reads it back", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    context.page = { ...context.page, border: { kind: "double", size: 0.75, color: "#1F4E79", space: 24, fromText: false } };
    context.borderChanged = true;
    const saved = await writeDocx(doc, context);
    const xml = await part(saved, "word/document.xml");
    expect(xml).toContain('<w:pgBorders w:offsetFrom="page"><w:top w:val="double" w:sz="6" w:space="24" w:color="1F4E79"/>');
    // In Word's order: after the margins, before the columns.
    expect(xml.indexOf("<w:pgBorders")).toBeGreaterThan(xml.indexOf("<w:pgMar"));
    expect((await readDocx(saved)).context.page.border).toEqual({ kind: "double", size: 0.75, color: "#1F4E79", space: 24, fromText: false });
  });
});

describe("change case", () => {
  it("goes lower → UPPER → Title → lower, across formatting", async () => {
    const { context } = await readDocx(await blankDocx());
    const para = schema.nodes.paragraph.create(null, [schema.text("the qui"), schema.text("ck fox", [schema.marks.bold.create()])]);
    let state = createState(schema.nodes.doc.create(null, [para]), () => context);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1, state.doc.content.size - 1)));
    const run = () => changeCase()(state, (tr) => (state = state.apply(tr)));
    run();
    expect(state.doc.textContent).toBe("THE QUICK FOX");
    run();
    expect(state.doc.textContent).toBe("The Quick Fox");
    run();
    expect(state.doc.textContent).toBe("the quick fox");
    expect(state.doc.firstChild!.lastChild!.marks.map((m) => m.type.name)).toEqual(["bold"]);
  });
});

describe("templates", () => {
  it("each make a document that saves and opens again", async () => {
    for (const t of TEMPLATES) {
      const { context } = await readDocx(await blankDocx());
      const doc = templateDoc(t.kind, context);
      const again = await readDocx(await writeDocx(doc, context));
      expect(again.doc.textContent.length).toBeGreaterThan(20);
      if (t.kind === "article") expect(again.doc.firstChild!.attrs.style).toBe("Title");
    }
  });
});

describe("watermarks", () => {
  it("are saved in the header as Word's watermark shape, and found again", async () => {
    const { setWatermark, watermarkText } = await import("../src/watermark");
    const { doc, context } = await readDocx(await blankDocx());
    setWatermark(context, "DRAFT & CO");
    expect(watermarkText(context)).toBe("DRAFT & CO");
    const saved = await writeDocx(doc, context);
    const again = await readDocx(saved);
    expect(watermarkText(again.context)).toBe("DRAFT & CO");
    setWatermark(again.context, "COPY");
    setWatermark(again.context, null);
    expect(watermarkText(again.context)).toBeNull();
  });
});
