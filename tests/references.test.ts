import JSZip from "jszip";
import type { Node } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { ensureStyles } from "../src/docx/builtins";
import type { DocContext } from "../src/docx/model";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { createState } from "../src/editor/editor";
import { fieldResults } from "../src/editor/fields";
import { listLabels } from "../src/editor/lists";
import { schema } from "../src/editor/schema";
import { toggleHeadingNumbers } from "../src/ui/references";

const text = (t: string) => schema.text(t);
const p = (attrs: Record<string, unknown>, ...content: Node[]) => schema.nodes.paragraph.create(attrs, content);
const seq = (label: string) => schema.nodes.field.create({ instr: ` SEQ ${label} \\* ARABIC `, result: "?" });
const bookmark = (id: string, name?: string) =>
  schema.nodes.raw_inline.create({ xml: name ? `<w:bookmarkStart w:id="${id}" w:name="${name}"/>` : `<w:bookmarkEnd w:id="${id}"/>`, hidden: true });
const caption = (label: string, id: string, name: string, words: string) => p({ style: "Caption" }, bookmark(id, name), text(`${label} `), seq(label), bookmark(id), text(`: ${words}`));
const ref = (name: string) => schema.nodes.field.create({ instr: ` REF ${name} \\h `, result: "" });

function results(doc: Node): string[] {
  const map = fieldResults(doc);
  const out: string[] = [];
  doc.descendants((n, pos) => {
    if (n.type === schema.nodes.field) out.push(map.get(pos) ?? n.attrs.result);
  });
  return out;
}

describe("captions and cross-references", () => {
  it("numbers figures and tables separately, in order", () => {
    const doc = schema.nodes.doc.create(null, [
      caption("Figure", "1", "_Ref1", "Setup"),
      caption("Table", "2", "_Ref2", "Results"),
      caption("Figure", "3", "_Ref3", "Graph"),
    ]);
    expect(results(doc)).toEqual(["1", "1", "2"]);
  });

  it("a reference shows the current label and number of what it points to", () => {
    const doc = schema.nodes.doc.create(null, [
      p({}, text("As "), ref("_Ref3"), text(" shows…")),
      caption("Figure", "1", "_Ref1", "Setup"),
      caption("Figure", "3", "_Ref3", "Graph"),
    ]);
    expect(results(doc)[0]).toBe("Figure 2");
  });

  it("keeps the numbers right while editing", async () => {
    const { context } = await readDocx(await blankDocx());
    let state = createState(schema.nodes.doc.create(null, [caption("Figure", "1", "_Ref1", "A"), p({}, ref("_Ref1"))]), () => context);
    // Another figure before it: the first becomes Figure 2, and so does the reference.
    state = state.apply(state.tr.insert(0, caption("Figure", "9", "_Ref9", "New")));
    const shown: string[] = [];
    state.doc.descendants((n) => {
      if (n.type === schema.nodes.field) shown.push(n.attrs.result);
    });
    expect(shown).toEqual(["1", "2", "Figure 2"]);
  });
});

describe("numbered headings", () => {
  it("numbers headings 1, 1.1, 1.1.1 through their styles, and saves it so Word sees it", async () => {
    const { context } = await readDocx(await blankDocx());
    const doc = schema.nodes.doc.create(null, [
      p({ style: "Heading1" }, text("Introduction")),
      p({ style: "Heading2" }, text("Background")),
      p({ style: "Heading2" }, text("Aim")),
      p({ style: "Heading1" }, text("Method")),
      p({ style: "Heading3" }, text("Detail")),
    ]);
    const state = createState(doc, () => context);
    const view = { state, dispatch: () => {} } as unknown as import("prosemirror-view").EditorView;
    toggleHeadingNumbers(view, context);
    expect([...listLabels(context, doc).values()].map((l) => l.label)).toEqual(["1", "1.1", "1.2", "2", "2.0.1"]);

    const saved = await JSZip.loadAsync(await writeDocx(doc, context));
    const styles = await saved.file("word/styles.xml")!.async("string");
    expect(/w:styleId="Heading2"[\s\S]*?<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="\d+"\/><\/w:numPr>[\s\S]*?<\/w:style>/.test(styles)).toBe(true);
    const again = await readDocx(await saved.generateAsync({ type: "uint8array" }));
    expect([...listLabels(again.context, again.doc).values()].map((l) => l.label)).toEqual(["1", "1.1", "1.2", "2", "2.0.1"]);
  });
});

describe("table of contents", () => {
  async function withToc(): Promise<{ doc: Node; context: DocContext }> {
    const { context } = await readDocx(await blankDocx());
    ensureStyles(context, "TOCHeading", "TOC1", "TOC2", "TOC3");
    const entries = [
      { level: 1, text: "Introduction", page: 1, pos: 0 },
      { level: 2, text: "Background", page: 2, pos: 0 },
    ];
    const doc = schema.nodes.doc.create(null, [schema.nodes.toc.create({ entries }), p({ style: "Heading1" }, text("Introduction")), p({ style: "Heading2" }, text("Background"))]);
    return { doc, context };
  }

  it("saves as a Word table of contents field, with the entries and page numbers", async () => {
    const { doc, context } = await withToc();
    const xml = await (await JSZip.loadAsync(await writeDocx(doc, context))).file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:pStyle w:val="TOCHeading"/>');
    expect(xml).toContain('<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>');
    expect(xml).toMatch(/<w:pStyle w:val="TOC2"\/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9026"\/><\/w:tabs><\/w:pPr><w:r><w:t xml:space="preserve">Background<\/w:t><\/w:r><w:r><w:tab\/><\/w:r><w:r><w:t>2<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"\/>/);
  });

  it("opens again as a table of contents", async () => {
    const { doc, context } = await withToc();
    const again = await readDocx(await writeDocx(doc, context));
    const toc = again.doc.firstChild!;
    expect(toc.type.name).toBe("toc");
    expect(toc.attrs.title).toBe("Contents");
    expect(toc.attrs.entries.map((e: { text: string; page: number; level: number }) => [e.level, e.text, e.page])).toEqual([
      [1, "Introduction", 1],
      [2, "Background", 2],
    ]);
    expect(again.doc.childCount).toBe(3);
  });
});
