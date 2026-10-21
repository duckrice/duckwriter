import JSZip from "jszip";
import type { Node } from "prosemirror-model";
import { AllSelection, EditorState, TextSelection, type Command } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import type { DocContext } from "../src/docx/model";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { formattingAt, newParagraph, setMark, setStyle, toggleFormat } from "../src/editor/commands";
import { createState } from "../src/editor/editor";
import { listLabels, toggleList, changeListLevel } from "../src/editor/lists";
import { schema } from "../src/editor/schema";

async function blank() {
  const { context } = await readDocx(await blankDocx());
  return context;
}

const p = (text: string, attrs: Record<string, unknown> = {}) => schema.nodes.paragraph.create(attrs, text ? schema.text(text) : null);

function stateWith(ctx: DocContext, paras: Node[]) {
  return createState(schema.nodes.doc.create(null, paras), () => ctx);
}

function run(state: EditorState, command: Command): EditorState {
  let next = state;
  command(state, (tr) => (next = state.apply(tr)));
  return next;
}

/** Selects from `from` to `to` (document positions). */
const select = (state: EditorState, from: number, to = from) => state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to)));

describe("formatting", () => {
  it("bold in a bold heading turns bold off, and back on", async () => {
    const ctx = await blank();
    ctx.styles.byId.get("Heading1")!.run.bold = true;
    let state = stateWith(ctx, [p("Heading text", { style: "Heading1" })]);
    state = select(state, 1, 8);
    expect(formattingAt(ctx, state).run.bold).toBe(true);
    state = run(state, toggleFormat(ctx, "bold"));
    const marks = state.doc.firstChild!.firstChild!.marks;
    expect(marks.map((m) => [m.type.name, m.attrs.on])).toEqual([["bold", false]]);
    state = run(state, toggleFormat(ctx, "bold"));
    // Back to what the style says: no direct formatting left.
    expect(state.doc.firstChild!.firstChild!.marks).toHaveLength(0);
  });

  it("knows the font and size at the cursor from the style and direct formatting", async () => {
    const ctx = await blank();
    let state = stateWith(ctx, [p("Some text")]);
    state = select(state, 2);
    expect(formattingAt(ctx, state).run).toMatchObject({ font: "Calibri", size: 11 });
    state = select(state, 1, 5);
    state = run(state, setMark(schema.marks.font, { family: "Georgia" }));
    state = select(state, 2);
    expect(formattingAt(ctx, state).run.font).toBe("Georgia");
  });

  it("Enter after a heading starts a normal paragraph; elsewhere it keeps the paragraph's formatting", async () => {
    const ctx = await blank();
    let state = stateWith(ctx, [p("Title here", { style: "Heading1" })]);
    state = select(state, 11);
    state = run(state, newParagraph(ctx));
    expect(state.doc.childCount).toBe(2);
    expect(state.doc.child(1).attrs.style).toBeNull();

    state = stateWith(ctx, [p("Centred", { align: "center" })]);
    state = select(state, 8);
    state = run(state, newParagraph(ctx));
    expect(state.doc.child(1).attrs.align).toBe("center");
  });

  it("choosing a style clears direct paragraph formatting", async () => {
    const ctx = await blank();
    let state = stateWith(ctx, [p("Text", { align: "right", spaceAfter: 30 })]);
    state = select(state, 2);
    state = run(state, setStyle(ctx, "Heading2"));
    expect(state.doc.firstChild!.attrs).toMatchObject({ style: "Heading2", align: null, spaceAfter: null });
  });
});

describe("lists", () => {
  it("numbers items, nests levels, and continues a list after a nested item", async () => {
    const ctx = await blank();
    let state = stateWith(ctx, [p("One"), p("Two"), p("Three"), p("Four")]);
    state = select(state, 1, state.doc.content.size - 1);
    state = run(state, toggleList(ctx, "number"));
    // Make "Two" and "Three" a level deeper.
    state = select(state, 7, 13);
    state = run(state, changeListLevel(ctx, 1));
    const labels = [...listLabels(ctx, state.doc).values()].map((l) => l.label);
    expect(labels).toEqual(["1.", "a.", "b.", "2."]);
  });

  it("switches bullets off again, and a new list after a gap starts at 1", async () => {
    const ctx = await blank();
    let state = stateWith(ctx, [p("A"), p("B"), p("gap"), p("C")]);
    state = select(state, 1, 5);
    state = run(state, toggleList(ctx, "number"));
    state = select(state, state.doc.content.size - 1);
    state = run(state, toggleList(ctx, "number"));
    expect([...listLabels(ctx, state.doc).values()].map((l) => l.label)).toEqual(["1.", "2.", "1."]);
    state = select(state, 1, 5);
    state = run(state, toggleList(ctx, "number"));
    expect([...listLabels(ctx, state.doc).values()].map((l) => l.label)).toEqual(["1."]);
  });

  it("Enter at the end of a list item starts the next item", async () => {
    const ctx = await blank();
    let state = stateWith(ctx, [p("Item")]);
    state = run(select(state, 1, 5), toggleList(ctx, "bullet"));
    state = run(select(state, 5), newParagraph(ctx));
    expect(state.doc.child(1).attrs).toMatchObject({ style: "ListParagraph", numId: state.doc.child(0).attrs.numId });
  });

  it("Enter in an empty list item ends the list", async () => {
    const ctx = await blank();
    let state = stateWith(ctx, [p("Item"), p("")]);
    state = select(state, 1, 8);
    state = run(state, toggleList(ctx, "bullet"));
    state = select(state, 7);
    state = run(state, newParagraph(ctx));
    expect(state.doc.child(1).attrs.numId).toBeNull();
    expect(listLabels(ctx, state.doc).size).toBe(1);
  });
});

describe("saving edits", () => {
  it("saves new lists and links in a way that opens again the same", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    let state = createState(doc, () => context);
    state = state.apply(state.tr.insertText("Visit the site", 1));
    state = run(select(state, 1, 15), toggleList(context, "bullet"));
    state = run(select(state, 11, 15), setMark(schema.marks.link, { href: "https://duckrice.example/" }));
    state = run(select(state, 1, 6), toggleFormat(context, "bold"));

    const saved = await writeDocx(state.doc, context);
    const zip = await JSZip.loadAsync(saved);
    expect(zip.file("word/numbering.xml")).not.toBeNull();
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain("/word/numbering.xml");
    expect(await zip.file("word/_rels/document.xml.rels")!.async("string")).toContain('Target="https://duckrice.example/" TargetMode="External"');

    const again = await readDocx(saved);
    // The link gets its relationship id when saved.
    const withoutIds = (doc: Node) => JSON.stringify(doc.toJSON()).replace(/"rid":("[^"]*"|null)/g, "");
    expect(withoutIds(again.doc)).toBe(withoutIds(state.doc));
    expect([...listLabels(again.context, again.doc).values()].map((l) => l.label)).toEqual(["•"]);

    // Saving twice doesn't add the list or the link twice.
    const twice = await JSZip.loadAsync(await writeDocx(state.doc, context));
    const numbering = await twice.file("word/numbering.xml")!.async("string");
    expect(numbering.match(/<w:abstractNum /g)).toHaveLength(1);
    const rels = await twice.file("word/_rels/document.xml.rels")!.async("string");
    expect(rels.match(/duckrice\.example/g)).toHaveLength(1);
  });
});

describe("the paragraph mark keeps formatting, as in Word", () => {
  const typeText = (state: EditorState, text: string) => state.apply(state.tr.insertText(text));
  const fontAt = (state: EditorState, pos: number) => schema.marks.font.isInSet(state.doc.resolve(pos).marks())?.attrs.family;

  it("an empty line keeps the font everything was changed to, even after clicking away", async () => {
    const ctx = await blank();
    let state = stateWith(ctx, [p("First"), p(""), p("Third")]);
    state = state.apply(state.tr.setSelection(new AllSelection(state.doc)));
    state = run(state, setMark(schema.marks.font, { family: "Times New Roman" }));
    state = run(state, setMark(schema.marks.size, { pt: 12 }));
    // Somewhere else, then into the empty line.
    state = select(state, 2);
    state = select(state, 8);
    expect(state.selection.$from.parent.content.size).toBe(0);
    state = typeText(state, "new");
    const typed = state.doc.child(1).firstChild!;
    expect(schema.marks.font.isInSet(typed.marks)?.attrs.family).toBe("Times New Roman");
    expect(schema.marks.size.isInSet(typed.marks)?.attrs.pt).toBe(12);
  });

  it("Enter carries the font on, also from a line that's still empty", async () => {
    const ctx = await blank();
    let state = stateWith(ctx, [p("Text")]);
    state = state.apply(state.tr.setSelection(new AllSelection(state.doc)));
    state = run(state, setMark(schema.marks.font, { family: "Times New Roman" }));
    state = select(state, 5);
    state = run(state, newParagraph(ctx));
    state = run(state, newParagraph(ctx));
    // Away and back: the new empty lines still know their font.
    state = select(state, 2);
    state = select(state, state.doc.content.size - 1);
    state = typeText(state, "more");
    expect(fontAt(state, state.doc.content.size - 2)).toBe("Times New Roman");
  });

  it("is saved to Word and read back", async () => {
    const ctx = await blank();
    const tnr = schema.marks.font.create({ family: "Times New Roman" });
    const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create({ mark: [tnr.toJSON()] })]);
    const bytes = await writeDocx(doc, ctx);
    const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
    expect(xml).toMatch(/<w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman"/);
    const back = await readDocx(bytes);
    expect(back.doc.firstChild!.attrs.mark?.[0]?.attrs.family).toBe("Times New Roman");
  });
});
