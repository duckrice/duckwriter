import { readFileSync } from "node:fs";
import { join } from "node:path";
import JSZip from "jszip";
import { TextSelection, type EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import type { DocContext } from "../src/docx/model";
import { createState } from "../src/editor/editor";
import { addComment, addReply, changes, changesAt, commentAnchors, decide, decideAll, removeComment } from "../src/editor/review";
import { schema } from "../src/editor/schema";

const sample = () => new Uint8Array(readFileSync(join(__dirname, "..", "samples", "review.docx")));

async function open(bytes: Uint8Array) {
  const { doc, context } = await readDocx(bytes);
  return { state: createState(doc, () => context), ctx: context };
}

const body = async (bytes: Uint8Array) => {
  const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
  return xml.slice(xml.indexOf("<w:body>"), xml.indexOf("<w:sectPr"));
};
const part = async (bytes: Uint8Array, path: string) => (await JSZip.loadAsync(bytes)).file(path)?.async("string");
const paras = (state: EditorState) => state.doc.content.content.map((p) => p.textContent);

/** The text as it reads once every change is accepted. */
const visible = (state: EditorState) => {
  let out = "";
  state.doc.descendants((n) => {
    if (n.isText && !n.marks.some((m) => m.type === schema.marks.deletion)) out += n.text;
    if (n.isTextblock && out) out += out.endsWith("|") ? "" : "|";
    return true;
  });
  return out;
};

describe("tracked changes and comments in a Word file", () => {
  it("are shown, not accepted, when the file is opened", async () => {
    const { state, ctx } = await open(sample());
    expect(paras(state)[0]).toBe("The rate wasis proportional to the temperature.");
    expect(changes(state.doc).map((c) => `${c.kind}:${state.doc.textBetween(c.from, c.to)}:${c.author}`)).toEqual([
      "del:was:Ada Lovelace",
      "ins:is:Ada Lovelace",
      "ins:clearly :Max Planck",
      "del: This sentence goes.:Max Planck",
    ]);
    expect([...ctx.comments.items.values()].map((c) => [c.author, c.text])).toEqual([
      ["Ada Lovelace", "Is this the right figure?"],
      ["Max Planck", "Please cite the source.\nAny review paper will do."],
    ]);
    const a = commentAnchors(state.doc);
    expect(a.map((x) => state.doc.textBetween(x.from, x.to))).toEqual(["figure 3", "Earlier work found the same."]);
  });

  it("are saved as they were", async () => {
    const bytes = sample();
    const { state, ctx } = await open(bytes);
    const saved = await writeDocx(state.doc, ctx);
    expect(await body(saved)).toBe((await body(bytes)).replace(/<w:p>/g, "<w:p>"));
    expect(await part(saved, "word/comments.xml")).toBe(await part(bytes, "word/comments.xml"));
  });

  it("can be accepted or rejected, all at once", async () => {
    const { state } = await open(sample());
    const accepted = state.apply(decideAll(state, true)!);
    expect(paras(accepted)).toEqual(["The rate is proportional to the temperature.", "As shown in figure 3, the curve clearly flattens out.", "Earlier work found the same."]);
    expect(changes(accepted.doc)).toEqual([]);
    const rejected = state.apply(decideAll(state, false)!);
    expect(paras(rejected)).toEqual(["The rate was proportional to the temperature.", "As shown in figure 3, the curve flattens out.", "Earlier work found the same. This sentence goes."]);
  });

  it("can be accepted one at a time, where the cursor is", async () => {
    let { state } = await open(sample());
    const at = changes(state.doc)[2]; // "clearly "
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, at.from + 2)));
    expect(changesAt(state).map((c) => c.kind)).toEqual(["ins"]);
    state = state.apply(decide(state, changesAt(state), false)!);
    expect(paras(state)[1]).toBe("As shown in figure 3, the curve flattens out.");
    expect(changes(state.doc)).toHaveLength(3);
  });

  it("comments can be added, answered and removed, and Word's parts are written", async () => {
    let { state, ctx } = await open(sample());
    const pos = state.doc.child(0).nodeSize + 3; // in "As"
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
    const added = addComment(state, ctx, "Kari Nordmann");
    state = state.apply(added.tr);
    ctx.comments.items.get(added.id)!.text = "Which figure?";
    const anchors = commentAnchors(state.doc).find((a) => a.id === added.id)!;
    expect(state.doc.textBetween(anchors.from, anchors.to)).toBe("As");
    const reply = addReply(state, ctx, added.id, "Ola Nordmann")!;
    state = state.apply(reply.tr);
    ctx.comments.items.get(reply.id)!.text = "Figure 3.";
    ctx.comments.items.get("0")!.done = true;
    ctx.comments.version++;

    const saved = await writeDocx(state.doc, ctx);
    const comments = (await part(saved, "word/comments.xml"))!;
    expect(comments).toContain('w:author="Kari Nordmann"');
    expect(comments).toContain("Which figure?");
    expect(await body(saved)).toContain(`<w:commentRangeStart w:id="${added.id}"/>`);
    const again = await open(saved);
    const items = [...again.ctx.comments.items.values()];
    expect(items.find((c) => c.text === "Figure 3.")?.parent).toBe(added.id);
    expect(items.find((c) => c.id === "0")?.done).toBe(true);
    expect((await part(saved, "[Content_Types].xml"))!).toContain("commentsExtended");

    state = state.apply(removeComment(state, ctx, added.id));
    expect(commentAnchors(state.doc).map((a) => a.id)).toEqual(["0", "1"]);
  });
});

describe("tracking changes while typing", () => {
  async function tracked(text: string) {
    const { doc, context } = await readDocx(await blankDocx());
    const ctx: DocContext = { ...context, trackChanges: true };
    localStorage.setItem("duckwriter.author", "Kari");
    let state = createState(schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, schema.text(text))]), () => ctx);
    void doc;
    const put = (pos: number) => (state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos))));
    const type = (t: string) => (state = state.apply(state.tr.insertText(t)));
    const backspace = () => {
      const { from, empty } = state.selection;
      state = state.apply(empty ? state.tr.delete(from - 1, from) : state.tr.deleteSelection());
    };
    return { get state() { return state; }, put, type, backspace, ctx, set state(s: EditorState) { state = s; } };
  }

  it("marks typed text as inserted, and keeps deleted text, struck through", async () => {
    const t = await tracked("The rate was high.");
    t.put(13); // after "was"
    t.backspace();
    t.backspace();
    t.backspace();
    expect(t.state.doc.textContent).toBe("The rate was high.");
    expect(t.state.selection.from).toBe(10);
    t.type("is");
    expect(t.state.doc.textContent).toBe("The rate iswas high.");
    expect(changes(t.state.doc).map((c) => `${c.kind}:${t.state.doc.textBetween(c.from, c.to)}`)).toEqual(["ins:is", "del:was"]);
    expect(visible(t.state)).toBe("The rate is high.");
    // Taking back your own typing just removes it.
    t.backspace();
    expect(t.state.doc.textContent).toBe("The rate iwas high.");
  });

  it("tracks new paragraphs and joined ones, and accepting makes it final", async () => {
    const t = await tracked("One two");
    t.put(4);
    t.state = t.state.apply(t.state.tr.split(4));
    expect(t.state.doc.childCount).toBe(2);
    expect(t.state.doc.child(0).attrs.change?.kind).toBe("ins");
    // Backspace at the start of the new paragraph: it was ours, so it simply joins again.
    t.put(t.state.doc.child(0).nodeSize + 1);
    t.state = t.state.apply(t.state.tr.delete(t.state.doc.child(0).nodeSize - 1, t.state.doc.child(0).nodeSize + 1));
    expect(t.state.doc.childCount).toBe(1);
    expect(t.state.doc.child(0).attrs.change).toBeNull();
    const accepted = t.state.apply(decideAll(t.state, true) ?? t.state.tr);
    expect(accepted.doc.textContent).toBe("One two");
  });

  it("tracks formatting, which can be rejected back to how it was", async () => {
    const t = await tracked("The rate was high.");
    t.state = t.state.apply(t.state.tr.addMark(5, 9, schema.marks.bold.create()));
    expect(changes(t.state.doc).map((c) => c.kind + ":" + t.state.doc.textBetween(c.from, c.to))).toEqual(["fmt:rate"]);
    const saved = await writeDocx(t.state.doc, t.ctx);
    expect(await body(saved)).toContain('<w:rPr><w:b/><w:bCs/><w:rPrChange w:id="');
    expect(await body(saved)).toMatch(/w:author="Kari" w:date="[^"]+"><w:rPr><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t xml:space="preserve">rate</);
    const again = await open(saved);
    expect(changes(again.state.doc).map((c) => c.kind)).toEqual(["fmt"]);
    const rejected = t.state.apply(decideAll(t.state, false, t.ctx)!);
    expect(rejected.doc.rangeHasMark(5, 9, schema.marks.bold)).toBe(false);
    expect(changes(rejected.doc)).toEqual([]);
    const accepted = t.state.apply(decideAll(t.state, true, t.ctx)!);
    expect(accepted.doc.rangeHasMark(5, 9, schema.marks.bold)).toBe(true);
    expect(changes(accepted.doc)).toEqual([]);
  });

  it("changes aren't tracked when it is off", async () => {
    const t = await tracked("Text");
    t.ctx.trackChanges = false;
    t.put(5);
    t.backspace();
    expect(t.state.doc.textContent).toBe("Tex");
  });
});
