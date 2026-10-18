import JSZip from "jszip";
import type { EditorState, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { createState } from "../src/editor/editor";
import { schema } from "../src/editor/schema";
import { convertNotes } from "../src/ui/endnotes";

const part = async (bytes: Uint8Array, path: string) => (await JSZip.loadAsync(bytes)).file(path)?.async("string");

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

const note = (text: string, style: string) => schema.nodes.doc.create(null, schema.nodes.paragraph.create({ style }, schema.text(text)));

describe("footnotes to endnotes and back", () => {
  it("moves every note, keeps its text, and saves the right parts", async () => {
    const { context: ctx } = await readDocx(await blankDocx());
    ctx.footnotes.notes.set("1", note("First note.", "FootnoteText"));
    ctx.footnotes.notes.set("2", note("Second note.", "FootnoteText"));
    const ref = schema.marks.charStyle.create({ id: "FootnoteReference" });
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(null, [schema.text("One"), schema.nodes.footnote.create({ id: "1" }, null, [ref])]),
      schema.nodes.paragraph.create(null, [schema.text("Two"), schema.nodes.footnote.create({ id: "2" }, null, [ref])]),
    ]);
    const view = fakeView(createState(doc, () => ctx));

    expect(convertNotes(view, ctx, "endnote")).toBe(2);
    const marks: string[] = [];
    view.state.doc.descendants((n) => {
      if (n.type === schema.nodes.endnote) marks.push(n.attrs.id);
      expect(n.type).not.toBe(schema.nodes.footnote);
    });
    expect(marks).toHaveLength(2);
    expect(view.state.doc.lastChild!.type).toBe(schema.nodes.endnotes);
    expect(marks.map((id) => ctx.endnotes.notes.get(id)!.textContent)).toEqual(["First note.", "Second note."]);
    expect(ctx.endnotes.notes.get(marks[0])!.firstChild!.attrs.style).toBe("EndnoteText");

    let saved = await writeDocx(view.state.doc, ctx);
    const body = (await part(saved, "word/document.xml"))!;
    expect(body).toContain("<w:endnoteReference");
    expect(body).not.toContain("<w:footnoteReference");
    expect(await part(saved, "word/endnotes.xml")).toContain("Second note.");

    // And back: footnotes again, and no list of endnotes left at the end.
    expect(convertNotes(view, ctx, "footnote")).toBe(2);
    expect(view.state.doc.lastChild!.type).toBe(schema.nodes.paragraph);
    const back: string[] = [];
    view.state.doc.descendants((n) => {
      if (n.type === schema.nodes.footnote) {
        back.push(ctx.footnotes.notes.get(n.attrs.id)!.textContent);
        expect(n.marks.some((m) => m.attrs.id === "FootnoteReference")).toBe(true);
      }
    });
    expect(back).toEqual(["First note.", "Second note."]);
    saved = await writeDocx(view.state.doc, ctx);
    expect(await part(saved, "word/document.xml")).not.toContain("<w:endnoteReference");
    const again = await readDocx(saved);
    const ids: string[] = [];
    again.doc.descendants((n) => {
      if (n.type === schema.nodes.footnote) ids.push(again.context.footnotes.notes.get(n.attrs.id)!.textContent);
    });
    expect(ids).toEqual(["First note.", "Second note."]);
  });
});
