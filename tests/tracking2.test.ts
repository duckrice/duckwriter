import JSZip from "jszip";
import { TextSelection, type EditorState } from "prosemirror-state";
import { addRowAfter, deleteRow } from "prosemirror-tables";
import { describe, expect, it } from "vitest";
import type { DocContext } from "../src/docx/model";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { createState } from "../src/editor/editor";
import { changes, decideAll } from "../src/editor/review";
import { schema } from "../src/editor/schema";

async function tracked(...blocks: ReturnType<typeof schema.nodes.paragraph.create>[]) {
  const { context } = await readDocx(await blankDocx());
  const ctx: DocContext = { ...context, trackChanges: true };
  localStorage.setItem("duckwriter.author", "Kari");
  const state = createState(schema.nodes.doc.create(null, blocks), () => ctx);
  return { state, ctx };
}

const para = (text: string, attrs: Record<string, unknown> = {}) => schema.nodes.paragraph.create(attrs, text ? schema.text(text) : null);

describe("tracking paragraph formatting", () => {
  it("keeps what it was, writes Word's pPrChange, and can be rejected back", async () => {
    let { state, ctx } = await tracked(para("Centre me", { spaceAfter: 6 }));
    state = state.apply(state.tr.setNodeMarkup(0, null, { ...state.doc.firstChild!.attrs, align: "center", style: "Heading1" }));
    const c = changes(state.doc);
    expect(c.map((x) => [x.kind, x.props])).toEqual([["fmt", true]]);
    const xml = await (await JSZip.loadAsync(await writeDocx(state.doc, ctx))).file("word/document.xml")!.async("string");
    expect(xml).toMatch(/<w:pPrChange w:id="\d+" w:author="Kari" w:date="[^"]+"><w:pPr><w:spacing w:after="120"\/><\/w:pPr><\/w:pPrChange>/);
    const again = await readDocx(await writeDocx(state.doc, ctx));
    expect(changes(again.doc).map((x) => x.kind)).toEqual(["fmt"]);
    const rejected = state.apply(decideAll(state, false, ctx)!);
    expect(rejected.doc.firstChild!.attrs).toMatchObject({ align: null, style: null, spaceAfter: 6, propsChange: null });
    const accepted = state.apply(decideAll(state, true, ctx)!);
    expect(accepted.doc.firstChild!.attrs).toMatchObject({ align: "center", style: "Heading1", propsChange: null });
  });
});

describe("tracking table rows", () => {
  const cell = (text: string) => schema.nodes.table_cell.create(null, para(text));
  const table = () => schema.nodes.table.create(null, [schema.nodes.table_row.create(null, [cell("a"), cell("b")]), schema.nodes.table_row.create(null, [cell("c"), cell("d")])]);
  const inFirstCell = (state: EditorState) => {
    let at = -1;
    state.doc.descendants((n, p) => {
      if (at < 0 && n.isText && n.text === "a") at = p;
    });
    return state.apply(state.tr.setSelection(TextSelection.create(state.doc, at)));
  };
  const rows = (state: EditorState) => {
    const out: string[] = [];
    state.doc.firstChild!.forEach((r) => out.push(r.textContent));
    return out;
  };

  it("marks a new row as inserted, which rejecting removes", async () => {
    let { state, ctx } = await tracked(para(""));
    // Tracking applies from here on.
    state = state.apply(state.tr.replaceWith(0, state.doc.content.size, [table(), para("")]).setMeta("noTrack", true));
    state = inFirstCell(state);
    addRowAfter(state, (tr) => (state = state.apply(tr)));
    expect(rows(state)).toEqual(["ab", "", "cd"]);
    expect(state.doc.firstChild!.child(1).attrs.trPr).toMatch(/^<w:ins w:id="\d+" w:author="Kari"/);
    const xml = await (await JSZip.loadAsync(await writeDocx(state.doc, ctx))).file("word/document.xml")!.async("string");
    expect(xml).toMatch(/<w:trPr><w:ins w:id="\d+" w:author="Kari"/);
    const rejected = state.apply(decideAll(state, false, ctx)!);
    expect(rows(rejected)).toEqual(["ab", "cd"]);
  });

  it("keeps a deleted row, marked, until the deletion is accepted", async () => {
    let { state, ctx } = await tracked(para(""));
    state = state.apply(state.tr.replaceWith(0, state.doc.content.size, [table(), para("")]).setMeta("noTrack", true));
    state = inFirstCell(state);
    deleteRow(state, (tr) => (state = state.apply(tr)));
    expect(rows(state)).toEqual(["ab", "cd"]);
    expect(state.doc.firstChild!.child(0).attrs.trPr).toMatch(/^<w:del /);
    expect(changes(state.doc).some((c) => c.row && c.kind === "del")).toBe(true);
    const accepted = state.apply(decideAll(state, true, ctx)!);
    expect(rows(accepted)).toEqual(["cd"]);
    const rejected = state.apply(decideAll(state, false, ctx)!);
    expect(rows(rejected)).toEqual(["ab", "cd"]);
    expect(rejected.doc.firstChild!.child(0).attrs.trPr).toBe("");
  });
});
