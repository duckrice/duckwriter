import { describe, expect, it } from "vitest";
import { EditorState } from "prosemirror-state";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { schema } from "../src/editor/schema";
import { insertSectionBreak, removeSectionBreak, sectionsOf, setSectionNumbering, setSectionStart } from "../src/editor/commands";
import { formatMarksPlugin, formatMarksOn } from "../src/editor/formatmarks";
import { TextSelection } from "prosemirror-state";

const para = (text: string) => schema.nodes.paragraph.create(null, schema.text(text));

async function report() {
  const { context } = await readDocx(await blankDocx());
  const doc = schema.nodes.doc.create(null, [para("My report"), para("Foreword"), para("Chapter 1")]);
  let state = EditorState.create({ doc, plugins: [formatMarksPlugin] });
  // Section breaks before "Foreword" and before "Chapter 1", as from the Breaks menu.
  for (const text of ["Foreword", "Chapter 1"]) {
    let at = -1;
    state.doc.descendants((n, pos) => {
      if (n.textContent === text && n.isTextblock) at = pos + 1;
    });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, at)));
    insertSectionBreak(context, "nextPage")(state, (tr) => (state = state.apply(tr)));
  }
  return { context, get: () => state, run: (cmd: (s: EditorState, d: (tr: import("prosemirror-state").Transaction) => void) => boolean) => cmd(state, (tr) => (state = state.apply(tr))) };
}

describe("the page numbers window", () => {
  it("lists every section with its first words", async () => {
    const r = await report();
    const sections = sectionsOf(r.get().doc, r.context);
    expect(sections.map((s) => s.preview)).toEqual(["My report", "Foreword", "Chapter 1"]);
    expect(sections.map((s) => s.start)).toEqual(["nextPage", "nextPage", "nextPage"]);
    expect(sections[2].end).toBe(-1);
  });

  it("sets each section's numbering, start and title page on its own", async () => {
    const r = await report();
    let s = sectionsOf(r.get().doc, r.context);
    r.run(setSectionNumbering(r.context, s[0].end, { format: "decimal", start: null, firstBlank: true }));
    s = sectionsOf(r.get().doc, r.context);
    r.run(setSectionNumbering(r.context, s[1].end, { format: "lowerRoman", start: 1, firstBlank: false }));
    s = sectionsOf(r.get().doc, r.context);
    r.run(setSectionNumbering(r.context, s[2].end, { format: "decimal", start: 1, firstBlank: false }));
    s = sectionsOf(r.get().doc, r.context);
    r.run(setSectionStart(r.context, s[2].end, "oddPage"));
    const after = sectionsOf(r.get().doc, r.context);
    expect(after.map((x) => x.numbering)).toEqual([
      { format: "decimal", start: null, firstBlank: true },
      { format: "lowerRoman", start: 1, firstBlank: false },
      { format: "decimal", start: 1, firstBlank: false },
    ]);
    expect(after[2].start).toBe("oddPage");
    expect(r.context.sectPr).toContain('w:type w:val="oddPage"');
  });

  it("takes a section break out: its text joins the next section and takes its numbering", async () => {
    const r = await report();
    let s = sectionsOf(r.get().doc, r.context);
    r.run(setSectionNumbering(r.context, s[1].end, { format: "lowerRoman", start: 1, firstBlank: false }));
    s = sectionsOf(r.get().doc, r.context);
    r.run(removeSectionBreak(s[0].end));
    s = sectionsOf(r.get().doc, r.context);
    expect(s).toHaveLength(2);
    expect(s[0].preview).toBe("My report");
    expect(s[0].numbering.format).toBe("lowerRoman");
  });
});

describe("show formatting (¶)", () => {
  it("marks paragraph ends, spaces and section breaks, only when on", async () => {
    const r = await report();
    const decorations = (state: EditorState) => formatMarksPlugin.props.decorations!.call(formatMarksPlugin, state);
    expect(formatMarksOn(r.get())).toBe(false);
    expect(decorations(r.get())).toBeNull();
    const on = r.get().apply(r.get().tr.setMeta(formatMarksPlugin, true));
    const set = decorations(on) as import("prosemirror-view").DecorationSet;
    const found = set.find();
    // A ¶ for each of the paragraphs (three, and the two holding the breaks), a dot for each space.
    expect(found.filter((d) => (d as unknown as { type: { side?: number } }).type.side === 1)).toHaveLength(5);
    expect(found.filter((d) => d.from + 1 === d.to)).toHaveLength(2);
  });
});
