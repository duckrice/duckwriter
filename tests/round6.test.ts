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
import { isRis, parseRis, toBibtex } from "../src/docx/ris";
import { parseBibtex } from "../src/docx/bibliography";
import { texEscape, toLatex } from "../src/convert/latex";
import { citationInstr } from "../src/editor/citations";

describe("RIS and BibTeX", () => {
  const ris = [
    "TY  - JOUR",
    "AU  - Smith, Anna",
    "AU  - Jones, Bob",
    "TI  - Rates of things",
    "T2  - Journal of Ducks",
    "PY  - 2020/05/01",
    "VL  - 3",
    "IS  - 2",
    "SP  - 1",
    "EP  - 10",
    "DO  - https://doi.org/10.1/x",
    "ER  - ",
    "TY  - BOOK",
    "AU  - Hansen, Kari",
    "TI  - A book",
    "PY  - 2019",
    "PB  - Duck Press",
    "CY  - Oslo",
    "ER  - ",
  ].join("\r\n");

  it("reads RIS records", () => {
    expect(isRis(ris)).toBe(true);
    expect(isRis("@article{a, title={x}}")).toBe(false);
    const [a, b] = parseRis(ris);
    expect(a).toMatchObject({ type: "JournalArticle", title: "Rates of things", container: "Journal of Ducks", year: "2020", volume: "3", issue: "2", pages: "1–10", doi: "10.1/x" });
    expect(a.authors).toEqual([{ last: "Smith", first: "Anna" }, { last: "Jones", first: "Bob" }]);
    expect(b).toMatchObject({ type: "Book", publisher: "Duck Press", city: "Oslo", year: "2019" });
  });

  it("writes BibTeX that reads back the same", () => {
    const sources = parseRis(ris).map((s, i) => ({ ...s, tag: ["Smi20", "Han19"][i], title: i ? "Ducks & DNA" : s.title }));
    const bib = toBibtex(sources);
    expect(bib).toContain("@article{Smi20,");
    expect(bib).toContain("pages = {1--10}");
    expect(bib).toContain("title = {{Ducks \\& DNA}}");
    const back = parseBibtex(bib);
    expect(back.map((s) => s.tag)).toEqual(["Smi20", "Han19"]);
    expect(back[0]).toMatchObject({ title: "Rates of things", container: "Journal of Ducks", pages: "1–10", doi: "10.1/x" });
    expect(back[1]).toMatchObject({ type: "Book", title: "Ducks & DNA", publisher: "Duck Press" });
  });
});

describe("LaTeX", () => {
  it("escapes what TeX treats specially", () => {
    expect(texEscape("50% of $5 & #1_a {x} ~ ^ \\")).toBe("50\\% of \\$5 \\& \\#1\\_a \\{x\\} \\textasciitilde{} \\textasciicircum{} \\textbackslash{}");
  });

  it("makes sections, lists, tables, equations, footnotes and citations", async () => {
    const { context: ctx } = await readDocx(await blankDocx());
    ctx.bibliography.sources.push(...parseRis("TY  - JOUR\nAU  - Smith, Anna\nTI  - Rates\nPY  - 2020\nER  - ").map((s) => ({ ...s, tag: "Smi20" })));
    ctx.footnotes.notes.set("1", note("A note.", "FootnoteText"));
    const bold = schema.marks.bold.create({ on: true });
    const cite = schema.nodes.field.create({ instr: citationInstr({ tags: ["Smi20"], page: "" }), result: "(Smith, 2020)" });
    const p = (content: Parameters<typeof schema.nodes.paragraph.create>[1], style: string | null = null) => schema.nodes.paragraph.create({ style }, content);
    const cell = (t: string) => schema.nodes.table_cell.create(null, p(schema.text(t)));
    const doc = schema.nodes.doc.create(null, [
      p(schema.text("My paper"), "Title"),
      p(schema.text("Introduction"), "Heading1"),
      p([schema.text("Rates rise "), schema.text("fast", [bold]), schema.text(" "), cite, schema.text(" at 50%."), schema.nodes.footnote.create({ id: "1" })]),
      p([schema.text("So "), schema.nodes.math.create({ latex: "E = mc^2", display: false })]),
      p(schema.text("Methods"), "Heading2"),
      schema.nodes.table.create(null, [schema.nodes.table_row.create(null, [cell("a"), cell("b")]), schema.nodes.table_row.create(null, [cell("1"), cell("2")])]),
    ]);
    const { tex, bib } = await toLatex(doc, ctx, "paper");
    expect(tex).toContain("\\title{My paper}");
    expect(tex).toContain("\\maketitle");
    expect(tex).toContain("\\section{Introduction}");
    expect(tex).toContain("\\subsection{Methods}");
    expect(tex).toContain("Rates rise \\textbf{fast} \\cite{Smi20} at 50\\%.\\footnote{A note.}");
    expect(tex).toContain("So $E = mc^2$");
    expect(tex).toContain("\\begin{tabular}{|l|l|}");
    expect(tex).toContain("a & b \\\\ \\hline");
    expect(tex).toContain("\\bibliography{paper}");
    expect(tex.trim().endsWith("\\end{document}")).toBe(true);
    expect(bib).toContain("@article{Smi20,");
  });
});

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
