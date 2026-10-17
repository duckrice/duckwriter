import * as docx from "docx";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { ensureStyles } from "../src/docx/builtins";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { schema } from "../src/editor/schema";

async function withFootnotes(): Promise<Uint8Array> {
  const d = new docx.Document({
    footnotes: {
      1: { children: [new docx.Paragraph("Measured at 20 °C.")] },
      2: { children: [new docx.Paragraph("See the appendix.")] },
    },
    sections: [
      {
        children: [
          new docx.Paragraph({ children: [new docx.TextRun("The rate doubled"), new docx.FootnoteReferenceRun(1), new docx.TextRun(" and then fell"), new docx.FootnoteReferenceRun(2), new docx.TextRun(".")] }),
        ],
      },
    ],
  });
  return new Uint8Array(await docx.Packer.toBuffer(d));
}

const refs = (doc: import("prosemirror-model").Node) => {
  const ids: string[] = [];
  doc.descendants((n) => {
    if (n.type.name === "footnote") ids.push(n.attrs.id);
  });
  return ids;
};

describe("footnotes", () => {
  it("reads the marks in the text and the notes themselves", async () => {
    const { doc, context } = await readDocx(await withFootnotes());
    expect(refs(doc)).toEqual(["1", "2"]);
    expect(context.footnotes.notes.get("1")!.textContent.trim()).toBe("Measured at 20 °C.");
    expect(context.footnotes.notes.get("2")!.textContent.trim()).toBe("See the appendix.");
    expect(context.footnotes.special).toHaveLength(2);
  });

  it("leaves the footnotes alone when they weren't changed", async () => {
    const original = await withFootnotes();
    const { doc, context } = await readDocx(original);
    const before = await (await JSZip.loadAsync(original)).file("word/footnotes.xml")!.async("string");
    const after = await (await JSZip.loadAsync(await writeDocx(doc, context))).file("word/footnotes.xml")!.async("string");
    expect(after).toBe(before);
  });

  it("saves an edited note, and drops a note whose mark was deleted", async () => {
    const { doc, context } = await readDocx(await withFootnotes());
    context.footnotes.notes.set("1", schema.nodes.doc.create(null, schema.nodes.paragraph.create({ style: "FootnoteText" }, schema.text(" Measured at 25 °C."))));
    // Take out the second mark.
    let without = doc;
    doc.descendants((n, pos) => {
      if (n.type.name === "footnote" && n.attrs.id === "2") without = doc.replace(pos, pos + 1, doc.slice(pos, pos));
    });
    const again = await readDocx(await writeDocx(without, context));
    expect(refs(again.doc)).toEqual(["1"]);
    expect([...again.context.footnotes.notes.keys()]).toEqual(["1"]);
    expect(again.context.footnotes.notes.get("1")!.textContent.trim()).toBe("Measured at 25 °C.");
  });

  it("adds the first footnote to a document that had none, with Word's footnote styles", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    ensureStyles(context, "FootnoteText", "FootnoteReference");
    expect(context.styles.byId.get("FootnoteReference")!.run.vertAlign).toBe("superscript");
    const para = schema.nodes.paragraph.create(null, [schema.text("Water boils at 100 °C"), schema.nodes.footnote.create({ id: "1" })]);
    context.footnotes.notes.set("1", schema.nodes.doc.create(null, schema.nodes.paragraph.create({ style: "FootnoteText" }, schema.text(" At sea level."))));
    const saved = await writeDocx(schema.nodes.doc.create(null, para), context);
    void doc;
    const zip = await JSZip.loadAsync(saved);
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain("wordprocessingml.footnotes+xml");
    expect(await zip.file("word/settings.xml")!.async("string")).toContain('<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>');
    expect(await zip.file("word/styles.xml")!.async("string")).toContain('w:styleId="FootnoteText"');
    const notes = await zip.file("word/footnotes.xml")!.async("string");
    expect(notes).toContain('w:type="separator"');
    expect(notes).toContain('<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>');
    const again = await readDocx(saved);
    expect(refs(again.doc)).toEqual(["1"]);
    expect(again.context.footnotes.notes.get("1")!.textContent).toBe(" At sea level.");
  });
});
