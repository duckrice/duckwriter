import * as docx from "docx";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { schema } from "../src/editor/schema";

async function withHeaderAndFooter(): Promise<Uint8Array> {
  const d = new docx.Document({
    sections: [
      {
        properties: { titlePage: true },
        headers: {
          default: new docx.Header({ children: [new docx.Paragraph({ children: [new docx.TextRun("Lab report — group 4")] })] }),
          first: new docx.Header({ children: [new docx.Paragraph("Cover")] }),
        },
        footers: {
          default: new docx.Footer({
            children: [new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, children: [new docx.TextRun({ children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES] })] })],
          }),
        },
        children: [new docx.Paragraph("Body text")],
      },
    ],
  });
  return new Uint8Array(await docx.Packer.toBuffer(d));
}

const fieldsIn = (doc: import("prosemirror-model").Node) => {
  const out: string[] = [];
  doc.descendants((n) => {
    if (n.type.name === "field") out.push(n.attrs.instr.trim().split(/\s+/)[0]);
  });
  return out;
};

describe("headers and footers", () => {
  it("reads them, with page number fields, and knows the first page has its own", async () => {
    const { context } = await readDocx(await withHeaderAndFooter());
    const header = context.stories.find((s) => s.kind === "header" && s.type === "default")!;
    const first = context.stories.find((s) => s.kind === "header" && s.type === "first")!;
    const footer = context.stories.find((s) => s.kind === "footer")!;
    expect(header.doc.textContent).toBe("Lab report — group 4");
    expect(first.doc.textContent).toBe("Cover");
    expect(fieldsIn(footer.doc)).toEqual(["PAGE", "NUMPAGES"]);
    expect(context.titlePage).toBe(true);
  });

  it("saves an edited header and leaves the others exactly as they were", async () => {
    const original = await withHeaderAndFooter();
    const { doc, context } = await readDocx(original);
    const header = context.stories.find((s) => s.kind === "header" && s.type === "default")!;
    header.doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text("Final report")));
    const saved = await writeDocx(doc, context);
    const before = await JSZip.loadAsync(original);
    const after = await JSZip.loadAsync(saved);
    const footerPath = context.stories.find((s) => s.kind === "footer")!.path!;
    expect(await after.file(footerPath)!.async("string")).toBe(await before.file(footerPath)!.async("string"));
    const again = await readDocx(saved);
    expect(again.context.stories.find((s) => s.kind === "header" && s.type === "default")!.doc.textContent).toBe("Final report");
    expect(again.context.stories.find((s) => s.kind === "header" && s.type === "first")!.doc.textContent).toBe("Cover");
  });

  it("adds a new header and footer with a page number to a document that had none", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    expect(context.stories).toHaveLength(0);
    const page = schema.nodes.field.create({ instr: " PAGE ", result: "1" });
    context.stories.push(
      { kind: "header", type: "default", path: null, rid: null, rootStart: null, prefixes: new Set(), rels: new Map(), doc: schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text("My thesis"))), saved: null },
      { kind: "footer", type: "default", path: null, rid: null, rootStart: null, prefixes: new Set(), rels: new Map(), doc: schema.nodes.doc.create(null, schema.nodes.paragraph.create({ align: "center" }, page)), saved: null },
    );
    const saved = await writeDocx(doc, context);
    const zip = await JSZip.loadAsync(saved);
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain("wordprocessingml.header+xml");
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toMatch(/<w:sectPr><w:headerReference w:type="default" r:id="rId\d+"\/><w:footerReference w:type="default" r:id="rId\d+"\/><w:pgSz/);
    const again = await readDocx(saved);
    expect(again.context.stories.map((s) => s.kind)).toEqual(["header", "footer"]);
    expect(again.context.stories[0].doc.textContent).toBe("My thesis");
    expect(again.context.stories[1].doc.firstChild!.firstChild!.attrs).toMatchObject({ instr: " PAGE ", result: "1" });
    expect(fieldsIn(again.context.stories[1].doc)).toEqual(["PAGE"]);
    // Saving again doesn't add them twice.
    const twice = await JSZip.loadAsync(await writeDocx(again.doc, again.context));
    expect((await twice.file("word/document.xml")!.async("string")).match(/headerReference/g)).toHaveLength(1);
  });

  it("switches a separate first page on and off", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    context.titlePage = true;
    let xml = await (await JSZip.loadAsync(await writeDocx(doc, context))).file("word/document.xml")!.async("string");
    expect(xml).toMatch(/<w:cols [^>]*\/><w:titlePg\/><\/w:sectPr>/);
    context.titlePage = false;
    xml = await (await JSZip.loadAsync(await writeDocx(doc, context))).file("word/document.xml")!.async("string");
    expect(xml).not.toContain("titlePg");
  });
});

describe("Word's page number box", () => {
  it("in a footer is read as its text, with the page number counting each page", async () => {
    const zip = await JSZip.loadAsync(await withHeaderAndFooter());
    const name = Object.keys(zip.files).find((f) => /^word\/footer\d*\.xml$/.test(f))!;
    const xml = await zip.file(name)!.async("string");
    // As Word's Insert → Page Number → Bottom of Page writes it: the paragraph in a gallery content control.
    zip.file(name, xml.replace(/(<w:p\b[\s\S]*<\/w:p>)/, '<w:sdt><w:sdtPr><w:id w:val="1"/><w:docPartObj><w:docPartGallery w:val="Page Numbers (Bottom of Page)"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>$1</w:sdtContent></w:sdt>'));
    const { context } = await readDocx(await zip.generateAsync({ type: "uint8array" }));
    const footer = context.stories.find((s) => s.kind === "footer")!;
    const kinds: string[] = [];
    footer.doc.forEach((n) => kinds.push(n.type.name));
    expect([kinds, fieldsIn(footer.doc)]).toEqual([["paragraph"], ["PAGE", "NUMPAGES"]]);
  });
});

describe("content controls holding paragraphs", () => {
  const SDT_PR = '<w:sdtPr><w:id w:val="7"/><w:placeholder><w:docPart w:val="DefaultPlaceholder"/></w:placeholder><w:showingPlcHdr/></w:sdtPr>';
  async function withField(): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(await withHeaderAndFooter());
    const xml = await zip.file("word/document.xml")!.async("string");
    // A form field as a school template has one: a paragraph in a content control, showing its placeholder.
    zip.file("word/document.xml", xml.replace(/(<w:body>)/, `$1<w:sdt>${SDT_PR}<w:sdtContent><w:p><w:r><w:t>Click here to enter text.</w:t></w:r></w:p></w:sdtContent></w:sdt>`));
    return zip.generateAsync({ type: "uint8array" });
  }

  it("are paragraphs to edit, and keep their control when saved; the page number box too", async () => {
    const { doc, context } = await readDocx(await withField());
    expect(doc.firstChild!.textContent).toBe("Click here to enter text.");
    const zip = await JSZip.loadAsync(await writeDocx(doc, context));
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain(`<w:sdt>${SDT_PR}<w:sdtContent><w:p`);
  });

  it("typed over, the placeholder is no longer a placeholder", async () => {
    const { doc, context } = await readDocx(await withField());
    const first = doc.firstChild!;
    const { Fragment, Slice } = await import("prosemirror-model");
    // The paragraph, typed over (it keeps its attributes, the control among them).
    const edited = doc.replace(0, first.nodeSize, new Slice(Fragment.from(first.type.create(first.attrs, schema.text("Kari Nordmann"))), 0, 0));
    const xml = await (await JSZip.loadAsync(await writeDocx(edited, context))).file("word/document.xml")!.async("string");
    expect([xml.includes("Kari Nordmann"), xml.includes("<w:showingPlcHdr/>"), /<w:sdt><w:sdtPr><w:id w:val="7"\/>/.test(xml)]).toEqual([true, false, true]);
  });
});
