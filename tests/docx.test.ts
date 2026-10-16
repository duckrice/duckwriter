import * as docx from "docx";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";

/** A Word file written by another program, with a bit of everything. */
async function sample(): Promise<Uint8Array> {
  const d = new docx.Document({
    numbering: {
      config: [
        {
          reference: "numbers",
          levels: [
            { level: 0, format: docx.LevelFormat.DECIMAL, text: "%1.", alignment: docx.AlignmentType.START, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
            { level: 1, format: docx.LevelFormat.LOWER_LETTER, text: "%2)", alignment: docx.AlignmentType.START, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
          ],
        },
      ],
    },
    sections: [
      {
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1800, bottom: 1440, left: 1800 } } },
        children: [
          new docx.Paragraph({ heading: docx.HeadingLevel.HEADING_1, children: [new docx.TextRun("Trip to Bergen")] }),
          new docx.Paragraph({
            alignment: docx.AlignmentType.JUSTIFIED,
            spacing: { before: 120, after: 240, line: 360 },
            children: [
              new docx.TextRun("Plain, "),
              new docx.TextRun({ text: "bold", bold: true }),
              new docx.TextRun(", "),
              new docx.TextRun({ text: "red italic", italics: true, color: "C00000" }),
              new docx.TextRun({ text: " big Georgia", size: 32, font: "Georgia" }),
              new docx.TextRun({ text: "\tafter a tab", highlight: "yellow" }),
              new docx.TextRun({ text: " æøå", underline: {} }),
            ],
          }),
          new docx.Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new docx.TextRun("First")] }),
          new docx.Paragraph({ numbering: { reference: "numbers", level: 1 }, children: [new docx.TextRun("Nested")] }),
          new docx.Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new docx.TextRun("Second")] }),
          new docx.Paragraph({ children: [new docx.ExternalHyperlink({ link: "https://example.com/", children: [new docx.TextRun({ text: "a link", style: "Hyperlink" })] })] }),
          new docx.Table({
            rows: [
              new docx.TableRow({ children: [new docx.TableCell({ children: [new docx.Paragraph("A1")] }), new docx.TableCell({ children: [new docx.Paragraph("B1")] })] }),
            ],
          }),
          new docx.Paragraph({ children: [new docx.TextRun("Before the break"), new docx.PageBreak(), new docx.TextRun("After")] }),
        ],
      },
    ],
  });
  return new Uint8Array(await docx.Packer.toBuffer(d));
}

const json = (bytes: Uint8Array) => readDocx(bytes).then(({ doc }) => doc.toJSON());

describe("reading a .docx", () => {
  it("reads text, formatting, lists, links, the page and what it can't edit", async () => {
    const { doc, context } = await readDocx(await sample());
    const paras = doc.content.content;
    expect(paras[0].attrs.style).toBe("Heading1");
    expect(paras[0].textContent).toBe("Trip to Bergen");

    const p = paras[1];
    expect(p.attrs).toMatchObject({ align: "justify", spaceBefore: 6, spaceAfter: 12, lineSpacing: 1.5 });
    const marksOf = (text: string) => {
      let found: string[] = [];
      p.forEach((n) => {
        if (n.text?.includes(text)) found = n.marks.map((m) => `${m.type.name}${m.attrs.on === false ? ":off" : ""}${m.attrs.color ? `:${m.attrs.color}` : ""}${m.attrs.pt ? `:${m.attrs.pt}` : ""}${m.attrs.family ? `:${m.attrs.family}` : ""}`);
      });
      return found.sort();
    };
    expect(marksOf("bold")).toEqual(["bold"]);
    expect(marksOf("red italic")).toEqual(["color:#c00000", "italic"]);
    expect(marksOf("big Georgia")).toEqual(["font:Georgia", "size:16"]);
    expect(marksOf("after a tab")).toEqual(["highlight:yellow"]);
    expect(p.textContent).toContain("\tafter a tab");
    expect(marksOf("æøå")).toEqual(["underline"]);

    expect(paras[2].attrs.numId).not.toBeNull();
    expect(paras[3].attrs.ilvl).toBe(1);
    const list = context.numbering.nums.get(paras[2].attrs.numId)!;
    expect(context.numbering.abstracts.get(list.abstractId)![1]).toMatchObject({ format: "lowerLetter", text: "%2)" });

    let link = "";
    paras[5].forEach((n) => (link = n.marks.find((m) => m.type.name === "link")?.attrs.href ?? link));
    expect(link).toBe("https://example.com/");

    expect(paras[6].type.name).toBe("raw_block");
    expect(paras[6].attrs.label).toBe("Table");
    expect(paras[6].attrs.preview).toContain("A1  |  B1");

    const types: string[] = [];
    paras[7].forEach((n) => types.push(n.type.name));
    expect(types).toEqual(["text", "page_break", "text"]);

    expect(context.page).toMatchObject({ width: 612, height: 792, marginLeft: 90 });
    // The file names no default font or size: Word then uses Times New Roman at 10 points.
    expect(context.styles.defaultRun).toMatchObject({ font: "Times New Roman", size: 10 });
  });

  it("says plainly when a file isn't a Word document", async () => {
    await expect(readDocx(new TextEncoder().encode("hello"))).rejects.toThrow(/isn't a Word document/);
  });
});

describe("saving a .docx", () => {
  it("saves exactly what it read when nothing changed", async () => {
    const original = await sample();
    const { doc, context } = await readDocx(original);
    const saved = await writeDocx(doc, context);
    expect(await json(saved)).toEqual(doc.toJSON());
  });

  it("keeps what it can't edit byte for byte", async () => {
    const original = await sample();
    const before = await (await JSZip.loadAsync(original)).file("word/document.xml")!.async("string");
    const { doc, context } = await readDocx(original);
    const after = await (await JSZip.loadAsync(await writeDocx(doc, context))).file("word/document.xml")!.async("string");
    const table = before.slice(before.indexOf("<w:tbl>"), before.indexOf("</w:tbl>") + 8);
    expect(table.length).toBeGreaterThan(100);
    expect(after).toContain(table);
    const sectPr = before.slice(before.lastIndexOf("<w:sectPr"), before.lastIndexOf("</w:sectPr>") + 11);
    expect(after).toContain(sectPr);
  });

  it("keeps every other part of the file", async () => {
    const original = await sample();
    const { doc, context } = await readDocx(original);
    const saved = await JSZip.loadAsync(await writeDocx(doc, context));
    const before = await JSZip.loadAsync(original);
    for (const name of Object.keys(before.files)) {
      if (name === "word/document.xml" || before.files[name].dir) continue;
      expect(await saved.file(name)?.async("string"), name).toBe(await before.file(name)!.async("string"));
    }
  });

  it("writes valid XML with properties in the order Word requires", async () => {
    const { doc, context } = await readDocx(await sample());
    const saved = await JSZip.loadAsync(await writeDocx(doc, context));
    const xml = await saved.file("word/document.xml")!.async("string");
    const parsed = new DOMParser().parseFromString(xml, "application/xml");
    expect(parsed.getElementsByTagName("parsererror")).toHaveLength(0);
    // pStyle first, jc last among what we write; rFonts before sz.
    expect(xml).toMatch(/<w:pPr><w:spacing[^>]*\/><w:jc w:val="both"\/><\/w:pPr>/);
    expect(xml).toMatch(/<w:rFonts[^>]*\/><w:sz w:val="32"\/>/);
    expect(xml).toContain('<w:tab/><w:t xml:space="preserve">after a tab</w:t>');
  });

  it("starts new documents from a blank one that reads and saves like any other", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    expect(doc.childCount).toBe(1);
    expect(context.styles.defaultParagraph).toBe("Normal");
    expect([...context.styles.byId.keys()]).toEqual(expect.arrayContaining(["Title", "Heading1", "Heading2", "Heading3", "Quote"]));
    expect(context.styles.defaultRun).toMatchObject({ font: "Calibri", size: 11 });
    expect(context.page.width).toBeCloseTo(595.3, 1);
    const again = await readDocx(await writeDocx(doc, context));
    expect(again.doc.toJSON()).toEqual(doc.toJSON());
  });
});
