import { describe, expect, it } from "vitest";
import { PDFDocument, PDFName, PDFString, StandardFonts, rgb } from "pdf-lib";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import { familyName, layout, pdfToDocx, type Block, type PageText, type TextBit } from "../src/convert/pdf";
import { readDocx } from "../src/docx/read";
import { listOf } from "../src/editor/lists";

/** A piece of text for a made-up page; widths roughly as Helvetica's. */
const bit = (text: string, x: number, base: number, size = 11, extra: Partial<TextBit> = {}): TextBit => ({ text, x, base, w: text.length * size * 0.5, size, font: "Arial", bold: false, italic: false, ...extra });
const page = (bits: TextBit[], width = 595, height = 842): PageText => ({ width, height, bits, pictures: [] });
const words = (b: Block) => (b.kind === "paragraph" ? b.runs.map((r) => r.text).join("") : "");

describe("working out a document from a PDF's text", () => {
  it("joins lines into paragraphs and finds headings by size", () => {
    const blocks = layout([
      page([
        bit("Cell biology", 72, 90, 24, { bold: true }),
        bit("The cell is the smallest unit of life. All living things are made of", 72, 130),
        bit("cells, and every cell comes from another cell.", 72, 144),
        bit("Second paragraph starts here after a gap.", 72, 172),
      ]),
    ]);
    expect(blocks.map((b) => (b.kind === "paragraph" ? b.heading ?? 0 : -1))).toEqual([1, 0, 0]);
    expect(words(blocks[1])).toBe("The cell is the smallest unit of life. All living things are made of cells, and every cell comes from another cell.");
  });

  it("mends words broken over two lines", () => {
    const [p] = layout([page([bit("Photosynthesis turns light into chemical en-", 72, 100), bit("ergy that plants store.", 72, 114)])]);
    expect(words(p)).toBe("Photosynthesis turns light into chemical energy that plants store.");
  });

  it("finds bullet and numbered lists, with lines that carry on", () => {
    const blocks = layout([
      page([
        bit("Things to bring:", 72, 100),
        bit("•", 80, 120),
        bit("a notebook and a pen for the long", 92, 120),
        bit("lecture", 92, 134),
        bit("•", 80, 148),
        bit("the textbook", 92, 148),
        bit("1.", 80, 176),
        bit("Read chapter 4", 96, 176),
        bit("2.", 80, 190),
        bit("Do the exercises", 96, 190),
      ]),
    ]);
    const lists = blocks.filter((b) => b.kind === "paragraph" && b.list);
    expect(lists.map((b) => (b.kind === "paragraph" ? [b.list, words(b)] : []))).toEqual([
      ["bullet", "a notebook and a pen for the long lecture"],
      ["bullet", "the textbook"],
      ["number", "Read chapter 4"],
      ["number", "Do the exercises"],
    ]);
  });

  it("finds tables from text in aligned columns", () => {
    const [table] = layout([page([bit("Part", 72, 100), bit("Job", 250, 100), bit("Nucleus", 72, 116), bit("Holds the DNA", 250, 116), bit("Ribosome", 72, 132), bit("Makes proteins", 250, 132)])]);
    expect(table.kind).toBe("table");
    if (table.kind === "table") expect(table.rows.map((r) => r.map((c) => c.map((run) => run.text).join("")))).toEqual([["Part", "Job"], ["Nucleus", "Holds the DNA"], ["Ribosome", "Makes proteins"]]);
  });

  it("reads two columns one after the other", () => {
    const left = Array.from({ length: 8 }, (_, i) => bit(`left line ${i + 1} of the first column text`, 50, 100 + i * 14, 10));
    const right = Array.from({ length: 8 }, (_, i) => bit(`right line ${i + 1} of the second column`, 320, 100 + i * 14, 10));
    const blocks = layout([page([...left, ...right])]);
    const all = blocks.map(words).join(" | ");
    expect(all.indexOf("left line 8")).toBeLessThan(all.indexOf("right line 1"));
  });

  it("leaves out page numbers and running heads", () => {
    const pages = [1, 2, 3].map((n) => page([bit("Biology notes", 72, 40, 9), bit(`Text on page ${n}.`, 72, 100), bit(String(n), 290, 810, 9)]));
    const all = layout(pages).map(words);
    expect(all.join(" ")).not.toContain("Biology notes");
    expect(all.some((w) => /^\d$/.test(w.trim()))).toBe(false);
    expect(all.length).toBe(3);
  });

  it("marks raised text as superscript", () => {
    const [p] = layout([page([bit("E = mc", 72, 100), bit("2", 72 + 6 * 5.5, 96, 7)])]);
    expect(p.kind === "paragraph" && p.runs.find((r) => r.text.trim() === "2")?.sup).toBe(true);
  });

  it("names fonts as Word knows them", () => {
    expect(familyName("ABCDEF+TimesNewRomanPS-BoldMT")).toBe("Times New Roman");
    expect(familyName("ArialMT")).toBe("Arial");
    expect(familyName("Helvetica-Bold")).toBe("Arial");
    expect(familyName("Calibri,Italic")).toBe("Calibri");
    expect(familyName("SegoeUI")).toBe("Segoe UI");
    expect(familyName("F1")).toBe("");
  });
});

describe("a real PDF to a Word document", () => {
  it("keeps the headings, bold, lists, tables and links", async () => {
    const pdf = await PDFDocument.create();
    const p = pdf.addPage([595, 842]);
    const regular = await pdf.embedFont(StandardFonts.Helvetica);
    const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
    const draw = (text: string, x: number, top: number, size = 11, font = regular) => p.drawText(text, { x, y: 842 - top, size, font, color: rgb(0, 0, 0) });
    draw("Lecture 3: Cells", 72, 90, 24, bold);
    draw("Cells are the building blocks of life. This line is long enough", 72, 130);
    draw("to wrap onto a second line in the same paragraph.", 72, 144);
    draw("Key words", 72, 180, 11, bold);
    draw("•", 80, 200);
    draw("nucleus", 92, 200);
    draw("•", 80, 214);
    draw("membrane", 92, 214);
    draw("Part", 72, 250);
    draw("Job", 260, 250);
    draw("Nucleus", 72, 266);
    draw("Holds DNA", 260, 266);
    draw("More at duck.no/cells", 72, 300);
    const link = pdf.context.obj({
      Type: "Annot",
      Subtype: "Link",
      Rect: [72, 842 - 303, 72 + 120, 842 - 290],
      Border: [0, 0, 0],
      A: { Type: "Action", S: "URI", URI: PDFString.of("https://duck.no/cells") },
    });
    p.node.set(PDFName.of("Annots"), pdf.context.obj([pdf.context.register(link)]));
    const bytes = await pdf.save();

    const loaded = await pdfjs.getDocument({ data: bytes, fontExtraProperties: true, standardFontDataUrl: "node_modules/pdfjs-dist/standard_fonts/" }).promise;
    const docx = await pdfToDocx(loaded as never);
    const { doc, context } = await readDocx(docx);
    const paras = doc.content.content;
    const byText = (t: string) => paras.find((n) => n.textContent.startsWith(t))!;

    expect(byText("Lecture 3").attrs.style).toBe("Heading1");
    expect(byText("Cells are").textContent).toBe("Cells are the building blocks of life. This line is long enough to wrap onto a second line in the same paragraph.");
    expect(byText("Key words").attrs.style).toMatch(/^Heading\d$/);
    expect(listOf(context, byText("nucleus"))).toBeTruthy();
    expect(listOf(context, byText("membrane"))).toBeTruthy();
    let table = false;
    let linked = false;
    doc.descendants((node) => {
      if (node.type.name === "table") table = node.textContent.includes("Holds DNA");
      if (node.isText && node.marks.some((m) => m.type.name === "link" && m.attrs.href === "https://duck.no/cells")) linked = true;
    });
    expect(table).toBe(true);
    expect(linked).toBe(true);
    let boldTitle = false;
    byText("Lecture 3").descendants((n) => {
      if (n.isText && n.marks.some((m) => m.type.name === "bold" || (m.type.name === "run" && m.attrs.bold))) boldTitle = true;
    });
    expect(boldTitle || byText("Lecture 3").attrs.style === "Heading1").toBe(true);
  });
});
