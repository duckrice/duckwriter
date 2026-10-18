import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { checkPackage } from "../src/docx/check";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";

const samples = readdirSync(join(__dirname, "..", "samples")).filter((n) => n.endsWith(".docx") && n !== "not-a-docx.docx");

describe("the package check", () => {
  it("finds what Word would call damage", async () => {
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(await blankDocx());
    zip.file("word/document.xml", (await zip.file("word/document.xml")!.async("string")).replace("<w:body>", '<w:body><w:p><w:r><w:footnoteReference w:id="9"/></w:r></w:p><w:p><w:bookmarkStart w:id="1" w:name="a"/></w:p>'));
    zip.file("word/stray.bin", "x");
    zip.file("word/bad.xml", "<a><b></a>");
    const problems = await checkPackage(await zip.generateAsync({ type: "uint8array" }));
    expect(problems.some((p) => p.includes("footnote 9"))).toBe(true);
    expect(problems.some((p) => p.includes("bookmark a"))).toBe(true);
    expect(problems.some((p) => p.includes("stray.bin has no content type"))).toBe(true);
    expect(problems.some((p) => p.includes("bad.xml isn't well-formed"))).toBe(true);
  });

  it("passes a new blank document", async () => {
    expect(await checkPackage(await blankDocx())).toEqual([]);
  });

  it("passes a document whose pictures, notes and comments were copied", async () => {
    const docx = await import("docx");
    const png = readFileSync(join(__dirname, "..", "src-tauri", "icons", "32x32.png"));
    const file = new docx.Document({
      footnotes: { 1: { children: [new docx.Paragraph("A note.")] } },
      comments: { children: [{ id: 0, author: "Kari", date: new Date(), children: [new docx.Paragraph("Look")] }] },
      sections: [
        {
          headers: { default: new docx.Header({ children: [new docx.Paragraph({ children: [new docx.ImageRun({ data: png, type: "png", transformation: { width: 20, height: 20 } })] })] }) },
          children: [
            new docx.Paragraph({ children: [new docx.TextRun("Picture "), new docx.ImageRun({ data: png, type: "png", transformation: { width: 64, height: 48 } })] }),
            new docx.Paragraph({ children: [new docx.CommentRangeStart(0), new docx.TextRun("noted"), new docx.CommentRangeEnd(0), new docx.TextRun({ children: [new docx.CommentReference(0)] }), new docx.FootnoteReferenceRun(1)] }),
            new docx.Paragraph({ children: [new docx.ExternalHyperlink({ link: "https://duck.no", children: [new docx.TextRun("link")] })] }),
          ],
        },
      ],
    });
    const original = new Uint8Array(await docx.Packer.toBuffer(file));
    expect(await checkPackage(original)).toEqual([]);
    const { doc, context } = await readDocx(original);
    // Copy every paragraph (pasting the whole text again, as a user might).
    const twice = doc.type.create(doc.attrs, doc.content.append(doc.content));
    expect(await checkPackage(await writeDocx(twice, context))).toEqual([]);
  });

  for (const name of samples) {
    it(`passes ${name}, as it is and saved by DuckWriter`, async () => {
      const original = new Uint8Array(readFileSync(join(__dirname, "..", "samples", name)));
      expect(await checkPackage(original)).toEqual([]);
      const { doc, context } = await readDocx(original);
      expect(await checkPackage(await writeDocx(doc, context))).toEqual([]);
    });
  }
});
