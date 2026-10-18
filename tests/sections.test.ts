import * as docx from "docx";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { writeDocx } from "../src/docx/write";
import { checkPackage } from "../src/docx/check";
import { formatPageNumber, pageLabels, sectionBreak } from "../src/editor/pagination";

/** A thesis in two sections: front matter numbered i, ii, then the chapters from 1. */
async function thesis(): Promise<Uint8Array> {
  const footer = (name: string) => new docx.Footer({ children: [new docx.Paragraph({ children: [new docx.TextRun({ children: [`${name} `, docx.PageNumber.CURRENT] })] })] });
  const file = new docx.Document({
    sections: [
      { properties: { page: { pageNumbers: { start: 1, formatType: docx.NumberFormat.LOWER_ROMAN } } }, footers: { default: footer("Front") }, children: [new docx.Paragraph("Abstract")] },
      { properties: { type: docx.SectionType.NEXT_PAGE, page: { pageNumbers: { start: 1, formatType: docx.NumberFormat.DECIMAL } } }, footers: { default: footer("Chapter") }, children: [new docx.Paragraph("Chapter one")] },
    ],
  });
  return new Uint8Array(await docx.Packer.toBuffer(file));
}

describe("sections", () => {
  it("are found at the paragraph that ends them, and kept when saved", async () => {
    const { doc, context } = await readDocx(await thesis());
    const breaks = doc.content.content.map((n) => sectionBreak(n)).filter(Boolean);
    expect(breaks).toHaveLength(1);
    expect(breaks[0]!.xml).toContain('w:fmt="lowerRoman"');
    expect(context.sectPr).toContain('w:fmt="decimal"');
    // The front matter's footer is read too, apart from the last section's.
    expect([...(context.sectionStories?.values() ?? [])].map((s) => s.doc.textContent.trim())).toEqual(["Front"]);
    expect(context.stories.map((s) => s.doc.textContent.trim())).toEqual(["Chapter"]);
    const saved = await writeDocx(doc, context);
    expect(await checkPackage(saved)).toEqual([]);
    const again = await readDocx(saved);
    expect(again.doc.content.content.map((n) => sectionBreak(n)).filter(Boolean)).toHaveLength(1);
  });

  it("number their pages: i, ii, then 1, 2, 3", () => {
    const front = '<w:sectPr><w:pgNumType w:fmt="lowerRoman" w:start="1"/></w:sectPr>';
    const body = '<w:sectPr><w:pgNumType w:start="1"/></w:sectPr>';
    expect(pageLabels(5, [{ firstPage: 0, restarts: true, xml: front }, { firstPage: 2, restarts: true, xml: body }])).toEqual(["i", "ii", "1", "2", "3"]);
    // Without a start, numbering goes on; one section is 1, 2, 3.
    expect(pageLabels(3, [{ firstPage: 0, restarts: true, xml: "" }])).toEqual(["1", "2", "3"]);
    expect(pageLabels(3, [{ firstPage: 0, restarts: true, xml: front }, { firstPage: 1, restarts: true, xml: "<w:sectPr/>" }])).toEqual(["i", "2", "3"]);
    expect(formatPageNumber(4, "upperRoman")).toBe("IV");
    expect(formatPageNumber(28, "lowerLetter")).toBe("bb");
  });
});
