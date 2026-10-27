import { describe, expect, it } from "vitest";
import { PDFDocument } from "pdf-lib";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import { addBookmarks } from "../src/convert/bookmarks";

describe("PDF bookmarks", () => {
  it("are the headings, nested by level, each going to its page", async () => {
    const doc = await PDFDocument.create();
    for (let i = 0; i < 3; i++) doc.addPage([595, 842]);
    const bytes = await addBookmarks(await doc.save(), [
      { level: 1, title: "Introduction", page: 0, down: 0.1 },
      { level: 2, title: "Why ducks", page: 0, down: 0.5 },
      { level: 3, title: "Ørret og ål", page: 1, down: 0.2 },
      { level: 1, title: "Method", page: 2, down: 0 },
      { level: 1, title: "Past the end", page: 9, down: 0 },
    ]);
    const pdf = await pdfjs.getDocument({ data: bytes }).promise;
    const outline = await pdf.getOutline();
    const tree = (items: typeof outline): unknown[] => items.map((i) => [i.title, ...(i.items.length ? [tree(i.items)] : [])]);
    expect(tree(outline)).toEqual([["Introduction", [["Why ducks", [["Ørret og ål"]]]]], ["Method"]]);
    const [ref, , , top] = outline[1].dest as [{ num: number; gen: number }, unknown, unknown, number];
    expect([await pdf.getPageIndex(ref), top]).toEqual([2, 842]);
  });
});
