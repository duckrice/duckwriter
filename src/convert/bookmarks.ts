// Bookmarks in a saved PDF: the headings, nested by level, each going to its place on its page,
// as Word's "Create bookmarks using headings". The browser's print to PDF makes none, so they
// are added to the file afterwards.

import { PDFDocument, PDFHexString, PDFName, PDFNull, PDFNumber, type PDFDict, type PDFRef } from "pdf-lib";

export interface Bookmark {
  /** 1 for Heading 1, and so on. */
  level: number;
  title: string;
  /** The page, from 0, and how far down it the heading is, from 0 (top) to 1 (bottom). */
  page: number;
  down: number;
}

/** The PDF with the bookmarks in; one without pages for them is left as it is. */
export async function addBookmarks(bytes: Uint8Array, bookmarks: Bookmark[]): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false });
  const pages = doc.getPages();
  const wanted = bookmarks.filter((b) => b.title.trim() && b.page >= 0 && b.page < pages.length);
  if (!wanted.length) return bytes;
  const { context } = doc;
  const root = context.obj({ Type: "Outlines" }) as PDFDict;
  const rootRef = context.register(root);
  interface Entry {
    ref: PDFRef;
    dict: PDFDict;
    level: number;
    kids: Entry[];
  }
  const top: Entry = { ref: rootRef, dict: root, level: 0, kids: [] };
  const path: Entry[] = [top];
  for (const b of wanted) {
    const page = pages[b.page];
    const y = page.getHeight() * (1 - Math.min(1, Math.max(0, b.down)));
    const dict = context.obj({}) as PDFDict;
    dict.set(PDFName.of("Title"), PDFHexString.fromText(b.title.trim()));
    dict.set(PDFName.of("Dest"), context.obj([page.ref, PDFName.of("XYZ"), PDFNull, PDFNumber.of(Math.round(y)), PDFNull]));
    const entry: Entry = { ref: context.register(dict), dict, level: b.level, kids: [] };
    // Under the nearest heading above of a higher level (a Heading 3 straight under a Heading 1 goes under it too).
    while (path.length > 1 && path[path.length - 1].level >= b.level) path.pop();
    path[path.length - 1].kids.push(entry);
    path.push(entry);
  }
  const link = (parent: Entry): number => {
    let count = 0;
    parent.kids.forEach((kid, i) => {
      kid.dict.set(PDFName.of("Parent"), parent.ref);
      if (i > 0) kid.dict.set(PDFName.of("Prev"), parent.kids[i - 1].ref);
      if (i < parent.kids.length - 1) kid.dict.set(PDFName.of("Next"), parent.kids[i + 1].ref);
      count += 1 + link(kid);
    });
    if (parent.kids.length) {
      parent.dict.set(PDFName.of("First"), parent.kids[0].ref);
      parent.dict.set(PDFName.of("Last"), parent.kids[parent.kids.length - 1].ref);
      // The top level's count is everything; a heading's is shown as closed below level 2.
      parent.dict.set(PDFName.of("Count"), PDFNumber.of(parent === top || parent.level < 2 ? count : -count));
    }
    return count;
  };
  link(top);
  doc.catalog.set(PDFName.of("Outlines"), rootRef);
  doc.catalog.set(PDFName.of("PageMode"), PDFName.of("UseOutlines"));
  // Without object streams, as the browser wrote it, so simple readers still find the pages.
  return doc.save({ useObjectStreams: false });
}
