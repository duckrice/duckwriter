// A PDF printed on one big paper size (to fit pages of several sizes) cut back to each page's own
// size: each page sits in the top of its slot, centred across, as the layout puts it.

import { PDFDocument } from "pdf-lib";

/** The PDF with page i cut to sizes[i] (points), from paper this big. */
export async function cropPages(bytes: Uint8Array, sizes: { width: number; height: number }[], paper: { width: number; height: number }): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false });
  doc.getPages().forEach((page, i) => {
    const size = sizes[i];
    if (!size) return;
    // PDF measures from the bottom: the page is the top `height` of the paper.
    const x = (paper.width - size.width) / 2;
    const y = paper.height - size.height;
    page.setMediaBox(x, y, size.width, size.height);
    page.setCropBox(x, y, size.width, size.height);
  });
  // The browser may add an empty sheet after the last slot.
  for (let i = doc.getPageCount() - 1; i >= sizes.length; i--) doc.removePage(i);
  return doc.save({ useObjectStreams: false });
}
