// Opens a PDF with pdf.js and turns it into a Word document. Loaded only when
// a PDF is opened, so pdf.js doesn't slow down starting DuckWriter.

import { GlobalWorkerOptions, getDocument } from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { pdfToDocx } from "./pdf";

GlobalWorkerOptions.workerSrc = workerUrl;

const asset = (dir: string) => new URL(`pdfjs/${dir}/`, document.baseURI).href;

/** The .docx bytes for a PDF's bytes. */
export async function convertPdf(bytes: Uint8Array): Promise<Uint8Array> {
  const task = getDocument({
    // pdf.js takes ownership of the buffer it is given, so hand it a copy.
    data: bytes.slice(),
    cMapUrl: asset("cmaps"),
    cMapPacked: true,
    standardFontDataUrl: asset("standard_fonts"),
    wasmUrl: asset("wasm"),
    iccUrl: asset("iccs"),
    // Keeps the fonts' real names, for the fonts in the Word document.
    fontExtraProperties: true,
  });
  task.onPassword = () => void task.destroy();
  let pdf;
  try {
    pdf = await task.promise;
  } catch (error) {
    if (task.destroyed) throw new Error("it is protected with a password. Open it in Duckument, which asks for the password, and use Save as Word there");
    throw error;
  }
  try {
    return await pdfToDocx(pdf);
  } finally {
    void task.destroy();
  }
}
