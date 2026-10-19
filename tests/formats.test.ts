import JSZip from "jszip";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { checkPackage } from "../src/docx/check";
import { asPackage, convertPackage, fromFlatOpc, isFlatOpc, packageKind, toFlatOpc, wordXml } from "../src/docx/package";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";

const MACRO = "application/vnd.ms-word.document.macroEnabled.main+xml";
const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const TEMPLATE = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
const types = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file("[Content_Types].xml")!.async("string");

/** A .docm: a document with a (pretend) macro project. */
async function docm(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await blankDocx());
  const t = await zip.file("[Content_Types].xml")!.async("string");
  zip.file("[Content_Types].xml", t.replace(DOCX, MACRO).replace("</Types>", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/word/vbaData.xml" ContentType="application/vnd.ms-word.vbaData+xml"/></Types>'));
  const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
  zip.file("word/_rels/document.xml.rels", rels.replace("</Relationships>", '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'));
  zip.file("word/vbaProject.bin", new Uint8Array([1, 2, 3]));
  zip.file("word/_rels/vbaProject.bin.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/wordVbaData" Target="vbaData.xml"/></Relationships>');
  zip.file("word/vbaData.xml", '<?xml version="1.0"?><wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>');
  return zip.generateAsync({ type: "uint8array" });
}

describe("the .docx family", () => {
  it("knows each kind by its name", () => {
    expect(["a.docx", "a.DOCM", "a.dotx", "a.dotm", "a.xml", "a.doc", "a.pdf"].map(packageKind)).toEqual(["docx", "docm", "dotx", "dotm", "docx", null, null]);
  });

  it("opens a .docm and keeps its macros when saved as .docm", async () => {
    const bytes = await docm();
    const { doc, context } = await readDocx(bytes);
    const saved = await convertPackage(await writeDocx(doc, context), "docm");
    const zip = await JSZip.loadAsync(saved);
    expect(zip.file("word/vbaProject.bin")).toBeTruthy();
    expect(await types(saved)).toContain(MACRO);
    expect(await checkPackage(saved)).toEqual([]);
  });

  it("leaves the macros out when a .docm is saved as .docx, so Word opens it", async () => {
    const { doc, context } = await readDocx(await docm());
    const saved = await convertPackage(await writeDocx(doc, context), "docx");
    const zip = await JSZip.loadAsync(saved);
    expect(zip.file("word/vbaProject.bin")).toBeNull();
    expect(zip.file("word/vbaData.xml")).toBeNull();
    const t = await types(saved);
    expect(t).toContain(DOCX);
    expect(t).not.toContain(MACRO);
    expect(t).not.toContain('Extension="bin"');
    expect(await zip.file("word/_rels/document.xml.rels")!.async("string")).not.toContain("vbaProject");
    expect(await checkPackage(saved)).toEqual([]);
  });

  it("makes a template of a document, and leaves one already right alone", async () => {
    const blank = await blankDocx();
    expect(await types(await convertPackage(blank, "dotx"))).toContain(TEMPLATE);
    expect(await convertPackage(blank, "docx")).toBe(blank);
  });
});

describe("Word XML documents (one file)", () => {
  it("go to one XML file and back, with every part", async () => {
    const original = new Uint8Array(readFileSync(join(__dirname, "..", "samples", "letterhead.docx")));
    const xml = await toFlatOpc(original);
    expect(xml).toContain('<?mso-application progid="Word.Document"?>');
    expect(isFlatOpc(xml)).toBe(true);
    const back = await fromFlatOpc(xml);
    expect(await checkPackage(back)).toEqual([]);
    const a = await JSZip.loadAsync(original);
    const b = await JSZip.loadAsync(back);
    const files = (z: JSZip) => Object.keys(z.files).filter((f) => !z.files[f].dir && f !== "[Content_Types].xml").sort();
    expect(files(b)).toEqual(files(a));
    // Pictures come back byte for byte; the text reads the same.
    for (const f of files(a).filter((n) => !/\.(xml|rels)$/.test(n))) expect(await b.file(f)!.async("base64")).toBe(await a.file(f)!.async("base64"));
    expect((await readDocx(back)).doc.textContent).toBe((await readDocx(original)).doc.textContent);
  });

  it("open, and say plainly when they are something else", async () => {
    const xml = new TextEncoder().encode(await toFlatOpc(await blankDocx()));
    expect((await readDocx(await wordXml(xml, "a.xml"))).doc).toBeTruthy();
    expect(await asPackage(xml)).not.toBe(xml);
    const zip = await blankDocx();
    expect(await asPackage(zip)).toBe(zip);
    await expect(wordXml(new TextEncoder().encode('<?xml version="1.0"?><w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"/>'), "old.xml")).rejects.toThrow(/Word 2003 XML/);
    await expect(wordXml(new TextEncoder().encode("<notes/>"), "notes.xml")).rejects.toThrow(/isn't a Word XML document/);
  });
});
