import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { htmlToDoc } from "../src/convert/import";
import { odtToHtml } from "../src/convert/odt";
import { rtfToHtml } from "../src/convert/rtf";
import { docToHtml } from "../src/convert/doc";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { listOf } from "../src/editor/lists";
import { schema } from "../src/editor/schema";

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", name)));

/** What every reader must get from tests/fixtures/formats.* (made by scripts/make-format-fixtures.mjs). */
async function expectTheTestDocument(html: string) {
  const { context } = await readDocx(await blankDocx());
  const doc = htmlToDoc(html, context);
  const paras = doc.content.content;
  const byText = (t: string) => paras.find((p) => p.textContent.startsWith(t))!;
  expect(byText("Rates of reaction").attrs.style).toBe("Title");
  expect(byText("Introduction").attrs.style).toBe("Heading1");
  expect(byText("Methods").attrs.style).toBe("Heading2");
  const body = byText("Plain");
  expect(body.textContent).toContain("Blåbærsyltetøy.");
  const marked = (text: string, mark: string) => {
    let found = false;
    body.forEach((n) => {
      if (n.text === text && n.marks.some((m) => m.type.name === mark)) found = true;
    });
    return found;
  };
  expect(marked("bold", "bold")).toBe(true);
  expect(marked("italic", "italic")).toBe(true);
  expect(marked("underlined", "underline")).toBe(true);
  expect(body.textContent).toMatch(/H2O/);
  let sub = false;
  let sup = false;
  body.forEach((n) => {
    const v = n.marks.find((m) => m.type === schema.marks.vertAlign)?.attrs.value;
    if (n.text === "2" && v === "subscript") sub = true;
    if (n.text === "2" && v === "superscript") sup = true;
  });
  expect([sub, sup]).toEqual([true, true]);
  expect(listOf(context, byText("First point"))).toBeTruthy();
  expect(listOf(context, byText("Second point"))).toBeTruthy();
  expect(byText("Centred line").attrs.align).toBe("center");
  const table = paras.find((p) => p.type === schema.nodes.table)!;
  expect(table.textContent).toBe("RunRate10.5");
  expect(doc.textContent).toContain("Measured at noon.");
  expect(doc.textContent).toContain("The end.");
}

describe("older formats, opened as Word documents", () => {
  it("OpenDocument (.odt)", async () => {
    await expectTheTestDocument(await odtToHtml(fixture("formats.odt")));
  });
});

describe("Rich Text Format, opened as a Word document", () => {
  it("(.rtf)", async () => {
    await expectTheTestDocument(rtfToHtml(new TextDecoder("latin1").decode(fixture("formats.rtf"))));
  });

  it("refuses what isn't RTF", () => {
    expect(() => rtfToHtml("hello")).toThrow(/isn't a Rich Text Format/);
  });
});

describe("Word 97–2003 documents, opened as Word documents", () => {
  it("(.doc)", async () => {
    await expectTheTestDocument(docToHtml(fixture("formats.doc")));
  });

  it("refuse what isn't one, plainly", () => {
    expect(() => docToHtml(new TextEncoder().encode("hello"))).toThrow(/isn't a Word 97/);
  });
});
