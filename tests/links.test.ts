import JSZip from "jszip";
import { TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { createState } from "../src/editor/editor";
import { schema } from "../src/editor/schema";
import { linkAt, normalizeAddress } from "../src/ui/links";

describe("links", () => {
  it("turn typed addresses into web or e-mail links", () => {
    expect(normalizeAddress("duck.no")).toBe("https://duck.no");
    expect(normalizeAddress(" kari@duck.no ")).toBe("mailto:kari@duck.no");
    expect(normalizeAddress("http://x.org/a")).toBe("http://x.org/a");
  });

  it("are found whole around the cursor, across formatting", async () => {
    const { context } = await readDocx(await blankDocx());
    const link = schema.marks.link.create({ href: "https://duck.no" });
    const para = schema.nodes.paragraph.create(null, [schema.text("See "), schema.text("the ", [link]), schema.text("duck", [link, schema.marks.bold.create()]), schema.text(" site")]);
    let state = createState(schema.nodes.doc.create(null, [para]), () => context);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 7)));
    const found = linkAt(state)!;
    expect(state.doc.textBetween(found.from, found.to)).toBe("the duck");
  });

  it("to a heading are saved with Word's anchor", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    const link = schema.marks.link.create({ href: "", anchor: "_Ref1" });
    const d = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, schema.text("Go there", [link]))]);
    const saved = await writeDocx(d, context);
    const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:hyperlink w:anchor="_Ref1" w:history="1">');
    void doc;
  });
});
