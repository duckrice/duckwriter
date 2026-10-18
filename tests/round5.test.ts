import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { schema } from "../src/editor/schema";
import { inspect } from "../src/ui/inspect";

describe("check before sending", () => {
  it("finds undescribed pictures, broken references, missing sources and empty headings", async () => {
    const { context } = await readDocx(await blankDocx());
    const picture = schema.nodes.image.create({ media: "x", width: 10, height: 10 });
    const ref = schema.nodes.field.create({ instr: " REF _Ref999 \h ", result: "Figure 9" });
    const cite = schema.nodes.field.create({ instr: " CITATION Nobody20 \l 1033 ", result: "(?)" });
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create({ style: "Heading1" }),
      schema.nodes.paragraph.create(null, [schema.text("See "), ref, schema.text(" and "), cite, picture]),
    ]);
    const found = inspect(doc, context).map((f) => f.text);
    expect(found.some((t) => /1 picture has no description/.test(t))).toBe(true);
    expect(found.some((t) => /1 reference or link goes to a place/.test(t))).toBe(true);
    expect(found.some((t) => /1 citation refers to a source/.test(t))).toBe(true);
    expect(found.some((t) => /1 heading is empty/.test(t))).toBe(true);
  });

  it("finds nothing in a clean document", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    expect(inspect(doc, context)).toEqual([]);
  });
});

describe("picture descriptions", () => {
  it("are saved as Word's alt text and read back", async () => {
    const { context } = await readDocx(await blankDocx());
    context.media.set("pic", { bytes: new Uint8Array([137, 80, 78, 71]), ext: "png", rid: null });
    const doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.nodes.image.create({ media: "pic", width: 20, height: 10, alt: "A duck on a lake" })));
    const saved = await writeDocx(doc, context);
    const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("string");
    expect(xml).toContain('descr="A duck on a lake"');
    let alt = "";
    (await readDocx(saved)).doc.descendants((n) => {
      if (n.type === schema.nodes.image) alt = n.attrs.alt;
    });
    expect(alt).toBe("A duck on a lake");
  });
});

describe("AutoCorrect", () => {
  it("corrects slips and symbols as the word is finished, keeping capitals", async () => {
    const { autoCorrect, addCorrection } = await import("../src/editor/autocorrect");
    const { createState } = await import("../src/editor/editor");
    const { TextSelection } = await import("prosemirror-state");
    const { context } = await readDocx(await blankDocx());
    const fix = (text: string, typed = " ") => {
      let state = createState(schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text(text))), () => context);
      state = state.apply(state.tr.setSelection(TextSelection.atEnd(state.doc)));
      const tr = autoCorrect(state, typed);
      return tr ? state.apply(tr).doc.textContent : null;
    };
    expect(fix("I saw teh")).toBe("I saw the");
    expect(fix("Teh")).toBe("The");
    expect(fix("we recieve", ".")).toBe("we receive");
    expect(fix("Duck (c)")).toBe("Duck ©");
    expect(fix("wait...")).toBe("wait…");
    expect(fix("one --")).toBe("one –");
    expect(fix("the")).toBeNull();
    addCorrection("ducwriter", "DuckWriter");
    expect(fix("I use ducwriter")).toBe("I use DuckWriter");
  });
});
