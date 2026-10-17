import { readFileSync } from "node:fs";
import { join } from "node:path";
import JSZip from "jszip";
import { DOMParser as PMDOMParser, DOMSerializer } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { linkTypedAddress } from "../src/editor/autoformat";
import { createState } from "../src/editor/editor";
import { flattenPastedLists } from "../src/editor/paste";
import { schema } from "../src/editor/schema";
import { paperOf } from "../src/ui/pagesetup";
import { count } from "../src/ui/wordcount";

const part = async (bytes: Uint8Array, path: string) => (await JSZip.loadAsync(bytes)).file(path)!.async("string");

describe("page setup", () => {
  it("knows the paper, and saves a new size, orientation and margins for Word", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    expect(paperOf(context.page)).toEqual({ name: "A4", landscape: false });
    context.page = { ...context.page, width: context.page.height, height: context.page.width, marginLeft: 36, marginRight: 36 };
    context.pageChanged = true;
    const xml = await part(await writeDocx(doc, context), "word/document.xml");
    expect(xml).toMatch(/<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/);
    expect(xml).toMatch(/<w:pgMar [^>]*w:right="720"[^>]*w:left="720"/);
    expect(paperOf(context.page)).toEqual({ name: "A4", landscape: true });
  });
});

describe("copy and paste", () => {
  it("keeps everything when copying inside DuckWriter", () => {
    const para = schema.nodes.paragraph.create({ style: "Heading1", align: "center", spaceAfter: 6, extra: "<w:keepNext/>" }, [
      schema.text("Big ", [schema.marks.font.create({ family: "Cambria" }), schema.marks.size.create({ pt: 18 })]),
      schema.text("idea", [schema.marks.color.create({ color: "#c00000" }), schema.marks.charStyle.create({ id: "Strong" })]),
    ]);
    const doc = schema.nodes.doc.create(null, [para]);
    const dom = DOMSerializer.fromSchema(schema).serializeFragment(doc.content);
    const box = document.createElement("div");
    box.append(dom);
    const back = PMDOMParser.fromSchema(schema).parse(box);
    expect(back.eq(doc)).toBe(true);
  });

  it("reads Word's and the web's lists as list paragraphs", () => {
    const word = `<p class=MsoListParagraph style='mso-list:l0 level1 lfo1'><span style='mso-list:Ignore'>1.<span>  </span></span>First</p><p style='mso-list:l1 level2 lfo2'><span style='mso-list:Ignore'>·<span> </span></span>Dot</p>`;
    const box = document.createElement("div");
    box.innerHTML = flattenPastedLists(word);
    expect([...box.querySelectorAll("p")].map((p) => [p.dataset.list, p.dataset.level, p.textContent])).toEqual([
      ["number", "0", "First"],
      ["bullet", "1", "Dot"],
    ]);
    box.innerHTML = flattenPastedLists("<ul><li>One<ul><li>Inner</li></ul></li><li><p>Two</p></li></ul>");
    expect([...box.querySelectorAll("p")].map((p) => [p.dataset.list, p.dataset.level, p.textContent])).toEqual([
      ["bullet", "0", "One"],
      ["bullet", "1", "Inner"],
      ["bullet", "0", "Two"],
    ]);
    const parsed = PMDOMParser.fromSchema(schema).parse(box);
    expect(parsed.child(1).attrs).toMatchObject({ numId: "__bullet", ilvl: 1 });
  });

  it("takes headings, bold, links and superscripts from web pages", () => {
    const box = document.createElement("div");
    box.innerHTML = `<h2>Methods</h2><p>E = mc<sup>2</sup>, <b style="font-weight:normal">plain</b> <strong>bold</strong> <a href="https://duck.no">link</a> <a href="javascript:alert(1)">bad</a></p>`;
    const doc = PMDOMParser.fromSchema(schema).parse(box);
    expect(doc.child(0).attrs.style).toBe("Heading2");
    const marks = (text: string) => {
      let found: string[] = [];
      doc.descendants((n) => {
        if (n.isText && n.text === text) found = n.marks.map((m) => m.type.name);
      });
      return found;
    };
    expect(marks("2")).toEqual(["vertAlign"]);
    expect(marks("plain")).toEqual([]);
    expect(marks("bold")).toEqual(["bold"]);
    expect(marks("link")).toEqual(["link"]);
    expect(marks("bad")).toEqual([]);
  });
});

describe("AutoFormat as you type", () => {
  it("makes a finished web or e-mail address a link, without its punctuation", async () => {
    const { context } = await readDocx(await blankDocx());
    for (const [typed, linked, href] of [
      ["see www.duck.no.", "www.duck.no", "https://www.duck.no"],
      ["mail kari@duck.no", "kari@duck.no", "mailto:kari@duck.no"],
      ["just words", null, null],
    ]) {
      let state = createState(schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, schema.text(typed!))]), () => context);
      state = state.apply(state.tr.setSelection(TextSelection.atEnd(state.doc)));
      const tr = linkTypedAddress(state, context);
      if (!linked) {
        expect(tr).toBeNull();
        continue;
      }
      const after = state.apply(tr!);
      let text = "";
      let link = "";
      after.doc.descendants((n) => {
        const m = schema.marks.link.isInSet(n.marks);
        if (n.isText && m) {
          text += n.text;
          link = m.attrs.href;
        }
      });
      expect([text, link]).toEqual([linked, href]);
    }
  });
});

describe("word count", () => {
  it("counts words, characters and paragraphs, of the selection too, leaving out deleted text", () => {
    const del = schema.marks.deletion.create({ id: "1", author: "A" });
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(null, [schema.text("One two "), schema.text("gone ", [del]), schema.text("three.")]),
      schema.nodes.paragraph.create(),
      schema.nodes.paragraph.create(null, schema.text("Fire fem")),
    ]);
    expect(count(doc)).toEqual({ words: 5, characters: 22, charactersNoSpaces: 19, paragraphs: 2 });
    expect(count(doc, 1, 8).words).toBe(2);
  });
});

describe("pictures in headers", () => {
  it("are shown, and saved unchanged", async () => {
    const bytes = new Uint8Array(readFileSync(join(__dirname, "..", "samples", "letterhead.docx")));
    const { doc, context } = await readDocx(bytes);
    const header = context.stories.find((s) => s.kind === "header")!;
    const kinds: string[] = [];
    header.doc.descendants((n) => void kinds.push(n.type.name));
    expect(kinds).toContain("image");
    const saved = await writeDocx(doc, context);
    expect(await part(saved, header.path!)).toBe(await part(bytes, header.path!));
  });
});
