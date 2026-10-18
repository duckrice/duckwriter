import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { stamped } from "../src/docx/properties";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { inspect } from "../src/ui/inspect";
import { isCslJson, parseCslJson } from "../src/docx/ris";
import { schema } from "../src/editor/schema";
import { lixLevel, readability } from "../src/ui/wordcount";

describe("CSL-JSON", () => {
  const csl = JSON.stringify([
    {
      id: "smith2020",
      type: "article-journal",
      title: "Rates of things",
      author: [{ family: "Smith", given: "Anna" }, { literal: "The Duck Group" }],
      issued: { "date-parts": [[2020, 5]] },
      "container-title": "Journal of Ducks",
      volume: 3,
      issue: "2",
      page: "1-10",
      DOI: "10.1/x",
    },
    { id: "b", type: "book", title: "A book", editor: [{ family: "Hansen", given: "Kari" }], issued: { raw: "2019" }, publisher: "Duck Press", "publisher-place": "Oslo" },
  ]);

  it("is told apart from other text", () => {
    expect(isCslJson(csl)).toBe(true);
    expect(isCslJson("TY  - JOUR")).toBe(false);
    expect(isCslJson("[1, 2]")).toBe(false);
  });

  it("reads items, and Zotero's answers", () => {
    const [a, b] = parseCslJson(csl);
    expect(a).toMatchObject({ tag: "smith2020", type: "JournalArticle", title: "Rates of things", year: "2020", container: "Journal of Ducks", volume: "3", issue: "2", pages: "1–10", doi: "10.1/x" });
    expect(a.authors).toEqual([{ last: "Smith", first: "Anna" }, { last: "The Duck Group", first: "" }]);
    expect(b).toMatchObject({ type: "Book", year: "2019", publisher: "Duck Press", city: "Oslo" });
    expect(b.authors).toEqual([{ last: "Hansen", first: "Kari" }]);
    // Zotero's local API: { items: [...] }, or items carrying csljson.
    expect(parseCslJson(JSON.stringify({ items: JSON.parse(csl) }))).toHaveLength(2);
    expect(parseCslJson([{ key: "X", csljson: JSON.parse(csl)[0] }])[0].title).toBe("Rates of things");
  });
});

describe("document properties", () => {
  it("are read, edited, stamped when saved, and read back", async () => {
    const { doc, context: ctx } = await readDocx(await blankDocx());
    expect(ctx.properties.created).not.toBe("");
    ctx.properties.title = "Rates & ducks";
    ctx.properties.keywords = "ducks; rates";
    ctx.savedBy = "Kari Nordmann";
    const saved = await writeDocx(doc, ctx);
    const core = await (await JSZip.loadAsync(saved)).file("docProps/core.xml")!.async("string");
    expect(core).toContain("<dc:title>Rates &amp; ducks</dc:title>");
    expect(core).toContain("<cp:lastModifiedBy>Kari Nordmann</cp:lastModifiedBy>");
    const again = (await readDocx(saved)).context.properties;
    expect(again).toMatchObject({ title: "Rates & ducks", keywords: "ducks; rates", lastModifiedBy: "Kari Nordmann", revision: String((Number(ctx.properties.revision) || 0) + 1) });
    // Writing doesn't change the document's own copy (a recovery copy isn't a save).
    expect(ctx.properties.lastModifiedBy).not.toBe("Kari Nordmann");
    expect(stamped(ctx.properties, "Kari Nordmann").lastModifiedBy).toBe("Kari Nordmann");
  });

  it("can leave the names out", async () => {
    const { doc, context: ctx } = await readDocx(await blankDocx());
    ctx.properties.creator = "Ola";
    ctx.properties.anonymous = true;
    ctx.savedBy = "Kari";
    const again = (await readDocx(await writeDocx(doc, ctx))).context.properties;
    expect(again.creator).toBe("");
    expect(again.lastModifiedBy).toBe("");
    // Check before sending offers it while names are there.
    ctx.properties.anonymous = false;
    const finding = inspect(doc, ctx).find((f) => f.text.includes("properties"));
    expect(finding?.text).toContain("Ola");
    finding!.action!.run();
    expect(ctx.properties).toMatchObject({ creator: "", anonymous: true });
  });

  it("are added to a file that has none", async () => {
    const { doc, context: ctx } = await readDocx(await blankDocx());
    ctx.zip.remove("docProps/core.xml");
    ctx.zip.file("_rels/.rels", (await ctx.zip.file("_rels/.rels")!.async("string")).replace(/<Relationship [^>]*core-properties[^>]*\/>/, ""));
    ctx.properties.title = "New";
    const zip = await JSZip.loadAsync(await writeDocx(doc, ctx));
    expect(await zip.file("docProps/core.xml")!.async("string")).toContain("<dc:title>New</dc:title>");
    expect(await zip.file("_rels/.rels")!.async("string")).toContain('Target="docProps/core.xml"');
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain('PartName="/docProps/core.xml"');
  });
});

describe("readability", () => {
  const p = (text: string, style: string | null = null) => schema.nodes.paragraph.create({ style }, schema.text(text));

  it("counts sentences and works out LIX, leaving headings out", () => {
    const doc = schema.nodes.doc.create(null, [
      p("Introduction to measurements", "Heading1"),
      p("The duck swam. It was fast! Was it cold? Yes."),
      p("Temperature measurements demonstrated considerable variability"),
    ]);
    const r = readability(doc);
    // Four sentences in the first paragraph, one (without a full stop) in the second.
    expect(r.sentences).toBe(5);
    // 10 + 5 words; long words (more than six letters): Temperature, measurements, demonstrated, considerable, variability.
    expect(r.wordsPerSentence).toBe(3);
    expect(r.longWords).toBeCloseTo(5 / 15);
    expect(r.lix).toBe(Math.round(3 + (5 / 15) * 100));
    expect(lixLevel(r.lix)).toBe("medium, like a newspaper");
    expect(lixLevel(20)).toBe("very easy");
    expect(lixLevel(60)).toBe("very hard, like technical writing");
  });

  it("is empty for an empty document", () => {
    expect(readability(schema.nodes.doc.create(null, [schema.nodes.paragraph.create()]))).toMatchObject({ sentences: 0, lix: 0 });
  });
});
