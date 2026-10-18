import { describe, expect, it } from "vitest";
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
