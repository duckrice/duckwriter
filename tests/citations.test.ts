import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { emptySource, makeTag, parseBibtex, type Bibliography, type Source } from "../src/docx/bibliography";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { bibliographyEntries, citationInstr, formatCitation, formatEntry, parseCitation } from "../src/editor/citations";
import { createState } from "../src/editor/editor";
import { schema } from "../src/editor/schema";

const BIBTEX = `
@article{smith2020,
  author = {Smith, Anna and Jones, Bob},
  title = {{Rates} of things},
  journal = {Journal of Ducks},
  year = 2020, volume = {3}, number = {2}, pages = {1--10},
  doi = {10.1000/ducks.3}
}
@book{haaland,
  author = {Kari Haaland and Per {\\O}stby and Ola Nordmann},
  title = {Fjordar og {\\aa}sar},
  publisher = {Samlaget}, address = {Oslo}, year = {2019}
}`;

const sources = (): Source[] => {
  const list = parseBibtex(BIBTEX);
  for (const s of list) s.tag = makeTag(s, list);
  return list;
};

const bib = (style: Bibliography["style"], list = sources()): Bibliography => ({ path: null, style, sources: list, version: 0, savedVersion: 0 });

describe("BibTeX", () => {
  it("is read into sources, Norwegian letters and all", () => {
    const [a, b] = sources();
    expect(a).toMatchObject({ tag: "Smi20", type: "JournalArticle", title: "Rates of things", container: "Journal of Ducks", volume: "3", issue: "2", pages: "1–10", doi: "10.1000/ducks.3" });
    expect(a.authors).toEqual([{ last: "Smith", first: "Anna" }, { last: "Jones", first: "Bob" }]);
    expect(b).toMatchObject({ tag: "Haa19", type: "Book", title: "Fjordar og åsar", publisher: "Samlaget", city: "Oslo" });
    expect(b.authors.map((p) => p.last)).toEqual(["Haaland", "Østby", "Nordmann"]);
  });
});

describe("citations", () => {
  it("read and write Word's CITATION instruction", () => {
    const c = { tags: ["Smi20", "Haa19"], page: "23" };
    expect(parseCitation(citationInstr(c))).toEqual(c);
    expect(parseCitation(" CITATION Smi20 \\l 1033 ")).toEqual({ tags: ["Smi20"], page: "" });
  });

  it("show in the chosen style", () => {
    const order = ["Haa19", "Smi20"];
    expect(formatCitation({ tags: ["Smi20"], page: "" }, bib("APA"), order)).toBe("(Smith & Jones, 2020)");
    expect(formatCitation({ tags: ["Haa19", "Smi20"], page: "12" }, bib("APA"), order)).toBe("(Haaland et al., 2019, p. 12; Smith & Jones, 2020)");
    expect(formatCitation({ tags: ["Smi20"], page: "" }, bib("Harvard"), order)).toBe("(Smith and Jones 2020)");
    expect(formatCitation({ tags: ["Smi20", "Haa19"], page: "" }, bib("IEEE"), order)).toBe("[2], [1]");
    expect(formatCitation({ tags: ["Nobody"], page: "" }, bib("APA"), order)).toBe("(Nobody?)");
  });

  it("make reference list entries, with italic titles", () => {
    const [article, book] = sources();
    const text = (pieces: { text: string; italic?: boolean }[]) => pieces.map((p) => (p.italic ? `*${p.text}*` : p.text)).join("");
    expect(text(formatEntry(article, "APA", 1))).toBe("Smith, A., & Jones, B. (2020). Rates of things. *Journal of Ducks*, *3*(2), 1–10. https://doi.org/10.1000/ducks.3");
    expect(text(formatEntry(book, "APA", 1))).toBe("Haaland, K., Østby, P., & Nordmann, O. (2019). *Fjordar og åsar.* Samlaget.");
    expect(text(formatEntry(article, "IEEE", 1))).toBe("[1]\tA. Smith and B. Jones, “Rates of things,” *Journal of Ducks*, vol. 3, no. 2, pp. 1–10, 2020, doi: 10.1000/ducks.3.");
  });
});

describe("in a document", () => {
  it("cites, lists the cited sources, and saves them for Word", async () => {
    const { context } = await readDocx(await blankDocx());
    context.bibliography = bib("APA");
    context.bibliography.version++;
    const cite = schema.nodes.field.create({ instr: citationInstr({ tags: ["Smi20"], page: "" }), result: "" });
    const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, [schema.text("As found "), cite, schema.text(".")]), schema.nodes.bibliography.create()]);
    let state = createState(doc, () => context);
    // Any change makes citations and the list follow.
    state = state.apply(state.tr.insertText("!", 1));
    let result = "";
    state.doc.descendants((n) => {
      if (n.type === schema.nodes.field) result = n.attrs.result;
    });
    expect(result).toBe("(Smith & Jones, 2020)");
    expect((state.doc.lastChild!.attrs.entries as unknown[]).length).toBe(1);
    expect(bibliographyEntries(state.doc, context.bibliography)).toHaveLength(1);

    const saved = await writeDocx(state.doc, context);
    const zip = await JSZip.loadAsync(saved);
    const item = await zip.file("customXml/item1.xml")!.async("string");
    expect(item).toContain('StyleName="APA"');
    expect(item).toContain("<b:Tag>Smi20</b:Tag>");
    expect(await zip.file("word/_rels/document.xml.rels")!.async("string")).toContain('Target="../customXml/item1.xml"');
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain("/customXml/itemProps1.xml");
    const xml = await zip.file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:docPartGallery w:val="Bibliographies"/>');
    expect(xml).toContain(" CITATION Smi20 \\l 1033 ");

    const again = await readDocx(saved);
    expect(again.context.bibliography.sources.map((s) => s.tag)).toEqual(["Smi20", "Haa19"]);
    expect(again.context.bibliography.sources[1].authors[1]).toEqual({ last: "Østby", first: "Per" });
    expect(again.doc.lastChild!.type).toBe(schema.nodes.bibliography);
    void emptySource;
  });
});
