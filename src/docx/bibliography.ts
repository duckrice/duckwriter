// The sources a document cites, kept as Word keeps them: a customXml part
// holding <b:Sources>, which Word's References tab reads and writes too.

import type JSZip from "jszip";
import { escapeText } from "./xml";

export const BIB_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

export type SourceType = "JournalArticle" | "Book" | "BookSection" | "ConferenceProceedings" | "Report" | "InternetSite" | "Misc";

export const SOURCE_TYPES: { type: SourceType; label: string }[] = [
  { type: "JournalArticle", label: "Journal article" },
  { type: "Book", label: "Book" },
  { type: "BookSection", label: "Chapter in a book" },
  { type: "ConferenceProceedings", label: "Conference paper" },
  { type: "Report", label: "Report" },
  { type: "InternetSite", label: "Web page" },
  { type: "Misc", label: "Other" },
];

export interface Person {
  last: string;
  first: string;
}

export interface Source {
  /** Word's short name for it, used by citations: "Smi20". */
  tag: string;
  type: SourceType;
  authors: Person[];
  title: string;
  year: string;
  /** The journal, the book a chapter is in, or the conference. */
  container: string;
  volume: string;
  issue: string;
  pages: string;
  publisher: string;
  city: string;
  url: string;
  doi: string;
  /** The source as read; its other details are kept when it is written back. */
  xml: string | null;
}

export type CitationStyle = "APA" | "Harvard" | "IEEE";

export interface Bibliography {
  /** The customXml part, or null until sources are first saved. */
  path: string | null;
  style: CitationStyle;
  sources: Source[];
  /** Counts edits, for telling whether there is something to save. */
  version: number;
  savedVersion: number;
}

export const emptySource = (): Source => ({ tag: "", type: "JournalArticle", authors: [], title: "", year: "", container: "", volume: "", issue: "", pages: "", publisher: "", city: "", url: "", doi: "", xml: null });

const STYLES: Record<CitationStyle, { file: string; name: string }> = {
  APA: { file: "\\APASixthEditionOfficeOnline.xsl", name: "APA" },
  Harvard: { file: "\\HarvardAnglia2008OfficeOnline.xsl", name: "Harvard - Anglia" },
  IEEE: { file: "\\IEEE2006OfficeOnline.xsl", name: "IEEE" },
};

function text(el: Element | null | undefined, name: string): string {
  const found = el ? Array.from(el.children).find((c) => c.localName === name) : undefined;
  return found?.textContent?.trim() ?? "";
}

const childEl = (el: Element | null | undefined, name: string) => (el ? (Array.from(el.children).find((c) => c.localName === name) ?? null) : null);

/** Finds and reads the document's sources. */
export async function readBibliography(zip: JSZip): Promise<Bibliography> {
  const bib: Bibliography = { path: null, style: "APA", sources: [], version: 0, savedVersion: 0 };
  for (const path of Object.keys(zip.files).filter((p) => /^customXml\/item\d+\.xml$/i.test(p))) {
    const xml = await zip.file(path)!.async("string");
    if (!xml.includes(BIB_NS)) continue;
    const root = new DOMParser().parseFromString(xml, "application/xml").documentElement;
    if (root.localName !== "Sources") continue;
    bib.path = path;
    const styleName = root.getAttribute("StyleName") ?? root.getAttribute("SelectedStyle") ?? "";
    bib.style = /ieee/i.test(styleName) ? "IEEE" : /harvard/i.test(styleName) ? "Harvard" : "APA";
    const serializer = new XMLSerializer();
    for (const el of Array.from(root.children).filter((c) => c.localName === "Source")) {
      const type: string = text(el, "SourceType");
      const people = childEl(childEl(childEl(childEl(el, "Author"), "Author"), "NameList"), "Person") ? Array.from(childEl(childEl(childEl(el, "Author"), "Author"), "NameList")!.children) : [];
      bib.sources.push({
        tag: text(el, "Tag"),
        type: SOURCE_TYPES.some((t) => t.type === type) ? (type as SourceType) : type === "ArticleInAPeriodical" ? "JournalArticle" : type === "DocumentFromInternetSite" ? "InternetSite" : "Misc",
        authors: people.map((p) => ({ last: text(p, "Last"), first: [text(p, "First"), text(p, "Middle")].filter(Boolean).join(" ") })),
        title: text(el, "Title"),
        year: text(el, "Year"),
        container: text(el, "JournalName") || text(el, "BookTitle") || text(el, "ConferenceName") || text(el, "PeriodicalTitle") || text(el, "InternetSiteTitle"),
        volume: text(el, "Volume"),
        issue: text(el, "Issue"),
        pages: text(el, "Pages"),
        publisher: text(el, "Publisher"),
        city: text(el, "City"),
        url: text(el, "URL"),
        doi: text(el, "DOI"),
        xml: serializer.serializeToString(el),
      });
    }
    break;
  }
  return bib;
}

const CONTAINER: Partial<Record<SourceType, string>> = { JournalArticle: "JournalName", BookSection: "BookTitle", ConferenceProceedings: "ConferenceName", InternetSite: "InternetSiteTitle" };

/** A source as Word writes it. */
export function sourceXml(s: Source, order: number): string {
  const el = (name: string, value: string) => (value ? `<b:${name}>${escapeText(value)}</b:${name}>` : "");
  const people = s.authors.map((p) => `<b:Person>${el("Last", p.last)}${el("First", p.first)}</b:Person>`).join("");
  return (
    `<b:Source>${el("Tag", s.tag)}${el("SourceType", s.type)}` +
    (people ? `<b:Author><b:Author><b:NameList>${people}</b:NameList></b:Author></b:Author>` : "") +
    el("Title", s.title) +
    (CONTAINER[s.type] ? el(CONTAINER[s.type]!, s.container) : "") +
    el("Year", s.year) +
    el("Volume", s.volume) +
    el("Issue", s.issue) +
    el("Pages", s.pages) +
    el("Publisher", s.publisher) +
    el("City", s.city) +
    el("URL", s.url) +
    el("DOI", s.doi) +
    `<b:RefOrder>${order}</b:RefOrder></b:Source>`
  );
}

/** Writes the sources back (adding the part, as Word would, the first time). */
export async function writeBibliography(zip: JSZip, bib: Bibliography, addRel: (target: string) => void, addOverride: (part: string, type: string) => Promise<void>) {
  if (bib.version === bib.savedVersion) return;
  if (!bib.path) {
    if (!bib.sources.length) return;
    let n = 1;
    while (zip.file(`customXml/item${n}.xml`)) n++;
    bib.path = `customXml/item${n}.xml`;
    const guid = `{${crypto.randomUUID().toUpperCase()}}`;
    zip.file(
      `customXml/itemProps${n}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="${guid}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${BIB_NS}"/></ds:schemaRefs></ds:datastoreItem>`,
    );
    zip.file(
      `customXml/_rels/item${n}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps${n}.xml"/></Relationships>`,
    );
    await addOverride(`/customXml/itemProps${n}.xml`, "application/vnd.openxmlformats-officedocument.customXmlProperties+xml");
    addRel(`../customXml/item${n}.xml`);
  }
  const style = STYLES[bib.style];
  const body = bib.sources.map((s, i) => sourceXml(s, i + 1)).join("");
  zip.file(bib.path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<b:Sources SelectedStyle="${style.file}" StyleName="${style.name}" Version="6" xmlns:b="${BIB_NS}" xmlns="${BIB_NS}">${body}</b:Sources>`);
  bib.savedVersion = bib.version;
}

// ---- BibTeX --------------------------------------------------------------------------------

const BIBTEX_TYPES: Record<string, SourceType> = {
  article: "JournalArticle",
  book: "Book",
  incollection: "BookSection",
  inbook: "BookSection",
  inproceedings: "ConferenceProceedings",
  conference: "ConferenceProceedings",
  techreport: "Report",
  report: "Report",
  online: "InternetSite",
  misc: "Misc",
};

/** Sources from BibTeX (as reference managers and Google Scholar export it). */
export function parseBibtex(input: string): Source[] {
  const out: Source[] = [];
  const entry = /@(\w+)\s*\{\s*([^,\s]*)\s*,/g;
  let m: RegExpExecArray | null;
  while ((m = entry.exec(input))) {
    // The entry's fields run to its matching closing brace.
    let depth = 1;
    let i = entry.lastIndex;
    for (; i < input.length && depth > 0; i++) {
      if (input[i] === "{") depth++;
      else if (input[i] === "}") depth--;
    }
    const body = input.slice(entry.lastIndex, i - 1);
    entry.lastIndex = i;
    const fields = new Map<string, string>();
    const field = /(\w+)\s*=\s*/g;
    let f: RegExpExecArray | null;
    while ((f = field.exec(body))) {
      let at = field.lastIndex;
      let value = "";
      if (body[at] === "{" || body[at] === '"') {
        const close = body[at] === "{" ? "}" : '"';
        let d = 1;
        let j = at + 1;
        for (; j < body.length && d > 0; j++) {
          if (close === "}" && body[j] === "{") d++;
          else if (body[j] === close) d--;
        }
        value = body.slice(at + 1, j - 1);
        at = j;
      } else {
        const end = body.indexOf(",", at);
        value = body.slice(at, end < 0 ? body.length : end);
        at = end < 0 ? body.length : end;
      }
      field.lastIndex = at;
      fields.set(f[1].toLowerCase(), cleanTex(value.trim()));
    }
    const source = emptySource();
    source.type = BIBTEX_TYPES[m[1].toLowerCase()] ?? "Misc";
    source.authors = (fields.get("author") ?? fields.get("editor") ?? "")
      .split(/\s+and\s+/i)
      .filter(Boolean)
      .map((name) => {
        if (name.includes(",")) {
          const [last, first] = name.split(",", 2);
          return { last: last.trim(), first: (first ?? "").trim() };
        }
        const parts = name.trim().split(/\s+/);
        return { last: parts.pop() ?? "", first: parts.join(" ") };
      });
    source.title = fields.get("title") ?? "";
    source.year = fields.get("year") ?? (fields.get("date") ?? "").slice(0, 4);
    source.container = fields.get("journal") ?? fields.get("journaltitle") ?? fields.get("booktitle") ?? "";
    source.volume = fields.get("volume") ?? "";
    source.issue = fields.get("number") ?? fields.get("issue") ?? "";
    source.pages = (fields.get("pages") ?? "").replace(/-+/g, "–");
    source.publisher = fields.get("publisher") ?? fields.get("institution") ?? fields.get("organization") ?? "";
    source.city = fields.get("address") ?? "";
    source.url = fields.get("url") ?? "";
    source.doi = fields.get("doi") ?? "";
    source.tag = m[2] || "";
    out.push(source);
  }
  return out;
}

function cleanTex(value: string): string {
  return value
    .replace(/\\&/g, "&")
    .replace(/\\'\{?([aeiouAEIOU])\}?/g, (_m, c: string) => c + "́")
    .replace(/\\"\{?([aeiouAEIOU])\}?/g, (_m, c: string) => c + "̈")
    .replace(/\{\\o\}|\\o\b/g, "ø")
    .replace(/\{\\O\}|\\O\b/g, "Ø")
    .replace(/\{\\aa\}|\\aa\b/g, "å")
    .replace(/\{\\AA\}|\\AA\b/g, "Å")
    .replace(/\{\\ae\}|\\ae\b/g, "æ")
    .replace(/\{\\AE\}|\\AE\b/g, "Æ")
    .replace(/[{}]/g, "")
    .replace(/\s+/g, " ")
    .normalize("NFC");
}

/** A short tag for a new source, as Word makes them: "Smi20", unique in the list. */
export function makeTag(source: Source, others: Source[]): string {
  const base = `${(source.authors[0]?.last || source.title || "Src").replace(/[^\p{L}\p{N}]/gu, "").slice(0, 3)}${source.year.slice(-2)}` || "Src";
  let tag = base;
  for (let n = 1; others.some((o) => o !== source && o.tag === tag); n++) tag = `${base}${String.fromCharCode(96 + n)}`;
  return tag;
}
