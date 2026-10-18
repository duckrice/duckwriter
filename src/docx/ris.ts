// Sources from RIS (EndNote, Mendeley, Zotero, most journal sites: "Export
// citation") and CSL-JSON (Zotero, Mendeley, citation.js), and out as BibTeX,
// for LaTeX or another reference manager.

import { emptySource, type Source, type SourceType } from "./bibliography";

const CSL_TYPES: Record<string, SourceType> = {
  "article-journal": "JournalArticle",
  "article-magazine": "JournalArticle",
  "article-newspaper": "JournalArticle",
  article: "JournalArticle",
  book: "Book",
  chapter: "BookSection",
  "paper-conference": "ConferenceProceedings",
  report: "Report",
  thesis: "Report",
  webpage: "InternetSite",
  "post-weblog": "InternetSite",
};

interface CslName {
  family?: string;
  given?: string;
  literal?: string;
}

/** Whether text looks like CSL-JSON: a list of items (or one) with a type and a title. */
export function isCslJson(input: string): boolean {
  const t = input.trim();
  if (!t.startsWith("[") && !t.startsWith("{")) return false;
  try {
    const items = JSON.parse(t);
    return (Array.isArray(items) ? items : [items]).some((i) => i && typeof i === "object" && "type" in i && ("title" in i || "author" in i));
  } catch {
    return false;
  }
}

/** Reads CSL-JSON items (and Zotero's local API, whose items carry them as `csljson`). */
export function parseCslJson(input: string | unknown): Source[] {
  let data = typeof input === "string" ? JSON.parse(input) : input;
  // Zotero's format=csljson: { items: [...] }.
  if (data && typeof data === "object" && !Array.isArray(data) && Array.isArray((data as { items?: unknown }).items)) data = (data as { items: unknown[] }).items;
  const items: Record<string, unknown>[] = (Array.isArray(data) ? data : [data]).map((i) => (i && typeof i === "object" && "csljson" in i ? (i as { csljson: Record<string, unknown> }).csljson : i));
  return items
    .filter((i) => i && typeof i === "object")
    .map((i) => {
      const str = (v: unknown) => (typeof v === "string" || typeof v === "number" ? String(v).trim() : "");
      const s = emptySource();
      s.type = CSL_TYPES[str(i.type)] ?? "Misc";
      const names = ((i.author ?? i.editor ?? []) as CslName[]).filter(Boolean);
      s.authors = names.map((n) => (n.literal ? { last: n.literal, first: "" } : { last: n.family ?? "", first: n.given ?? "" }));
      s.title = str(i.title);
      const issued = i.issued as { "date-parts"?: unknown[][]; raw?: string } | undefined;
      s.year = str(issued?.["date-parts"]?.[0]?.[0]) || (/\d{4}/.exec(issued?.raw ?? "") ?? [""])[0];
      s.container = str(i["container-title"]) || str(i["collection-title"]);
      s.volume = str(i.volume);
      s.issue = str(i.issue);
      s.pages = str(i.page).replace(/-+/g, "–");
      s.publisher = str(i.publisher);
      s.city = str(i["publisher-place"]);
      s.url = str(i.URL);
      s.doi = str(i.DOI);
      s.tag = str(i["citation-key"]) || str(i.id);
      return s;
    });
}

const RIS_TYPES: Record<string, SourceType> = {
  JOUR: "JournalArticle",
  JFULL: "JournalArticle",
  MGZN: "JournalArticle",
  EJOUR: "JournalArticle",
  BOOK: "Book",
  EBOOK: "Book",
  EDBOOK: "Book",
  CHAP: "BookSection",
  ECHAP: "BookSection",
  CONF: "ConferenceProceedings",
  CPAPER: "ConferenceProceedings",
  RPRT: "Report",
  THES: "Report",
  ELEC: "InternetSite",
  WEB: "InternetSite",
  BLOG: "InternetSite",
};

const person = (name: string) => {
  if (name.includes(",")) {
    const [last, first] = name.split(",", 2);
    return { last: last.trim(), first: (first ?? "").trim() };
  }
  const parts = name.trim().split(/\s+/);
  return { last: parts.pop() ?? "", first: parts.join(" ") };
};

/** Whether text looks like RIS ("TY  - JOUR" lines). */
export const isRis = (input: string) => /^TY {2}- /m.test(input);

/** Reads RIS records: TY starts one, ER ends it. */
export function parseRis(input: string): Source[] {
  const out: Source[] = [];
  let fields: [string, string][] = [];
  const finish = () => {
    if (!fields.length) return;
    const get = (...tags: string[]) => tags.map((t) => fields.find(([k]) => k === t)?.[1]).find((v) => v) ?? "";
    const s = emptySource();
    s.type = RIS_TYPES[get("TY").toUpperCase()] ?? "Misc";
    s.authors = fields.filter(([k]) => k === "AU" || k === "A1").map(([, v]) => person(v));
    if (!s.authors.length) s.authors = fields.filter(([k]) => k === "A2" || k === "ED").map(([, v]) => person(v));
    s.title = get("TI", "T1", "CT");
    s.year = (/\d{4}/.exec(get("PY", "Y1", "DA")) ?? [""])[0];
    s.container = get("T2", "JO", "JF", "JA", "BT", "T3");
    s.volume = get("VL");
    s.issue = get("IS", "CP");
    const start = get("SP");
    const end = get("EP");
    s.pages = start.includes("-") ? start.replace(/-+/g, "–") : end ? `${start}–${end}` : start;
    s.publisher = get("PB");
    s.city = get("CY", "PP");
    s.url = get("UR", "L2");
    s.doi = get("DO").replace(/^https?:\/\/(dx\.)?doi\.org\//i, "");
    s.tag = get("ID");
    out.push(s);
    fields = [];
  };
  for (const raw of input.split(/\r?\n/)) {
    const m = /^([A-Z][A-Z0-9]) {2}-(?: (.*))?$/.exec(raw.trimEnd());
    if (!m) {
      // A long value can go on over the next lines.
      if (fields.length && raw.trim()) fields[fields.length - 1][1] += ` ${raw.trim()}`;
      continue;
    }
    const [, tag, value = ""] = m;
    if (tag === "TY") finish();
    if (tag === "ER") {
      finish();
      continue;
    }
    fields.push([tag, value.trim()]);
  }
  finish();
  return out;
}

const BIBTEX_TYPE: Record<SourceType, string> = {
  JournalArticle: "article",
  Book: "book",
  BookSection: "incollection",
  ConferenceProceedings: "inproceedings",
  Report: "techreport",
  InternetSite: "misc",
  Misc: "misc",
};

/** A value in braces, with the characters TeX treats specially escaped. */
function texValue(s: string): string {
  return s
    .replace(/\\/g, "\\textbackslash{}")
    .replace(/([&%$#_])/g, "\\$1")
    .replace(/[{}]/g, (c) => `\\${c}`)
    .replace(/~/g, "\\textasciitilde{}")
    .replace(/\^/g, "\\textasciicircum{}");
}

/** The sources as a .bib file; each entry's key is its tag, as cited. */
export function toBibtex(sources: Source[]): string {
  return sources
    .map((s) => {
      const type = BIBTEX_TYPE[s.type] ?? "misc";
      const container = type === "article" ? "journal" : type === "incollection" || type === "inproceedings" ? "booktitle" : type === "misc" ? "howpublished" : "series";
      const fields: [string, string][] = [
        ["author", s.authors.map((a) => (a.first ? `${a.last}, ${a.first}` : a.last)).join(" and ")],
        ["title", s.title],
        [container, s.container],
        ["year", s.year],
        ["volume", s.volume],
        ["number", s.issue],
        ["pages", s.pages.replace(/[–—]/g, "--")],
        [type === "techreport" ? "institution" : "publisher", s.publisher],
        ["address", s.city],
        ["doi", s.doi],
        ["url", s.url],
      ];
      const body = fields
        .filter(([, v]) => v)
        // The title keeps its capitals ({{...}}): BibTeX styles would make "DNA" "dna".
        .map(([k, v]) => `  ${k} = {${k === "title" ? `{${texValue(v)}}` : k === "url" || k === "doi" ? v : texValue(v)}}`)
        .join(",\n");
      return `@${type}{${s.tag || "source"},\n${body}\n}`;
    })
    .join("\n\n")
    .concat(sources.length ? "\n" : "");
}
