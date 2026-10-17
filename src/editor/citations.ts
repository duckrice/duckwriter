// Citations in the text (Word's CITATION fields: "(Smith, 2020)" or "[1]") and
// the reference list (a BIBLIOGRAPHY), in APA, Harvard or IEEE style. Both
// follow the sources and the order of citing as they change.

import type { Node } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import type { Bibliography, CitationStyle, Person, Source } from "../docx/bibliography";
import type { DocContext } from "../docx/model";
import { fieldKind, schema } from "./schema";

export interface Citation {
  tags: string[];
  /** A page or pages ("23", "23–25"), for the first source. */
  page: string;
}

/** Reads a CITATION field's instruction: " CITATION Smi20 \l 1033 \m Jon19 \p 23 ". */
export function parseCitation(instr: string): Citation {
  const words = instr.trim().split(/\s+/);
  const tags = words[1] ? [words[1]] : [];
  let page = "";
  for (let i = 2; i < words.length; i++) {
    if (words[i] === "\\m" && words[i + 1]) tags.push(words[++i]);
    else if (words[i] === "\\p" && words[i + 1]) page = words[++i];
    else if (/^\\[lfsyvt]$/.test(words[i])) i++;
  }
  return { tags, page };
}

export function citationInstr(c: Citation): string {
  return ` CITATION ${c.tags[0]} \\l 1033 ${c.tags
    .slice(1)
    .map((t) => `\\m ${t} `)
    .join("")}${c.page ? `\\p ${c.page} ` : ""}`;
}

const isCitation = (node: Node) => node.type === schema.nodes.field && fieldKind(node.attrs.instr) === "CITATION";

/** The sources in the order they are first cited (IEEE numbers them so). */
export function citedOrder(doc: Node): string[] {
  const order: string[] = [];
  doc.descendants((node) => {
    if (isCitation(node)) for (const tag of parseCitation(node.attrs.instr).tags) if (!order.includes(tag)) order.push(tag);
  });
  return order;
}

// ---- in the text ----------------------------------------------------------------------------

function authorsShort(authors: Person[], style: CitationStyle): string {
  const names = authors.map((a) => a.last).filter(Boolean);
  if (!names.length) return "";
  if (names.length === 1) return names[0];
  if (names.length === 2) return `${names[0]} ${style === "APA" ? "&" : "and"} ${names[1]}`;
  return `${names[0]} et al.`;
}

/** What a citation shows in the text. */
export function formatCitation(c: Citation, bib: Bibliography, order: string[]): string {
  const found = c.tags.map((t) => bib.sources.find((s) => s.tag === t));
  if (bib.style === "IEEE") {
    const parts = c.tags.map((t, i) => {
      const n = order.indexOf(t) + 1;
      return n ? `${n}${i === 0 && c.page ? `, p. ${c.page}` : ""}` : "?";
    });
    return parts.map((p) => `[${p}]`).join(", ");
  }
  if (bib.style === "Vancouver") {
    const numbers = c.tags.map((t) => order.indexOf(t) + 1).sort((a, b) => a - b);
    // Runs of three or more as a range: (1–3, 5).
    const parts: string[] = [];
    for (let i = 0; i < numbers.length; ) {
      let j = i;
      while (j + 1 < numbers.length && numbers[j + 1] === numbers[j] + 1) j++;
      parts.push(numbers[i] === 0 ? "?" : j - i >= 2 ? `${numbers[i]}–${numbers[j]}` : numbers.slice(i, j + 1).join(", "));
      i = j + 1;
    }
    return `(${parts.join(", ")}${c.page ? `, p. ${c.page}` : ""})`;
  }
  const parts = found.map((s, i) => {
    if (!s) return `${c.tags[i]}?`;
    const who = authorsShort(s.authors, bib.style) || s.title;
    const year = s.year || "n.d.";
    const page = i === 0 && c.page ? `, ${/[–-]/.test(c.page) ? "pp." : "p."} ${c.page}` : "";
    if (bib.style === "Chicago") return `${who} ${year}${i === 0 && c.page ? `, ${c.page}` : ""}`;
    return bib.style === "APA" ? `${who}, ${year}${page}` : `${who} ${year}${page}`;
  });
  return `(${parts.join("; ")})`;
}

// ---- the reference list -----------------------------------------------------------------------

/** A piece of an entry: italic for titles of books and journals. */
export interface Piece {
  text: string;
  italic?: boolean;
}

const initials = (first: string, dots = true) =>
  first
    .split(/[\s-]+/)
    .filter(Boolean)
    .map((n) => `${n[0].toUpperCase()}${dots ? "." : ""}`)
    .join(" ");

function namesApa(authors: Person[]): string {
  const names = authors.map((a) => [a.last, initials(a.first)].filter(Boolean).join(", "));
  if (names.length <= 1) return names[0] ?? "";
  if (names.length > 20) return `${names.slice(0, 19).join(", ")}, … ${names[names.length - 1]}`;
  return `${names.slice(0, -1).join(", ")}, & ${names[names.length - 1]}`;
}

function namesIeee(authors: Person[]): string {
  const names = authors.map((a) => [initials(a.first), a.last].filter(Boolean).join(" "));
  if (names.length > 6) return `${names[0]} et al.`;
  if (names.length <= 2) return names.join(" and ");
  return `${names.slice(0, -1).join(", ")}, and ${names[names.length - 1]}`;
}

const doiLink = (s: Source) => (s.doi ? `https://doi.org/${s.doi.replace(/^https?:\/\/(dx\.)?doi\.org\//, "")}` : s.url);
const end = (text: string) => (/[.?!]$/.test(text) ? text : `${text}.`);

/** One source in the reference list. */
export function formatEntry(s: Source, style: CitationStyle, n: number): Piece[] {
  const pieces: Piece[] = [];
  const add = (text: string, italic = false) => text && pieces.push({ text, italic });
  const pages = s.pages.replace(/-+/g, "–");
  if (style === "Vancouver") {
    // Smith A, Jones B. Title. Journal. 2020;3(2):1–10. doi:…
    const names = s.authors.map((a) => `${a.last} ${initials(a.first, false).replace(/\s+/g, "")}`.trim());
    add(`${n}.\t`);
    if (names.length) add(`${names.length > 6 ? `${names.slice(0, 6).join(", ")}, et al` : names.join(", ")}. `);
    add(end(s.title));
    if (s.type === "JournalArticle") add(` ${end(s.container)} ${s.year}${s.volume ? `;${s.volume}` : ""}${s.issue ? `(${s.issue})` : ""}${pages ? `:${pages}` : ""}.`);
    else if (s.type === "BookSection" || s.type === "ConferenceProceedings") add(` In: ${end(s.container)} ${[s.city && s.publisher ? `${s.city}: ${s.publisher}` : s.publisher, s.year].filter(Boolean).join("; ")}.${pages ? ` p. ${pages}.` : ""}`);
    else add(` ${[s.city && s.publisher ? `${s.city}: ${s.publisher}` : s.publisher, s.year].filter(Boolean).join("; ")}.`);
    if (s.doi) add(` doi:${s.doi}`);
    else if (s.url) add(` Available from: ${s.url}`);
    return pieces;
  }
  if (style === "Chicago") {
    // Smith, Anna, and Bob Jones. 2020. "Title." Journal 3 (2): 1–10. https://doi.org/…
    const names = s.authors.map((a, i) => (i === 0 ? [a.last, a.first].filter(Boolean).join(", ") : [a.first, a.last].filter(Boolean).join(" ")));
    const who = names.length <= 1 ? (names[0] ?? "") : names.length === 2 ? `${names[0]}, and ${names[1]}` : `${names.slice(0, -1).join(", ")}, and ${names[names.length - 1]}`;
    if (who) add(`${end(who)} `);
    add(`${s.year || "n.d."}. `);
    if (s.type === "Book" || s.type === "Report" || s.type === "Misc") {
      add(end(s.title), true);
      add(` ${end([s.city, s.publisher].filter(Boolean).join(": "))}`.replace(/^ \.$/, ""));
    } else {
      add(`“${end(s.title)}”`);
      if (s.container) {
        add(s.type === "JournalArticle" ? " " : " In ");
        add(s.container, true);
      }
      if (s.type === "JournalArticle") add(`${s.volume ? ` ${s.volume}` : ""}${s.issue ? ` (${s.issue})` : ""}${pages ? `: ${pages}` : ""}.`);
      else add(`${pages ? `, ${pages}` : ""}.${s.publisher ? ` ${end([s.city, s.publisher].filter(Boolean).join(": "))}` : ""}`);
    }
    const link = doiLink(s);
    if (link) add(` ${link}.`);
    return pieces;
  }
  if (style === "IEEE") {
    add(`[${n}]\t`);
    if (s.authors.length) add(`${namesIeee(s.authors)}, `);
    if (s.type === "Book" || s.type === "Report") {
      add(s.title, true);
      add(`. ${[s.city && s.publisher ? `${s.city}: ${s.publisher}` : s.publisher, s.year].filter(Boolean).join(", ")}.`);
    } else {
      add(`“${s.title},” `);
      if (s.container) {
        if (s.type === "ConferenceProceedings" || s.type === "BookSection") add("in ");
        add(s.container, true);
        add(", ");
      }
      add(
        `${[s.volume && `vol. ${s.volume}`, s.issue && `no. ${s.issue}`, pages && `pp. ${pages}`, s.year].filter(Boolean).join(", ")}${s.doi ? `, doi: ${s.doi}` : ""}.`,
      );
    }
    if (!s.doi && s.url) add(` [Online]. Available: ${s.url}`);
    return pieces;
  }
  // APA and Harvard: author, year, title; then where it was published.
  const who = style === "APA" ? namesApa(s.authors) : s.authors.map((a) => [a.last, initials(a.first, true)].filter(Boolean).join(", ")).join(style === "Harvard" ? ", " : ", ");
  const year = s.year || "n.d.";
  if (who) add(style === "APA" ? `${end(who)} (${year}). ` : `${who} (${year}) `);
  else add(style === "APA" ? "" : "");
  const italicTitle = s.type === "Book" || s.type === "Report" || s.type === "InternetSite" || s.type === "Misc";
  if (italicTitle) {
    add(end(s.title), true);
    if (!who) add(` (${year}).`);
  } else add(end(s.title));
  if (s.type === "JournalArticle") {
    add(" ");
    add(s.container, true);
    if (s.volume) {
      add(", ");
      add(s.volume, true);
    }
    add(`${s.issue ? `(${s.issue})` : ""}${pages ? `, ${pages}` : ""}.`);
  } else if (s.type === "BookSection" || s.type === "ConferenceProceedings") {
    add(" In ");
    add(s.container, true);
    add(`${pages ? ` (pp. ${pages})` : ""}.`);
    if (s.publisher) add(` ${end(s.publisher)}`);
  } else if (s.publisher) add(` ${end(style === "Harvard" && s.city ? `${s.city}: ${s.publisher}` : s.publisher)}`);
  const link = doiLink(s);
  if (link) add(style === "Harvard" && !s.doi ? ` Available at: ${link}` : ` ${link}`);
  return pieces;
}

/** The reference list's entries: cited sources, alphabetically (APA, Harvard) or in citing order (IEEE). */
export function bibliographyEntries(doc: Node, bib: Bibliography): Piece[][] {
  const order = citedOrder(doc);
  const cited = order.map((t) => bib.sources.find((s) => s.tag === t)).filter((s): s is Source => !!s);
  // Nothing cited yet: all the sources, so the list isn't empty.
  const list = cited.length ? cited : [...bib.sources];
  if (bib.style !== "IEEE" && bib.style !== "Vancouver") list.sort((a, b) => (a.authors[0]?.last || a.title).localeCompare(b.authors[0]?.last || b.title) || a.year.localeCompare(b.year));
  return list.map((s, i) => formatEntry(s, bib.style, i + 1));
}

/** Keeps citations and reference lists up to date as the text and the sources change. */
export const citationsKey = new PluginKey("citations");

export function citationsPlugin(ctx: () => DocContext) {
  return new Plugin({
    key: citationsKey,
    appendTransaction(trs, _old, state) {
      if (!trs.some((tr) => tr.docChanged || tr.getMeta(citationsKey))) return null;
      const bib = ctx().bibliography;
      const order = citedOrder(state.doc);
      const tr = state.tr;
      let entries: Piece[][] | null = null;
      state.doc.descendants((node, pos) => {
        if (isCitation(node)) {
          const result = formatCitation(parseCitation(node.attrs.instr), bib, order);
          if (result !== node.attrs.result) tr.setNodeMarkup(pos, undefined, { ...node.attrs, result });
        } else if (node.type === schema.nodes.bibliography) {
          entries ??= bibliographyEntries(state.doc, bib);
          if (JSON.stringify(entries) !== JSON.stringify(node.attrs.entries)) tr.setNodeMarkup(pos, undefined, { ...node.attrs, entries });
        }
        return true;
      });
      return tr.docChanged ? tr.setMeta("addToHistory", false).setMeta("noTrack", true) : null;
    },
  });
}
