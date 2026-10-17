// Opening Markdown, HTML and plain text: they become a new Word document (with
// DuckWriter's usual styles), read the same way pasted HTML is.

import { DOMParser as PMDOMParser, type Node } from "prosemirror-model";
import type { DocContext } from "../docx/model";
import { flattenPastedLists, realLists } from "../editor/paste";
import { schema } from "../editor/schema";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** An equation, as DuckWriter's own copied HTML holds one (see the schema's clipboard rules). */
const mathHtml = (latex: string, display: boolean) => `<span data-pm-node="math" data-pm="${esc(JSON.stringify({ latex: latex.trim(), display, xml: null, original: null }))}">${esc(latex)}</span>`;

/** Markdown's inline parts: emphasis, links, code, equations. */
function inline(text: string): string {
  const keep: string[] = [];
  const hold = (html: string) => `\u0000${keep.push(html) - 1}\u0000`;
  let s = text
    .replace(/`([^`]+)`/g, (_m, code: string) => hold(esc(code)))
    .replace(/\$([^$\n]+)\$/g, (_m, latex: string) => hold(mathHtml(latex, false)))
    .replace(/!\[([^\]]*)\]\(([^)\s]+)[^)]*\)/g, (_m, alt: string) => hold(esc(alt)))
    .replace(/\[([^\]]+)\]\(([^)\s]+)[^)]*\)/g, (_m, label: string, href: string) => hold(`<a href="${esc(href)}">${inline(label)}</a>`))
    .replace(/<((?:https?:\/\/|mailto:)[^>\s]+)>/g, (_m, href: string) => hold(`<a href="${esc(href)}">${esc(href)}</a>`));
  s = esc(s)
    .replace(/\*\*([^*]+)\*\*|__([^_]+)__/g, (_m, a?: string, b?: string) => `<strong>${a ?? b}</strong>`)
    .replace(/(^|[^*\w])\*([^*\s][^*]*)\*(?!\*)|(^|[^_\w])_([^_\s][^_]*)_(?!\w)/g, (_m, p1?: string, a?: string, p2?: string, b?: string) => `${p1 ?? p2 ?? ""}<em>${a ?? b}</em>`)
    .replace(/~~([^~]+)~~/g, "<s>$1</s>")
    .replace(/&lt;(\/?(?:sup|sub))&gt;/g, "<$1>")
    .replace(/\\([\\`*_[\]#|<>~$])/g, "$1");
  return s.replace(/\u0000(\d+)\u0000/g, (_m, i: string) => keep[Number(i)]);
}

/** Markdown (CommonMark's common parts, with tables and $equations$) as HTML DuckWriter reads. */
export function markdownToHtml(markdown: string): string {
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n");
  const out: string[] = [];
  let para: string[] = [];
  const flush = () => {
    if (para.length) out.push(`<p>${inline(para.join(" "))}</p>`);
    para = [];
  };
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (/^\s*$/.test(line)) {
      flush();
      continue;
    }
    // Code: each line as it is.
    if (/^\s*```/.test(line)) {
      flush();
      for (i++; i < lines.length && !/^\s*```/.test(lines[i]); i++) out.push(`<p>${esc(lines[i]) || "&nbsp;"}</p>`);
      continue;
    }
    const display = /^\s*\$\$(.*?)(\$\$)?\s*$/.exec(line);
    if (display) {
      flush();
      let latex = display[1];
      if (!display[2]) for (i++; i < lines.length && !/\$\$\s*$/.test(lines[i]); i++) latex += ` ${lines[i]}`;
      if (!display[2] && i < lines.length) latex += ` ${lines[i].replace(/\$\$\s*$/, "")}`;
      out.push(`<p style="text-align:center">${mathHtml(latex, true)}</p>`);
      continue;
    }
    const heading = /^(#{1,6})\s+(.*?)\s*#*\s*$/.exec(line);
    if (heading) {
      flush();
      const level = heading[1].length;
      // A single # is the document's title.
      if (level === 1 && !out.length) out.push(`<p class="MsoTitle">${inline(heading[2])}</p>`);
      else {
        // Under a title, ## is the first level of headings.
        const tag = `h${Math.max(1, level - (out.some((o) => o.includes("MsoTitle")) ? 1 : 0))}`;
        out.push(`<${tag}>${inline(heading[2])}</${tag}>`);
      }
      continue;
    }
    const item = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/.exec(line);
    if (item) {
      flush();
      const level = Math.floor(item[1].replace(/\t/g, "    ").length / 2);
      out.push(`<p data-list="${/\d/.test(item[2]) ? "number" : "bullet"}" data-level="${Math.min(level, 8)}">${inline(item[3])}</p>`);
      continue;
    }
    const quote = /^>\s?(.*)$/.exec(line);
    if (quote) {
      flush();
      out.push(`<p class="MsoQuote">${inline(quote[1])}</p>`);
      continue;
    }
    // A table: a row of cells, then a line of dashes.
    if (/^\s*\|.*\|\s*$/.test(line) && /^\s*\|?\s*:?-{3,}/.test(lines[i + 1] ?? "")) {
      flush();
      const cells = (l: string) => l.trim().replace(/^\||\|$/g, "").split(/(?<!\\)\|/).map((c) => inline(c.trim()));
      let html = `<table><tr>${cells(line).map((c) => `<th><p>${c}</p></th>`).join("")}</tr>`;
      for (i += 2; i < lines.length && /^\s*\|/.test(lines[i]); i++) html += `<tr>${cells(lines[i]).map((c) => `<td><p>${c}</p></td>`).join("")}</tr>`;
      i--;
      out.push(`${html}</table>`);
      continue;
    }
    if (/^\s*(-{3,}|\*{3,}|_{3,})\s*$/.test(line)) {
      flush();
      continue;
    }
    para.push(line.trim());
  }
  flush();
  return out.join("\n");
}

/** Plain text: a paragraph for each line. */
export function textToHtml(text: string): string {
  return text
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((line) => `<p>${esc(line).replace(/\t/g, "&#9;") || ""}</p>`)
    .join("");
}

/** A document from HTML, with lists made the document's own. */
export function htmlToDoc(html: string, ctx: DocContext): Node {
  const body = new DOMParser().parseFromString(flattenPastedLists(html), "text/html").body;
  const parsed = PMDOMParser.fromSchema(schema).parse(body, { preserveWhitespace: false });
  const content = realLists(parsed.content, ctx, {});
  const doc = schema.nodes.doc.create(null, content.size ? content : schema.nodes.paragraph.create());
  // A document ends with a paragraph to go on typing in.
  return doc.lastChild?.type === schema.nodes.paragraph ? doc : schema.nodes.doc.create(null, [...doc.content.content, schema.nodes.paragraph.create()]);
}

/** Which of the formats a file name is, if any. */
export function textFormat(name: string): "markdown" | "html" | "text" | null {
  if (/\.(md|markdown)$/i.test(name)) return "markdown";
  if (/\.html?$/i.test(name)) return "html";
  if (/\.txt$/i.test(name)) return "text";
  return null;
}
