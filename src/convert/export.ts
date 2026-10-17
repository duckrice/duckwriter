// The document as a web page (HTML, with real MathML equations), as Markdown,
// or as plain text: for the web, for wikis and notebooks, for e-mail. Deleted
// text (tracked changes) is left out, as if the changes were accepted.

import type { Mark, Node } from "prosemirror-model";
import type { DocContext } from "../docx/model";
import { kindOf, listLabels, listOf } from "../editor/lists";
import { fieldKind, schema } from "../editor/schema";
import { parseMath } from "../math/ast";
import { toMathML } from "../math/mathml";
import { headingLevel } from "../ui/references";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const deleted = (n: Node) => n.marks.some((m) => m.type === schema.marks.deletion);

/** A picture's bytes as a data: URL, from the file or from DuckWriter. */
async function pictureUrl(ctx: DocContext, attrs: Node["attrs"]): Promise<string> {
  let bytes: Uint8Array | undefined;
  let ext = "png";
  if (attrs.media) {
    const entry = ctx.media.get(attrs.media);
    bytes = entry?.bytes;
    ext = entry?.ext ?? ext;
  } else if (attrs.rid) {
    const rel = ctx.rels.get(attrs.rid);
    const dir = ctx.documentPath.slice(0, ctx.documentPath.lastIndexOf("/") + 1);
    if (rel) {
      const path = (dir + rel.target).replace(/[^/]+\/\.\.\//g, "");
      bytes = await ctx.zip.file(path)?.async("uint8array");
      ext = rel.target.split(".").pop() ?? ext;
    }
  }
  if (!bytes) return "";
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  const type = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "svg" ? "image/svg+xml" : `image/${ext}`;
  return `data:${type};base64,${btoa(s)}`;
}

/** What a paragraph is: a heading (and its level), a list item, the title, or plain. */
function kind(ctx: DocContext, para: Node): { heading: number; list: { ordered: boolean; level: number } | null } {
  if (para.attrs.style === "Title") return { heading: 1, list: null };
  const heading = headingLevel(ctx, para);
  if (heading) return { heading: Math.min(6, heading + (hasTitle ? 1 : 0)), list: null };
  const list = listOf(ctx, para);
  if (list) return { heading: 0, list: { ordered: kindOf(ctx, list.numId, list.ilvl) === "number", level: list.ilvl } };
  return { heading: 0, list: null };
}
/** With a title (h1), headings start at h2. */
let hasTitle = false;

// ---- HTML -----------------------------------------------------------------------------

export async function toHtml(doc: Node, ctx: DocContext, title: string): Promise<string> {
  hasTitle = doc.firstChild?.attrs.style === "Title";
  const notes: string[] = [];
  const inline = async (para: Node): Promise<string> => {
    let out = "";
    for (let i = 0; i < para.childCount; i++) {
      const n = para.child(i);
      if (deleted(n)) continue;
      let html = "";
      if (n.isText) html = esc(n.text!).replace(/\t/g, "&emsp;");
      else if (n.type === schema.nodes.hard_break) html = "<br>";
      else if (n.type === schema.nodes.math) {
        try {
          html = toMathML(parseMath(n.attrs.latex), n.attrs.display).outerHTML;
        } catch {
          html = esc(n.attrs.latex);
        }
      } else if (n.type === schema.nodes.field) html = esc(n.attrs.result);
      else if (n.type === schema.nodes.footnote) {
        const note = ctx.footnotes.notes.get(n.attrs.id);
        notes.push(note ? (await Promise.all(note.content.content.map(inline))).join(" ") : "");
        const k = notes.length;
        html = `<sup><a href="#note-${k}" id="ref-${k}">${k}</a></sup>`;
      } else if (n.type === schema.nodes.image) {
        const src = await pictureUrl(ctx, n.attrs);
        if (src) html = `<img src="${src}" width="${Math.round(n.attrs.width * (4 / 3))}" height="${Math.round(n.attrs.height * (4 / 3))}" alt="">`;
      }
      out += wrapHtml(html, n.marks);
    }
    return out;
  };
  const block = async (node: Node, level: number): Promise<string> => {
    if (node.type === schema.nodes.paragraph) return "";
    if (node.type === schema.nodes.table) {
      let rows = "";
      for (const row of node.content.content) {
        let cells = "";
        for (const cell of row.content.content) {
          const tag = cell.type === schema.nodes.table_header ? "th" : "td";
          const span = `${cell.attrs.colspan > 1 ? ` colspan="${cell.attrs.colspan}"` : ""}${cell.attrs.rowspan > 1 ? ` rowspan="${cell.attrs.rowspan}"` : ""}`;
          cells += `<${tag}${span}>${await blocks(cell, level + 1)}</${tag}>`;
        }
        rows += `<tr>${cells}</tr>`;
      }
      return `<table>${rows}</table>\n`;
    }
    if (node.type === schema.nodes.toc) {
      const entries = node.attrs.entries as { level: number; text: string }[];
      return `<nav>${node.attrs.title ? `<h2>${esc(node.attrs.title)}</h2>` : ""}<ul>${entries.map((e) => `<li class="level-${e.level}">${esc(e.text)}</li>`).join("")}</ul></nav>\n`;
    }
    if (node.type === schema.nodes.bibliography) {
      const entries = node.attrs.entries as { text: string; italic?: boolean }[][];
      return `<section class="references"><h2>${esc(node.attrs.title)}</h2>${entries.map((e) => `<p>${e.map((p) => (p.italic ? `<em>${esc(p.text)}</em>` : esc(p.text))).join("")}</p>`).join("")}</section>\n`;
    }
    return "";
  };
  /** The blocks of a document or a cell; list paragraphs become nested lists. */
  const blocks = async (parent: Node, level: number): Promise<string> => {
    let out = "";
    const open: ("ul" | "ol")[] = [];
    const closeTo = (depth: number) => {
      while (open.length > depth) out += `</li></${open.pop()}>`;
    };
    for (const node of parent.content.content) {
      if (node.type !== schema.nodes.paragraph) {
        closeTo(0);
        out += await block(node, level);
        continue;
      }
      const k = kind(ctx, node);
      const text = await inline(node);
      if (k.list) {
        const tag = k.list.ordered ? "ol" : "ul";
        if (open.length > k.list.level + 1) closeTo(k.list.level + 1);
        // A numbered list straight after a bulleted one (or the other way) is a list of its own.
        if (open.length === k.list.level + 1 && open[k.list.level] !== tag) closeTo(k.list.level);
        if (open.length === k.list.level + 1) out += "</li>";
        while (open.length < k.list.level + 1) {
          open.push(tag);
          out += `<${tag}>`;
        }
        out += `<li>${text}`;
        continue;
      }
      closeTo(0);
      const align = node.attrs.align && node.attrs.align !== "left" ? ` style="text-align:${node.attrs.align}"` : "";
      if (k.heading) out += `<h${k.heading}${align}>${text}</h${k.heading}>\n`;
      else if (text || level === 0) out += `<p${align}>${text || "&nbsp;"}</p>\n`;
    }
    closeTo(0);
    return out;
  };
  const body = await blocks(doc, 0);
  const footnotes = notes.length ? `<section class="notes"><hr><ol>${notes.map((n, i) => `<li id="note-${i + 1}">${n} <a href="#ref-${i + 1}">↩</a></li>`).join("")}</ol></section>\n` : "";
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)}</title>
<style>
body { max-width: 46rem; margin: 2rem auto; padding: 0 1rem; font: 17px/1.55 Georgia, "Times New Roman", serif; color: #1a1a1a; }
h1, h2, h3, h4 { font-family: "Segoe UI", system-ui, sans-serif; line-height: 1.25; }
table { border-collapse: collapse; margin: 1em 0; }
td, th { border: 1px solid #bbb; padding: 4px 8px; vertical-align: top; }
td p, th p { margin: 0; }
img { max-width: 100%; height: auto; }
.references p { padding-left: 2em; text-indent: -2em; }
.notes { font-size: 0.9em; }
</style>
</head>
<body>
${body}${footnotes}</body>
</html>
`;
}

function wrapHtml(html: string, marks: readonly Mark[]): string {
  let out = html;
  for (const m of marks) {
    const a = m.attrs;
    switch (m.type.name) {
      case "bold":
        if (a.on) out = `<strong>${out}</strong>`;
        break;
      case "italic":
        if (a.on) out = `<em>${out}</em>`;
        break;
      case "underline":
        if (a.on) out = `<u>${out}</u>`;
        break;
      case "strike":
        if (a.on) out = `<s>${out}</s>`;
        break;
      case "vertAlign":
        out = a.value === "superscript" ? `<sup>${out}</sup>` : `<sub>${out}</sub>`;
        break;
      case "color":
        out = `<span style="color:${esc(a.color)}">${out}</span>`;
        break;
      case "link":
        out = a.href ? `<a href="${esc(a.href)}">${out}</a>` : out;
        break;
    }
  }
  return out;
}

// ---- Markdown -----------------------------------------------------------------------------

const mdEscape = (s: string) => s.replace(/([\\`*_[\]#|<>])/g, "\\$1");

export async function toMarkdown(doc: Node, ctx: DocContext): Promise<string> {
  hasTitle = doc.firstChild?.attrs.style === "Title";
  const notes: string[] = [];
  const inline = async (para: Node, inTable = false): Promise<string> => {
    let out = "";
    for (let i = 0; i < para.childCount; i++) {
      const n = para.child(i);
      if (deleted(n)) continue;
      let md = "";
      if (n.isText) md = mdEscape(n.text!).replace(/\t/g, " ");
      else if (n.type === schema.nodes.hard_break) md = inTable ? " " : "  \n";
      else if (n.type === schema.nodes.math) md = n.attrs.display ? `\n$$${n.attrs.latex}$$\n` : `$${n.attrs.latex}$`;
      else if (n.type === schema.nodes.field) md = mdEscape(n.attrs.result);
      else if (n.type === schema.nodes.footnote) {
        const note = ctx.footnotes.notes.get(n.attrs.id);
        notes.push(note ? (await Promise.all(note.content.content.map((p) => inline(p)))).join(" ") : "");
        md = `[^${notes.length}]`;
      } else if (n.type === schema.nodes.image) {
        const src = await pictureUrl(ctx, n.attrs);
        if (src) md = `![](${src})`;
      }
      out += wrapMarkdown(md, n.marks);
    }
    return out;
  };
  const parts: string[] = [];
  const labels = listLabels(ctx, doc);
  let offset = 0;
  for (let i = 0; i < doc.childCount; offset += doc.child(i).nodeSize, i++) {
    const node = doc.child(i);
    if (node.type === schema.nodes.paragraph) {
      const k = kind(ctx, node);
      const text = await inline(node);
      if (k.heading) parts.push(`${"#".repeat(k.heading)} ${text}`);
      else if (k.list) {
        const marker = k.list.ordered ? `${/\d+/.exec(labels.get(offset)?.label ?? "")?.[0] ?? "1"}.` : "-";
        parts.push(`${"   ".repeat(k.list.level)}${marker} ${text}`);
      } else if (text.trim()) parts.push(text);
      else parts.push("");
    } else if (node.type === schema.nodes.table) {
      const rows: string[][] = [];
      for (const row of node.content.content) rows.push(await Promise.all(row.content.content.map(async (cell) => (await Promise.all(cell.content.content.map((p) => (p.type === schema.nodes.paragraph ? inline(p, true) : Promise.resolve(""))))).join(" ").replace(/\|/g, "\\|"))));
      if (rows.length) {
        const width = Math.max(...rows.map((r) => r.length));
        const line = (r: string[]) => `| ${Array.from({ length: width }, (_, j) => r[j] ?? "").join(" | ")} |`;
        parts.push([line(rows[0]), `|${" --- |".repeat(width)}`, ...rows.slice(1).map(line)].join("\n"));
      }
    } else if (node.type === schema.nodes.bibliography) {
      const entries = node.attrs.entries as { text: string; italic?: boolean }[][];
      parts.push(`## ${node.attrs.title}`, ...entries.map((e) => e.map((p) => (p.italic ? `*${mdEscape(p.text)}*` : mdEscape(p.text))).join("").replace(/\t/g, " ")));
    } else if (node.type === schema.nodes.toc) {
      const entries = node.attrs.entries as { level: number; text: string }[];
      parts.push(`## ${node.attrs.title}`, entries.map((e) => `${"   ".repeat(Math.max(0, e.level - 1))}- ${mdEscape(e.text)}`).join("\n"));
    }
  }
  // Paragraphs are separated by blank lines; list items of one list aren't.
  let out = "";
  parts.forEach((p, i) => {
    const prevList = i > 0 && /^\s*([-*]|\d+\.)\s/.test(parts[i - 1]);
    const thisList = /^\s*([-*]|\d+\.)\s/.test(p);
    out += i === 0 ? p : prevList && thisList ? `\n${p}` : `\n\n${p}`;
  });
  out = out.replace(/\n{3,}/g, "\n\n").trim();
  if (notes.length) out += `\n\n${notes.map((n, i) => `[^${i + 1}]: ${n}`).join("\n")}`;
  return `${out}\n`;
}

function wrapMarkdown(md: string, marks: readonly Mark[]): string {
  if (!md.trim()) return md;
  let out = md;
  for (const m of marks) {
    const a = m.attrs;
    if (m.type.name === "bold" && a.on) out = `**${out}**`;
    else if (m.type.name === "italic" && a.on) out = `*${out}*`;
    else if (m.type.name === "strike" && a.on) out = `~~${out}~~`;
    else if (m.type.name === "vertAlign") out = a.value === "superscript" ? `<sup>${out}</sup>` : `<sub>${out}</sub>`;
    else if (m.type.name === "link" && a.href) out = `[${out}](${a.href})`;
  }
  return out;
}

// ---- plain text ---------------------------------------------------------------------------

export function toText(doc: Node, ctx: DocContext): string {
  const labels = listLabels(ctx, doc);
  const textOf = (para: Node) => {
    let out = "";
    para.forEach((n) => {
      if (deleted(n)) return;
      if (n.isText) out += n.text;
      else if (n.type === schema.nodes.hard_break) out += "\n";
      else if (n.type === schema.nodes.field && fieldKind(n.attrs.instr) !== "") out += n.attrs.result;
      else if (n.type === schema.nodes.math) out += n.attrs.latex;
    });
    return out;
  };
  const lines: string[] = [];
  doc.forEach((node, pos) => {
    if (node.type === schema.nodes.paragraph) {
      const label = labels.get(pos)?.label;
      lines.push(`${label ? `${label}\t` : ""}${textOf(node)}`);
    } else if (node.type === schema.nodes.table) {
      node.forEach((row) => {
        const cells: string[] = [];
        row.forEach((cell) => cells.push(cell.textContent));
        lines.push(cells.join("\t"));
      });
    } else if (node.type === schema.nodes.bibliography || node.type === schema.nodes.toc) {
      lines.push(node.attrs.title, ...(node.attrs.entries as { text?: string }[] | { text: string }[][]).map((e) => (Array.isArray(e) ? e.map((p) => p.text).join("") : (e.text ?? ""))));
    }
  });
  return `${lines.join("\r\n")}\r\n`;
}
