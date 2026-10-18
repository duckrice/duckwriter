// The document as LaTeX, for journals that want it: sections from the
// headings, lists, tables, equations as they are written, footnotes,
// citations as \cite with the sources in a .bib file, and the pictures as
// files next to it. Deleted text (tracked changes) is left out.

import type { Mark, Node } from "prosemirror-model";
import type { CitationStyle } from "../docx/bibliography";
import type { DocContext } from "../docx/model";
import { toBibtex } from "../docx/ris";
import { parseCitation } from "../editor/citations";
import { kindOf, listOf } from "../editor/lists";
import { fieldKind, schema } from "../editor/schema";
import { headingLevel } from "../ui/references";
import { pictureFile } from "./export";

export interface LatexExport {
  tex: string;
  /** The sources, or "" when there are none. */
  bib: string;
  /** Picture files to write next to the .tex file. */
  pictures: { name: string; bytes: Uint8Array }[];
}

/** Text with the characters TeX treats specially escaped. */
export function texEscape(s: string): string {
  return s
    .replace(/\\/g, "\u0000")
    .replace(/([&%$#_{}])/g, "\\$1")
    .replace(/~/g, "\\textasciitilde{}")
    .replace(/\^/g, "\\textasciicircum{}")
    .replace(/\u0000/g, "\\textbackslash{}")
    .replace(/\u00a0/g, "~")
    .replace(/\u2011/g, "-")
    .replace(/\t/g, " ");
}

const deleted = (n: Node) => n.marks.some((m) => m.type === schema.marks.deletion);

const BIB_STYLES: Record<CitationStyle, string> = { APA: "apalike", Harvard: "apalike", Chicago: "apalike", IEEE: "ieeetr", Vancouver: "unsrt" };
const SECTIONS = ["section", "subsection", "subsubsection", "paragraph", "subparagraph"];

function wrap(tex: string, marks: readonly Mark[]): string {
  if (!tex.trim()) return tex;
  let out = tex;
  for (const m of marks) {
    const a = m.attrs;
    if (m.type.name === "bold" && a.on) out = `\\textbf{${out}}`;
    else if (m.type.name === "italic" && a.on) out = `\\emph{${out}}`;
    else if (m.type.name === "underline" && a.on) out = `\\uline{${out}}`;
    else if (m.type.name === "strike" && a.on) out = `\\sout{${out}}`;
    else if (m.type.name === "vertAlign") out = a.value === "superscript" ? `\\textsuperscript{${out}}` : `\\textsubscript{${out}}`;
    else if (m.type.name === "link" && a.href) out = `\\href{${a.href.replace(/([%#\\])/g, "\\$1")}}{${out}}`;
  }
  return out;
}

export async function toLatex(doc: Node, ctx: DocContext, base: string): Promise<LatexExport> {
  const pictures: LatexExport["pictures"] = [];
  const known = new Set(ctx.bibliography.sources.map((s) => s.tag));
  let cites = false;
  let endnotes = false;

  const inline = async (para: Node, skipSeq = false): Promise<string> => {
    let out = "";
    for (let i = 0; i < para.childCount; i++) {
      const n = para.child(i);
      if (deleted(n)) continue;
      let tex = "";
      if (n.isText) tex = texEscape(n.text!);
      else if (n.type === schema.nodes.hard_break) tex = "\\\\\n";
      else if (n.type === schema.nodes.page_break) tex = "\n\\newpage\n";
      else if (n.type === schema.nodes.math) tex = n.attrs.display ? `\n\\[ ${n.attrs.latex} \\]\n` : `$${n.attrs.latex}$`;
      else if (n.type === schema.nodes.field) {
        const kind = fieldKind(n.attrs.instr);
        if (kind === "CITATION") {
          const c = parseCitation(n.attrs.instr);
          const tags = c.tags.filter((t) => known.has(t));
          if (tags.length) {
            cites = true;
            tex = `\\cite${c.page ? `[p.~${texEscape(c.page)}]` : ""}{${tags.join(",")}}`;
          } else tex = texEscape(n.attrs.result);
        } else if (kind === "SEQ" && skipSeq) tex = "";
        else tex = texEscape(n.attrs.result);
      } else if (n.type === schema.nodes.footnote || n.type === schema.nodes.endnote) {
        const isEnd = n.type === schema.nodes.endnote;
        const note = (isEnd ? ctx.endnotes : ctx.footnotes).notes.get(n.attrs.id);
        const text = note ? (await Promise.all(note.content.content.map((p) => (p.type === schema.nodes.paragraph ? inline(p) : Promise.resolve(""))))).join("\n\n").trim() : "";
        if (isEnd) endnotes = true;
        tex = `\\${isEnd ? "endnote" : "footnote"}{${text}}`;
      } else if (n.type === schema.nodes.image) {
        const file = await pictureFile(ctx, n.attrs);
        if (file) {
          const name = `${base}-figure${pictures.length + 1}.${file.ext === "jpeg" ? "jpg" : file.ext}`;
          pictures.push({ name, bytes: file.bytes });
          tex = `\\includegraphics[width=${Math.round(n.attrs.width)}pt]{${name}}`;
        }
      }
      out += n.isText || n.type === schema.nodes.field ? wrap(tex, n.marks) : tex;
    }
    return out;
  };

  /** A caption's own text, without "Figure 1:" (LaTeX numbers it). */
  const captionText = async (para: Node) => (await inline(para, true)).replace(/^\s*[\p{L}]+\s*[:.\u2013\u2014-]?\s*/u, "").trim();
  const isCaption = (n: Node | null | undefined) => !!n && n.type === schema.nodes.paragraph && n.attrs.style === "Caption";
  const onlyPicture = (n: Node) => {
    let pictures = 0;
    let other = false;
    n.forEach((c) => {
      if (c.type === schema.nodes.image) pictures++;
      else if (!(c.isText && !c.text!.trim())) other = true;
    });
    return pictures === 1 && !other;
  };

  const out: string[] = [];
  let title = "";
  const lists: ("itemize" | "enumerate")[] = [];
  const closeLists = (to = 0) => {
    while (lists.length > to) out.push(`${"  ".repeat(lists.length - 1)}\\end{${lists.pop()}}`);
  };
  let usedCaption = -1;

  for (let i = 0; i < doc.childCount; i++) {
    const node = doc.child(i);
    if (i === usedCaption) continue;
    if (node.type === schema.nodes.paragraph) {
      // All of it deleted, while changes were tracked.
      if (node.childCount && node.content.content.every(deleted)) continue;
      const list = listOf(ctx, node);
      if (!list) closeLists();
      if (node.attrs.style === "Title" && !title) {
        title = await inline(node);
        continue;
      }
      const heading = headingLevel(ctx, node);
      if (heading) {
        out.push("", `\\${SECTIONS[Math.min(heading, SECTIONS.length) - 1]}{${(await inline(node)).trim()}}`);
        continue;
      }
      if (list) {
        const env = kindOf(ctx, list.numId, list.ilvl) === "number" ? "enumerate" : "itemize";
        const depth = list.ilvl + 1;
        closeLists(Math.min(lists.length, depth));
        if (lists.length === depth && lists[depth - 1] !== env) closeLists(depth - 1);
        while (lists.length < depth) {
          const kind = lists.length === depth - 1 ? env : "itemize";
          out.push(`${"  ".repeat(lists.length)}\\begin{${kind}}`);
          lists.push(kind);
        }
        out.push(`${"  ".repeat(depth)}\\item ${(await inline(node)).trim()}`);
        continue;
      }
      if (onlyPicture(node)) {
        const next = i + 1 < doc.childCount ? doc.child(i + 1) : null;
        const caption = isCaption(next) ? await captionText(next!) : "";
        if (caption) usedCaption = i + 1;
        out.push("", "\\begin{figure}[htbp]", "  \\centering", `  ${(await inline(node)).trim()}`, ...(caption ? [`  \\caption{${caption}}`] : []), "\\end{figure}");
        continue;
      }
      // A caption above a table goes into the table.
      if (isCaption(node) && i + 1 < doc.childCount && doc.child(i + 1).type === schema.nodes.table) continue;
      const text = (await inline(node)).trim();
      const align = node.attrs.align === "center" ? "center" : null;
      out.push("", align && text ? `\\begin{center}\n${text}\n\\end{center}` : text);
    } else if (node.type === schema.nodes.table) {
      closeLists();
      const rows: string[][] = [];
      for (const row of node.content.content) {
        if (/<w:del/.test(row.attrs.trPr ?? "")) continue;
        rows.push(await Promise.all(row.content.content.map(async (cell) => (await Promise.all(cell.content.content.map((p) => (p.type === schema.nodes.paragraph ? inline(p) : Promise.resolve(""))))).join(" ").trim())));
      }
      const width = Math.max(1, ...rows.map((r) => r.length));
      const prev = i > 0 ? doc.child(i - 1) : null;
      const caption = isCaption(prev) ? await captionText(prev!) : "";
      out.push(
        "",
        "\\begin{table}[htbp]",
        "  \\centering",
        ...(caption ? [`  \\caption{${caption}}`] : []),
        `  \\begin{tabular}{|${"l|".repeat(width)}}`,
        "    \\hline",
        ...rows.map((r) => `    ${Array.from({ length: width }, (_, j) => r[j] ?? "").join(" & ")} \\\\ \\hline`),
        "  \\end{tabular}",
        "\\end{table}",
      );
    } else if (node.type === schema.nodes.bibliography) {
      closeLists();
      out.push("", `\\bibliographystyle{${BIB_STYLES[ctx.bibliography.style] ?? "plain"}}`, `\\bibliography{${base}}`);
      cites = false;
    } else if (node.type === schema.nodes.toc) {
      closeLists();
      const label = node.attrs.label as string | null;
      out.push("", label === "Figure" ? "\\listoffigures" : label === "Table" ? "\\listoftables" : "\\tableofcontents");
    } else if (node.type === schema.nodes.endnotes) {
      closeLists();
      if (endnotes) out.push("", "\\theendnotes");
      endnotes = false;
    }
  }
  closeLists();
  // Cited, but no reference list in the document: one at the end.
  if (cites) out.push("", `\\bibliographystyle{${BIB_STYLES[ctx.bibliography.style] ?? "plain"}}`, `\\bibliography{${base}}`);
  if (endnotes) out.push("", "\\theendnotes");

  const body = out.join("\n").replace(/\n{3,}/g, "\n\n").trim();
  const usesEndnotes = body.includes("\\endnote{");
  const tex = [
    "\\documentclass[11pt]{article}",
    "\\usepackage[utf8]{inputenc}",
    "\\usepackage[T1]{fontenc}",
    "\\usepackage{amsmath,amssymb}",
    "\\usepackage{graphicx}",
    "\\usepackage[normalem]{ulem}",
    ...(usesEndnotes ? ["\\usepackage{endnotes}"] : []),
    "\\usepackage[hidelinks]{hyperref}",
    "",
    ...(title ? [`\\title{${title.trim()}}`, "\\author{}", "\\date{}", ""] : []),
    "\\begin{document}",
    ...(title ? ["\\maketitle"] : []),
    "",
    body,
    "",
    "\\end{document}",
    "",
  ].join("\n");
  const cited = ctx.bibliography.sources.filter((s) => s.tag);
  return { tex, bib: cited.length ? toBibtex(cited) : "", pictures };
}
