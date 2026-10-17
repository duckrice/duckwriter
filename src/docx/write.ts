// Saves an editor document back into its .docx. Only the body of the main
// document is rewritten (plus list definitions and links that were added);
// every other part of the original file is kept exactly as it was.

import type { Mark, Node } from "prosemirror-model";
import { schema } from "../editor/schema";
import { parseMath } from "../math/ast";
import { toOmml } from "../math/omml";
import { TWIPS_PER_POINT, type DocContext, type Relationship } from "./model";
import { escapeAttr, escapeText, tag } from "./xml";

/** The order Word requires for the elements of a paragraph's properties. */
const PPR_ORDER = [
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers",
  "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
  "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
  "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
  "rPr", "sectPr", "pPrChange",
];

/** The order Word requires for the elements of a run's properties. */
const RPR_ORDER = [
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss",
  "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
  "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
  "specVanish", "oMath",
];

/** Splits a run of sibling elements written as XML into one string per element. */
function splitElements(xml: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let start = 0;
  const re = /<(\/?)([\w:.-]+)[^>]*?(\/?)>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    if (m[1]) depth--;
    else if (!m[3]) depth++;
    if (depth === 0) {
      out.push(xml.slice(start, re.lastIndex));
      start = re.lastIndex;
    }
  }
  return out;
}

const localName = (element: string) => /^<(?:[\w.-]+:)?([\w.-]+)/.exec(element)?.[1] ?? "";

/** Puts property elements in the order Word requires; unknown ones go last. */
function ordered(elements: string[], order: string[]): string {
  const rank = (el: string) => {
    const i = order.indexOf(localName(el));
    return i < 0 ? order.length : i;
  };
  return elements
    .map((el, i) => ({ el, i }))
    .sort((a, b) => rank(a.el) - rank(b.el) || a.i - b.i)
    .map(({ el }) => el)
    .join("");
}

const twips = (points: number) => Math.round(points * TWIPS_PER_POINT);

export async function writeDocx(doc: Node, ctx: DocContext): Promise<Uint8Array> {
  // Headers and footers first: new ones change the section properties.
  await writeStories(ctx);
  await writeFootnotes(doc, ctx);
  await writeAddedStyles(ctx);
  await writeStyleNumbering(ctx);
  updateSectPr(ctx);
  const writer = new BodyWriter(ctx);
  const body = writer.body(doc);

  const xml = ctx.documentXml;
  const open = /<w:body(\s[^>]*)?>/.exec(xml);
  const close = xml.lastIndexOf("</w:body>");
  if (!open || close < 0) throw new Error("The original document could not be updated.");
  const documentXml = xml.slice(0, open.index + open[0].length) + body + xml.slice(close);
  ctx.zip.file(ctx.documentPath, documentXml);
  ctx.documentXml = documentXml;

  await writeNumbering(ctx);
  await writeRels(ctx);
  return ctx.zip.generateAsync({ type: "uint8array", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

class BodyWriter {
  /** @param story a header, footer or the footnotes being written, whose links live in their own part. */
  constructor(
    private ctx: DocContext,
    private story: { rels: Map<string, Relationship> } | null = null,
  ) {}

  body(doc: Node): string {
    return this.blocks(doc) + this.ctx.sectPr;
  }

  /** @param lead XML put at the start of the first paragraph (a footnote's own number). */
  blocks(doc: Node, lead = ""): string {
    let out = "";
    doc.forEach((node, _offset, index) => {
      if (node.type === schema.nodes.paragraph) out += this.paragraph(node, index === 0 ? lead : "");
      else if (node.type === schema.nodes.raw_block) out += node.attrs.xml;
      else if (node.type === schema.nodes.toc) out += this.toc(node);
    });
    return out;
  }

  /** A table of contents as Word writes one: a TOC field spanning its entries, so Word can update it too. */
  private toc(node: Node): string {
    const p = this.ctx.page;
    const tab = twips(p.width - p.marginLeft - p.marginRight);
    const entries = node.attrs.entries as { level: number; text: string; page: number }[];
    let out = `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t xml:space="preserve">${escapeText(node.attrs.title)}</w:t></w:r></w:p>`;
    const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    if (!entries.length) return `${out}<w:p>${begin}<w:r><w:t>No headings yet.</w:t></w:r>${end}</w:p>`;
    entries.forEach((e, i) => {
      const pPr = `<w:pPr><w:pStyle w:val="TOC${Math.min(e.level, 3)}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${tab}"/></w:tabs></w:pPr>`;
      out += `<w:p>${pPr}${i === 0 ? begin : ""}<w:r><w:t xml:space="preserve">${escapeText(e.text)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${e.page}</w:t></w:r>${i === entries.length - 1 ? end : ""}</w:p>`;
    });
    return out;
  }

  paragraph(node: Node, lead = ""): string {
    const a = node.attrs;
    const props: string[] = [];
    if (a.style) props.push(tag("w:pStyle", { "w:val": a.style }));
    if (a.numId !== null) props.push(tag("w:numPr", {}, tag("w:ilvl", { "w:val": a.ilvl }), tag("w:numId", { "w:val": a.numId })));
    if (a.spaceBefore !== null || a.spaceAfter !== null || a.lineSpacing !== null || a.lineHeight !== null) {
      props.push(
        tag("w:spacing", {
          "w:before": a.spaceBefore !== null ? twips(a.spaceBefore) : null,
          "w:after": a.spaceAfter !== null ? twips(a.spaceAfter) : null,
          "w:line": a.lineSpacing !== null ? Math.round(a.lineSpacing * 240) : a.lineHeight !== null ? twips(a.lineHeight) : null,
          "w:lineRule": a.lineSpacing !== null ? "auto" : a.lineHeight !== null ? a.lineRule ?? "exact" : null,
        }),
      );
    }
    if (a.indentLeft !== null || a.indentRight !== null || a.indentFirst !== null) {
      props.push(
        tag("w:ind", {
          "w:left": a.indentLeft !== null ? twips(a.indentLeft) : null,
          "w:right": a.indentRight !== null ? twips(a.indentRight) : null,
          "w:firstLine": a.indentFirst !== null && a.indentFirst >= 0 ? twips(a.indentFirst) : null,
          "w:hanging": a.indentFirst !== null && a.indentFirst < 0 ? twips(-a.indentFirst) : null,
        }),
      );
    }
    if (a.align) props.push(tag("w:jc", { "w:val": a.align === "justify" ? "both" : a.align }));
    props.push(...splitElements(a.extra));
    const pPr = props.length ? `<w:pPr>${ordered(props, PPR_ORDER)}</w:pPr>` : "";
    return `<w:p>${pPr}${lead}${this.inlines(node)}</w:p>`;
  }

  /** The paragraph's content; neighbouring pieces with the same link share one hyperlink element. */
  private inlines(para: Node): string {
    let out = "";
    let link: Mark | null = null;
    let inLink = "";
    const closeLink = () => {
      if (link) out += this.hyperlink(link, inLink);
      link = null;
      inLink = "";
    };
    para.forEach((node) => {
      const nodeLink = node.marks.find((m) => m.type === schema.marks.link) ?? null;
      if (link && (!nodeLink || !nodeLink.eq(link))) closeLink();
      const xml = this.inline(node);
      if (nodeLink) {
        link = nodeLink;
        inLink += xml;
      } else out += xml;
    });
    closeLink();
    return out;
  }

  private inline(node: Node): string {
    if (node.isText) return `<w:r>${this.runProps(node.marks)}${textXml(node.text!)}</w:r>`;
    if (node.type === schema.nodes.hard_break) return `<w:r>${this.runProps(node.marks)}<w:br/></w:r>`;
    if (node.type === schema.nodes.page_break) return `<w:r>${this.runProps(node.marks)}<w:br w:type="page"/></w:r>`;
    if (node.type === schema.nodes.raw_inline) return node.attrs.xml;
    if (node.type === schema.nodes.math) return this.equation(node);
    if (node.type === schema.nodes.field) return this.field(node);
    if (node.type === schema.nodes.footnote) {
      return node.attrs.xml ?? `<w:r>${this.runProps(node.marks) || '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'}<w:footnoteReference w:id="${escapeAttr(node.attrs.id)}"/></w:r>`;
    }
    return "";
  }

  private field(node: Node): string {
    const { instr, result, xml, original } = node.attrs;
    if (xml && original === `${instr}|${result}`) return xml;
    const rPr = this.runProps(node.marks);
    return `<w:fldSimple w:instr="${escapeAttr(instr)}"><w:r>${rPr}${textXml(result || " ")}</w:r></w:fldSimple>`;
  }

  private equation(node: Node): string {
    const { latex, display, xml, original } = node.attrs;
    if (xml && latex === original) return xml;
    let tree;
    try {
      tree = parseMath(latex);
    } catch {
      // An equation that doesn't parse is saved as its text, so nothing is lost.
      tree = { t: "text" as const, v: latex };
    }
    return toOmml(tree, display, !this.ctx.prefixes.has("m"));
  }

  private hyperlink(mark: Mark, content: string): string {
    const { href, anchor } = mark.attrs;
    if (anchor && !href) return `<w:hyperlink w:anchor="${escapeAttr(anchor)}" w:history="1">${content}</w:hyperlink>`;
    let rid: string | null = mark.attrs.rid;
    if (this.story) {
      // A link in a header keeps its relationship; a new one is written as plain text for now.
      const known = rid ? this.story.rels.get(rid) : undefined;
      if (!known || known.target !== href) return content;
      return `<w:hyperlink r:id="${escapeAttr(rid!)}" w:history="1">${content}</w:hyperlink>`;
    }
    const known = rid ? (this.ctx.rels.get(rid) ?? this.ctx.addedRels.get(rid)) : undefined;
    if (!known || known.target !== href) rid = this.addLinkRel(href);
    const ns = this.ctx.prefixes.has("r") ? "" : ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
    return `<w:hyperlink r:id="${escapeAttr(rid!)}" w:history="1"${ns}>${content}</w:hyperlink>`;
  }

  private addLinkRel(href: string): string {
    for (const [id, rel] of [...this.ctx.rels, ...this.ctx.addedRels]) {
      if (rel.external && rel.target === href && rel.type.endsWith("/hyperlink")) return id;
    }
    let n = this.ctx.rels.size + this.ctx.addedRels.size + 1;
    while (this.ctx.rels.has(`rId${n}`) || this.ctx.addedRels.has(`rId${n}`)) n++;
    const id = `rId${n}`;
    this.ctx.addedRels.set(id, { type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target: href, external: true });
    return id;
  }

  runProps(marks: readonly Mark[]): string {
    const props: string[] = [];
    const toggle = (name: string, on: boolean) => (on ? tag(`w:${name}`) : tag(`w:${name}`, { "w:val": "0" }));
    for (const mark of marks) {
      const a = mark.attrs;
      switch (mark.type.name) {
        case "charStyle":
          props.push(tag("w:rStyle", { "w:val": a.id }));
          break;
        case "font":
          props.push(a.xml ?? tag("w:rFonts", { "w:ascii": a.family, "w:hAnsi": a.family, "w:cs": a.family, "w:eastAsia": a.family }));
          break;
        case "bold":
          props.push(toggle("b", a.on), toggle("bCs", a.on));
          break;
        case "italic":
          props.push(toggle("i", a.on), toggle("iCs", a.on));
          break;
        case "strike":
          props.push(toggle("strike", a.on));
          break;
        case "underline":
          props.push(tag("w:u", { "w:val": a.on ? a.kind || "single" : "none" }));
          break;
        case "color":
          props.push(tag("w:color", { "w:val": String(a.color).replace("#", "").toUpperCase() }));
          break;
        case "size":
          props.push(tag("w:sz", { "w:val": Math.round(a.pt * 2) }), tag("w:szCs", { "w:val": Math.round(a.pt * 2) }));
          break;
        case "highlight":
          props.push(tag("w:highlight", { "w:val": a.color }));
          break;
        case "vertAlign":
          props.push(tag("w:vertAlign", { "w:val": a.value }));
          break;
        case "runExtra":
          props.push(...splitElements(a.xml));
          break;
      }
    }
    return props.length ? `<w:rPr>${ordered(props, RPR_ORDER)}</w:rPr>` : "";
  }
}

/** Text as run content: tabs and special hyphens become their own elements. */
function textXml(text: string): string {
  let out = "";
  for (const piece of text.split(/([\t‑­])/)) {
    if (piece === "\t") out += "<w:tab/>";
    else if (piece === "‑") out += "<w:noBreakHyphen/>";
    else if (piece === "­") out += "<w:softHyphen/>";
    else if (piece) out += `<w:t xml:space="preserve">${escapeText(piece)}</w:t>`;
  }
  return out;
}

// ---- headers and footers -----------------------------------------------------------

const STORY = {
  header: { root: "w:hdr", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", content: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml" },
  footer: { root: "w:ftr", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", content: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml" },
};

const NEW_ROOT = (tag: string) =>
  `<${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">`;

function newRelId(ctx: DocContext): string {
  let n = ctx.rels.size + ctx.addedRels.size + 1;
  while (ctx.rels.has(`rId${n}`) || ctx.addedRels.has(`rId${n}`)) n++;
  return `rId${n}`;
}

async function writeStories(ctx: DocContext) {
  const dir = documentDir(ctx);
  for (const story of ctx.stories) {
    if (story.path && story.saved && story.doc.eq(story.saved)) continue;
    const info = STORY[story.kind];
    if (!story.path) {
      let n = 1;
      while (ctx.zip.file(`${dir}${story.kind}${n}.xml`)) n++;
      story.path = `${dir}${story.kind}${n}.xml`;
      story.rid = newRelId(ctx);
      const rel: Relationship = { type: info.type, target: `${story.kind}${n}.xml`, external: false };
      ctx.addedRels.set(story.rid, rel);
      await addContentType(ctx, `/${story.path}`, info.content);
    }
    const start = story.rootStart ?? NEW_ROOT(info.root);
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
${start}${new BodyWriter(ctx, story).blocks(story.doc)}</${info.root}>`;
    ctx.zip.file(story.path, xml);
    story.rootStart = start;
    story.saved = story.doc;
  }
}

// ---- footnotes ----------------------------------------------------------------------

const FOOTNOTES = {
  type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
  content: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
};
const SEPARATORS = [
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>',
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
];
/** The number at the start of a note. */
const NOTE_NUMBER = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>';

/** The footnotes the text refers to, in order. */
function referencedNotes(doc: Node): string[] {
  const ids: string[] = [];
  doc.descendants((n) => {
    if (n.type === schema.nodes.footnote && !ids.includes(n.attrs.id)) ids.push(n.attrs.id);
  });
  return ids;
}

async function writeFootnotes(doc: Node, ctx: DocContext) {
  const fn = ctx.footnotes;
  const ids = referencedNotes(doc).filter((id) => fn.notes.has(id));
  const unchanged =
    ids.length === fn.saved.size && ids.every((id) => fn.saved.get(id) && fn.notes.get(id)!.eq(fn.saved.get(id)!));
  if (unchanged && (fn.path || !ids.length)) return;
  if (!fn.path) {
    const dir = documentDir(ctx);
    fn.path = `${dir}footnotes.xml`;
    ctx.addedRels.set(newRelId(ctx), { type: FOOTNOTES.type, target: "footnotes.xml", external: false });
    await addContentType(ctx, `/${fn.path}`, FOOTNOTES.content);
    await addFootnoteSettings(ctx);
  }
  const special = fn.special.length ? fn.special : SEPARATORS;
  const writer = new BodyWriter(ctx, fn);
  const notes = ids.map((id) => `<w:footnote w:id="${escapeAttr(id)}">${writer.blocks(fn.notes.get(id)!, NOTE_NUMBER)}</w:footnote>`);
  const start = fn.rootStart ?? NEW_ROOT("w:footnotes");
  ctx.zip.file(fn.path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${start}${special.join("")}${notes.join("")}</w:footnotes>`);
  fn.rootStart = start;
  fn.special = special;
  fn.saved = new Map(ids.map((id) => [id, fn.notes.get(id)!]));
}

/** Word lists the separator notes in the settings. */
async function addFootnoteSettings(ctx: DocContext) {
  const path = `${documentDir(ctx)}settings.xml`;
  const xml = await ctx.zip.file(path)?.async("string");
  if (!xml || xml.includes("<w:footnotePr")) return;
  const props = '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>';
  const at = xml.search(/<w:(endnotePr|compat|docVars|rsids|mathPr|attachedSchema|themeFontLang|clrSchemeMapping|doNotIncludeSubdocsInStats|doNotAutoCompressPictures|forceUpgrade|captions|readModeInkLockDown|smartTagType|schemaLibrary|shapeDefaults|doNotEmbedSmartTags|decimalSymbol|listSeparator)\b/);
  const end = xml.lastIndexOf("</w:settings>");
  ctx.zip.file(path, at >= 0 ? xml.slice(0, at) + props + xml.slice(at) : xml.slice(0, end) + props + xml.slice(end));
}

/** Adds the style definitions DuckWriter needed (for footnotes, captions…) to the document's styles. */
async function writeAddedStyles(ctx: DocContext) {
  if (!ctx.addedStyles.length) return;
  const rel = [...ctx.rels.values(), ...ctx.addedRels.values()].find((r) => r.type.endsWith("/styles"));
  if (!rel) return;
  const path = documentDir(ctx) + rel.target.replace(/^\//, "");
  let xml = await ctx.zip.file(path)?.async("string");
  if (!xml) return;
  for (const style of ctx.addedStyles) {
    const id = /w:styleId="([^"]+)"/.exec(style)?.[1];
    if (id && xml.includes(`w:styleId="${id}"`)) continue;
    const end = xml.lastIndexOf("</w:styles>");
    xml = xml.slice(0, end) + style + xml.slice(end);
  }
  ctx.zip.file(path, xml);
  ctx.addedStyles = [];
}

/** Numbers (or stops numbering) paragraph styles, for numbered headings. */
async function writeStyleNumbering(ctx: DocContext) {
  if (!ctx.styleNumbering.size) return;
  const rel = [...ctx.rels.values(), ...ctx.addedRels.values()].find((r) => r.type.endsWith("/styles"));
  if (!rel) return;
  const path = documentDir(ctx) + rel.target.replace(/^\//, "");
  let xml = await ctx.zip.file(path)?.async("string");
  if (!xml) return;
  for (const [id, numbering] of ctx.styleNumbering) {
    const re = new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`);
    const match = re.exec(xml);
    if (!match) continue;
    let style = match[0];
    const numPr = numbering ? `<w:numPr><w:ilvl w:val="${numbering.ilvl}"/><w:numId w:val="${numbering.numId}"/></w:numPr>` : "";
    const pPr = /<w:pPr>([\s\S]*?)<\/w:pPr>|<w:pPr\/>/.exec(style);
    if (pPr) {
      const parts = splitElements(pPr[1] ?? "").filter((p) => localName(p) !== "numPr");
      if (numPr) parts.push(numPr);
      style = style.replace(pPr[0], parts.length ? `<w:pPr>${ordered(parts, PPR_ORDER)}</w:pPr>` : "");
    } else if (numPr) {
      const at = style.search(/<w:rPr\b|<w:tblPr\b|<\/w:style>/);
      style = style.slice(0, at) + `<w:pPr>${numPr}</w:pPr>` + style.slice(at);
    }
    xml = xml.replace(match[0], style);
  }
  ctx.zip.file(path, xml);
  ctx.styleNumbering.clear();
}

/** The order Word requires for the elements of section properties. */
const SECTPR_ORDER = [
  "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders",
  "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter",
  "docGrid", "printerSettings", "sectPrChange",
];

/** Points the section at its headers and footers, and says whether the first page has its own. */
function updateSectPr(ctx: DocContext) {
  const xml = ctx.sectPr || "<w:sectPr></w:sectPr>";
  const start = /^<w:sectPr\b[^>]*?(\/?)>/.exec(xml);
  if (!start) return;
  const open = start[0].replace(/\/>$/, ">");
  const inner = start[1] ? "" : xml.slice(start[0].length, xml.lastIndexOf("</w:sectPr>"));
  let parts = splitElements(inner);
  const ns = ctx.prefixes.has("r") ? "" : ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
  for (const story of ctx.stories) {
    if (!story.rid || parts.some((p) => p.includes(`"${story.rid}"`))) continue;
    const name = `${story.kind}Reference`;
    parts = parts.filter((p) => !(localName(p) === name && p.includes(`w:type="${story.type}"`)));
    parts.push(`<w:${name} w:type="${story.type}" r:id="${story.rid}"${ns}/>`);
  }
  parts = parts.filter((p) => localName(p) !== "titlePg");
  if (ctx.titlePage) parts.push("<w:titlePg/>");
  ctx.sectPr = `${open}${ordered(parts, SECTPR_ORDER)}</w:sectPr>`;
}

// ---- other parts ------------------------------------------------------------------

const NUMBERING_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const NUMBERING_CONTENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

function documentDir(ctx: DocContext) {
  return ctx.documentPath.slice(0, ctx.documentPath.lastIndexOf("/") + 1);
}

async function writeNumbering(ctx: DocContext) {
  const { addedAbstracts, addedNums } = ctx.numbering;
  if (!addedAbstracts.length && !addedNums.length) return;
  const dir = documentDir(ctx);
  let path: string | null = null;
  for (const rel of [...ctx.rels.values(), ...ctx.addedRels.values()]) {
    if (rel.type === NUMBERING_TYPE) path = dir + rel.target.replace(/^\//, "");
  }
  let xml = path ? await ctx.zip.file(path)?.async("string") : undefined;
  if (!path || !xml) {
    path = `${dir}numbering.xml`;
    xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>`;
    ctx.addedRels.set(newRelId(ctx), { type: NUMBERING_TYPE, target: "numbering.xml", external: false });
    await addContentType(ctx, `/${path}`, NUMBERING_CONTENT);
  }
  // All abstract definitions come before the list instances.
  const firstNum = xml.search(/<w:num[\s>]/);
  const end = xml.lastIndexOf("</w:numbering>");
  const abstracts = addedAbstracts.join("");
  xml = firstNum >= 0 ? xml.slice(0, firstNum) + abstracts + xml.slice(firstNum, end) + addedNums.join("") + xml.slice(end) : xml.slice(0, end) + abstracts + addedNums.join("") + xml.slice(end);
  ctx.zip.file(path, xml);
  ctx.numbering.addedAbstracts = [];
  ctx.numbering.addedNums = [];
}

async function writeRels(ctx: DocContext) {
  if (!ctx.addedRels.size) return;
  const dir = documentDir(ctx);
  const path = `${dir}_rels/${ctx.documentPath.slice(dir.length)}.rels`;
  let xml = await ctx.zip.file(path)?.async("string");
  if (!xml) xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const added = [...ctx.addedRels]
    .map(([id, rel]) => `<Relationship Id="${escapeAttr(id)}" Type="${escapeAttr(rel.type)}" Target="${escapeAttr(rel.target)}"${rel.external ? ' TargetMode="External"' : ""}/>`)
    .join("");
  const end = xml.lastIndexOf("</Relationships>");
  ctx.zip.file(path, xml.slice(0, end) + added + xml.slice(end));
  for (const [id, rel] of ctx.addedRels) ctx.rels.set(id, rel);
  ctx.addedRels.clear();
}

async function addContentType(ctx: DocContext, partName: string, contentType: string) {
  const path = "[Content_Types].xml";
  const xml = (await ctx.zip.file(path)?.async("string")) ?? "";
  if (xml.includes(`PartName="${partName}"`)) return;
  const end = xml.lastIndexOf("</Types>");
  ctx.zip.file(path, xml.slice(0, end) + `<Override PartName="${partName}" ContentType="${contentType}"/>` + xml.slice(end));
}
