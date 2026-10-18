// Saves an editor document back into its .docx. Only the body of the main
// document is rewritten (plus list definitions and links that were added);
// every other part of the original file is kept exactly as it was.

import type { Mark, Node } from "prosemirror-model";
import { schema } from "../editor/schema";
import { parseMath } from "../math/ast";
import { adjustPictureXml, imageType, pictureXml, withDescription, type PictureLayout } from "./drawing";
import { toOmml } from "../math/omml";
import { TableMap } from "prosemirror-tables";
import { TWIPS_PER_POINT, originalXml, type CommentData, type DocContext, type Relationship } from "./model";
import { escapeAttr, escapeText, tag } from "./xml";
import { writeBibliography } from "./bibliography";
import { writeProperties } from "./properties";

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
  "specVanish", "oMath", "rPrChange",
];

const TBLPR_ORDER = [
  "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
  "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription",
];
const TRPR_ORDER = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden"];
const TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"];

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
  await writeEndnotes(doc, ctx);
  await writeComments(doc, ctx);
  await writeTrackSetting(ctx);
  if (ctx.properties) await writeProperties(ctx.zip, ctx.properties, ctx.savedBy ?? null);
  await writeBibliography(
    ctx.zip,
    ctx.bibliography,
    (target) => ctx.addedRels.set(newRelId(ctx), { type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml", target, external: false }),
    (part, type) => addContentType(ctx, part, type),
  );
  await writeAddedStyles(ctx);
  await writeStyleNumbering(ctx);
  await writeStyleChanges(ctx);
  updateSectPr(ctx);
  const writer = new BodyWriter(ctx);
  const body = writer.body(doc);
  await writeMediaTypes(ctx, writer.mediaExtensions);

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
  /** @param story a header, footer or the footnotes being written, whose links and pictures live in their own part. */
  constructor(
    private ctx: DocContext,
    private story: { rels: Map<string, Relationship> } | null = null,
  ) {}

  /** Relationships this part (a header, the footnotes) got while being written. */
  readonly storyRels = new Map<string, Relationship>();

  private addStoryRel(rel: Relationship): string {
    const rels = this.story!.rels;
    for (const [id, r] of rels) if (r.type === rel.type && r.target === rel.target && r.external === rel.external) return id;
    let n = rels.size + 1;
    while (rels.has(`rId${n}`)) n++;
    const id = `rId${n}`;
    rels.set(id, rel);
    this.storyRels.set(id, rel);
    return id;
  }

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
      else if (node.type === schema.nodes.bibliography) out += this.bibliography(node);
      // The endnotes list is only shown; the notes themselves go in endnotes.xml.
      else if (node.type === schema.nodes.table) out += this.table(node);
    });
    return out;
  }

  /**
   * A table. One that wasn't changed is written exactly as it was read; an
   * edited or new one is written from the editor, with its own properties kept.
   */
  private table(node: Node): string {
    const original = originalXml.get(node);
    if (original) return original;
    const map = TableMap.get(node);
    const p = this.ctx.page;
    const fallback = Math.round(twips(p.width - p.marginLeft - p.marginRight) / Math.max(1, map.width));
    // Column widths: from the cells (resized by dragging), else as read, else equal.
    const grid: number[] = [];
    for (let c = 0; c < map.width; c++) {
      let width: number | null = null;
      for (let r = 0; r < map.height && width === null; r++) {
        const pos = map.map[r * map.width + c];
        const cell = node.nodeAt(pos)!;
        const rect = map.findCell(pos);
        const px = cell.attrs.colwidth?.[c - rect.left];
        if (px) width = Math.round(px * 15);
      }
      grid.push(width ?? node.attrs.grid[c] ?? fallback);
    }

    const props = splitElements(node.attrs.tblPr.replace(/^<w:tblPr\b[^>]*>|<\/w:tblPr>$/g, "")).filter((e) => !["tblStyle", "tblBorders", "tblGrid"].includes(localName(e)));
    const style = node.attrs.style ?? (node.attrs.tblPr ? null : "TableGrid");
    if (style) props.push(tag("w:tblStyle", { "w:val": style }));
    if (!props.some((e) => localName(e) === "tblW")) props.push(tag("w:tblW", { "w:w": 0, "w:type": "auto" }));
    if (node.attrs.borders !== null) {
      const line = node.attrs.borders ? { "w:val": "single", "w:sz": 4, "w:space": 0, "w:color": "auto" } : { "w:val": "none", "w:sz": 0, "w:space": 0, "w:color": "auto" };
      props.push(tag("w:tblBorders", {}, ...["top", "left", "bottom", "right", "insideH", "insideV"].map((side) => tag(`w:${side}`, line))));
    } else {
      const kept = /<w:tblBorders\b[\s\S]*?<\/w:tblBorders>/.exec(node.attrs.tblPr)?.[0];
      if (kept) props.push(kept);
    }
    if (!props.some((e) => localName(e) === "tblLook")) props.push('<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>');
    let out = `<w:tbl><w:tblPr>${ordered(props, TBLPR_ORDER)}</w:tblPr><w:tblGrid>${grid.map((w) => tag("w:gridCol", { "w:w": w })).join("")}</w:tblGrid>`;

    node.forEach((row, rowOffset, r) => {
      const header = row.childCount > 0 && row.firstChild!.type === schema.nodes.table_header;
      const trParts = splitElements(row.attrs.trPr);
      if (header) trParts.push("<w:tblHeader/>");
      out += `<w:tr>${trParts.length ? `<w:trPr>${ordered(trParts, TRPR_ORDER)}</w:trPr>` : ""}`;
      for (let c = 0; c < map.width; ) {
        const pos = map.map[r * map.width + c];
        const rect = map.findCell(pos);
        const cell = node.nodeAt(pos)!;
        const span = rect.right - rect.left;
        const width = grid.slice(rect.left, rect.right).reduce((a, b) => a + b, 0);
        if (rect.top === r) {
          out += this.cell(cell, width, span, cell.attrs.rowspan > 1 ? "restart" : null);
        } else {
          // Covered by a cell from a row above: Word needs a cell here saying so.
          out += this.cell(null, width, span, "continue", cell);
        }
        c = rect.right;
      }
      out += "</w:tr>";
      void rowOffset;
    });
    return out + "</w:tbl>";
  }

  private cell(cell: Node | null, width: number, span: number, vMerge: "restart" | "continue" | null, owner: Node | null = cell): string {
    const a = (cell ?? owner)!.attrs;
    const props = splitElements(a.tcPr ?? "");
    props.push(tag("w:tcW", { "w:w": width, "w:type": "dxa" }));
    if (span > 1) props.push(tag("w:gridSpan", { "w:val": span }));
    if (vMerge === "restart") props.push(tag("w:vMerge", { "w:val": "restart" }));
    else if (vMerge === "continue") props.push(tag("w:vMerge"));
    if (a.background) props.push(tag("w:shd", { "w:val": "clear", "w:color": "auto", "w:fill": String(a.background).replace("#", "").toUpperCase() }));
    if (a.valign) props.push(tag("w:vAlign", { "w:val": a.valign === "middle" ? "center" : a.valign }));
    let content = cell ? this.blocks(cell) : "<w:p/>";
    // A cell must end with a paragraph.
    if (cell && cell.lastChild?.type !== schema.nodes.paragraph) content += "<w:p/>";
    return `<w:tc><w:tcPr>${ordered(props, TCPR_ORDER)}</w:tcPr>${content}</w:tc>`;
  }

  /** A table of contents as Word writes one: a TOC field spanning its entries, so Word can update it too. */
  private toc(node: Node): string {
    const p = this.ctx.page;
    const tab = twips(p.width - p.marginLeft - p.marginRight);
    const entries = node.attrs.entries as { level: number; text: string; page: number }[];
    const label = node.attrs.label as string | null;
    let out = node.attrs.title ? `<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t xml:space="preserve">${escapeText(node.attrs.title)}</w:t></w:r></w:p>` : "";
    // A list of figures is Word's table of figures: a TOC field over the captions with that label.
    const instr = label ? ` TOC \\h \\z \\c "${escapeText(label)}" ` : ' TOC \\o "1-3" \\h \\z \\u ';
    const begin = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    if (!entries.length) return `${out}<w:p>${begin}<w:r><w:t>${label ? "No captions yet." : "No headings yet."}</w:t></w:r>${end}</w:p>`;
    entries.forEach((e, i) => {
      const pPr = `<w:pPr><w:pStyle w:val="${label ? "TableofFigures" : `TOC${Math.min(e.level, 3)}`}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${tab}"/></w:tabs></w:pPr>`;
      out += `<w:p>${pPr}${i === 0 ? begin : ""}<w:r><w:t xml:space="preserve">${escapeText(e.text)}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${e.page}</w:t></w:r>${i === entries.length - 1 ? end : ""}</w:p>`;
    });
    return out;
  }

  /** A reference list as Word writes one: a Bibliographies content control around a BIBLIOGRAPHY field. */
  private bibliography(node: Node): string {
    const entries = node.attrs.entries as { text: string; italic?: boolean }[][];
    const heading = this.ctx.styles.byId.has("Heading1") ? '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' : "";
    const title = `<w:p>${heading}<w:r><w:t xml:space="preserve">${escapeText(node.attrs.title)}</w:t></w:r></w:p>`;
    const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> BIBLIOGRAPHY </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const pPr = '<w:pPr><w:pStyle w:val="Bibliography"/></w:pPr>';
    const runs = (pieces: { text: string; italic?: boolean }[]) => pieces.map((p) => `<w:r>${p.italic ? "<w:rPr><w:i/><w:iCs/></w:rPr>" : ""}${textXml(p.text)}</w:r>`).join("");
    const paras = entries.length
      ? entries.map((e, i) => `<w:p>${pPr}${i === 0 ? begin : ""}${runs(e)}${i === entries.length - 1 ? end : ""}</w:p>`).join("")
      : `<w:p>${pPr}${begin}<w:r><w:t>No sources cited yet.</w:t></w:r>${end}</w:p>`;
    return `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Bibliographies"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>${title}<w:sdt><w:sdtPr><w:bibliography/></w:sdtPr><w:sdtContent>${paras}</w:sdtContent></w:sdt></w:sdtContent></w:sdt>`;
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
    const extras = splitElements(a.extra);
    if (a.change) {
      // A tracked change to the paragraph mark lives in the mark's run properties, first.
      const c = a.change as { kind: "ins" | "del"; id: string; author: string; date: string | null };
      const record = tag(`w:${c.kind}`, { "w:id": c.id, "w:author": c.author, "w:date": c.date });
      const i = extras.findIndex((e) => localName(e) === "rPr");
      if (i < 0) extras.push(`<w:rPr>${record}</w:rPr>`);
      else extras[i] = extras[i].replace(/^<w:rPr\b([^>]*?)\/>$/, "<w:rPr$1></w:rPr>").replace(/^<w:rPr\b[^>]*>/, (open) => open + record);
    }
    props.push(...extras);
    if (a.propsChange) {
      const c = a.propsChange as { id: string; author: string; date: string | null; before: string };
      props.push(`<w:pPrChange w:id="${escapeAttr(c.id)}" w:author="${escapeAttr(c.author)}"${c.date ? ` w:date="${escapeAttr(c.date)}"` : ""}><w:pPr>${c.before}</w:pPr></w:pPrChange>`);
    }
    const pPr = props.length ? `<w:pPr>${ordered(props, PPR_ORDER)}</w:pPr>` : "";
    return `<w:p>${pPr}${lead}${this.inlines(node)}</w:p>`;
  }

  /**
   * The paragraph's content. Neighbouring pieces with the same link share one
   * hyperlink element, and inside it, those of the same tracked change one
   * w:ins or w:del.
   */
  private inlines(para: Node): string {
    const nodes: Node[] = [];
    para.forEach((node) => nodes.push(node));
    return this.wrapped(nodes, 0);
  }

  private wrapped(nodes: Node[], level: number): string {
    const type = [schema.marks.link, schema.marks.insertion, schema.marks.deletion][level];
    if (!type) return nodes.map((n) => this.inline(n)).join("");
    let out = "";
    for (let i = 0; i < nodes.length; ) {
      const mark = type.isInSet(nodes[i].marks) ?? null;
      let j = i + 1;
      while (j < nodes.length && sameMark(type.isInSet(nodes[j].marks) ?? null, mark)) j++;
      const inner = this.wrapped(nodes.slice(i, j), level + 1);
      out += mark ? this.wrap(mark, inner) : inner;
      i = j;
    }
    return out;
  }

  private wrap(mark: Mark, content: string): string {
    if (mark.type === schema.marks.link) return this.hyperlink(mark, content);
    const a = mark.attrs;
    const name = mark.type === schema.marks.insertion ? (a.move ? "w:moveTo" : "w:ins") : a.move ? "w:moveFrom" : "w:del";
    return `<${name} w:id="${escapeAttr(a.id)}" w:author="${escapeAttr(a.author)}"${a.date ? ` w:date="${escapeAttr(a.date)}"` : ""}>${content}</${name}>`;
  }

  private inline(node: Node): string {
    if (node.isText) {
      const deleted = node.marks.some((m) => m.type === schema.marks.deletion);
      return `<w:r>${this.runProps(node.marks)}${textXml(node.text!, deleted)}</w:r>`;
    }
    if (node.type === schema.nodes.comment) return node.attrs.xml ?? commentAnchorXml(node.attrs.kind, node.attrs.id);
    if (node.type === schema.nodes.hard_break) return `<w:r>${this.runProps(node.marks)}<w:br/></w:r>`;
    if (node.type === schema.nodes.page_break) return `<w:r>${this.runProps(node.marks)}<w:br w:type="page"/></w:r>`;
    if (node.type === schema.nodes.raw_inline) return node.attrs.xml;
    if (node.type === schema.nodes.math) return this.equation(node);
    if (node.type === schema.nodes.field) return this.field(node);
    if (node.type === schema.nodes.image) return this.image(node);
    if (node.type === schema.nodes.endnote) {
      return node.attrs.xml ?? `<w:r>${this.runProps(node.marks) || '<w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>'}<w:endnoteReference w:id="${escapeAttr(node.attrs.id)}"/></w:r>`;
    }
    if (node.type === schema.nodes.footnote) {
      return node.attrs.xml ?? `<w:r>${this.runProps(node.marks) || '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'}<w:footnoteReference w:id="${escapeAttr(node.attrs.id)}"/></w:r>`;
    }
    return "";
  }

  /** File extensions of pictures added in this save, for the content types. */
  readonly mediaExtensions = new Set<string>();
  private drawingId = 0;

  private image(node: Node): string {
    const a = node.attrs;
    const layout: PictureLayout = { width: a.width, height: a.height, floating: a.floating, x: a.x, y: a.y, relX: a.relX, relY: a.relY, wrap: a.wrap };
    if (a.xml && a.original) {
      // A new description is put in the picture as read.
      const xml = a.alt !== a.altRead ? withDescription(a.xml, a.alt) : a.xml;
      const before = JSON.parse(a.original) as PictureLayout;
      if (JSON.stringify(before) === JSON.stringify(layout)) return xml;
      const adjusted = adjustPictureXml(xml, before, layout);
      if (adjusted) return adjusted;
    }
    const rid = a.rid ?? this.addMedia(a.media);
    if (!rid) return "";
    if (!this.drawingId) this.drawingId = Math.max(0, ...[...this.ctx.documentXml.matchAll(/docPr id="(\d+)"/g)].map((m) => Number(m[1]))) + 1;
    const picture = pictureXml(rid, layout, this.drawingId++, !this.ctx.prefixes.has("r"));
    return `<w:r>${this.runProps(node.marks)}${a.alt ? withDescription(picture, a.alt) : picture}</w:r>`;
  }

  /** Puts a picture added in DuckWriter into the file's media folder, once, and relates this part to it. */
  private addMedia(key: string | null): string | null {
    const entry = key ? this.ctx.media.get(key) : undefined;
    if (!entry) return null;
    const dir = documentDir(this.ctx);
    if (!entry.file) {
      let n = 1;
      while (this.ctx.zip.file(`${dir}media/duckwriter${n}.${entry.ext}`)) n++;
      entry.file = `media/duckwriter${n}.${entry.ext}`;
      this.ctx.zip.file(dir + entry.file, entry.bytes);
      this.mediaExtensions.add(entry.ext.toLowerCase());
    }
    const rel = { type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: entry.file, external: false };
    // A header's picture is related from the header's own part.
    if (this.story) return this.addStoryRel(rel);
    if (entry.rid) return entry.rid;
    entry.rid = newRelId(this.ctx);
    this.ctx.addedRels.set(entry.rid, rel);
    return entry.rid;
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
      // A link in a header is related from the header's own part.
      const known = rid ? this.story.rels.get(rid) : undefined;
      if (!known || known.target !== href) rid = this.addStoryRel({ type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", target: href, external: true });
      const ns = this.ctx.prefixes.has("r") ? "" : ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
      return `<w:hyperlink r:id="${escapeAttr(rid!)}" w:history="1"${ns}>${content}</w:hyperlink>`;
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
        case "formatChange":
          props.push(`<w:rPrChange w:id="${escapeAttr(a.id)}" w:author="${escapeAttr(a.author)}"${a.date ? ` w:date="${escapeAttr(a.date)}"` : ""}><w:rPr>${a.before}</w:rPr></w:rPrChange>`);
          break;
      }
    }
    return props.length ? `<w:rPr>${ordered(props, RPR_ORDER)}</w:rPr>` : "";
  }
}

/** A paragraph's properties as the inside of a w:pPr (without its mark's formatting or tracked changes). */
export function paraPropsXml(attrs: Node["attrs"], ctx: DocContext): string {
  const node = schema.nodes.paragraph.create({ ...attrs, change: null, propsChange: null });
  const xml = new BodyWriter(ctx).paragraph(node);
  const inner = /^<w:p><w:pPr>([\s\S]*)<\/w:pPr>/.exec(xml)?.[1] ?? "";
  return inner.replace(/<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>/g, "");
}

/** Formatting (editor marks) as the inside of a w:rPr. */
export function runPropsXml(marks: readonly Mark[], ctx: DocContext): string {
  return new BodyWriter(ctx)
    .runProps(marks.filter((m) => m.type !== schema.marks.formatChange))
    .replace(/^<w:rPr>|<\/w:rPr>$/g, "");
}

/** Text as run content: tabs and special hyphens become their own elements. Deleted text is w:delText. */
function textXml(text: string, deleted = false): string {
  let out = "";
  const t = deleted ? "w:delText" : "w:t";
  for (const piece of text.split(/([\t‑­])/)) {
    if (piece === "\t") out += "<w:tab/>";
    else if (piece === "‑") out += "<w:noBreakHyphen/>";
    else if (piece === "­") out += "<w:softHyphen/>";
    else if (piece) out += `<${t} xml:space="preserve">${escapeText(piece)}</${t}>`;
  }
  return out;
}

const sameMark = (a: Mark | null, b: Mark | null) => (a && b ? a.eq(b) : a === b);

function commentAnchorXml(kind: string, id: string): string {
  if (kind === "start") return `<w:commentRangeStart w:id="${escapeAttr(id)}"/>`;
  if (kind === "end") return `<w:commentRangeEnd w:id="${escapeAttr(id)}"/>`;
  return `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${escapeAttr(id)}"/></w:r>`;
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
    const writer = new BodyWriter(ctx, story);
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
${start}${writer.blocks(story.doc)}</${info.root}>`;
    ctx.zip.file(story.path, xml);
    await writePartRels(ctx, story.path, writer.storyRels);
    await writeMediaTypes(ctx, writer.mediaExtensions);
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
  await writePartRels(ctx, fn.path, writer.storyRels);
  await writeMediaTypes(ctx, writer.mediaExtensions);
  fn.rootStart = start;
  fn.special = special;
  fn.saved = new Map(ids.map((id) => [id, fn.notes.get(id)!]));
}

// ---- comments -----------------------------------------------------------------------

const COMMENTS = {
  type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  content: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
};
const COMMENTS_EXTENDED = {
  type: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
  content: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
};
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

/** The comments the document (with its headers, footers and footnotes) still has anchors for, in order. */
export function referencedComments(doc: Node, ctx: DocContext): string[] {
  const ids: string[] = [];
  const visit = (d: Node) =>
    d.descendants((n) => {
      if (n.type === schema.nodes.comment && !ids.includes(n.attrs.id)) ids.push(n.attrs.id);
    });
  visit(doc);
  for (const story of ctx.stories) visit(story.doc);
  for (const note of ctx.footnotes.notes.values()) visit(note);
  return ids.filter((id) => ctx.comments.items.has(id));
}

/** A new Word paragraph id: eight hex digits, below 80000000. */
const newParaId = () => Math.floor(Math.random() * 0x7fffffff).toString(16).toUpperCase().padStart(8, "0");

/** Gives the last paragraph of a comment written by another program its id. */
function withParaId(xml: string, id: string): string {
  const at = Math.max(xml.lastIndexOf("<w:p>"), xml.lastIndexOf("<w:p "));
  return at < 0 ? xml : `${xml.slice(0, at)}<w:p w14:paraId="${id}"${xml.slice(at + 4)}`;
}

function commentXml(c: CommentData): string {
  if (c.xml && c.text === c.original) return c.xml;
  const lines = c.text.split("\n");
  c.paraId ??= newParaId();
  const paras = lines.map((line, i) => {
    const id = i === lines.length - 1 ? c.paraId! : newParaId();
    const mark = i === 0 ? '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' : "";
    return `<w:p w14:paraId="${id}" w14:textId="77777777"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>${mark}<w:r>${textXml(line)}</w:r></w:p>`;
  });
  return `<w:comment w:id="${escapeAttr(c.id)}" w:author="${escapeAttr(c.author)}"${c.date ? ` w:date="${escapeAttr(c.date)}"` : ""} w:initials="${escapeAttr(c.initials)}">${paras.join("")}</w:comment>`;
}

async function writeComments(doc: Node, ctx: DocContext) {
  const cm = ctx.comments;
  const ids = referencedComments(doc, ctx);
  if (cm.version === cm.savedVersion && ids.join(",") === cm.savedIds) return;
  if (!ids.length && !cm.path) return;
  const dir = documentDir(ctx);
  if (!cm.path) {
    cm.path = `${dir}comments.xml`;
    ctx.addedRels.set(newRelId(ctx), { type: COMMENTS.type, target: "comments.xml", external: false });
    await addContentType(ctx, `/${cm.path}`, COMMENTS.content);
  }
  let start = cm.rootStart ?? `${NEW_ROOT("w:comments").slice(0, -1)} xmlns:w14="${W14}">`;
  // New and edited comments have paragraph ids, for replies and "resolved".
  if (!/xmlns:w14=/.test(start)) start = start.replace(/>$/, ` xmlns:w14="${W14}">`);
  const items = ids.map((id) => cm.items.get(id)!);
  // Resolved comments and those with replies need a paragraph id for Word to say so.
  for (const c of items) {
    const needed = c.done || c.parent || items.some((r) => r.parent === c.id);
    if (needed && !c.paraId) {
      c.paraId = newParaId();
      if (c.xml) c.xml = withParaId(c.xml, c.paraId);
    }
  }
  const body = items.map(commentXml).join("");
  ctx.zip.file(cm.path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${start}${body}</w:comments>`);
  cm.rootStart = start;

  // Replies and resolved comments: Word's commentsExtended part.
  const threaded = items.filter((c) => c.paraId);
  if (cm.extendedPath || threaded.some((c) => c.parent || c.done)) {
    if (!cm.extendedPath) {
      cm.extendedPath = `${dir}commentsExtended.xml`;
      ctx.addedRels.set(newRelId(ctx), { type: COMMENTS_EXTENDED.type, target: "commentsExtended.xml", external: false });
      await addContentType(ctx, `/${cm.extendedPath}`, COMMENTS_EXTENDED.content);
    }
    const exStart = cm.extendedRootStart ?? '<w15:commentsEx xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w15">';
    const entries = threaded.map((c) => {
      const parent = c.parent ? cm.items.get(c.parent)?.paraId : null;
      return `<w15:commentEx w15:paraId="${c.paraId}"${parent ? ` w15:paraIdParent="${parent}"` : ""} w15:done="${c.done ? 1 : 0}"/>`;
    });
    ctx.zip.file(cm.extendedPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${exStart}${entries.join("")}</w15:commentsEx>`);
    cm.extendedRootStart = exStart;
  }
  // Written as they are now: unchanged from here on.
  for (const c of items) {
    if (c.text !== c.original) {
      c.xml = commentXml(c);
      c.original = c.text;
    }
  }
  cm.savedVersion = cm.version;
  cm.savedIds = ids.join(",");
}

/** Whether Word should go on tracking changes when it opens the document. */
async function writeTrackSetting(ctx: DocContext) {
  if (ctx.trackChanges === ctx.trackChangesSaved) return;
  const path = `${documentDir(ctx)}settings.xml`;
  const xml = await ctx.zip.file(path)?.async("string");
  if (!xml) return;
  let out = xml.replace(/<w:trackRevisions\b[^>]*\/>/g, "");
  if (ctx.trackChanges) {
    const at = out.search(/<w:(doNotTrackMoves|doNotTrackFormatting|documentProtection|autoFormatOverride|styleLockTheme|styleLockQFSet|defaultTabStop|autoHyphenation|consecutiveHyphenLimit|hyphenationZone|doNotHyphenateCaps|showEnvelope|summaryLength|clickAndTypeStyle|defaultTableStyle|evenAndOddHeaders|bookFoldRevPrinting|bookFoldPrinting|bookFoldPrintingSheets|drawingGridHorizontalSpacing|drawingGridVerticalSpacing|displayHorizontalDrawingGridEvery|displayVerticalDrawingGridEvery|doNotUseMarginsForDrawingGridOrigin|drawingGridHorizontalOrigin|drawingGridVerticalOrigin|doNotShadeFormData|noPunctuationKerning|characterSpacingControl|printTwoOnOne|strictFirstAndLastChars|noLineBreaksAfter|noLineBreaksBefore|savePreviewPicture|doNotValidateAgainstSchema|saveInvalidXml|ignoreMixedContent|alwaysShowPlaceholderText|doNotDemarcateInvalidXml|saveXmlDataOnly|useXSLTWhenSaving|saveThroughXslt|showXMLTags|alwaysMergeEmptyNamespace|updateFields|hdrShapeDefaults|footnotePr|endnotePr|compat|docVars|rsids|mathPr|attachedSchema|themeFontLang|clrSchemeMapping|doNotIncludeSubdocsInStats|doNotAutoCompressPictures|forceUpgrade|captions|readModeInkLockDown|smartTagType|schemaLibrary|shapeDefaults|doNotEmbedSmartTags|decimalSymbol|listSeparator)\b/);
    const end = out.lastIndexOf("</w:settings>");
    out = at >= 0 ? out.slice(0, at) + "<w:trackRevisions/>" + out.slice(at) : out.slice(0, end) + "<w:trackRevisions/>" + out.slice(end);
  }
  ctx.zip.file(path, out);
  ctx.trackChangesSaved = ctx.trackChanges;
}

// ---- endnotes -------------------------------------------------------------------------

async function writeEndnotes(doc: Node, ctx: DocContext) {
  const en = ctx.endnotes;
  const ids: string[] = [];
  doc.descendants((n) => {
    if (n.type === schema.nodes.endnote && en.notes.has(n.attrs.id) && !ids.includes(n.attrs.id)) ids.push(n.attrs.id);
  });
  const unchanged = ids.length === en.saved.size && ids.every((id) => en.saved.get(id) && en.notes.get(id)!.eq(en.saved.get(id)!));
  if (unchanged && (en.path || !ids.length)) return;
  if (!en.path) {
    en.path = `${documentDir(ctx)}endnotes.xml`;
    ctx.addedRels.set(newRelId(ctx), { type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes", target: "endnotes.xml", external: false });
    await addContentType(ctx, `/${en.path}`, "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml");
    await addNoteSettings(ctx, "endnotePr", '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>');
  }
  const special = en.special.length
    ? en.special
    : [
        '<w:endnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:endnote>',
        '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>',
      ];
  const writer = new BodyWriter(ctx, en);
  const number = '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>';
  const notes = ids.map((id) => `<w:endnote w:id="${escapeAttr(id)}">${writer.blocks(en.notes.get(id)!, number)}</w:endnote>`);
  const start = en.rootStart ?? NEW_ROOT("w:endnotes");
  ctx.zip.file(en.path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${start}${special.join("")}${notes.join("")}</w:endnotes>`);
  await writePartRels(ctx, en.path, writer.storyRels);
  await writeMediaTypes(ctx, writer.mediaExtensions);
  en.rootStart = start;
  en.special = special;
  en.saved = new Map(ids.map((id) => [id, en.notes.get(id)!]));
}

/** Adds the notes' separators to the settings (Word lists them there). */
async function addNoteSettings(ctx: DocContext, name: string, props: string) {
  const path = `${documentDir(ctx)}settings.xml`;
  const xml = await ctx.zip.file(path)?.async("string");
  if (!xml || xml.includes(`<w:${name}`)) return;
  const at = xml.search(/<w:(compat|docVars|rsids|mathPr|attachedSchema|themeFontLang|clrSchemeMapping|doNotIncludeSubdocsInStats|doNotAutoCompressPictures|forceUpgrade|captions|readModeInkLockDown|smartTagType|schemaLibrary|shapeDefaults|doNotEmbedSmartTags|decimalSymbol|listSeparator)\b/);
  const end = xml.lastIndexOf("</w:settings>");
  ctx.zip.file(path, at >= 0 ? xml.slice(0, at) + props + xml.slice(at) : xml.slice(0, end) + props + xml.slice(end));
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

/** The order of a style's own elements. */
const STYLE_ORDER = ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"];

/** Writes styles changed in DuckWriter (Modify style) into the document's styles. */
async function writeStyleChanges(ctx: DocContext) {
  if (!ctx.styleChanges.size) return;
  const rel = [...ctx.rels.values(), ...ctx.addedRels.values()].find((r) => r.type.endsWith("/styles"));
  if (!rel) return;
  const path = documentDir(ctx) + rel.target.replace(/^\//, "");
  let xml = await ctx.zip.file(path)?.async("string");
  if (!xml) return;
  const writer = new BodyWriter(ctx);
  for (const [id, change] of ctx.styleChanges) {
    const match = new RegExp(`<w:style\\b[^>]*w:styleId="${id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*>([\\s\\S]*?)</w:style>`).exec(xml);
    if (!match) continue;
    const open = match[0].slice(0, match[0].indexOf(">") + 1);
    const parts = splitElements(match[1]);
    const inner = (name: string) => {
      const el = parts.find((p) => localName(p) === name);
      return el ? splitElements(el.replace(new RegExp(`^<w:${name}\\b[^>]*>|</w:${name}>$`, "g"), "").replace(new RegExp(`^<w:${name}\\b[^>]*/>$`), "")) : [];
    };
    // Run properties: the changed ones replace what was there.
    const r = change.run;
    const runMarks = [
      r.font ? schema.marks.font.create({ family: r.font }) : null,
      r.size !== undefined ? schema.marks.size.create({ pt: r.size }) : null,
      r.bold !== undefined ? schema.marks.bold.create({ on: r.bold }) : null,
      r.italic !== undefined ? schema.marks.italic.create({ on: r.italic }) : null,
      r.underline !== undefined ? schema.marks.underline.create({ on: r.underline }) : null,
      r.color ? schema.marks.color.create({ color: r.color }) : null,
    ].filter((m): m is Mark => !!m);
    const newRun = splitElements(writer.runProps(runMarks).replace(/^<w:rPr>|<\/w:rPr>$/g, ""));
    const replaced = new Set(newRun.map(localName));
    if (replaced.has("b")) replaced.add("bCs");
    if (replaced.has("i")) replaced.add("iCs");
    if (replaced.has("sz")) replaced.add("szCs");
    const rPr = [...inner("rPr").filter((e) => !replaced.has(localName(e))), ...newRun];
    // Paragraph properties: alignment and spacing.
    const p = change.para;
    const pPrParts = inner("pPr").filter((e) => !(p.align !== undefined && localName(e) === "jc"));
    if (p.align !== undefined) pPrParts.push(tag("w:jc", { "w:val": p.align === "justify" ? "both" : p.align }));
    if (p.spaceBefore !== undefined || p.spaceAfter !== undefined || p.lineSpacing !== undefined) {
      const old = pPrParts.find((e) => localName(e) === "spacing") ?? "";
      const attrs = new Map([...old.matchAll(/(w:\w+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));
      if (p.spaceBefore !== undefined) attrs.set("w:before", String(twips(p.spaceBefore)));
      if (p.spaceAfter !== undefined) attrs.set("w:after", String(twips(p.spaceAfter)));
      if (p.lineSpacing !== undefined) {
        attrs.set("w:line", String(Math.round(p.lineSpacing * 240)));
        attrs.set("w:lineRule", "auto");
      }
      const i = pPrParts.findIndex((e) => localName(e) === "spacing");
      const spacing = tag("w:spacing", Object.fromEntries(attrs));
      if (i >= 0) pPrParts[i] = spacing;
      else pPrParts.push(spacing);
    }
    const others = parts.filter((e) => localName(e) !== "rPr" && localName(e) !== "pPr");
    if (pPrParts.length) others.push(`<w:pPr>${ordered(pPrParts, PPR_ORDER)}</w:pPr>`);
    if (rPr.length) others.push(`<w:rPr>${ordered(rPr, RPR_ORDER)}</w:rPr>`);
    xml = xml.replace(match[0], `${open}${ordered(others, STYLE_ORDER)}</w:style>`);
  }
  ctx.zip.file(path, xml);
  ctx.styleChanges.clear();
}

/** Says what kind of file each new picture is, so Word can open it. */
async function writeMediaTypes(ctx: DocContext, extensions: Set<string>) {
  if (!extensions.size) return;
  const path = "[Content_Types].xml";
  let xml = (await ctx.zip.file(path)?.async("string")) ?? "";
  for (const ext of extensions) {
    if (new RegExp(`<Default Extension="${ext}"`, "i").test(xml)) continue;
    const end = xml.lastIndexOf("</Types>");
    xml = xml.slice(0, end) + `<Default Extension="${ext}" ContentType="${imageType(ext)}"/>` + xml.slice(end);
  }
  ctx.zip.file(path, xml);
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
  if (ctx.pageChanged) {
    const p = ctx.page;
    const kept = (name: string) => {
      const el = parts.find((e) => localName(e) === name) ?? "";
      const attrs = new Map([...el.matchAll(/(w:\w+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));
      parts = parts.filter((e) => localName(e) !== name);
      return attrs;
    };
    const size = kept("pgSz");
    size.set("w:w", String(twips(p.width)));
    size.set("w:h", String(twips(p.height)));
    if (p.width > p.height) size.set("w:orient", "landscape");
    else size.delete("w:orient");
    const margin = kept("pgMar");
    for (const [name, value] of [["w:top", p.marginTop], ["w:right", p.marginRight], ["w:bottom", p.marginBottom], ["w:left", p.marginLeft], ["w:header", p.headerDistance], ["w:footer", p.footerDistance]] as const) margin.set(name, String(twips(value)));
    if (!margin.has("w:gutter")) margin.set("w:gutter", "0");
    parts.push(tag("w:pgSz", Object.fromEntries(size)), tag("w:pgMar", Object.fromEntries(margin)));
    parts = parts.filter((e) => localName(e) !== "lnNumType");
    if (p.lineNumbers) parts.push(tag("w:lnNumType", { "w:countBy": p.lineNumbers.countBy, "w:restart": p.lineNumbers.restart }));
    ctx.pageChanged = false;
  }
  if (ctx.lineGridChanged) {
    parts = parts.filter((p) => localName(p) !== "docGrid");
    parts.push(ctx.page.lineGrid ? `<w:docGrid w:type="lines" w:linePitch="${Math.round(ctx.page.lineGrid * TWIPS_PER_POINT)}"/>` : '<w:docGrid w:linePitch="360"/>');
    ctx.lineGridChanged = false;
  }
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

/** Adds relationships to a part's own .rels (a header's, the footnotes'). */
async function writePartRels(ctx: DocContext, partPath: string, added: Map<string, Relationship>) {
  if (!added.size) return;
  const dir = partPath.slice(0, partPath.lastIndexOf("/") + 1);
  const path = `${dir}_rels/${partPath.slice(dir.length)}.rels`;
  let xml = await ctx.zip.file(path)?.async("string");
  if (!xml) xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const rels = [...added]
    .filter(([id]) => !xml!.includes(`Id="${id}"`))
    .map(([id, rel]) => `<Relationship Id="${escapeAttr(id)}" Type="${escapeAttr(rel.type)}" Target="${escapeAttr(rel.target)}"${rel.external ? ' TargetMode="External"' : ""}/>`)
    .join("");
  const end = xml.lastIndexOf("</Relationships>");
  ctx.zip.file(path, xml.slice(0, end) + rels + xml.slice(end));
}

async function addContentType(ctx: DocContext, partName: string, contentType: string) {
  const path = "[Content_Types].xml";
  const xml = (await ctx.zip.file(path)?.async("string")) ?? "";
  if (xml.includes(`PartName="${partName}"`)) return;
  const end = xml.lastIndexOf("</Types>");
  ctx.zip.file(path, xml.slice(0, end) + `<Override PartName="${partName}" ContentType="${contentType}"/>` + xml.slice(end));
}
