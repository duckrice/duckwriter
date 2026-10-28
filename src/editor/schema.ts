// The shape of a document in the editor. It follows Word's own model closely
// (paragraphs with properties, runs of formatted text) so that opening and
// saving a .docx keeps what was there. Anything DuckWriter can't edit yet is
// kept as a "raw" node holding the original XML.

import { Schema, type DOMOutputSpec, type Mark, type MarkSpec, type Node, type NodeSpec, type TagParseRule, type StyleParseRule } from "prosemirror-model";
import { tableNodes } from "prosemirror-tables";
import { HIGHLIGHTS } from "../docx/model";

const pt = (value: number) => `${Math.round(value * 100) / 100}pt`;

/** The CSS for a paragraph's own (direct) formatting. Style formatting comes from the generated stylesheet. */
/** A Word line shape (a straight connector, as templates rule a cover page with), in points; null for other drawings. */
export function lineShape(xml: string): { width: number; height: number; thickness: number; color: string; x: number; y: number } | null {
  if (!/wordprocessingShape/.test(xml) || !/<a:prstGeom prst="(line|straightConnector1)"/.test(xml) || (xml.match(/<wps:wsp\b/g) ?? []).length !== 1) return null;
  const emu = (v: string | undefined) => (Number(v) || 0) / 12700;
  const extent = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml);
  if (!extent) return null;
  const offset = (axis: "H" | "V", from: RegExp) => {
    const m = new RegExp(`<wp:position${axis} relativeFrom="(\\w+)">\\s*<wp:posOffset>(-?\\d+)</wp:posOffset>`).exec(xml);
    return m && from.test(m[1]) ? emu(m[2]) : 0;
  };
  const line = /<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>/.exec(xml);
  if (line && /<a:noFill\/>/.test(line[2])) return null;
  return {
    width: emu(extent[1]),
    height: emu(extent[2]),
    // Word's default line is three-quarters of a point.
    thickness: Math.max(0.5, emu(/w="(\d+)"/.exec(line?.[1] ?? "")?.[1] ?? "9525")),
    color: `#${/<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(line?.[2] ?? "")?.[1] ?? "000000"}`,
    // Across from the column or margin, and down from the paragraph or line it is anchored in.
    x: offset("H", /^(column|margin)$/),
    y: offset("V", /^(paragraph|line)$/),
  };
}

export function paragraphCss(attrs: Node["attrs"]): string {
  const css: string[] = [];
  if (attrs.align) css.push(`text-align:${attrs.align}`);
  if (attrs.spaceBefore !== null) css.push(`padding-top:${pt(attrs.spaceBefore)}`);
  if (attrs.spaceAfter !== null) css.push(`padding-bottom:${pt(attrs.spaceAfter)}`);
  if (attrs.indentLeft !== null) css.push(`padding-left:${pt(attrs.indentLeft)}`);
  if (attrs.indentRight !== null) css.push(`padding-right:${pt(attrs.indentRight)}`);
  if (attrs.indentFirst !== null) css.push(`text-indent:${pt(attrs.indentFirst)}`);
  if (attrs.lineSpacing !== null) css.push(`line-height:${lineHeightFor(attrs.lineSpacing)}`);
  else if (attrs.lineHeight !== null) css.push(attrs.lineRule === "atLeast" ? `line-height:max(${pt(attrs.lineHeight)}, ${lineHeightFor(1)})` : `line-height:${pt(attrs.lineHeight)}`);
  return css.join(";");
}

/**
 * Word's single line spacing follows the font's own line height; for the
 * common Windows fonts that is close to 1.17 times the font size.
 */
export function lineHeightFor(multiple: number): string {
  return String(Math.round(multiple * 1.17 * 1000) / 1000);
}

/**
 * Tables, Word style: cells hold paragraphs (or more tables), may span rows
 * and columns, and have a width, a shading and a vertical alignment. Table,
 * row and cell properties DuckWriter doesn't handle are kept as XML.
 */
const tables = tableNodes({
  tableGroup: "block",
  cellContent: "block+",
  cellAttributes: {
    background: {
      default: null,
      getFromDOM: (dom) => (dom as HTMLElement).style.backgroundColor || null,
      setDOMAttr: (value, attrs) => {
        if (value) attrs.style = `${attrs.style ?? ""}background-color:${value};`;
      },
    },
    valign: {
      default: null,
      getFromDOM: (dom) => (dom as HTMLElement).style.verticalAlign || null,
      setDOMAttr: (value, attrs) => {
        if (value) attrs.style = `${attrs.style ?? ""}vertical-align:${value};`;
      },
    },
    /** Other cell properties (borders, margins…), as XML. */
    tcPr: {
      default: "",
      getFromDOM: (dom) => (dom as HTMLElement).getAttribute("data-tcpr") ?? "",
      setDOMAttr: (value, attrs) => {
        if (value) attrs["data-tcpr"] = value;
      },
    },
  },
});
const tableSpec: NodeSpec = {
  ...tables.table,
  attrs: {
    /** The table's properties as read (width, alignment, look…), as XML. */
    tblPr: { default: "" },
    style: { default: null },
    /** Lines around the cells: true, false, or null to follow the table's style. */
    borders: { default: null },
    /** Column widths in twentieths of a point, from the file. */
    grid: { default: [] },
  },
};
const rowSpec: NodeSpec = { ...tables.table_row, attrs: { trPr: { default: "" } } };

export const schema = new Schema(clipboardReady({
  nodes: {
    doc: { content: "block+" },

    paragraph: {
      group: "block",
      content: "inline*",
      attrs: {
        style: { default: null },
        align: { default: null },
        spaceBefore: { default: null },
        spaceAfter: { default: null },
        lineSpacing: { default: null },
        lineHeight: { default: null },
        lineRule: { default: null },
        indentLeft: { default: null },
        indentRight: { default: null },
        indentFirst: { default: null },
        numId: { default: null },
        ilvl: { default: 0 },
        /** Paragraph properties DuckWriter doesn't handle, as XML, written back unchanged. */
        extra: { default: "" },
        /**
         * The paragraph mark's own formatting (Word's w:pPr/w:rPr), as marks in JSON: an empty
         * line keeps the font it was given, and Enter passes it on. Null for none.
         */
        mark: { default: null },
        /** A tracked change to the paragraph mark: { kind: "ins" | "del", id, author, date }, or null. */
        change: { default: null },
        /** Its formatting changed while changes were tracked (w:pPrChange): { id, author, date, before } (before: the old pPr, as XML). */
        propsChange: { default: null },
        /**
         * The Word content control it's in (a form field, a page number box): { id, pr, text }, where
         * pr is the control's own settings as XML and text what it held when read. Paragraphs with
         * the same id are written back inside one control.
         */
        sdt: { default: null },
      },
      parseDOM: [
        { tag: "p", getAttrs: (el) => pastedParagraph(el as HTMLElement) },
        ...[1, 2, 3, 4, 5, 6].map((n) => ({ tag: `h${n}`, getAttrs: (el: HTMLElement | string) => ({ ...pastedParagraph(el as HTMLElement), style: `Heading${Math.min(n, 3)}` }) })),
      ],
      toDOM(node): DOMOutputSpec {
        const attrs: Record<string, string> = { class: "para" };
        if (node.attrs.style) attrs["data-style"] = node.attrs.style;
        const css = paragraphCss(node.attrs);
        if (css) attrs.style = css;
        return ["p", attrs, 0];
      },
    },

    // After the paragraph: an empty cell (or document) is filled with the first kind of block listed.
    table: tableSpec,
    table_row: rowSpec,
    table_cell: tables.table_cell,
    table_header: tables.table_header,

    /** Something at the top level DuckWriter can't edit yet (a table, a content control…), kept as is. */
    raw_block: {
      group: "block",
      atom: true,
      selectable: true,
      attrs: { xml: {}, label: { default: "" }, preview: { default: "" }, hidden: { default: false } },
      toDOM(node): DOMOutputSpec {
        if (node.attrs.hidden) return ["div", { class: "raw-block hidden" }];
        return ["div", { class: "raw-block", title: `${node.attrs.label} — kept as it is` }, ["div", { class: "raw-label" }, node.attrs.label], ["div", { class: "raw-preview" }, node.attrs.preview]];
      },
    },

    /**
     * A table of contents, made from the headings. Its entries (text and page)
     * are kept up to date as the text and the pages change.
     */
    toc: {
      group: "block",
      atom: true,
      selectable: true,
      /** label: null for the headings, or a caption label ("Figure", "Table") for a list of those. */
      attrs: { title: { default: "Contents" }, entries: { default: [] }, label: { default: null } },
      toDOM(node): DOMOutputSpec {
        const entries = node.attrs.entries as { level: number; text: string; page: number | string; pos: number }[];
        return [
          "div",
          { class: "toc", title: node.attrs.label ? `List of ${String(node.attrs.label).toLowerCase()}s — it updates by itself` : "Table of contents — it updates by itself" },
          ["div", { class: "toc-title" }, node.attrs.title],
          ...entries.map((e): DOMOutputSpec => [
            "div",
            { class: `toc-entry level-${e.level}`, "data-pos": String(e.pos) },
            ["span", { class: "toc-text" }, e.text],
            ["span", { class: "toc-dots" }],
            ["span", { class: "toc-page" }, String(e.page)],
          ]),
          ...(entries.length ? [] : [["div", { class: "toc-empty" }, node.attrs.label ? `Captions (${node.attrs.label} 1, ${node.attrs.label} 2…) will be listed here.` : "Headings will be listed here."] as DOMOutputSpec]),
        ];
      },
    },

    /**
     * The reference list (Word's bibliography): the cited sources, in the
     * document's citation style, kept up to date as citations are added.
     */
    bibliography: {
      group: "block",
      atom: true,
      selectable: true,
      attrs: { title: { default: "References" }, entries: { default: [] } },
      toDOM(node): DOMOutputSpec {
        const entries = node.attrs.entries as { text: string; italic?: boolean }[][];
        return [
          "div",
          { class: "bibliography", title: "Reference list — it follows the citations by itself" },
          ["div", { class: "bibliography-title" }, node.attrs.title],
          ...entries.map((pieces): DOMOutputSpec => ["p", { class: "bibliography-entry" }, ...pieces.map((p): DOMOutputSpec => (p.italic ? ["i", p.text] : ["span", p.text]))]),
          ...(entries.length ? [] : [["div", { class: "toc-empty" }, "Sources you cite are listed here (Alt+Shift+C cites one)."] as DOMOutputSpec]),
        ];
      },
    },

    text: { group: "inline" },

    hard_break: {
      group: "inline",
      inline: true,
      selectable: false,
      attrs: { xml: { default: null } },
      parseDOM: [{ tag: "br" }],
      toDOM: () => ["br"],
    },

    page_break: {
      group: "inline",
      inline: true,
      atom: true,
      toDOM: () => ["span", { class: "page-break", title: "Page break" }],
    },

    /**
     * An equation, kept as its shorthand (LaTeX-like). `xml` is the Word
     * equation it was read from, written back unchanged unless edited.
     */
    math: {
      group: "inline",
      inline: true,
      atom: true,
      selectable: true,
      attrs: { latex: { default: "" }, display: { default: false }, xml: { default: null }, original: { default: null } },
      toDOM: (node) => ["span", { class: "math", "data-latex": node.attrs.latex }, node.attrs.latex],
    },

    /**
     * A field: text Word works out itself, like a page number (PAGE), the
     * number of pages (NUMPAGES), a figure number (SEQ Figure) or a reference
     * to one (REF). `result` is the text shown; `xml` is the field as read,
     * written back unchanged while the instruction and result stay the same.
     */
    field: {
      group: "inline",
      inline: true,
      atom: true,
      attrs: { instr: {}, result: { default: "" }, xml: { default: null }, original: { default: null } },
      toDOM: (node) => ["span", { class: "field", "data-instr": fieldKind(node.attrs.instr), title: node.attrs.instr.trim() }, node.attrs.result || " "],
    },

    /**
     * The mark of a footnote in the text. Its number comes from its place in
     * the document; the note itself lives with the document's footnotes.
     */
    footnote: {
      group: "inline",
      inline: true,
      atom: true,
      attrs: { id: {}, xml: { default: null } },
      toDOM: (node) => ["span", { class: "footnote-ref", "data-id": node.attrs.id, title: "Footnote" }],
    },

    /**
     * A picture, in line with the text or floating on the page. `rid` is its
     * image in the file; `media` a picture added since opening. `xml` is the
     * picture as read, adjusted rather than rewritten when it is resized or moved.
     */
    image: {
      group: "inline",
      inline: true,
      atom: true,
      draggable: true,
      attrs: {
        rid: { default: null },
        media: { default: null },
        src: { default: "" },
        width: {},
        height: {},
        floating: { default: false },
        x: { default: 0 },
        y: { default: 0 },
        relX: { default: "column" },
        relY: { default: "paragraph" },
        wrap: { default: "square" },
        /** A description of the picture for people who can't see it (alt text), and as read. */
        alt: { default: "" },
        altRead: { default: "" },
        xml: { default: null },
        original: { default: null },
      },
      toDOM: (node) => ["img", { class: "picture", src: node.attrs.src, alt: node.attrs.alt, style: `width:${node.attrs.width}pt;height:${node.attrs.height}pt` }],
      // Pasted pictures; the paste plugin gives them their own copy (or leaves out ones it can't read).
      parseDOM: [
        {
          tag: "img[src]",
          getAttrs: (dom) => {
            const img = dom as HTMLImageElement;
            // Points from the style DuckWriter writes, else pixels (three quarters of a point).
            const size = (name: "width" | "height") => {
              const style = new RegExp(`${name}:\\s*([\\d.]+)(pt|px)`).exec(img.getAttribute("style") ?? "");
              if (style) return Number(style[1]) * (style[2] === "px" ? 0.75 : 1);
              return Number(img.getAttribute(name)) * 0.75 || (name === "width" ? img.naturalWidth : img.naturalHeight) * 0.75 || 150;
            };
            return { src: img.getAttribute("src"), alt: img.getAttribute("alt") ?? "", width: size("width"), height: size("height") };
          },
        },
      ],
    },

    /**
     * Where a comment starts or ends in the text, or its mark (Word's
     * commentReference, after the end). The comment itself lives with the
     * document's comments.
     */
    comment: {
      group: "inline",
      inline: true,
      atom: true,
      selectable: false,
      attrs: { id: {}, kind: {}, xml: { default: null } },
      toDOM: (node) => ["span", { class: `comment-anchor comment-${node.attrs.kind}`, "data-id": node.attrs.id }],
    },

    /** The mark of an endnote in the text; numbered i, ii, iii by its place. */
    endnote: {
      group: "inline",
      inline: true,
      atom: true,
      attrs: { id: {}, xml: { default: null } },
      toDOM: (node) => ["span", { class: "endnote-ref", "data-id": node.attrs.id, title: "Endnote: click to edit" }],
    },

    /** The list of endnotes at the end of the document (drawn by its node view; not part of the file's text). */
    endnotes: {
      group: "block",
      atom: true,
      selectable: false,
      attrs: { version: { default: 0 } },
      toDOM: () => ["div", { class: "endnotes" }],
    },

    /** A part of a paragraph DuckWriter can't edit yet (a picture, a field, a bookmark…), kept as is. */
    raw_inline: {
      group: "inline",
      inline: true,
      atom: true,
      attrs: { xml: {}, label: { default: "" }, hidden: { default: false } },
      toDOM(node): DOMOutputSpec {
        if (node.attrs.hidden) return ["span", { class: "raw-inline hidden" }];
        // A line is drawn as one, where it sits below its paragraph's start (it keeps its place in the file as is).
        const line = lineShape(node.attrs.xml ?? "");
        if (line) {
          const across = line.width >= line.height;
          return ["span", { class: "shape-line", title: "A line (as in Word)", style: `margin-left:${pt(line.x)};margin-top:${pt(line.y)};width:${pt(across ? line.width : line.thickness)};height:${pt(across ? line.thickness : line.height)};background:${line.color}` }];
        }
        return ["span", { class: "raw-inline", title: `${node.attrs.label} — kept as it is` }, node.attrs.label];
      },
    },
  },

  marks: {
    bold: {
      attrs: { on: { default: true } },
      toDOM: (mark) => ["span", { style: `font-weight:${mark.attrs.on ? "bold" : "normal"}` }, 0],
      parseDOM: [
        // Google Docs wraps everything in <b style="font-weight:normal">.
        { tag: "b", getAttrs: (el) => ((el as HTMLElement).style.fontWeight === "normal" ? false : null) },
        { tag: "strong" },
        { style: "font-weight", getAttrs: (v) => (/^(bold|bolder|[6-9]00)$/.test(v as string) ? null : false) },
      ],
    },
    italic: {
      attrs: { on: { default: true } },
      toDOM: (mark) => ["span", { style: `font-style:${mark.attrs.on ? "italic" : "normal"}` }, 0],
      parseDOM: [{ tag: "i" }, { tag: "em" }, { style: "font-style=italic" }],
    },
    underline: {
      attrs: { on: { default: true }, kind: { default: "single" } },
      toDOM: (mark) => ["span", { class: mark.attrs.on ? `u u-${mark.attrs.kind}` : "no-u" }, 0],
      parseDOM: [{ tag: "u" }, { style: "text-decoration-line", getAttrs: (v) => ((v as string).includes("underline") ? null : false) }],
    },
    strike: {
      attrs: { on: { default: true } },
      toDOM: (mark) => ["span", { class: mark.attrs.on ? "s" : "no-s" }, 0],
      parseDOM: [{ tag: "s" }, { tag: "strike" }, { style: "text-decoration-line", getAttrs: (v) => ((v as string).includes("line-through") ? null : false) }],
    },
    font: {
      attrs: { family: {}, xml: { default: null } },
      toDOM: (mark) => ["span", { style: `font-family:${cssFontFamily(mark.attrs.family)}` }, 0],
    },
    size: {
      attrs: { pt: {} },
      toDOM: (mark) => ["span", { style: `font-size:${pt(mark.attrs.pt)}` }, 0],
    },
    color: {
      attrs: { color: {} },
      toDOM: (mark) => ["span", { style: `color:${mark.attrs.color}` }, 0],
    },
    highlight: {
      attrs: { color: {} },
      toDOM: (mark) => ["span", { style: `background-color:${HIGHLIGHTS[mark.attrs.color] ?? mark.attrs.color}` }, 0],
    },
    vertAlign: {
      attrs: { value: {} },
      toDOM: (mark) => [mark.attrs.value === "superscript" ? "sup" : "sub", 0],
      parseDOM: [
        { tag: "sup", attrs: { value: "superscript" } },
        { tag: "sub", attrs: { value: "subscript" } },
      ],
    },
    link: {
      attrs: { href: { default: "" }, anchor: { default: null }, rid: { default: null } },
      inclusive: false,
      // Links from the web: only to web pages and e-mail addresses.
      parseDOM: [{ tag: "a[href]", getAttrs: (el) => (/^(https?:|mailto:)/i.test((el as HTMLElement).getAttribute("href") ?? "") ? { href: (el as HTMLElement).getAttribute("href") } : false) }],
      toDOM: (mark) => ["a", { href: mark.attrs.href || `#${mark.attrs.anchor ?? ""}`, title: `${mark.attrs.href.replace(/^mailto:/, "") || "A place in this document"}\nCtrl+click to follow the link` }, 0],
    },
    /**
     * A content control around text (Word's w:sdt: a template's "Title" or
     * "Author" box). `pr` is its properties as read; `placeholder` its
     * placeholder text, while it shows one.
     */
    contentControl: {
      attrs: { pr: {}, placeholder: { default: null } },
      toDOM: (mark) => ["span", { class: "content-control", title: /w:alias w:val="([^"]*)"/.exec(mark.attrs.pr)?.[1] ?? /w:tag w:val="([^"]*)"/.exec(mark.attrs.pr)?.[1] ?? "Content control" }, 0],
    },
    /** A character style (Word's "rStyle"). */
    charStyle: {
      attrs: { id: {} },
      toDOM: (mark) => ["span", { "data-cstyle": mark.attrs.id }, 0],
    },
    /** Text added while changes were tracked (Word's w:ins, or w:moveTo with move). */
    insertion: {
      attrs: { id: {}, author: { default: "" }, date: { default: null }, move: { default: false } },
      inclusive: false,
      toDOM: (mark) => ["ins", { class: `tc-ins ${authorClass(mark.attrs.author)}`, title: changeTitle(mark, "Inserted") }, 0],
    },
    /** Text removed while changes were tracked: still there until the change is accepted (w:del, w:moveFrom). */
    deletion: {
      attrs: { id: {}, author: { default: "" }, date: { default: null }, move: { default: false } },
      inclusive: false,
      toDOM: (mark) => ["del", { class: `tc-del ${authorClass(mark.attrs.author)}`, title: changeTitle(mark, "Deleted") }, 0],
    },
    /** Formatting changed while changes were tracked (w:rPrChange): `before` is the old formatting, as XML. */
    formatChange: {
      attrs: { id: {}, author: { default: "" }, date: { default: null }, before: { default: "" } },
      inclusive: false,
      toDOM: (mark) => ["span", { class: `tc-fmt ${authorClass(mark.attrs.author)}`, title: changeTitle(mark, "Formatted") }, 0],
    },
    /** Character properties DuckWriter doesn't handle, as XML, written back unchanged. */
    runExtra: {
      attrs: { xml: {} },
      toDOM: () => ["span", 0],
    },
  },
}));

/** A pasted paragraph's alignment, Word's title styles, and list (marked by flattenPastedLists). */
function pastedParagraph(el: HTMLElement): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};
  const align = el.style.textAlign;
  if (align === "center" || align === "right" || align === "justify") attrs.align = align;
  const cls = el.className ?? "";
  if (/\bMsoTitle\b/.test(cls)) attrs.style = "Title";
  else if (/\bMsoSubtitle\b/.test(cls)) attrs.style = "Subtitle";
  else if (/\bMso(Intense)?Quote\b/.test(cls)) attrs.style = "Quote";
  const list = el.getAttribute("data-list");
  if (list) {
    // Made into a real list of the document when pasted (see the paste plugin).
    attrs.numId = list === "number" ? "__number" : "__bullet";
    attrs.ilvl = Math.max(0, Math.min(8, Number(el.getAttribute("data-level")) || 0));
  }
  return attrs;
}

/**
 * Copying inside DuckWriter keeps everything: each node and mark is written
 * with its attributes (data-pm) and read back from them, before any other rule.
 * Comment anchors aren't copied, so a comment isn't doubled.
 */
function clipboardReady(spec: { nodes: Record<string, NodeSpec>; marks: Record<string, MarkSpec> }) {
  const withData = (out: DOMOutputSpec, kind: string, name: string, attrs: object): DOMOutputSpec => {
    if (!Array.isArray(out)) return out;
    const [tag, second, ...rest] = out as [string, unknown, ...unknown[]];
    const data = { [kind]: name, "data-pm": JSON.stringify(attrs) };
    if (second && typeof second === "object" && !Array.isArray(second)) return [tag, { ...(second as object), ...data }, ...rest] as DOMOutputSpec;
    return [tag, data, ...(out.length > 1 ? [second] : []), ...rest] as DOMOutputSpec;
  };
  const read = (el: HTMLElement | string) => JSON.parse((el as HTMLElement).getAttribute("data-pm") ?? "{}");
  for (const [name, node] of Object.entries(spec.nodes)) {
    if (!node.toDOM) continue;
    const toDOM = node.toDOM;
    node.toDOM = (n) => withData(toDOM(n), "data-pm-node", name, n.attrs);
    // Comment anchors and footnote marks would be doubled; a bookmark would have its name twice.
    const own: TagParseRule =
      name === "comment" || name === "footnote" || name === "endnote" || name === "endnotes"
        ? { tag: `[data-pm-node="${name}"]`, ignore: true }
        : { tag: `[data-pm-node="${name}"]`, getAttrs: (el) => (name === "raw_inline" && /^<w:bookmark/.test(read(el).xml ?? "") ? false : read(el)), priority: 100 };
    node.parseDOM = [own, ...(node.parseDOM ?? [])];
  }
  for (const [name, mark] of Object.entries(spec.marks)) {
    if (!mark.toDOM) continue;
    const toDOM = mark.toDOM;
    mark.toDOM = (m, inline) => withData(toDOM(m, inline), "data-pm-mark", name, m.attrs);
    // A content control isn't copied along with its text: there would be two with one id.
    const getAttrs = name === "contentControl" ? () => false as const : read;
    mark.parseDOM = [{ tag: `[data-pm-mark="${name}"]`, getAttrs, priority: 100 } as TagParseRule, ...((mark.parseDOM ?? []) as (TagParseRule | StyleParseRule)[])];
  }
  return spec;
}

/** Each reviewer gets a colour of their own, as in Word. */
export function authorClass(author: string): string {
  let hash = 0;
  for (const ch of author) hash = (hash * 31 + ch.charCodeAt(0)) >>> 0;
  return `author-${hash % 6}`;
}

function changeTitle(mark: Mark, what: string): string {
  const when = mark.attrs.date ? new Date(mark.attrs.date) : null;
  const date = when && !Number.isNaN(when.getTime()) ? `, ${when.toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" })}` : "";
  return `${what}${mark.attrs.move ? " (moved)" : ""} by ${mark.attrs.author || "someone"}${date}`;
}

/** The kind of a field, from its instruction: " PAGE  \* MERGEFORMAT " is "PAGE". */
export function fieldKind(instr: string): string {
  return (instr.trim().split(/\s+/)[0] ?? "").toUpperCase();
}

/** A font name as CSS, with a fallback of the same kind. */
export function cssFontFamily(family: string): string {
  const serif = /times|georgia|cambria|garamond|book|serif|palatino|constantia/i.test(family);
  const mono = /courier|consol|mono|lucida console/i.test(family);
  return `"${family.replace(/"/g, "")}", ${mono ? "monospace" : serif ? "serif" : "sans-serif"}`;
}

export const isMark = (mark: Mark, name: keyof typeof schema.marks) => mark.type === schema.marks[name];
