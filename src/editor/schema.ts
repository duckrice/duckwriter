// The shape of a document in the editor. It follows Word's own model closely
// (paragraphs with properties, runs of formatted text) so that opening and
// saving a .docx keeps what was there. Anything DuckWriter can't edit yet is
// kept as a "raw" node holding the original XML.

import { Schema, type DOMOutputSpec, type Mark, type Node, type NodeSpec } from "prosemirror-model";
import { tableNodes } from "prosemirror-tables";
import { HIGHLIGHTS } from "../docx/model";

const pt = (value: number) => `${Math.round(value * 100) / 100}pt`;

/** The CSS for a paragraph's own (direct) formatting. Style formatting comes from the generated stylesheet. */
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
      setDOMAttr: (value, attrs) => {
        if (value) attrs.style = `${attrs.style ?? ""}vertical-align:${value};`;
      },
    },
    /** Other cell properties (borders, margins…), as XML. */
    tcPr: { default: "" },
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

export const schema = new Schema({
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
      },
      parseDOM: [{ tag: "p" }],
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
      attrs: { title: { default: "Contents" }, entries: { default: [] } },
      toDOM(node): DOMOutputSpec {
        const entries = node.attrs.entries as { level: number; text: string; page: number; pos: number }[];
        return [
          "div",
          { class: "toc", title: "Table of contents — it updates by itself" },
          ["div", { class: "toc-title" }, node.attrs.title],
          ...entries.map((e): DOMOutputSpec => [
            "div",
            { class: `toc-entry level-${e.level}`, "data-pos": String(e.pos) },
            ["span", { class: "toc-text" }, e.text],
            ["span", { class: "toc-dots" }],
            ["span", { class: "toc-page" }, String(e.page)],
          ]),
          ...(entries.length ? [] : [["div", { class: "toc-empty" }, "Headings will be listed here."] as DOMOutputSpec]),
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
        xml: { default: null },
        original: { default: null },
      },
      toDOM: (node) => ["img", { class: "picture", src: node.attrs.src, style: `width:${node.attrs.width}pt;height:${node.attrs.height}pt` }],
    },

    /** A part of a paragraph DuckWriter can't edit yet (a picture, a field, a bookmark…), kept as is. */
    raw_inline: {
      group: "inline",
      inline: true,
      atom: true,
      attrs: { xml: {}, label: { default: "" }, hidden: { default: false } },
      toDOM(node): DOMOutputSpec {
        if (node.attrs.hidden) return ["span", { class: "raw-inline hidden" }];
        return ["span", { class: "raw-inline", title: `${node.attrs.label} — kept as it is` }, node.attrs.label];
      },
    },
  },

  marks: {
    bold: {
      attrs: { on: { default: true } },
      toDOM: (mark) => ["span", { style: `font-weight:${mark.attrs.on ? "bold" : "normal"}` }, 0],
      parseDOM: [{ tag: "b" }, { tag: "strong" }],
    },
    italic: {
      attrs: { on: { default: true } },
      toDOM: (mark) => ["span", { style: `font-style:${mark.attrs.on ? "italic" : "normal"}` }, 0],
      parseDOM: [{ tag: "i" }, { tag: "em" }],
    },
    underline: {
      attrs: { on: { default: true }, kind: { default: "single" } },
      toDOM: (mark) => ["span", { class: mark.attrs.on ? `u u-${mark.attrs.kind}` : "no-u" }, 0],
      parseDOM: [{ tag: "u" }],
    },
    strike: {
      attrs: { on: { default: true } },
      toDOM: (mark) => ["span", { class: mark.attrs.on ? "s" : "no-s" }, 0],
      parseDOM: [{ tag: "s" }],
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
    },
    link: {
      attrs: { href: { default: "" }, anchor: { default: null }, rid: { default: null } },
      inclusive: false,
      toDOM: (mark) => ["a", { href: mark.attrs.href || `#${mark.attrs.anchor ?? ""}`, title: mark.attrs.href || mark.attrs.anchor || "" }, 0],
    },
    /** A character style (Word's "rStyle"). */
    charStyle: {
      attrs: { id: {} },
      toDOM: (mark) => ["span", { "data-cstyle": mark.attrs.id }, 0],
    },
    /** Character properties DuckWriter doesn't handle, as XML, written back unchanged. */
    runExtra: {
      attrs: { xml: {} },
      toDOM: () => ["span", 0],
    },
  },
});

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
