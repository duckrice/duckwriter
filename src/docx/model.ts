// What DuckWriter knows about a Word document beyond its text: the styles,
// list definitions, page setup and the original file, kept so that saving
// puts back everything that wasn't changed.

import type JSZip from "jszip";
import type { Node } from "prosemirror-model";

export type Align = "left" | "center" | "right" | "justify";

/** Character formatting. Sizes are in points; colours are "#rrggbb". */
export interface RunProps {
  font?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  highlight?: string;
  vertAlign?: "superscript" | "subscript";
}

/** Paragraph formatting. Lengths are in points; a negative first-line indent is a hanging indent. */
export interface ParaProps {
  align?: Align;
  spaceBefore?: number;
  spaceAfter?: number;
  /** Line spacing as a multiple of single spacing (lineRule "auto"). */
  lineSpacing?: number;
  /** A fixed line height in points (lineRule "exact" or "atLeast"). */
  lineHeight?: number;
  lineRule?: "exact" | "atLeast";
  indentLeft?: number;
  indentRight?: number;
  indentFirst?: number;
  numId?: string;
  ilvl?: number;
  /** No space between neighbouring paragraphs of the same style. */
  contextual?: boolean;
  /** Stay on the same page as the next paragraph (headings). */
  keepNext?: boolean;
  /** Always start on a new page. */
  pageBreakBefore?: boolean;
  /** Outline level: 0 for a top-level heading … 8; 9 means body text. */
  outline?: number;
}

export interface StyleDef {
  id: string;
  name: string;
  type: "paragraph" | "character" | "table" | "numbering";
  basedOn?: string;
  next?: string;
  /** Shown among the quick styles in Word's gallery. */
  quick: boolean;
  hidden: boolean;
  priority: number;
  run: RunProps;
  para: ParaProps;
}

export interface Styles {
  defaultRun: RunProps;
  defaultPara: ParaProps;
  byId: Map<string, StyleDef>;
  /** The paragraph style used when none is given (normally "Normal"). */
  defaultParagraph: string | null;
}

export interface ListLevel {
  /** "decimal", "bullet", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", "none"… */
  format: string;
  /** Like "%1." or "•": %n stands for the number at level n. */
  text: string;
  start: number;
  indentLeft: number;
  hanging: number;
  font?: string;
}

export interface Numbering {
  /** Abstract list definitions by abstractNumId: one entry per level. */
  abstracts: Map<string, ListLevel[]>;
  /** List instances by numId. */
  nums: Map<string, { abstractId: string; startOverrides: Map<number, number> }>;
  /** Definitions DuckWriter added, as XML, to write into numbering.xml. */
  addedAbstracts: string[];
  addedNums: string[];
}

export interface PageSetup {
  width: number;
  height: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  marginLeft: number;
  /** Distance from the top of the page to the header, and from the bottom to the footer. */
  headerDistance: number;
  footerDistance: number;
  /** Word's document grid for lines: every line takes a whole number of these (points), or null. */
  lineGrid: number | null;
}

/** A header or footer: its own little document, in its own part of the file. */
export interface Story {
  kind: "header" | "footer";
  /** Which pages it is for: all ("default"), the first page, or even pages. */
  type: "default" | "first" | "even";
  /** Path of its part in the file; null until it is first saved. */
  path: string | null;
  rid: string | null;
  /** The part's root start tag as read, with its namespace declarations. */
  rootStart: string | null;
  prefixes: Set<string>;
  rels: Map<string, Relationship>;
  doc: Node;
  /** The content as last saved, to tell whether it changed. */
  saved: Node | null;
}

export interface Relationship {
  type: string;
  target: string;
  external: boolean;
}

/** The document's footnotes: a part of their own, each note a little document. */
export interface Footnotes {
  /** Path of the part; null until the first footnote is saved. */
  path: string | null;
  rootStart: string | null;
  prefixes: Set<string>;
  rels: Map<string, Relationship>;
  /** The separator lines and the like (ids -1 and 0), kept as XML. */
  special: string[];
  notes: Map<string, Node>;
  /** The notes as last saved, to tell whether they changed. */
  saved: Map<string, Node>;
}

export interface DocContext {
  zip: JSZip;
  /** The original word/document.xml, which saving rewrites the body of. */
  documentXml: string;
  /** Path of the main document part inside the zip (normally word/document.xml). */
  documentPath: string;
  /** Namespace prefixes the document's root declares. */
  prefixes: Set<string>;
  styles: Styles;
  numbering: Numbering;
  themeFonts: { major: string | null; minor: string | null };
  rels: Map<string, Relationship>;
  /** Relationships added since opening, by id. */
  addedRels: Map<string, Relationship>;
  /** The document's final section properties, as XML. */
  sectPr: string;
  page: PageSetup;
  /** Headers and footers. */
  stories: Story[];
  /** The first page has its own header and footer. */
  titlePage: boolean;
  /** The line grid was switched on or off, so the section properties need updating. */
  lineGridChanged?: boolean;
  footnotes: Footnotes;
  /** Style definitions (XML) to add to the document's styles when it is saved. */
  addedStyles: string[];
  /** Pictures added in DuckWriter, by key, until saved into the file's media folder. */
  media: Map<string, { bytes: Uint8Array; ext: string; rid: string | null }>;
  /** Where to show the file's own pictures from, by relationship id. */
  imageSources: Map<string, string>;
  /** Styles whose list numbering changed (numbered headings), by style id: null switches it off. */
  styleNumbering: Map<string, { numId: string; ilvl: number } | null>;
}

export const TWIPS_PER_POINT = 20;

/** A4 with Word's default 2.54 cm margins. */
export const DEFAULT_PAGE: PageSetup = {
  width: 595.3,
  height: 841.9,
  marginTop: 72,
  marginRight: 72,
  marginBottom: 72,
  marginLeft: 72,
  headerDistance: 35.4,
  footerDistance: 35.4,
  lineGrid: null,
};

/** Follows a style's basedOn chain, nearest last, so later entries win. */
export function styleChain(styles: Styles, id: string | null | undefined): StyleDef[] {
  const chain: StyleDef[] = [];
  const seen = new Set<string>();
  let current = id ? styles.byId.get(id) : undefined;
  while (current && !seen.has(current.id)) {
    chain.unshift(current);
    seen.add(current.id);
    current = current.basedOn ? styles.byId.get(current.basedOn) : undefined;
  }
  return chain;
}

/** The paragraph formatting a style gives, with the document defaults underneath. */
export function resolvedPara(styles: Styles, id: string | null | undefined): ParaProps {
  return Object.assign({}, styles.defaultPara, ...styleChain(styles, id ?? styles.defaultParagraph).map((s) => s.para));
}

/** The character formatting a paragraph style gives, with the document defaults underneath. */
export function resolvedRun(styles: Styles, id: string | null | undefined): RunProps {
  return Object.assign({}, styles.defaultRun, ...styleChain(styles, id ?? styles.defaultParagraph).map((s) => s.run));
}

/** Formats a list number the way a level asks for. */
export function formatNumber(n: number, format: string): string {
  switch (format) {
    case "lowerLetter":
      return letters(n).toLowerCase();
    case "upperLetter":
      return letters(n);
    case "lowerRoman":
      return roman(n).toLowerCase();
    case "upperRoman":
      return roman(n);
    case "decimalZero":
      return n < 10 ? `0${n}` : String(n);
    case "none":
      return "";
    default:
      return String(n);
  }
}

function letters(n: number): string {
  // Word repeats the letter: a … z, aa … zz, aaa …
  const letter = String.fromCharCode(65 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

function roman(n: number): string {
  const table: [number, string][] = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
    [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ];
  let out = "";
  for (const [value, symbol] of table) {
    while (n >= value) {
      out += symbol;
      n -= value;
    }
  }
  return out;
}

/** Bullet characters in Symbol and Wingdings sit in the private use area; these are their Unicode look-alikes. */
const SYMBOL_BULLETS: Record<number, string> = {
  0xf0b7: "•",
  0xf0a7: "▪",
  0xf0a8: "◆",
  0xf076: "❖",
  0xf0d8: "➢",
  0xf0fc: "✓",
  0xf06e: "■",
  0xf06f: "□",
  0xf071: "❑",
  0xf0e0: "➔",
};

export function bulletText(text: string): string {
  return [...text].map((ch) => {
    const code = ch.codePointAt(0)!;
    if (SYMBOL_BULLETS[code]) return SYMBOL_BULLETS[code];
    return code >= 0xf000 && code <= 0xf0ff ? String.fromCharCode(code - 0xf000) : ch;
  }).join("");
}

/** Word's named highlight colours and how they look. */
export const HIGHLIGHTS: Record<string, string> = {
  yellow: "#ffff00",
  green: "#00ff00",
  cyan: "#00ffff",
  magenta: "#ff00ff",
  blue: "#0000ff",
  red: "#ff0000",
  darkBlue: "#000080",
  darkCyan: "#008080",
  darkGreen: "#008000",
  darkMagenta: "#800080",
  darkRed: "#800000",
  darkYellow: "#808000",
  darkGray: "#808080",
  lightGray: "#c0c0c0",
  black: "#000000",
  white: "#ffffff",
};
