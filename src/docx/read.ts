// Opens a .docx: the text and its formatting become an editor document; the
// rest of the file (styles, lists, pictures, headers…) is kept for saving.

import JSZip from "jszip";
import type { Mark, Node } from "prosemirror-model";
import { schema } from "../editor/schema";
import {
  DEFAULT_PAGE,
  TWIPS_PER_POINT,
  type DocContext,
  type ListLevel,
  type Numbering,
  type PageSetup,
  type ParaProps,
  type Relationship,
  type RunProps,
  type StyleDef,
  type Styles,
} from "./model";
import { NS, attr, child, childrenNamed, children, declaredPrefixes, isW, numAttr, parseXml, rawXml, startTag, toggle } from "./xml";

export interface OpenedDocument {
  doc: Node;
  context: DocContext;
}

const twips = (value: number | null) => (value === null ? undefined : value / TWIPS_PER_POINT);

export async function readDocx(bytes: Uint8Array): Promise<OpenedDocument> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    throw new Error("This isn't a Word document (.docx), or the file is damaged.");
  }
  const documentPath = await mainDocumentPath(zip);
  const documentXml = await zip.file(documentPath)?.async("string");
  if (!documentXml) throw new Error("This file has no document inside. Is it a Word document (.docx)?");

  const dir = documentPath.slice(0, documentPath.lastIndexOf("/") + 1);
  const rels = await readRels(zip, `${dir}_rels/${documentPath.slice(dir.length)}.rels`);
  const partFor = (type: string) => {
    for (const rel of rels.values()) if (rel.type.endsWith(`/${type}`) && !rel.external) return resolvePath(dir, rel.target);
    return null;
  };
  const text = async (path: string | null) => (path ? ((await zip.file(path)?.async("string")) ?? null) : null);

  const themeFonts = readTheme(await text(partFor("theme")));
  const styles = readStyles(await text(partFor("styles")), themeFonts);
  const numbering = readNumbering(await text(partFor("numbering")));

  const xml = parseXml(documentXml);
  const root = xml.documentElement;
  const body = child(root, "body");
  if (!body) throw new Error("This Word document has no body.");
  const prefixes = declaredPrefixes(root);

  const reader = new BodyReader(prefixes, rels, themeFonts);
  const blocks: Node[] = [];
  let sectPr = "";
  for (const el of children(body)) {
    if (isW(el, "sectPr")) sectPr = rawXml(el, prefixes);
    else blocks.push(...reader.block(el));
  }
  if (!blocks.some((b) => b.type === schema.nodes.paragraph)) blocks.push(schema.nodes.paragraph.create());

  const context: DocContext = {
    zip,
    documentXml,
    documentPath,
    prefixes,
    styles,
    numbering,
    themeFonts,
    rels,
    addedRels: new Map(),
    sectPr,
    page: readPage(child(body, "sectPr")),
  };
  return { doc: schema.nodes.doc.create(null, blocks), context };
}

async function mainDocumentPath(zip: JSZip): Promise<string> {
  const rels = await readRels(zip, "_rels/.rels");
  for (const rel of rels.values()) if (rel.type.endsWith("/officeDocument")) return rel.target.replace(/^\//, "");
  return "word/document.xml";
}

async function readRels(zip: JSZip, path: string): Promise<Map<string, Relationship>> {
  const out = new Map<string, Relationship>();
  const text = await zip.file(path)?.async("string");
  if (!text) return out;
  for (const el of Array.from(parseXml(text).getElementsByTagNameNS(NS.rel, "Relationship"))) {
    out.set(el.getAttribute("Id") ?? "", {
      type: el.getAttribute("Type") ?? "",
      target: el.getAttribute("Target") ?? "",
      external: el.getAttribute("TargetMode") === "External",
    });
  }
  return out;
}

function resolvePath(dir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = (dir + target).split("/");
  const out: string[] = [];
  for (const part of parts) {
    if (part === "..") out.pop();
    else if (part !== ".") out.push(part);
  }
  return out.join("/");
}

// ---- theme, styles, lists, page ---------------------------------------------------

function readTheme(text: string | null): DocContext["themeFonts"] {
  if (!text) return { major: null, minor: null };
  const xml = parseXml(text);
  const latin = (scheme: string) => {
    const el = xml.getElementsByTagNameNS(NS.a, scheme)[0];
    const font = el && Array.from(el.children).find((c) => c.localName === "latin");
    return font?.getAttribute("typeface") || null;
  };
  return { major: latin("majorFont"), minor: latin("minorFont") };
}

/** The font a w:rFonts element names, resolving theme references. */
export function fontOf(el: Element | null, theme: DocContext["themeFonts"]): string | undefined {
  if (!el) return undefined;
  const direct = attr(el, "ascii") ?? attr(el, "hAnsi");
  if (direct) return direct;
  const themed = attr(el, "asciiTheme") ?? attr(el, "hAnsiTheme");
  // Without a theme, Word uses the Office theme fonts.
  if (themed) return (themed.startsWith("major") ? (theme.major ?? "Calibri Light") : (theme.minor ?? "Calibri"));
  return undefined;
}

export function readRunProps(rPr: Element | null, theme: DocContext["themeFonts"]): RunProps {
  const out: RunProps = {};
  if (!rPr) return out;
  const font = fontOf(child(rPr, "rFonts"), theme);
  if (font) out.font = font;
  const size = numAttr(child(rPr, "sz"), "val");
  if (size !== null) out.size = size / 2;
  const bold = toggle(rPr, "b");
  if (bold !== null) out.bold = bold;
  const italic = toggle(rPr, "i");
  if (italic !== null) out.italic = italic;
  const strike = toggle(rPr, "strike");
  if (strike !== null) out.strike = strike;
  const u = child(rPr, "u");
  if (u) out.underline = attr(u, "val") !== "none";
  const color = attr(child(rPr, "color"), "val");
  if (color && /^[0-9a-f]{6}$/i.test(color)) out.color = `#${color.toLowerCase()}`;
  const highlight = attr(child(rPr, "highlight"), "val");
  if (highlight && highlight !== "none") out.highlight = highlight;
  const vert = attr(child(rPr, "vertAlign"), "val");
  if (vert === "superscript" || vert === "subscript") out.vertAlign = vert;
  return out;
}

export function readParaProps(pPr: Element | null): ParaProps {
  const out: ParaProps = {};
  if (!pPr) return out;
  const jc = attr(child(pPr, "jc"), "val");
  if (jc) out.align = jc === "both" || jc === "distribute" ? "justify" : jc === "end" ? "right" : jc === "start" ? "left" : (jc as ParaProps["align"]);
  const spacing = child(pPr, "spacing");
  if (spacing) {
    const before = twips(numAttr(spacing, "before"));
    const after = twips(numAttr(spacing, "after"));
    if (before !== undefined && attr(spacing, "beforeAutospacing") !== "1") out.spaceBefore = before;
    if (after !== undefined && attr(spacing, "afterAutospacing") !== "1") out.spaceAfter = after;
    const line = numAttr(spacing, "line");
    const rule = attr(spacing, "lineRule");
    if (line !== null) {
      if (!rule || rule === "auto") out.lineSpacing = line / 240;
      else {
        out.lineHeight = line / TWIPS_PER_POINT;
        out.lineRule = rule === "exact" ? "exact" : "atLeast";
      }
    }
  }
  const ind = child(pPr, "ind");
  if (ind) {
    const left = numAttr(ind, "left") ?? numAttr(ind, "start");
    const right = numAttr(ind, "right") ?? numAttr(ind, "end");
    const first = numAttr(ind, "firstLine");
    const hanging = numAttr(ind, "hanging");
    if (left !== null) out.indentLeft = left / TWIPS_PER_POINT;
    if (right !== null) out.indentRight = right / TWIPS_PER_POINT;
    if (hanging !== null) out.indentFirst = -hanging / TWIPS_PER_POINT;
    else if (first !== null) out.indentFirst = first / TWIPS_PER_POINT;
  }
  const contextual = toggle(pPr, "contextualSpacing");
  if (contextual !== null) out.contextual = contextual;
  const numPr = child(pPr, "numPr");
  if (numPr) {
    const numId = attr(child(numPr, "numId"), "val");
    if (numId !== null) out.numId = numId;
    const ilvl = numAttr(child(numPr, "ilvl"), "val");
    if (ilvl !== null) out.ilvl = ilvl;
  }
  return out;
}

function readStyles(text: string | null, theme: DocContext["themeFonts"]): Styles {
  const styles: Styles = { defaultRun: { font: "Times New Roman", size: 10 }, defaultPara: {}, byId: new Map(), defaultParagraph: null };
  if (!text) return styles;
  const root = parseXml(text).documentElement;
  const defaults = child(root, "docDefaults");
  // With nothing said, Word uses Times New Roman at 10 points.
  styles.defaultRun = { font: "Times New Roman", size: 10, ...readRunProps(child(child(defaults, "rPrDefault"), "rPr"), theme) };
  styles.defaultPara = readParaProps(child(child(defaults, "pPrDefault"), "pPr"));
  for (const el of childrenNamed(root, "style")) {
    const id = attr(el, "styleId");
    if (!id) continue;
    const type = (attr(el, "type") ?? "paragraph") as StyleDef["type"];
    const style: StyleDef = {
      id,
      name: attr(child(el, "name"), "val") ?? id,
      type,
      basedOn: attr(child(el, "basedOn"), "val") ?? undefined,
      next: attr(child(el, "next"), "val") ?? undefined,
      quick: !!child(el, "qFormat"),
      hidden: !!child(el, "semiHidden") || !!child(el, "hidden"),
      priority: numAttr(child(el, "uiPriority"), "val") ?? 99,
      run: readRunProps(child(el, "rPr"), theme),
      para: readParaProps(child(el, "pPr")),
    };
    styles.byId.set(id, style);
    if (type === "paragraph" && attr(el, "default") === "1") styles.defaultParagraph = id;
  }
  return styles;
}

function readNumbering(text: string | null): Numbering {
  const numbering: Numbering = { abstracts: new Map(), nums: new Map(), addedAbstracts: [], addedNums: [] };
  if (!text) return numbering;
  const root = parseXml(text).documentElement;
  for (const el of childrenNamed(root, "abstractNum")) {
    const levels: ListLevel[] = [];
    for (const lvl of childrenNamed(el, "lvl")) {
      const ind = child(child(lvl, "pPr"), "ind");
      levels[numAttr(lvl, "ilvl") ?? levels.length] = {
        format: attr(child(lvl, "numFmt"), "val") ?? "decimal",
        text: attr(child(lvl, "lvlText"), "val") ?? "",
        start: numAttr(child(lvl, "start"), "val") ?? 1,
        indentLeft: (numAttr(ind, "left") ?? numAttr(ind, "start") ?? 0) / TWIPS_PER_POINT,
        hanging: (numAttr(ind, "hanging") ?? 0) / TWIPS_PER_POINT,
        font: attr(child(child(lvl, "rPr"), "rFonts"), "ascii") ?? undefined,
      };
    }
    numbering.abstracts.set(attr(el, "abstractNumId") ?? "", levels);
  }
  for (const el of childrenNamed(root, "num")) {
    const startOverrides = new Map<number, number>();
    for (const o of childrenNamed(el, "lvlOverride")) {
      const start = numAttr(child(o, "startOverride"), "val");
      if (start !== null) startOverrides.set(numAttr(o, "ilvl") ?? 0, start);
    }
    numbering.nums.set(attr(el, "numId") ?? "", { abstractId: attr(child(el, "abstractNumId"), "val") ?? "", startOverrides });
  }
  return numbering;
}

function readPage(sectPr: Element | null): PageSetup {
  if (!sectPr) return { ...DEFAULT_PAGE };
  const size = child(sectPr, "pgSz");
  const margin = child(sectPr, "pgMar");
  const get = (el: Element | null, name: string, fallback: number) => twips(numAttr(el, name)) ?? fallback;
  return {
    width: get(size, "w", DEFAULT_PAGE.width),
    height: get(size, "h", DEFAULT_PAGE.height),
    marginTop: Math.abs(get(margin, "top", DEFAULT_PAGE.marginTop)),
    marginRight: get(margin, "right", DEFAULT_PAGE.marginRight),
    marginBottom: Math.abs(get(margin, "bottom", DEFAULT_PAGE.marginBottom)),
    marginLeft: get(margin, "left", DEFAULT_PAGE.marginLeft),
  };
}

// ---- the body -------------------------------------------------------------------

/** The paragraph properties DuckWriter understands; the rest are kept as XML. */
const KNOWN_PPR = new Set(["pStyle", "jc", "spacing", "ind", "numPr", "rPr"]);
/** The character properties DuckWriter understands; the rest are kept as XML. */
const KNOWN_RPR = new Set(["rStyle", "rFonts", "b", "bCs", "i", "iCs", "strike", "u", "color", "sz", "szCs", "highlight", "vertAlign"]);

class BodyReader {
  constructor(
    private prefixes: Set<string>,
    private rels: Map<string, Relationship>,
    private theme: DocContext["themeFonts"],
  ) {}

  private raw = (el: Element) => rawXml(el, this.prefixes);

  block(el: Element): Node[] {
    if (isW(el, "p")) return [this.paragraph(el)];
    if (isW(el, "tbl")) return [this.rawBlock(el, "Table", tablePreview(el))];
    if (isW(el, "sdt")) return [this.rawBlock(el, "Content control", textOf(el).slice(0, 400))];
    if (isW(el, "bookmarkStart") || isW(el, "bookmarkEnd") || isW(el, "proofErr")) return [schema.nodes.raw_block.create({ xml: this.raw(el), hidden: true })];
    return [this.rawBlock(el, el.localName, textOf(el).slice(0, 400))];
  }

  private rawBlock(el: Element, label: string, preview: string): Node {
    return schema.nodes.raw_block.create({ xml: this.raw(el), label, preview });
  }

  paragraph(p: Element): Node {
    const pPr = child(p, "pPr");
    const props = readParaProps(pPr);
    const extra = pPr ? children(pPr).filter((c) => !(c.namespaceURI === NS.w && KNOWN_PPR.has(c.localName))).map(this.raw).join("") : "";
    // The paragraph mark's own formatting (pPr/rPr) is kept with the extras.
    const markRPr = child(pPr, "rPr");
    const attrs = {
      style: attr(child(pPr, "pStyle"), "val"),
      align: props.align ?? null,
      spaceBefore: props.spaceBefore ?? null,
      spaceAfter: props.spaceAfter ?? null,
      lineSpacing: props.lineSpacing ?? null,
      lineHeight: props.lineHeight ?? null,
      lineRule: props.lineRule ?? null,
      indentLeft: props.indentLeft ?? null,
      indentRight: props.indentRight ?? null,
      indentFirst: props.indentFirst ?? null,
      numId: props.numId ?? null,
      ilvl: props.ilvl ?? 0,
      extra: extra + (markRPr ? this.raw(markRPr) : ""),
    };
    const content: Node[] = [];
    this.inlines(p, [], content);
    return schema.nodes.paragraph.create(attrs, mergeText(content));
  }

  /** Reads the runs (and things that hold runs) inside `parent`. */
  private inlines(parent: Element, marks: readonly Mark[], out: Node[]) {
    for (const el of children(parent)) {
      if (el.namespaceURI !== NS.w) {
        // Equations and other foreign content.
        out.push(this.rawInline(el, el.localName === "oMath" || el.localName === "oMathPara" ? "Equation" : el.localName, marks));
        continue;
      }
      switch (el.localName) {
        case "pPr":
        case "proofErr":
          break;
        case "r":
          this.run(el, marks, out);
          break;
        case "hyperlink": {
          const rid = el.getAttributeNS(NS.r, "id");
          const href = rid ? (this.rels.get(rid)?.target ?? "") : "";
          const link = schema.marks.link.create({ href, anchor: attr(el, "anchor"), rid });
          this.inlines(el, link.addToSet(marks), out);
          break;
        }
        case "ins":
        case "smartTag":
        case "customXml":
        case "dir":
        case "bdo":
          // Wrappers: their runs are ordinary text. (A tracked insertion is shown as accepted.)
          this.inlines(el, marks, out);
          break;
        case "sdt":
          this.inlines(child(el, "sdtContent") ?? el, marks, out);
          break;
        case "del":
        case "moveFrom":
          // Text a tracked change deleted is no longer part of the document.
          break;
        case "fldSimple":
          out.push(this.rawInline(el, textOf(el) || "Field", marks));
          break;
        default:
          // Bookmarks, comment anchors and the like: invisible, but kept.
          out.push(schema.nodes.raw_inline.create({ xml: this.raw(el), hidden: true }));
      }
    }
  }

  private rawInline(el: Element, label: string, marks: readonly Mark[]): Node {
    return schema.nodes.raw_inline.create({ xml: this.raw(el), label }, null, marks);
  }

  private run(r: Element, outer: readonly Mark[], out: Node[]) {
    const rPr = child(r, "rPr");
    const marks = this.runMarks(rPr, outer);
    let text = "";
    const flush = () => {
      if (text) out.push(schema.text(text, marks));
      text = "";
    };
    for (const el of children(r)) {
      if (el.namespaceURI !== NS.w) continue;
      switch (el.localName) {
        case "rPr":
        case "lastRenderedPageBreak":
          break;
        case "t":
          text += el.textContent ?? "";
          break;
        case "tab":
          text += "\t";
          break;
        case "noBreakHyphen":
          text += "‑";
          break;
        case "softHyphen":
          text += "­";
          break;
        case "cr":
          flush();
          out.push(schema.nodes.hard_break.create(null, null, marks));
          break;
        case "br": {
          flush();
          const type = attr(el, "type");
          if (type === "page") out.push(schema.nodes.page_break.create(null, null, marks));
          else if (!type || type === "textWrapping") out.push(schema.nodes.hard_break.create(null, null, marks));
          else out.push(this.runPiece(r, el, "Column break", marks));
          break;
        }
        default: {
          flush();
          const label =
            el.localName === "drawing" || el.localName === "pict" || el.localName === "object"
              ? "Picture"
              : el.localName === "footnoteReference" || el.localName === "endnoteReference"
                ? "Note"
                : el.localName === "sym"
                  ? String.fromCharCode(parseInt(attr(el, "char") ?? "3F", 16) & 0xff)
                  : "";
          out.push(this.runPiece(r, el, label, marks));
        }
      }
    }
    flush();
  }

  /** One special piece of a run, kept as a run of its own with the run's formatting. */
  private runPiece(r: Element, el: Element, label: string, marks: readonly Mark[]): Node {
    const rPr = child(r, "rPr");
    const xml = `${startTag(r)}>${rPr ? this.raw(rPr) : ""}${this.raw(el)}</${r.tagName}>`;
    return schema.nodes.raw_inline.create({ xml, label, hidden: !label }, null, marks);
  }

  runMarks(rPr: Element | null, outer: readonly Mark[]): readonly Mark[] {
    let marks = outer;
    if (!rPr) return marks;
    const props = readRunProps(rPr, this.theme);
    const add = (mark: Mark) => (marks = mark.addToSet(marks));
    const style = attr(child(rPr, "rStyle"), "val");
    if (style) add(schema.marks.charStyle.create({ id: style }));
    if (props.font) add(schema.marks.font.create({ family: props.font, xml: this.raw(child(rPr, "rFonts")!) }));
    if (props.size !== undefined) add(schema.marks.size.create({ pt: props.size }));
    if (props.bold !== undefined) add(schema.marks.bold.create({ on: props.bold }));
    if (props.italic !== undefined) add(schema.marks.italic.create({ on: props.italic }));
    if (props.strike !== undefined) add(schema.marks.strike.create({ on: props.strike }));
    if (props.underline !== undefined) add(schema.marks.underline.create({ on: props.underline, kind: attr(child(rPr, "u"), "val") ?? "single" }));
    if (props.color) add(schema.marks.color.create({ color: props.color }));
    if (props.highlight) add(schema.marks.highlight.create({ color: props.highlight }));
    if (props.vertAlign) add(schema.marks.vertAlign.create({ value: props.vertAlign }));
    const extra = children(rPr).filter((c) => !(c.namespaceURI === NS.w && KNOWN_RPR.has(c.localName)));
    if (extra.length) add(schema.marks.runExtra.create({ xml: extra.map(this.raw).join("") }));
    return marks;
  }
}

/** Joins neighbouring text with the same formatting into one text node. */
function mergeText(nodes: Node[]): Node[] {
  const out: Node[] = [];
  for (const node of nodes) {
    const last = out[out.length - 1];
    if (last?.isText && node.isText && last.sameMarkup(node)) out[out.length - 1] = schema.text(last.text! + node.text!, node.marks);
    else out.push(node);
  }
  return out;
}

function textOf(el: Element): string {
  let out = "";
  const walk = (node: Element) => {
    for (const c of children(node)) {
      if (isW(c, "t")) out += c.textContent ?? "";
      else if (isW(c, "tab")) out += " ";
      else if (isW(c, "p") && out && !out.endsWith("\n")) {
        walk(c);
        out += "\n";
      } else walk(c);
    }
  };
  walk(el);
  return out.trim();
}

function tablePreview(tbl: Element): string {
  return childrenNamed(tbl, "tr")
    .slice(0, 12)
    .map((tr) => childrenNamed(tr, "tc").map((tc) => textOf(tc).replace(/\s+/g, " ")).join("  |  "))
    .join("\n");
}
