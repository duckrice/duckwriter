// Opens a .docx: the text and its formatting become an editor document; the
// rest of the file (styles, lists, pictures, headers…) is kept for saving.

import JSZip from "jszip";
import type { Mark, Node } from "prosemirror-model";
import { schema } from "../editor/schema";
import { toLatex } from "../math/ast";
import { MATH_NS, fromOmml } from "../math/omml";
import { imageType, readPicture } from "./drawing";
import {
  DEFAULT_PAGE,
  originalXml,
  type CommentData,
  type Comments,
  type Footnotes,
  type TableBorders,
  type Story,
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

  // The pictures, ready to show.
  const imageSources = new Map<string, string>();
  for (const [rid, rel] of rels) {
    if (rel.external || !rel.type.endsWith("/image")) continue;
    const path = resolvePath(dir, rel.target);
    const bytes = await zip.file(path)?.async("uint8array");
    if (!bytes) continue;
    const type = imageType(path.split(".").pop() ?? "");
    imageSources.set(rid, imageUrl(bytes, type));
  }

  const reader = new BodyReader(prefixes, rels, themeFonts, imageSources);
  const blocks: Node[] = [];
  let sectPr = "";
  const bodyChildren = children(body);
  for (let i = 0; i < bodyChildren.length; i++) {
    const el = bodyChildren[i];
    if (isW(el, "sectPr")) {
      sectPr = rawXml(el, prefixes);
      continue;
    }
    // A table of contents becomes one that keeps itself up to date.
    const toc = tableOfContents(bodyChildren, i);
    if (toc) {
      let title = toc.title;
      const last = blocks[blocks.length - 1];
      if (!title && last?.type === schema.nodes.paragraph && /^TOCHeading$/i.test(last.attrs.style ?? "")) {
        title = last.textContent;
        blocks.pop();
      }
      blocks.push(schema.nodes.toc.create({ title: title || "Contents", entries: toc.entries }));
      i = toc.end;
      continue;
    }
    blocks.push(...reader.block(el));
  }
  if (!blocks.some((b) => b.type === schema.nodes.paragraph)) blocks.push(schema.nodes.paragraph.create());

  const finalSectPr = child(body, "sectPr");
  const stories: Story[] = [];
  for (const ref of finalSectPr ? children(finalSectPr) : []) {
    const kind = isW(ref, "headerReference") ? "header" : isW(ref, "footerReference") ? "footer" : null;
    const rid = ref.getAttributeNS(NS.r, "id");
    const rel = rid ? rels.get(rid) : undefined;
    if (!kind || !rid || !rel) continue;
    const path = resolvePath(dir, rel.target);
    const partXml = await text(path);
    if (!partXml) continue;
    const partDir = path.slice(0, path.lastIndexOf("/") + 1);
    const partRels = await readRels(zip, `${partDir}_rels/${path.slice(partDir.length)}.rels`);
    const partRoot = parseXml(partXml).documentElement;
    const partPrefixes = declaredPrefixes(partRoot);
    const partReader = new BodyReader(partPrefixes, partRels, themeFonts);
    const content = children(partRoot).flatMap((el) => partReader.block(el));
    if (!content.length) content.push(schema.nodes.paragraph.create());
    const doc = schema.nodes.doc.create(null, content);
    const type = (attr(ref, "type") ?? "default") as Story["type"];
    stories.push({ kind, type, path, rid, rootStart: /<w:(?:hdr|ftr)\b[^>]*>/.exec(partXml)?.[0] ?? null, prefixes: partPrefixes, rels: partRels, doc, saved: doc });
  }

  const footnotes: Footnotes = { path: null, rootStart: null, prefixes: new Set(), rels: new Map(), special: [], notes: new Map(), saved: new Map() };
  const footnotesPath = partFor("footnotes");
  const footnotesXml = await text(footnotesPath);
  if (footnotesPath && footnotesXml) {
    const partDir = footnotesPath.slice(0, footnotesPath.lastIndexOf("/") + 1);
    const partRoot = parseXml(footnotesXml).documentElement;
    footnotes.path = footnotesPath;
    footnotes.rootStart = /<w:footnotes\b[^>]*>/.exec(footnotesXml)?.[0] ?? null;
    footnotes.prefixes = declaredPrefixes(partRoot);
    footnotes.rels = await readRels(zip, `${partDir}_rels/${footnotesPath.slice(partDir.length)}.rels`);
    const noteReader = new BodyReader(footnotes.prefixes, footnotes.rels, themeFonts);
    for (const note of childrenNamed(partRoot, "footnote")) {
      const id = attr(note, "id") ?? "";
      const type = attr(note, "type");
      if (type && type !== "normal") {
        footnotes.special.push(rawXml(note, footnotes.prefixes));
        continue;
      }
      const content = children(note).flatMap((el) => noteReader.block(el));
      if (!content.length) content.push(schema.nodes.paragraph.create());
      const doc = schema.nodes.doc.create(null, content);
      footnotes.notes.set(id, doc);
      footnotes.saved.set(id, doc);
    }
  }

  const comments = await readComments(zip, partFor("comments"), partFor("commentsExtended"));
  const settings = await text(partFor("settings"));
  const trackChanges = !!settings && /<w:trackRevisions(?:\s*\/>|\s+w:val="(?:1|true|on)")/.test(settings);

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
    page: readPage(finalSectPr),
    stories,
    titlePage: !!toggle(finalSectPr, "titlePg"),
    footnotes,
    addedStyles: [],
    styleNumbering: new Map(),
    media: new Map(),
    imageSources,
    comments,
    trackChanges,
    trackChangesSaved: trackChanges,
  };
  return { doc: schema.nodes.doc.create(null, blocks), context };
}

/** The comments, with who wrote them and when, and which answer which (from commentsExtended). */
async function readComments(zip: JSZip, path: string | null, extendedPath: string | null): Promise<Comments> {
  const comments: Comments = { path: null, rootStart: null, items: new Map(), extendedPath: null, extendedRootStart: null, version: 0, savedVersion: 0, savedIds: "" };
  const xml = path ? await zip.file(path)?.async("string") : undefined;
  if (!path || !xml) return comments;
  const root = parseXml(xml).documentElement;
  const prefixes = declaredPrefixes(root);
  comments.path = path;
  comments.rootStart = /<w:comments\b[^>]*>/.exec(xml)?.[0] ?? null;
  const byParaId = new Map<string, CommentData>();
  for (const el of childrenNamed(root, "comment")) {
    const paras = childrenNamed(el, "p");
    const text = paras.map(commentParagraphText).join("\n");
    const paraId = paras.length ? paras[paras.length - 1].getAttributeNS(NS.w14, "paraId") : null;
    const item: CommentData = {
      id: attr(el, "id") ?? "",
      author: attr(el, "author") ?? "",
      initials: attr(el, "initials") ?? "",
      date: attr(el, "date"),
      text,
      xml: rawXml(el, prefixes),
      original: text,
      paraId,
      parent: null,
      done: false,
    };
    comments.items.set(item.id, item);
    if (paraId) byParaId.set(paraId.toUpperCase(), item);
  }
  comments.savedIds = [...comments.items.keys()].join(",");
  const extended = extendedPath ? await zip.file(extendedPath)?.async("string") : undefined;
  if (extendedPath && extended) {
    comments.extendedPath = extendedPath;
    comments.extendedRootStart = /<w15:commentsEx\b[^>]*>/.exec(extended)?.[0] ?? null;
    for (const ex of Array.from(parseXml(extended).getElementsByTagNameNS(NS.w15, "commentEx"))) {
      const item = byParaId.get((ex.getAttributeNS(NS.w15, "paraId") ?? "").toUpperCase());
      if (!item) continue;
      item.done = ex.getAttributeNS(NS.w15, "done") === "1";
      const parent = byParaId.get((ex.getAttributeNS(NS.w15, "paraIdParent") ?? "").toUpperCase());
      if (parent && parent !== item) item.parent = parent.id;
    }
  }
  return comments;
}

function commentParagraphText(p: Element): string {
  let out = "";
  const walk = (node: Element) => {
    for (const c of children(node)) {
      if (isW(c, "t")) out += c.textContent ?? "";
      else if (isW(c, "tab")) out += "\t";
      else if (isW(c, "br")) out += " ";
      else if (!isW(c, "instrText") && !isW(c, "pPr") && !isW(c, "rPr")) walk(c);
    }
  };
  walk(p);
  return out;
}

/** A URL to show a picture from. (Outside a browser, as in tests, there may be none.) */
function imageUrl(bytes: Uint8Array, type: string): string {
  try {
    return URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  } catch {
    return "";
  }
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
  const keepNext = toggle(pPr, "keepNext");
  if (keepNext !== null) out.keepNext = keepNext;
  const pageBreakBefore = toggle(pPr, "pageBreakBefore");
  if (pageBreakBefore !== null) out.pageBreakBefore = pageBreakBefore;
  const outline = numAttr(child(pPr, "outlineLvl"), "val");
  if (outline !== null) out.outline = outline;
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
    const style = readStyle(el, theme);
    if (!style) continue;
    styles.byId.set(style.id, style);
    if (style.type === "paragraph" && attr(el, "default") === "1") styles.defaultParagraph = style.id;
  }
  return styles;
}

function readStyle(el: Element, theme: DocContext["themeFonts"]): StyleDef | null {
  const id = attr(el, "styleId");
  if (!id) return null;
  return {
    id,
    name: attr(child(el, "name"), "val") ?? id,
    type: (attr(el, "type") ?? "paragraph") as StyleDef["type"],
    basedOn: attr(child(el, "basedOn"), "val") ?? undefined,
    next: attr(child(el, "next"), "val") ?? undefined,
    quick: !!child(el, "qFormat"),
    hidden: !!child(el, "semiHidden") || !!child(el, "hidden"),
    priority: numAttr(child(el, "uiPriority"), "val") ?? 99,
    run: readRunProps(child(el, "rPr"), theme),
    para: readParaProps(child(el, "pPr")),
    tableBorders: tableBorders(child(child(el, "tblPr"), "tblBorders")),
  };
}

/** Whether table borders show lines, and how: undefined when the element isn't there. */
export function tableBorders(borders: Element | null): TableBorders | undefined {
  if (!borders) return undefined;
  for (const side of ["insideH", "insideV", "top", "left", "bottom", "right", "start", "end"]) {
    const el = child(borders, side);
    const val = attr(el, "val");
    if (!el || !val || val === "none" || val === "nil") continue;
    const color = attr(el, "color");
    return { size: (numAttr(el, "sz") ?? 4) / 8, color: color && /^[0-9a-f]{6}$/i.test(color) ? `#${color}` : "#000000" };
  }
  return null;
}

/** A style written as XML (a <w:style> element), as DuckWriter understands it. */
export function styleFromXml(xml: string, theme: DocContext["themeFonts"]): StyleDef | null {
  const root = parseXml(`<w:styles xmlns:w="${NS.w}">${xml}</w:styles>`).documentElement;
  const el = childrenNamed(root, "style")[0];
  return el ? readStyle(el, theme) : null;
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
    headerDistance: get(margin, "header", DEFAULT_PAGE.headerDistance),
    footerDistance: get(margin, "footer", DEFAULT_PAGE.footerDistance),
    lineGrid: lineGridOf(child(sectPr, "docGrid")),
  };
}

/** The line pitch of a document grid that lines snap to ("lines" or "linesAndChars"). */
function lineGridOf(grid: Element | null): number | null {
  const type = attr(grid, "type");
  const pitch = numAttr(grid, "linePitch");
  return (type === "lines" || type === "linesAndChars") && pitch ? pitch / TWIPS_PER_POINT : null;
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
    /** Pictures this part can show, by relationship id. */
    private images: Map<string, string> = new Map(),
  ) {}

  private raw = (el: Element) => rawXml(el, this.prefixes);

  /** A field being read, which may run across several runs. */
  private capture: { depth: number; instr: string; result: string; separated: boolean; xml: string; marks: readonly Mark[] | null } | null = null;

  /** Keeps a field that can't be held together as it was, unseen. */
  private abandonField(out: Node[]) {
    if (this.capture?.xml) out.push(schema.nodes.raw_inline.create({ xml: this.capture.xml, hidden: true }));
    this.capture = null;
  }

  block(el: Element): Node[] {
    if (isW(el, "p")) return [this.paragraph(el)];
    if (isW(el, "tbl")) {
      try {
        return [this.table(el)];
      } catch (error) {
        console.warn("Table kept as it is", error);
        return [this.rawBlock(el, "Table", tablePreview(el))];
      }
    }
    if (isW(el, "sdt")) return [this.rawBlock(el, "Content control", textOf(el).slice(0, 400))];
    if (isW(el, "bookmarkStart") || isW(el, "bookmarkEnd") || isW(el, "proofErr")) return [schema.nodes.raw_block.create({ xml: this.raw(el), hidden: true })];
    return [this.rawBlock(el, el.localName, textOf(el).slice(0, 400))];
  }

  /** A Word table, with merged cells (gridSpan, vMerge) turned into column and row spans. */
  table(tbl: Element): Node {
    const tblPr = child(tbl, "tblPr");
    const grid = childrenNamed(child(tbl, "tblGrid"), "gridCol").map((g) => numAttr(g, "w") ?? 0);
    const direct = tableBorders(child(tblPr, "tblBorders"));
    interface Cell {
      el: Element;
      col: number;
      span: number;
      rowspan: number;
      covered: boolean;
    }
    const rows = childrenNamed(tbl, "tr");
    const cells: Cell[][] = [];
    // Which cell owns each column of the row above, to extend merged cells downwards.
    let above: (Cell | undefined)[] = [];
    for (const tr of rows) {
      const row: Cell[] = [];
      const here: (Cell | undefined)[] = [];
      let col = numAttr(child(child(tr, "trPr"), "gridBefore"), "val") ?? 0;
      for (const tc of childrenNamed(tr, "tc")) {
        const tcPr = child(tc, "tcPr");
        const span = Math.max(1, numAttr(child(tcPr, "gridSpan"), "val") ?? 1);
        const vMerge = child(tcPr, "vMerge");
        const continues = !!vMerge && attr(vMerge, "val") !== "restart";
        const owner = continues ? above[col] : undefined;
        const cell: Cell = { el: tc, col, span, rowspan: 1, covered: !!owner };
        if (owner) owner.rowspan++;
        for (let c = col; c < col + span; c++) here[c] = owner ?? cell;
        row.push(cell);
        col += span;
      }
      cells.push(row);
      above = here;
    }
    const rowNodes = rows.map((tr, r) => {
      const trPr = child(tr, "trPr");
      const header = !!toggle(trPr, "tblHeader");
      const trExtra = trPr ? children(trPr).filter((c) => !isW(c, "tblHeader")).map(this.raw).join("") : "";
      const cellNodes = cells[r]
        .filter((cell) => !cell.covered)
        .map((cell) => {
          const tcPr = child(cell.el, "tcPr");
          const widths = grid.slice(cell.col, cell.col + cell.span);
          const fill = attr(child(tcPr, "shd"), "fill");
          const valign = attr(child(tcPr, "vAlign"), "val");
          const known = new Set(["tcW", "gridSpan", "vMerge", "shd", "vAlign"]);
          const tcExtra = tcPr ? children(tcPr).filter((c) => !(c.namespaceURI === NS.w && known.has(c.localName))).map(this.raw).join("") : "";
          const content = children(cell.el)
            .filter((c) => !isW(c, "tcPr"))
            .flatMap((c) => this.block(c));
          if (!content.some((n) => n.type !== schema.nodes.raw_block || !n.attrs.hidden)) content.push(schema.nodes.paragraph.create());
          const type = header ? schema.nodes.table_header : schema.nodes.table_cell;
          return type.create(
            {
              colspan: cell.span,
              rowspan: cell.rowspan,
              colwidth: widths.length === cell.span && widths.every((w) => w > 0) ? widths.map((w) => Math.round(w / 15)) : null,
              background: fill && /^[0-9a-f]{6}$/i.test(fill) ? `#${fill}` : null,
              valign: valign === "center" ? "middle" : valign === "bottom" ? "bottom" : null,
              tcPr: tcExtra,
            },
            content,
          );
        });
      return schema.nodes.table_row.create({ trPr: trExtra }, cellNodes);
    });
    const node = schema.nodes.table.create(
      { tblPr: tblPr ? this.raw(tblPr) : "", style: attr(child(tblPr, "tblStyle"), "val"), borders: direct === undefined ? null : direct !== null, grid },
      rowNodes,
    );
    originalXml.set(node, this.raw(tbl));
    return node;
  }

  private rawBlock(el: Element, label: string, preview: string): Node {
    return schema.nodes.raw_block.create({ xml: this.raw(el), label, preview });
  }

  paragraph(p: Element): Node {
    const pPr = child(p, "pPr");
    const props = readParaProps(pPr);
    const extra = pPr ? children(pPr).filter((c) => !(c.namespaceURI === NS.w && KNOWN_PPR.has(c.localName))).map(this.raw).join("") : "";
    // The paragraph mark's own formatting (pPr/rPr) is kept with the extras; a tracked change to it is read apart.
    let markRPr = child(pPr, "rPr");
    let change: { kind: "ins" | "del"; id: string; author: string; date: string | null } | null = null;
    const changeEl = markRPr && children(markRPr).find((c) => isW(c, "ins") || isW(c, "del") || isW(c, "moveTo") || isW(c, "moveFrom"));
    if (markRPr && changeEl) {
      change = { kind: isW(changeEl, "ins") || isW(changeEl, "moveTo") ? "ins" : "del", id: attr(changeEl, "id") ?? "", author: attr(changeEl, "author") ?? "", date: attr(changeEl, "date") };
      markRPr = markRPr.cloneNode(true) as Element;
      for (const c of children(markRPr)) if (isW(c, "ins") || isW(c, "del") || isW(c, "moveTo") || isW(c, "moveFrom")) markRPr.removeChild(c);
      if (!children(markRPr).length) markRPr = null;
    }
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
      change,
    };
    const content: Node[] = [];
    this.capture = null;
    this.inlines(p, [], content);
    // A field that goes on into the next paragraph (a table of contents) is kept as it is.
    if (this.capture) this.abandonField(content);
    return schema.nodes.paragraph.create(attrs, mergeText(content));
  }

  /** Reads the runs (and things that hold runs) inside `parent`. */
  private inlines(parent: Element, marks: readonly Mark[], out: Node[]) {
    for (const el of children(parent)) {
      if (this.capture) {
        // Bookmarks may sit inside a field; anything else ends what we can keep together.
        if (isW(el, "bookmarkStart") || isW(el, "bookmarkEnd") || isW(el, "proofErr")) {
          this.capture.xml += this.raw(el);
          continue;
        }
        if (!isW(el, "r")) this.abandonField(out);
      }
      if (el.namespaceURI === MATH_NS && (el.localName === "oMath" || el.localName === "oMathPara")) {
        out.push(this.equation(el, marks));
        continue;
      }
      if (el.namespaceURI !== NS.w) {
        out.push(this.rawInline(el, el.localName, marks));
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
        case "moveTo":
        case "del":
        case "moveFrom": {
          // A tracked change: shown, and kept until it is accepted or rejected.
          const kind = el.localName === "ins" || el.localName === "moveTo" ? schema.marks.insertion : schema.marks.deletion;
          const mark = kind.create({ id: attr(el, "id") ?? "", author: attr(el, "author") ?? "", date: attr(el, "date"), move: el.localName.startsWith("move") });
          this.inlines(el, mark.addToSet(marks), out);
          break;
        }
        case "commentRangeStart":
        case "commentRangeEnd":
          out.push(schema.nodes.comment.create({ id: attr(el, "id") ?? "", kind: el.localName === "commentRangeStart" ? "start" : "end", xml: this.raw(el) }));
          break;
        case "smartTag":
        case "customXml":
        case "dir":
        case "bdo":
          // Wrappers: their runs are ordinary text.
          this.inlines(el, marks, out);
          break;
        case "sdt":
          this.inlines(child(el, "sdtContent") ?? el, marks, out);
          break;
        case "fldSimple": {
          const firstRun = childrenNamed(el, "r")[0];
          const fieldMarks = this.runMarks(child(firstRun, "rPr"), marks);
          const instr = attr(el, "instr") ?? "";
          out.push(schema.nodes.field.create({ instr, result: textOf(el), xml: this.raw(el), original: `${instr}|${textOf(el)}` }, null, fieldMarks));
          break;
        }
        default:
          // Bookmarks, comment anchors and the like: invisible, but kept.
          out.push(schema.nodes.raw_inline.create({ xml: this.raw(el), hidden: true }));
      }
    }
  }

  private equation(el: Element, marks: readonly Mark[]): Node {
    let latex = "";
    try {
      latex = toLatex(fromOmml(el));
    } catch {
      // Shown as its text; saved unchanged unless edited.
      latex = `\\text{${(el.textContent ?? "").replace(/[{}]/g, "")}}`;
    }
    const display = el.localName === "oMathPara";
    return schema.nodes.math.create({ latex, display, xml: this.raw(el), original: latex }, null, marks);
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
      if (el.namespaceURI !== NS.w || isW(el, "rPr") || isW(el, "lastRenderedPageBreak")) continue;
      // A field (begin, instruction, separate, result, end) may span several runs, or sit in one.
      const fieldChar = isW(el, "fldChar") ? attr(el, "fldCharType") : null;
      if (this.capture || fieldChar === "begin") {
        if (!this.capture) {
          flush();
          this.capture = { depth: 0, instr: "", result: "", separated: false, xml: "", marks: null };
        }
        const c = this.capture;
        c.xml += `${startTag(r)}>${rPr ? this.raw(rPr) : ""}${this.raw(el)}</${r.tagName}>`;
        if (fieldChar === "begin") c.depth++;
        else if (fieldChar === "separate" && c.depth === 1) c.separated = true;
        else if (fieldChar === "end") c.depth--;
        else if ((isW(el, "instrText") || isW(el, "delInstrText")) && c.depth === 1 && !c.separated) c.instr += el.textContent ?? "";
        else if (c.separated && (isW(el, "t") || isW(el, "delText"))) {
          c.result += el.textContent ?? "";
          c.marks ??= marks;
        } else if (c.separated && isW(el, "tab")) c.result += "\t";
        if (c.depth === 0) {
          out.push(schema.nodes.field.create({ instr: c.instr, result: c.result, xml: c.xml, original: `${c.instr}|${c.result}` }, null, c.marks ?? marks));
          this.capture = null;
        }
        continue;
      }
      switch (el.localName) {
        case "t":
        case "delText":
          text += el.textContent ?? "";
          break;
        case "commentReference": {
          flush();
          const xml = `${startTag(r)}>${rPr ? this.raw(rPr) : ""}${this.raw(el)}</${r.tagName}>`;
          out.push(schema.nodes.comment.create({ id: attr(el, "id") ?? "", kind: "ref", xml }));
          break;
        }
        case "annotationRef":
          break;
        case "tab":
          text += "\t";
          break;
        case "drawing": {
          const picture = readPicture(el);
          if (picture && this.images.has(picture.rid)) {
            flush();
            const xml = `${startTag(r)}>${rPr ? this.raw(rPr) : ""}${this.raw(el)}</${r.tagName}>`;
            const { rid, ...layout } = picture;
            out.push(schema.nodes.image.create({ ...layout, rid, src: this.images.get(rid), xml, original: JSON.stringify(layout) }, null, marks));
            break;
          }
          flush();
          out.push(this.runPiece(r, el, "Picture", marks));
          break;
        }
        case "footnoteRef":
          // The note's own number; written again when the note is saved.
          break;
        case "footnoteReference": {
          flush();
          const xml = `${startTag(r)}>${rPr ? this.raw(rPr) : ""}${this.raw(el)}</${r.tagName}>`;
          out.push(schema.nodes.footnote.create({ id: attr(el, "id") ?? "", xml }, null, marks));
          break;
        }
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

interface TocFound {
  /** Index of its last element among the body's children. */
  end: number;
  title: string;
  entries: { level: number; text: string; page: number; pos: number }[];
}

/** A table of contents starting at body child `i`: Word's content control, or a TOC field across paragraphs. */
function tableOfContents(siblings: Element[], i: number): TocFound | null {
  const el = siblings[i];
  const entriesOf = (paras: Element[]) =>
    paras
      .map((p) => {
        const style = attr(child(child(p, "pPr"), "pStyle"), "val") ?? "";
        const level = Number(/^TOC([1-9])$/i.exec(style)?.[1] ?? 0);
        const text = tocText(p);
        const tab = text.lastIndexOf("\t");
        return { level, text: (tab >= 0 ? text.slice(0, tab) : text).trim(), page: Number(text.slice(tab + 1).trim()) || 0, pos: 0 };
      })
      .filter((e) => e.level && e.text);
  if (isW(el, "sdt")) {
    const gallery = el.getElementsByTagNameNS(NS.w, "docPartGallery")[0];
    if (attr(gallery, "val") !== "Table of Contents") return null;
    const paras = childrenNamed(child(el, "sdtContent"), "p");
    const heading = paras.find((p) => /^TOCHeading$/i.test(attr(child(child(p, "pPr"), "pStyle"), "val") ?? ""));
    return { end: i, title: heading ? tocText(heading).trim() : "Contents", entries: entriesOf(paras) };
  }
  if (!isW(el, "p")) return null;
  const instr = Array.from(el.getElementsByTagNameNS(NS.w, "instrText")).map((t) => t.textContent ?? "").join("");
  if (!/^\s*TOC\b/.test(instr)) return null;
  // The field ends in this paragraph or a later one.
  for (let j = i; j < siblings.length; j++) {
    if (!isW(siblings[j], "p")) return null;
    const ends = Array.from(siblings[j].getElementsByTagNameNS(NS.w, "fldChar")).some((f) => attr(f, "fldCharType") === "end");
    if (ends) return { end: j, title: "", entries: entriesOf(siblings.slice(i, j + 1)) };
  }
  return null;
}

/** The visible text of a contents entry: field instructions left out, tabs kept. */
function tocText(p: Element): string {
  let out = "";
  let inInstr = 0;
  const walk = (node: Element) => {
    for (const c of children(node)) {
      if (isW(c, "fldChar")) {
        const type = attr(c, "fldCharType");
        if (type === "begin") inInstr++;
        else if (type === "separate" || type === "end") inInstr = Math.max(0, inInstr - (type === "separate" ? 1 : 0));
      } else if (isW(c, "t") && !inInstr) out += c.textContent ?? "";
      else if (isW(c, "tab") && !inInstr) out += "\t";
      else if (!isW(c, "instrText")) walk(c);
    }
  };
  walk(p);
  return out;
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
