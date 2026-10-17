// The tools for reports and papers: captions that number themselves,
// cross-references to them ("see Figure 2"), numbered headings (1, 1.1,
// 1.1.1) and a table of contents that keeps itself up to date.

import type { Node } from "prosemirror-model";
import { TextSelection, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { h } from "../dom";
import { ensureStyles } from "../docx/builtins";
import { resolvedPara, type DocContext, type ListLevel } from "../docx/model";
import { tag } from "../docx/xml";
import { bookmarkOf, fieldWords } from "../editor/fields";
import { listLabels, listPluginKey } from "../editor/lists";
import type { PageLayout } from "../editor/pagination";
import { fieldKind, schema } from "../editor/schema";

export const CAPTION_LABELS = ["Figure", "Table", "Equation"];

/** The heading level of a paragraph (1 for Heading 1 …), or 0 for ordinary text. */
export function headingLevel(ctx: DocContext, para: Node): number {
  if (para.type !== schema.nodes.paragraph || !para.attrs.style) return 0;
  const outline = resolvedPara(ctx.styles, para.attrs.style).outline;
  if (outline !== undefined && outline < 9) return outline + 1;
  const m = /^Heading([1-9])$/.exec(para.attrs.style);
  return m ? Number(m[1]) : 0;
}

// ---- bookmarks ----------------------------------------------------------------------

function bookmarkIds(doc: Node): number[] {
  const ids: number[] = [];
  doc.descendants((n) => {
    const mark = bookmarkOf(n);
    if (mark) ids.push(Number(mark.id));
  });
  return ids.filter(Number.isFinite);
}

function bookmarkNodes(doc: Node): [Node, Node] {
  const id = String(Math.max(0, ...bookmarkIds(doc)) + 1);
  const name = `_Ref${String(Math.floor(Math.random() * 1e9)).padStart(9, "0")}`;
  return [
    schema.nodes.raw_inline.create({ xml: `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`, hidden: true }),
    schema.nodes.raw_inline.create({ xml: `<w:bookmarkEnd w:id="${id}"/>`, hidden: true }),
  ];
}

/** The name of a bookmark that starts at the beginning of a paragraph, if any. */
function bookmarkAtStart(para: Node): string | null {
  let name: string | null = null;
  para.forEach((child, _offset, index) => {
    if (name !== null || index > 2) return;
    const mark = bookmarkOf(child);
    if (mark?.kind === "start" && mark.name?.startsWith("_Ref")) name = mark.name;
  });
  return name;
}

// ---- captions -----------------------------------------------------------------------

/** Adds a caption paragraph below (or above) the paragraph with the cursor: "Figure 1: text". */
export function insertCaption(view: EditorView, ctx: DocContext, label: string, text: string, above: boolean) {
  ensureStyles(ctx, "Caption");
  const { $from } = view.state.selection;
  const depth = $from.depth > 0 ? 1 : 0;
  const at = depth ? (above ? $from.before(1) : $from.after(1)) : $from.pos;
  const [start, end] = bookmarkNodes(view.state.doc);
  const seq = schema.nodes.field.create({ instr: ` SEQ ${label} \\* ARABIC `, result: "1" });
  const content = [start, schema.text(`${label} `), seq, end, ...(text ? [schema.text(`: ${text}`)] : [])];
  const para = schema.nodes.paragraph.create({ style: "Caption" }, content);
  const tr = view.state.tr.insert(at, para);
  tr.setSelection(TextSelection.create(tr.doc, at + para.nodeSize - 1));
  view.dispatch(tr.scrollIntoView());
  view.focus();
}

/** Asks for a caption's label and text, then adds it. */
export function captionDialog(view: EditorView, ctx: DocContext) {
  const label = h("select", { class: "tb-select" }, ...CAPTION_LABELS.map((l) => h("option", { value: l }, l)));
  const text = h("input", { class: "tb-input dialog-input", placeholder: "What it shows (optional)" });
  const where = h("select", { class: "tb-select" }, h("option", { value: "below" }, "Below the selected item"), h("option", { value: "above" }, "Above the selected item"));
  label.addEventListener("change", () => {
    // Tables are captioned above, figures below, by convention.
    where.value = label.value === "Table" ? "above" : "below";
  });
  const dialog = h(
    "dialog",
    { class: "dialog" },
    h("h2", {}, "Insert caption"),
    h("div", { class: "dialog-form" }, h("label", {}, "Label", label), h("label", {}, "Caption", text), h("label", {}, "Position", where)),
    h(
      "div",
      { class: "dialog-footer" },
      h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Cancel"),
      h(
        "button",
        {
          type: "button",
          class: "btn primary",
          onclick: () => {
            dialog.close();
            insertCaption(view, ctx, label.value, text.value.trim(), where.value === "above");
          },
        },
        "Insert",
      ),
    ),
  );
  text.addEventListener("keydown", (event) => {
    if (event.key === "Enter") (dialog.querySelector(".btn.primary") as HTMLButtonElement).click();
  });
  dialog.addEventListener("close", () => dialog.remove());
  document.body.append(dialog);
  dialog.showModal();
  text.focus();
}

// ---- cross-references ----------------------------------------------------------------

export interface ReferenceTarget {
  kind: "caption" | "heading";
  /** What the menu shows: "Figure 1: The setup" or a heading's text. */
  label: string;
  pos: number;
}

/** Captions and headings that can be referred to. */
export function referenceTargets(ctx: DocContext, doc: Node): ReferenceTarget[] {
  const out: ReferenceTarget[] = [];
  const labels = listLabels(ctx, doc);
  doc.forEach((para, pos) => {
    if (para.type !== schema.nodes.paragraph) return;
    let hasSeq = false;
    para.descendants((n) => {
      if (n.type === schema.nodes.field && fieldKind(n.attrs.instr) === "SEQ") hasSeq = true;
    });
    const text = captionText(para);
    if (hasSeq) out.push({ kind: "caption", label: text, pos });
    else if (headingLevel(ctx, para) && text.trim()) out.push({ kind: "heading", label: `${labels.get(pos)?.label ?? ""} ${text}`.trim(), pos });
  });
  return out;
}

function captionText(para: Node): string {
  let text = "";
  para.forEach((child) => {
    if (child.isText) text += child.text;
    else if (child.type === schema.nodes.field) text += child.attrs.result;
  });
  return text;
}

/**
 * Puts a reference to a caption ("Figure 2") or a heading at the cursor.
 * The target gets a bookmark if it doesn't have one yet.
 */
export function insertReference(view: EditorView, target: ReferenceTarget) {
  const tr = view.state.tr;
  const name = bookmarkFor(tr, target);
  const ref = schema.nodes.field.create({ instr: ` REF ${name} \\h `, result: "" });
  tr.replaceSelectionWith(ref, true);
  view.dispatch(tr.scrollIntoView());
  view.focus();
}

/** The name of the target's bookmark, adding one (to `tr`) if it has none yet. */
export function bookmarkFor(tr: Transaction, target: ReferenceTarget): string {
  const para = tr.doc.nodeAt(target.pos)!;
  const known = bookmarkAtStart(para);
  if (known) return known;
  const [start, end] = bookmarkNodes(tr.doc);
  // A caption's bookmark covers its label and number; a heading's covers all of it.
  let endAt = target.pos + para.nodeSize - 1;
  if (target.kind === "caption") {
    para.forEach((child, offset) => {
      if (child.type === schema.nodes.field && fieldKind(child.attrs.instr) === "SEQ" && endAt === target.pos + para.nodeSize - 1) endAt = target.pos + 1 + offset + child.nodeSize;
    });
  }
  tr.insert(endAt, end).insert(target.pos + 1, start);
  return bookmarkOf(start)!.name!;
}

/** Where a bookmark starts, or null. */
export function bookmarkPosition(doc: Node, name: string): number | null {
  let found: number | null = null;
  doc.descendants((node, pos) => {
    if (found !== null) return false;
    if (bookmarkOf(node)?.name === name) found = pos;
    return true;
  });
  return found;
}

// ---- numbered headings --------------------------------------------------------------

const HEADINGS = ["Heading1", "Heading2", "Heading3"];

/** Whether the headings are numbered (1, 1.1, 1.1.1). */
export function headingsNumbered(ctx: DocContext): boolean {
  const numId = resolvedPara(ctx.styles, "Heading1").numId;
  return !!numId && numId !== "0";
}

/** Numbers the headings 1, 1.1, 1.1.1 (or stops numbering them). */
export function toggleHeadingNumbers(view: EditorView, ctx: DocContext) {
  const on = !headingsNumbered(ctx);
  const present = HEADINGS.filter((id) => ctx.styles.byId.has(id));
  if (on) {
    const numId = createHeadingList(ctx, present);
    present.forEach((id, i) => {
      const style = ctx.styles.byId.get(id)!;
      style.para = { ...style.para, numId, ilvl: i };
      ctx.styleNumbering.set(id, { numId, ilvl: i });
    });
  } else {
    for (const id of present) {
      const style = ctx.styles.byId.get(id)!;
      const { numId: _numId, ilvl: _ilvl, ...rest } = style.para;
      style.para = rest;
      ctx.styleNumbering.set(id, null);
    }
  }
  // Redraw the numbers.
  view.dispatch(view.state.tr.setMeta(listPluginKey, true));
}

function createHeadingList(ctx: DocContext, styles: string[]): string {
  const { numbering } = ctx;
  const nextId = (ids: Iterable<string>) => String(Math.max(0, ...[...ids].map(Number).filter(Number.isFinite)) + 1);
  const abstractId = nextId(numbering.abstracts.keys());
  const numId = nextId(numbering.nums.keys());
  const levels: ListLevel[] = [];
  let xml = "";
  for (let i = 0; i < 9; i++) {
    const text = Array.from({ length: i + 1 }, (_, n) => `%${n + 1}`).join(".");
    const indent = 432 + 144 * i;
    levels.push({ format: "decimal", text, start: 1, indentLeft: indent / 20, hanging: indent / 20 });
    xml += tag(
      "w:lvl",
      { "w:ilvl": i },
      tag("w:start", { "w:val": 1 }),
      tag("w:numFmt", { "w:val": "decimal" }),
      ...(styles[i] ? [tag("w:pStyle", { "w:val": styles[i] })] : []),
      tag("w:lvlText", { "w:val": text }),
      tag("w:lvlJc", { "w:val": "left" }),
      tag("w:pPr", {}, tag("w:ind", { "w:left": indent, "w:hanging": indent })),
    );
  }
  numbering.abstracts.set(abstractId, levels);
  numbering.nums.set(numId, { abstractId, startOverrides: new Map() });
  numbering.addedAbstracts.push(tag("w:abstractNum", { "w:abstractNumId": abstractId }, tag("w:multiLevelType", { "w:val": "multilevel" }), xml));
  numbering.addedNums.push(tag("w:num", { "w:numId": numId }, tag("w:abstractNumId", { "w:val": abstractId })));
  return numId;
}

// ---- table of contents ----------------------------------------------------------------

export function insertToc(view: EditorView, ctx: DocContext) {
  ensureStyles(ctx, "TOCHeading", "TOC1", "TOC2", "TOC3");
  const { $from } = view.state.selection;
  const at = $from.depth > 0 ? $from.before(1) : $from.pos;
  const tr = view.state.tr.insert(at, schema.nodes.toc.create());
  view.dispatch(tr.scrollIntoView());
  view.focus();
}

/** Brings tables of contents up to date with the headings and the pages they are on. */
export function updateTocs(view: EditorView, ctx: DocContext, layout: PageLayout) {
  const { doc } = view.state;
  let hasToc = false;
  doc.forEach((n) => {
    if (n.type === schema.nodes.toc) hasToc = true;
  });
  if (!hasToc) return;
  const labels = listLabels(ctx, doc);
  const entries: { level: number; text: string; page: number; pos: number }[] = [];
  doc.forEach((para, pos) => {
    const level = headingLevel(ctx, para);
    if (!level || level > 3 || !para.textContent.trim()) return;
    entries.push({ level, text: `${labels.get(pos)?.label ?? ""} ${para.textContent}`.trim(), page: (layout.pageOfBlock.get(pos) ?? 0) + 1, pos });
  });
  const tr = view.state.tr;
  doc.forEach((node, pos) => {
    if (node.type === schema.nodes.toc && JSON.stringify(node.attrs.entries) !== JSON.stringify(entries)) {
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, entries });
    }
  });
  if (tr.docChanged) view.dispatch(tr.setMeta("addToHistory", false));
}

/** What a SEQ field counts (its label), for tests. */
export const seqLabel = (instr: string) => fieldWords(instr)[1];
