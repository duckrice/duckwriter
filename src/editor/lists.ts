// Bulleted and numbered lists, the Word way: a list item is an ordinary
// paragraph that points at a list definition (numId) and a level. The numbers
// and bullets are worked out here and drawn in front of each item.

import type { Node } from "prosemirror-model";
import { Plugin, PluginKey, type EditorState, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { bulletText, formatNumber, resolvedPara, type DocContext, type ListLevel } from "../docx/model";
import { tag } from "../docx/xml";

export type ListKind = "bullet" | "number";

/** The list a paragraph belongs to, directly or through its style; null when it isn't in one. */
export function listOf(ctx: DocContext, para: Node): { numId: string; ilvl: number } | null {
  let numId: string | null = para.attrs.numId;
  let ilvl: number = para.attrs.ilvl ?? 0;
  if (numId === null) {
    const fromStyle = resolvedPara(ctx.styles, para.attrs.style);
    if (fromStyle.numId === undefined) return null;
    numId = fromStyle.numId;
    ilvl = fromStyle.ilvl ?? 0;
  }
  // numId 0 switches a list off.
  if (numId === "0" || !ctx.numbering.nums.has(numId)) return null;
  return { numId, ilvl };
}

export function levelOf(ctx: DocContext, numId: string, ilvl: number): ListLevel | undefined {
  const num = ctx.numbering.nums.get(numId);
  return num ? ctx.numbering.abstracts.get(num.abstractId)?.[ilvl] : undefined;
}

export function kindOf(ctx: DocContext, numId: string, ilvl = 0): ListKind | null {
  const level = levelOf(ctx, numId, ilvl);
  if (!level) return null;
  return level.format === "bullet" ? "bullet" : "number";
}

/** The label in front of every list item, by position of the paragraph. */
export function listLabels(ctx: DocContext, doc: Node): Map<number, { label: string; level: ListLevel }> {
  const out = new Map<number, { label: string; level: ListLevel }>();
  // Lists sharing a definition share their count, as in Word.
  const counts = new Map<string, number[]>();
  const seenNums = new Set<string>();
  doc.forEach((para, pos) => {
    if (para.type.name !== "paragraph") return;
    const list = listOf(ctx, para);
    if (!list) return;
    const num = ctx.numbering.nums.get(list.numId)!;
    const levels = ctx.numbering.abstracts.get(num.abstractId);
    const level = levels?.[list.ilvl];
    if (!levels || !level) return;
    let count = counts.get(num.abstractId);
    if (!count) counts.set(num.abstractId, (count = []));
    if (!seenNums.has(list.numId)) {
      seenNums.add(list.numId);
      for (const [lvl, start] of num.startOverrides) {
        count[lvl] = start - 1;
        count.length = lvl + 1;
      }
    }
    count[list.ilvl] = (count[list.ilvl] ?? (level.start ?? 1) - 1) + 1;
    count.length = list.ilvl + 1;
    let label: string;
    if (level.format === "bullet") label = bulletText(level.text);
    else {
      label = level.text.replace(/%(\d)/g, (_m, n: string) => {
        const i = Number(n) - 1;
        const l = levels[i];
        return l ? formatNumber(count![i] ?? l.start, i === list.ilvl ? level.format : l.format) : "";
      });
    }
    out.set(pos, { label, level });
  });
  return out;
}

export const listPluginKey = new PluginKey<DecorationSet>("lists");

/** Draws list numbers and bullets, and gives list items their level's indent. */
export function listPlugin(getContext: () => DocContext) {
  const build = (doc: Node) => {
    const ctx = getContext();
    const decorations: Decoration[] = [];
    for (const [pos, { label, level }] of listLabels(ctx, doc)) {
      const para = doc.nodeAt(pos)!;
      const css: string[] = [];
      if (para.attrs.indentLeft === null) css.push(`padding-left:${level.indentLeft}pt`);
      if (para.attrs.indentFirst === null) css.push(`text-indent:${-level.hanging}pt`);
      decorations.push(Decoration.node(pos, pos + para.nodeSize, { class: "list-item", style: css.join(";") }));
      decorations.push(
        Decoration.widget(
          pos + 1,
          () => {
            const el = document.createElement("span");
            el.className = level.format === "bullet" ? "list-label bullet" : "list-label";
            el.textContent = label;
            el.style.minWidth = `${Math.max(level.hanging, 0)}pt`;
            if (level.font && level.format === "bullet" && !/symbol|wingdings/i.test(level.font)) el.style.fontFamily = `"${level.font}"`;
            return el;
          },
          { side: -1, key: `${label}|${level.hanging}|${level.format}`, ignoreSelection: true },
        ),
      );
    }
    return DecorationSet.create(doc, decorations);
  };
  return new Plugin<DecorationSet>({
    key: listPluginKey,
    state: {
      init: (_config, state) => build(state.doc),
      apply: (tr, old) => (tr.docChanged || tr.getMeta(listPluginKey) ? build(tr.doc) : old),
    },
    props: { decorations: (state) => listPluginKey.getState(state) },
  });
}

// ---- creating lists ------------------------------------------------------------

const BULLETS = ["•", "◦", "▪"];
const NUMBER_FORMATS: [string, string][] = [
  ["decimal", "%1."],
  ["lowerLetter", "%2."],
  ["lowerRoman", "%3."],
];

/** Adds a new list definition and instance to the document, returning its numId. */
export function createList(ctx: DocContext, kind: ListKind): string {
  const { numbering } = ctx;
  const nextId = (ids: Iterable<string>) => String(Math.max(0, ...[...ids].map(Number).filter(Number.isFinite)) + 1);
  const abstractId = nextId(numbering.abstracts.keys());
  const numId = nextId(numbering.nums.keys());
  const levels: ListLevel[] = [];
  let xml = "";
  for (let i = 0; i < 9; i++) {
    const [format, pattern] = kind === "bullet" ? ["bullet", BULLETS[i % 3]] : NUMBER_FORMATS[i % 3];
    const text = kind === "bullet" ? pattern : pattern.replace(/\d/, String(i + 1));
    const level: ListLevel = { format, text, start: 1, indentLeft: 36 * (i + 1), hanging: 18 };
    levels.push(level);
    xml += tag(
      "w:lvl",
      { "w:ilvl": i },
      tag("w:start", { "w:val": 1 }),
      tag("w:numFmt", { "w:val": format }),
      tag("w:lvlText", { "w:val": text }),
      tag("w:lvlJc", { "w:val": "left" }),
      tag("w:pPr", {}, tag("w:ind", { "w:left": 720 * (i + 1), "w:hanging": 360 })),
    );
  }
  numbering.abstracts.set(abstractId, levels);
  numbering.nums.set(numId, { abstractId, startOverrides: new Map() });
  numbering.addedAbstracts.push(tag("w:abstractNum", { "w:abstractNumId": abstractId }, tag("w:multiLevelType", { "w:val": "hybridMultilevel" }), xml));
  numbering.addedNums.push(tag("w:num", { "w:numId": numId }, tag("w:abstractNumId", { "w:val": abstractId })));
  return numId;
}

/** Paragraphs the selection touches, with their positions. */
export function selectedParagraphs(state: EditorState): { node: Node; pos: number }[] {
  const out: { node: Node; pos: number }[] = [];
  const { from, to } = state.selection;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name === "paragraph") out.push({ node, pos });
    return node.type.name === "doc";
  });
  return out;
}

/**
 * Turns a list on or off for the selected paragraphs. A list right before the
 * selection of the same kind is continued; otherwise a new list starts.
 */
export function toggleList(ctx: DocContext, kind: ListKind) {
  return (state: EditorState, dispatch?: (tr: Transaction) => void): boolean => {
    const paras = selectedParagraphs(state);
    if (!paras.length) return false;
    const allInKind = paras.every(({ node }) => {
      const list = listOf(ctx, node);
      return list && kindOf(ctx, list.numId, list.ilvl) === kind;
    });
    if (!dispatch) return true;
    const tr = state.tr;
    // Word gives list items the List Paragraph style, which has no gaps between items.
    const listStyle = ctx.styles.byId.has("ListParagraph") ? "ListParagraph" : null;
    if (allInKind) {
      for (const { node, pos } of paras) {
        // Off: "0" also switches off a list the paragraph's style would give it.
        const fromStyle = resolvedPara(ctx.styles, node.attrs.style).numId;
        const style = node.attrs.style === listStyle ? null : node.attrs.style;
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, style, numId: fromStyle ? "0" : null, ilvl: 0 });
      }
    } else {
      const before = state.doc.resolve(paras[0].pos).nodeBefore;
      const previous = before?.type.name === "paragraph" ? listOf(ctx, before) : null;
      const numId = previous && kindOf(ctx, previous.numId) === kind ? previous.numId : createList(ctx, kind);
      for (const { node, pos } of paras) {
        const current = listOf(ctx, node);
        const style = node.attrs.style === null && listStyle ? listStyle : node.attrs.style;
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, style, numId, ilvl: current?.ilvl ?? 0 });
      }
    }
    dispatch(tr.setMeta(listPluginKey, true).scrollIntoView());
    return true;
  };
}

/** Moves the selected list items a level in (+1) or out (-1). Returns false outside lists. */
export function changeListLevel(ctx: DocContext, delta: 1 | -1) {
  return (state: EditorState, dispatch?: (tr: Transaction) => void): boolean => {
    const paras = selectedParagraphs(state).filter(({ node }) => listOf(ctx, node));
    if (!paras.length) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    for (const { node, pos } of paras) {
      const list = listOf(ctx, node)!;
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, numId: list.numId, ilvl: Math.min(8, Math.max(0, list.ilvl + delta)) });
    }
    dispatch(tr.setMeta(listPluginKey, true));
    return true;
  };
}
