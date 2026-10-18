// Reviewing: tracked changes (text inserted or deleted while Word, or
// DuckWriter, tracked changes) and comments. Changes can be accepted or
// rejected one at a time or all at once; comments are anchored in the text by
// a start, an end and a mark, as in Word.

import type { Node } from "prosemirror-model";
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { ensureStyles } from "../docx/builtins";
import { newAnnotationId, type CommentData, type DocContext } from "../docx/model";
import { marksFromRunProps, paraAttrsFromPPr } from "../docx/read";
import { authorClass, schema } from "./schema";

/** One change as a reader sees it: neighbouring text with the same kind of change by the same person. */
export interface Change {
  /** Inserted, deleted, or formatted. */
  kind: "ins" | "del" | "fmt";
  from: number;
  to: number;
  author: string;
  date: string | null;
  /** A change to a paragraph mark (a paragraph added or joined to the next), at the end of the paragraph. */
  paragraph?: boolean;
  /** A change to a paragraph's formatting (from is the paragraph's position). */
  props?: boolean;
  /** A table row inserted or deleted (from and to span the row). */
  row?: boolean;
}

/** A row's tracked change, from its properties. */
function rowChange(trPr: string): { kind: "ins" | "del"; author: string; date: string | null } | null {
  const m = /<w:(ins|del)\b([^>]*)\/?>/.exec(trPr);
  if (!m) return null;
  const attr = (name: string) => new RegExp(`w:${name}="([^"]*)"`).exec(m[2])?.[1] ?? null;
  return { kind: m[1] as "ins" | "del", author: (attr("author") ?? "").replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&amp;/g, "&"), date: attr("date") };
}

/** Every change in the document, in order. */
export function changes(doc: Node): Change[] {
  const out: Change[] = [];
  doc.descendants((block, blockPos) => {
    if (block.type === schema.nodes.table_row) {
      const c = rowChange(block.attrs.trPr);
      if (c) out.push({ kind: c.kind, from: blockPos, to: blockPos + block.nodeSize, author: c.author, date: c.date, row: true });
      return true;
    }
    if (!block.isTextblock) return true;
    const props = block.attrs.propsChange as { author: string; date: string | null } | null;
    if (props) out.push({ kind: "fmt", from: blockPos + 1, to: blockPos + 1, author: props.author, date: props.date, props: true });
    const open: Partial<Record<"ins" | "del" | "fmt", Change>> = {};
    const types = { ins: schema.marks.insertion, del: schema.marks.deletion, fmt: schema.marks.formatChange };
    block.forEach((child, offset) => {
      const from = blockPos + 1 + offset;
      const to = from + child.nodeSize;
      for (const kind of ["ins", "del", "fmt"] as const) {
        const mark = types[kind].isInSet(child.marks);
        const current = open[kind];
        if (!mark) {
          delete open[kind];
          continue;
        }
        if (current && current.to === from && current.author === mark.attrs.author) current.to = to;
        else {
          const change: Change = { kind, from, to, author: mark.attrs.author, date: mark.attrs.date };
          open[kind] = change;
          out.push(change);
        }
      }
    });
    const c = block.attrs.change as { kind: "ins" | "del"; author: string; date: string | null } | null;
    if (c) {
      const end = blockPos + block.nodeSize - 1;
      out.push({ kind: c.kind, from: end, to: end, author: c.author, date: c.date, paragraph: true });
    }
    return false;
  });
  return out.sort((a, b) => a.from - b.from || a.to - b.to);
}

/** The changes the selection touches; with just a cursor, the one it is in. */
export function changesAt(state: EditorState): Change[] {
  const { from, to, empty } = state.selection;
  const all = changes(state.doc);
  if (!empty) return all.filter((c) => (c.paragraph ? c.from >= from && c.from < to : c.from < to && c.to > from));
  const inside = all.filter((c) => (c.paragraph ? c.from === from : c.from < from && c.to > from));
  if (inside.length) return inside;
  // At the very start or end of a change.
  const around = all.filter((c) => !c.paragraph && !c.row && !c.props && (c.from === from || c.to === from)).slice(0, 1);
  if (around.length) return around;
  // In a paragraph whose formatting changed, or a row that was inserted or deleted.
  const $from = state.selection.$from;
  return all.filter((c) => (c.props && $from.depth > 0 && c.from === $from.start()) || (c.row && c.from < from && c.to > from)).slice(0, 1);
}

/** Accepts (true) or rejects (false) the given changes. Rejecting formatting needs the document, for its styles. */
export function decide(state: EditorState, list: Change[], accept: boolean, ctx?: DocContext): Transaction | null {
  if (!list.length) return null;
  const tr = state.tr;
  // From the end, so that what is still to do keeps its place.
  for (const c of [...list].sort((a, b) => b.from - a.from || b.to - a.to)) {
    const from = tr.mapping.map(c.from);
    const to = tr.mapping.map(c.to);
    if (c.props) {
      const $para = tr.doc.resolve(from);
      const para = $para.parent;
      const pos = $para.before();
      if (!para.attrs.propsChange) continue;
      const before = !accept && ctx ? paraAttrsFromPPr(para.attrs.propsChange.before, ctx) : {};
      tr.setNodeMarkup(pos, null, { ...para.attrs, ...before, change: para.attrs.change, propsChange: null });
      continue;
    }
    if (c.row) {
      const row = tr.doc.nodeAt(from);
      if (!row || row.type !== schema.nodes.table_row) continue;
      if ((c.kind === "ins") === accept) tr.setNodeMarkup(from, null, { ...row.attrs, trPr: row.attrs.trPr.replace(/<w:(ins|del)\b[^>]*\/>/g, "") });
      else {
        const $row = tr.doc.resolve(from);
        // The table's last row goes with the table.
        if ($row.parent.childCount === 1) tr.delete($row.before(), $row.after());
        else tr.delete(from, from + row.nodeSize);
      }
      continue;
    }
    if (c.kind === "fmt") {
      if (!accept && ctx) {
        // Back to the formatting before, piece by piece.
        const pieces: { a: number; b: number; before: string }[] = [];
        tr.doc.nodesBetween(from, to, (node, pos) => {
          const change = schema.marks.formatChange.isInSet(node.marks);
          if (node.isInline && change) pieces.push({ a: Math.max(pos, from), b: Math.min(pos + node.nodeSize, to), before: change.attrs.before });
          return !node.isInline;
        });
        for (const p of pieces) {
          for (const name of FORMATTING_MARKS) tr.removeMark(p.a, p.b, schema.marks[name]);
          for (const mark of marksFromRunProps(p.before, ctx)) if ((FORMATTING_MARKS as readonly string[]).includes(mark.type.name)) tr.addMark(p.a, p.b, mark);
        }
      }
      tr.removeMark(from, to, schema.marks.formatChange);
      continue;
    }
    const keep = (c.kind === "ins") === accept;
    if (c.paragraph) {
      const $end = tr.doc.resolve(from);
      const para = $end.parent;
      if (!para.attrs.change) continue;
      const paraPos = $end.before();
      tr.setNodeMarkup(paraPos, null, { ...para.attrs, change: null });
      // A paragraph mark that goes joins the paragraph to the next one.
      const after = paraPos + para.nodeSize;
      const next = tr.doc.resolve(after).nodeAfter;
      if (!keep && next?.type === schema.nodes.paragraph) tr.join(after);
      continue;
    }
    if (keep) tr.removeMark(from, to, c.kind === "ins" ? schema.marks.insertion : schema.marks.deletion);
    else tr.delete(from, to);
  }
  return tr.setMeta("noTrack", true);
}

const FORMATTING_MARKS = ["bold", "italic", "underline", "strike", "font", "size", "color", "highlight", "vertAlign", "charStyle", "runExtra"] as const;

/** Accepts or rejects every change, and drops Word's records of paragraph formatting changes too. */
export function decideAll(state: EditorState, accept: boolean, ctx?: DocContext): Transaction | null {
  const tr = decide(state, changes(state.doc), accept, ctx) ?? state.tr;
  // Formatting changes are accepted as they are (DuckWriter can't undo them).
  tr.doc.descendants((node, pos) => {
    if (node.isTextblock && /<w:pPrChange\b/.test(node.attrs.extra)) {
      tr.setNodeMarkup(pos, null, { ...node.attrs, extra: stripChange(node.attrs.extra, "pPrChange") });
    }
    return true;
  });
  const withRunChanges: { from: number; to: number; xml: string; mark: import("prosemirror-model").Mark }[] = [];
  tr.doc.descendants((node, pos) => {
    const extra = node.marks.find((m) => m.type === schema.marks.runExtra && /<w:rPrChange\b/.test(m.attrs.xml));
    if (extra) withRunChanges.push({ from: pos, to: pos + node.nodeSize, xml: stripChange(extra.attrs.xml, "rPrChange"), mark: extra });
  });
  for (const r of withRunChanges) {
    tr.removeMark(r.from, r.to, r.mark);
    if (r.xml) tr.addMark(r.from, r.to, schema.marks.runExtra.create({ xml: r.xml }));
  }
  // Word's markers of where moved text came from and went to.
  const markers: { from: number; to: number }[] = [];
  tr.doc.descendants((node, pos) => {
    if (node.type === schema.nodes.raw_inline && /^<w:move(?:From|To)Range(?:Start|End)\b/.test(node.attrs.xml)) markers.push({ from: pos, to: pos + node.nodeSize });
  });
  for (const m of markers.reverse()) tr.delete(m.from, m.to);
  return tr.docChanged ? tr.setMeta("noTrack", true) : null;
}

function stripChange(xml: string, name: string): string {
  return xml.replace(new RegExp(`<w:${name}\\b[^>]*?(?:/>|>[\\s\\S]*?</w:${name}>)`, "g"), "");
}

/** The next (1) or previous (-1) change after the selection, wrapping round. */
export function nextChange(state: EditorState, direction: 1 | -1): Change | null {
  const all = changes(state.doc);
  if (!all.length) return null;
  const { from, to } = state.selection;
  if (direction > 0) return all.find((c) => c.from >= to && !(c.from === from && c.to === to)) ?? all[0];
  return [...all].reverse().find((c) => c.to <= from && !(c.from === from && c.to === to)) ?? all[all.length - 1];
}

/** Selects a change (a paragraph mark change: puts the cursor there). */
export function selectChange(state: EditorState, change: Change): Transaction {
  return state.tr.setSelection(TextSelection.create(state.doc, change.from, change.to)).scrollIntoView();
}

// ---- comments ---------------------------------------------------------------------------

export interface CommentAnchors {
  id: string;
  /** Where the commented text starts and ends (the anchors' positions). */
  from: number;
  to: number;
  /** Position of the comment's mark. */
  ref: number;
}

/** Where each comment is in the text, in the order they start. */
export function commentAnchors(doc: Node): CommentAnchors[] {
  const found = new Map<string, Partial<CommentAnchors>>();
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.comment) return true;
    const a: Partial<CommentAnchors> = found.get(node.attrs.id) ?? { id: node.attrs.id };
    if (node.attrs.kind === "start") a.from ??= pos;
    else if (node.attrs.kind === "end") a.to = pos;
    else a.ref = pos;
    found.set(node.attrs.id, a);
    return false;
  });
  const out: CommentAnchors[] = [];
  for (const a of found.values()) {
    const ref = a.ref ?? a.to ?? a.from;
    if (ref === undefined) continue;
    out.push({ id: a.id!, from: a.from ?? a.to ?? ref, to: a.to ?? ref, ref });
  }
  return out.sort((a, b) => a.from - b.from || a.ref - b.ref);
}

/** The word around an empty cursor, which a new comment is about. */
export function wordAround(state: EditorState): { from: number; to: number } {
  const { from, to, empty, $from } = state.selection;
  if (!empty || !$from.parent.isTextblock) return { from, to };
  const text = $from.parent.textBetween(0, $from.parent.content.size, undefined, "￼");
  const at = $from.parentOffset;
  let s = at;
  let e = at;
  while (s > 0 && /[\p{L}\p{N}_'’-]/u.test(text[s - 1])) s--;
  while (e < text.length && /[\p{L}\p{N}_'’-]/u.test(text[e])) e++;
  const start = $from.start();
  return { from: start + s, to: start + e };
}

const nowIso = () => new Date().toISOString().replace(/\.\d+Z$/, "Z");

export function initialsOf(name: string): string {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .map((w) => w[0].toUpperCase())
    .join("")
    .slice(0, 3);
}

function newComment(ctx: DocContext, author: string, parent: string | null): CommentData {
  ensureStyles(ctx, "CommentText", "CommentReference");
  const item: CommentData = { id: newAnnotationId(ctx), author, initials: initialsOf(author), date: nowIso(), text: "", xml: null, original: null, paraId: null, parent, done: false };
  ctx.comments.items.set(item.id, item);
  ctx.comments.version++;
  return item;
}

const anchor = (id: string, kind: "start" | "end" | "ref") => schema.nodes.comment.create({ id, kind });

/** Adds a comment on the selection (or the word at the cursor). */
export function addComment(state: EditorState, ctx: DocContext, author: string): { tr: Transaction; id: string } {
  const { from, to } = wordAround(state);
  const item = newComment(ctx, author, null);
  const tr = state.tr;
  tr.insert(to, [anchor(item.id, "end"), anchor(item.id, "ref")]);
  tr.insert(from, anchor(item.id, "start"));
  return { tr: tr.setMeta("noTrack", true), id: item.id };
}

/** Adds a reply to a comment, anchored where the comment is (as Word does). */
export function addReply(state: EditorState, ctx: DocContext, parentId: string, author: string): { tr: Transaction; id: string } | null {
  const parent = commentAnchors(state.doc).find((a) => a.id === parentId);
  if (!parent) return null;
  const item = newComment(ctx, author, parentId);
  const tr = state.tr;
  tr.insert(parent.ref + 1, [anchor(item.id, "end"), anchor(item.id, "ref")]);
  tr.insert(parent.from + 1, anchor(item.id, "start"));
  return { tr: tr.setMeta("noTrack", true), id: item.id };
}

/** Removes a comment and its replies from the text. */
export function removeComment(state: EditorState, ctx: DocContext, id: string): Transaction {
  const ids = new Set([id, ...[...ctx.comments.items.values()].filter((c) => c.parent === id).map((c) => c.id)]);
  const tr = state.tr;
  const found: number[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type === schema.nodes.comment && ids.has(node.attrs.id)) found.push(pos);
    return node.isInline ? false : true;
  });
  for (const pos of found.reverse()) tr.delete(pos, pos + 1);
  return tr.setMeta("noTrack", true);
}

/** The comment whose text the cursor is in (the innermost one). */
export function commentAt(state: EditorState, ctx: DocContext): string | null {
  const pos = state.selection.from;
  const around = commentAnchors(state.doc).filter((a) => a.from <= pos && a.to >= pos && !ctx.comments.items.get(a.id)?.parent);
  around.sort((a, b) => a.to - a.from - (b.to - b.from));
  return around[0]?.id ?? null;
}

// ---- showing them -----------------------------------------------------------------------

interface ReviewState {
  /** The comment being looked at, shown more strongly. */
  active: string | null;
  decorations: DecorationSet;
}

export const reviewKey = new PluginKey<ReviewState>("review");

function decorate(doc: Node, active: string | null, ctx: DocContext | null): DecorationSet {
  const decos: Decoration[] = [];
  for (const a of commentAnchors(doc)) {
    const item = ctx?.comments.items.get(a.id);
    if (item?.parent) continue;
    const cls = `comment-range${a.id === active ? " active" : ""}${item?.done ? " done" : ""}`;
    if (a.to > a.from) decos.push(Decoration.inline(a.from, a.to, { class: cls, "data-comment": a.id }));
  }
  doc.descendants((node, pos) => {
    if (node.type === schema.nodes.table_row) {
      const rc = rowChange(node.attrs.trPr);
      if (rc) decos.push(Decoration.node(pos, pos + node.nodeSize, { class: `tc-row tc-row-${rc.kind} ${authorClass(rc.author)}`, title: `Row ${rc.kind === "ins" ? "inserted" : "deleted"} by ${rc.author || "someone"}` }));
      return true;
    }
    if (!node.isTextblock) return true;
    const pc = node.attrs.propsChange as { author: string } | null;
    if (pc) decos.push(Decoration.node(pos, pos + node.nodeSize, { class: `tc-props ${authorClass(pc.author)}`, title: `Formatting changed by ${pc.author || "someone"}` }));
    const c = node.attrs.change as { kind: string; author: string } | null;
    if (c) {
      const end = pos + node.nodeSize - 1;
      decos.push(
        Decoration.widget(end, () => {
          const el = document.createElement("span");
          el.className = `tc-para tc-${c.kind} ${authorClass(c.author)}`;
          el.textContent = "¶";
          el.title = `${c.kind === "ins" ? "New paragraph" : "Paragraphs joined"} by ${c.author || "someone"}`;
          return el;
        }, { side: 1, key: `para-${c.kind}-${c.author}` }),
      );
    }
    return false;
  });
  return DecorationSet.create(doc, decos);
}

export function reviewPlugin(ctx: () => DocContext | null) {
  return new Plugin<ReviewState>({
    key: reviewKey,
    state: {
      init: (_config, state) => ({ active: null, decorations: decorate(state.doc, null, ctx()) }),
      apply(tr, old, _oldState, state) {
        const meta = tr.getMeta(reviewKey) as { active?: string | null; refresh?: boolean } | undefined;
        let active = meta && "active" in meta ? (meta.active ?? null) : old.active;
        // Moving the cursor into a comment's text makes it the one looked at.
        if (!meta && tr.selectionSet) {
          const c = ctx();
          if (c) active = commentAt(state, c);
        }
        if (!tr.docChanged && active === old.active && !meta?.refresh) return old;
        return { active, decorations: decorate(state.doc, active, ctx()) };
      },
    },
    props: { decorations: (state) => reviewKey.getState(state)?.decorations },
  });
}

export const activeComment = (state: EditorState) => reviewKey.getState(state)?.active ?? null;
