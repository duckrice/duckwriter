// Working with the document's outline: a heading and everything under it (its
// section) moves as one, up or down past the section beside it, or dragged in
// the headings pane; and a section's headings go a level up or down.

import type { Node } from "prosemirror-model";
import { TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import { ensureStyles } from "../docx/builtins";
import type { DocContext } from "../docx/model";
import { headingLevel } from "../ui/references";
import { schema } from "./schema";

const levelOf = (ctx: DocContext, node: Node) => (node.type === schema.nodes.paragraph ? headingLevel(ctx, node) : 0);

/** Where the section of the heading at `pos` ends: at the next heading as high or higher, or the end. */
export function sectionEnd(doc: Node, ctx: DocContext, pos: number): number {
  const level = levelOf(ctx, doc.nodeAt(pos)!);
  let end = -1;
  doc.forEach((node, at) => {
    if (end >= 0 || at <= pos) return;
    const l = levelOf(ctx, node);
    // The list of endnotes always stays last.
    if ((l && l <= level) || node.type === schema.nodes.endnotes) end = at;
  });
  return end >= 0 ? end : doc.content.size;
}

/** The top-level heading the cursor is in or under, if any. */
export function headingAround(state: EditorState, ctx: DocContext): number | null {
  const at = state.selection.$from.before(1);
  let found: number | null = null;
  state.doc.forEach((node, pos) => {
    if (pos <= at && levelOf(ctx, node)) found = pos;
  });
  return found;
}

/** As a command: the section of the heading at `from` to before `to` (for dragging in the headings pane). */
export function moveSectionCommand(ctx: DocContext, from: number, to: number): Command {
  return (state, dispatch) => {
    const tr = moveSectionTo(state, ctx, from, to);
    if (!tr) return false;
    dispatch?.(tr);
    return true;
  };
}

/** Moves the section of the heading at `from` to `to` (a top-level position outside it). */
export function moveSectionTo(state: EditorState, ctx: DocContext, from: number, to: number): Transaction | null {
  const end = sectionEnd(state.doc, ctx, from);
  if (to >= from && to <= end) return null;
  const offset = state.selection.from - from;
  const inside = state.selection.from >= from && state.selection.from < end;
  const slice = state.doc.slice(from, end);
  const tr = state.tr.delete(from, end);
  const target = tr.mapping.map(to, -1);
  tr.insert(target, slice.content);
  // The cursor goes along with the text it was in.
  const cursor = inside ? target + offset : target + 1;
  tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(cursor, tr.doc.content.size))));
  return tr.scrollIntoView();
}

/** Moves the section the cursor is in past the section above or below it (Alt+Shift+Up/Down). */
export function moveSection(ctx: DocContext, direction: -1 | 1, at?: number): Command {
  return (state, dispatch) => {
    const from = at ?? headingAround(state, ctx);
    if (from === null) return false;
    const doc = state.doc;
    const level = levelOf(ctx, doc.nodeAt(from)!);
    let to: number | null = null;
    if (direction < 0) {
      // Before the closest heading above as high or higher.
      doc.forEach((node, pos) => {
        const l = levelOf(ctx, node);
        if (pos < from && l && l <= level) to = pos;
      });
    } else {
      const end = sectionEnd(doc, ctx, from);
      const next = doc.nodeAt(end);
      if (!next || !levelOf(ctx, next)) return false;
      to = sectionEnd(doc, ctx, end);
    }
    if (to === null) return false;
    const tr = moveSectionTo(state, ctx, from, to);
    if (!tr) return false;
    dispatch?.(tr);
    return true;
  };
}

/** Takes the section's headings a level up (-1) or down (1) (Alt+Shift+Left/Right). */
export function shiftSection(ctx: DocContext, delta: -1 | 1, at?: number): Command {
  return (state, dispatch) => {
    const from = at ?? headingAround(state, ctx);
    if (from === null) return false;
    // The section keeps its shape: when its heading can't go further, nothing does.
    const top = levelOf(ctx, state.doc.nodeAt(from)!);
    if (top + delta < 1 || top + delta > 9) return false;
    const end = sectionEnd(state.doc, ctx, from);
    const changes: { pos: number; node: Node; level: number }[] = [];
    state.doc.nodesBetween(from, end, (node, pos) => {
      const l = levelOf(ctx, node);
      if (l && node.attrs.style !== "Title") changes.push({ pos, node, level: Math.max(1, Math.min(9, l + delta)) });
      return false;
    });
    if (!changes.length || changes.every((c) => c.level === levelOf(ctx, c.node))) return false;
    if (dispatch) {
      ensureStyles(ctx, ...new Set(changes.map((c) => `Heading${c.level}`)));
      const tr = state.tr;
      for (const c of changes) tr.setNodeMarkup(c.pos, undefined, { ...c.node.attrs, style: `Heading${c.level}` });
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}
