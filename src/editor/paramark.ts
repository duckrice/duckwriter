// The paragraph mark (¶) has formatting of its own, as in Word. It is how an
// empty line remembers the font it was given: the cursor in it picks the
// font up, Enter passes it on to the next line, and changing the font of a
// selection that reaches a paragraph's end changes its mark too.

import type { Mark, Node } from "prosemirror-model";
import { Plugin, Selection, type EditorState, type Transaction } from "prosemirror-state";
import { schema } from "./schema";

/** The kinds of formatting a paragraph mark keeps (not links, comments or tracked changes). */
const KEPT = new Set(["font", "size", "bold", "italic", "underline", "strike", "color", "highlight", "vertAlign", "charStyle", "runExtra"]);

export const keptOnMark = (marks: readonly Mark[]) => marks.filter((m) => KEPT.has(m.type.name));

/** A paragraph's mark formatting. */
export function paraMarks(attrs: { mark?: unknown }): readonly Mark[] {
  const json = attrs.mark as object[] | null | undefined;
  if (!json?.length) return [];
  try {
    return json.map((m) => schema.markFromJSON(m));
  } catch {
    return [];
  }
}

/** The attribute value for a paragraph mark with these marks. */
export const markAttr = (marks: readonly Mark[]) => {
  const kept = keptOnMark(marks);
  return kept.length ? kept.map((m) => m.toJSON()) : null;
};

/**
 * Changes the mark formatting of every paragraph in [from, to] whose end the
 * range reaches (an empty paragraph inside it, or the last one if it's
 * selected to its end), as Word does.
 */
export function updateParaMarks(tr: Transaction, from: number, to: number, change: (marks: readonly Mark[], para: Node) => readonly Mark[]) {
  // Changing a paragraph's settings would move the selection; nothing changes size, so it goes back as it was.
  const selection = tr.selection.toJSON();
  let changed = false;
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    const end = pos + node.nodeSize - 1;
    if (end < from || end > to) return false;
    const next = markAttr(change(paraMarks(node.attrs), node));
    if (JSON.stringify(next) !== JSON.stringify(node.attrs.mark)) {
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, mark: next });
      changed = true;
    }
    return false;
  });
  if (changed) tr.setSelection(Selection.fromJSON(tr.doc, selection));
}

/** With the cursor in an empty paragraph, its mark's formatting is what typing will use. */
export const paraMarkPlugin = new Plugin({
  appendTransaction(trs, _old, state: EditorState) {
    if (!trs.some((tr) => tr.selectionSet || tr.docChanged)) return null;
    const { empty, $from } = state.selection;
    if (!empty || state.storedMarks || $from.parent.type !== schema.nodes.paragraph || $from.parent.content.size) return null;
    const marks = paraMarks($from.parent.attrs);
    return marks.length ? state.tr.setStoredMarks(marks) : null;
  },
});
