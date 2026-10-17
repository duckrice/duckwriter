// AutoFormat as you type, as Word does: a web or e-mail address becomes a link
// when it is finished (space or Enter), and "1. ", "- " or "* " at the start of
// a paragraph starts a numbered or bulleted list. Ctrl+Z takes it back.

import { Plugin, type EditorState, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { ensureStyles } from "../docx/builtins";
import type { DocContext } from "../docx/model";
import { listOf, toggleList } from "./lists";
import { withoutLink } from "./commands";
import { schema } from "./schema";
import { SYMBOLS } from "../math/ast";

const KEY = "duckwriter.autoformat";

export function autoFormatOn(): boolean {
  try {
    return localStorage.getItem(KEY) !== "off";
  } catch {
    return true;
  }
}

export function setAutoFormat(on: boolean) {
  try {
    localStorage.setItem(KEY, on ? "on" : "off");
  } catch {
    // For now only.
  }
}

const URL = /^(?:https?:\/\/|www\.)[^\s<>"]+$/i;
const EMAIL = /^[^\s@<>"()]+@[^\s@<>"()]+\.[a-z]{2,}$/i;

/** Makes the address just typed before the cursor a link. */
export function linkTypedAddress(state: EditorState, ctx: DocContext): Transaction | null {
  const { $from, empty } = state.selection;
  if (!empty || !$from.parent.isTextblock) return null;
  const before = $from.parent.textBetween(0, $from.parentOffset, undefined, "￼");
  const word = /\S+$/.exec(before)?.[0];
  if (!word) return null;
  // Punctuation after an address isn't part of it.
  const address = word.replace(/[.,;:!?)\]'"»]+$/, "");
  if (!URL.test(address) && !EMAIL.test(address)) return null;
  const start = $from.pos - word.length;
  const end = start + address.length;
  if (state.doc.rangeHasMark(start, end, schema.marks.link)) return null;
  const href = EMAIL.test(address) ? `mailto:${address}` : /^www\./i.test(address) ? `https://${address}` : address;
  ensureStyles(ctx, "Hyperlink");
  return state.tr.addMark(start, end, schema.marks.link.create({ href })).addMark(start, end, schema.marks.charStyle.create({ id: "Hyperlink" }));
}

/** Arrows and signs typed with ordinary keys. */
const SHORTHAND: [string, string][] = [
  ["<->", "↔"],
  ["<=>", "⇔"],
  ["->", "→"],
  ["<-", "←"],
  ["=>", "⇒"],
  ["<=", "≤"],
  [">=", "≥"],
  ["!=", "≠"],
  ["+-", "±"],
  ["~=", "≈"],
];

/**
 * Maths shorthand in the text, as typed is finished (a space or punctuation
 * after it): \alpha → α, \infty → ∞, -> → →, <= → ≤, +- → ±.
 */
export function mathShorthand(state: EditorState, typed: string): Transaction | null {
  const { $from, empty } = state.selection;
  if (!empty || !$from.parent.isTextblock || !/^[\s.,;:!?)\]]$/.test(typed)) return null;
  // Not inside an equation's own text or code-like runs; only plain text before the cursor.
  const before = $from.parent.textBetween(0, $from.parentOffset, undefined, "\uFFFC");
  const named = /\\([A-Za-z]+)$/.exec(before);
  if (named && SYMBOLS[named[1]]) return state.tr.insertText(SYMBOLS[named[1]], $from.pos - named[0].length, $from.pos);
  for (const [from, to] of SHORTHAND) {
    if (!before.endsWith(from)) continue;
    // Only between words or numbers, not as part of something longer ("-->" or "<!--").
    const prev = before[before.length - from.length - 1];
    if (prev && !/[\s\p{L}\p{N})]/u.test(prev)) return null;
    return state.tr.insertText(to, $from.pos - from.length, $from.pos);
  }
  return null;
}

/** "1." or "-" typed at the start of a paragraph, then a space: a list. */
function startList(view: EditorView, ctx: DocContext): boolean {
  const { state } = view;
  const { $from, empty } = state.selection;
  if (!empty || $from.parent.type !== schema.nodes.paragraph || listOf(ctx, $from.parent)) return false;
  if ($from.parent.childCount > 1 || $from.parentOffset !== $from.parent.content.size) return false;
  const typed = $from.parent.textContent;
  const kind = /^1[.)]$/.test(typed) ? "number" : /^[-*•]$/.test(typed) ? "bullet" : null;
  if (!kind) return false;
  view.dispatch(state.tr.delete($from.start(), $from.pos));
  return toggleList(ctx, kind)(view.state, view.dispatch);
}

export function autoFormatPlugin(ctx: () => DocContext) {
  return new Plugin({
    props: {
      handleTextInput(view, from, to, text) {
        if (autoFormatOn()) {
          const math = mathShorthand(view.state, text);
          if (math) {
            const at = math.mapping.map(from);
            view.dispatch(math.insertText(text, at, at + (to - from)));
            return true;
          }
        }
        if (autoFormatOn() && text === " ") {
          if (startList(view, ctx())) return true;
          const link = linkTypedAddress(view.state, ctx());
          if (link) {
            view.dispatch(link.setStoredMarks(withoutLink(link.doc.resolve(from).marks())).insertText(text, from, to));
            return true;
          }
        }
        // Typing just after a link isn't part of it (nor blue and underlined), as in Word.
        const marks = view.state.storedMarks ?? view.state.doc.resolve(from).marks();
        const plain = withoutLink(marks);
        if (plain.length === marks.length || view.state.storedMarks) return false;
        // Inside the link's text, typing stays part of it.
        const after = view.state.doc.resolve(to).nodeAfter;
        if (after && withoutLink(after.marks).length !== after.marks.length) return false;
        view.dispatch(view.state.tr.setStoredMarks(plain).insertText(text, from, to));
        return true;
      },
      handleKeyDown(view, event) {
        if (event.key !== "Enter" || !autoFormatOn() || event.ctrlKey || event.altKey) return false;
        const link = linkTypedAddress(view.state, ctx());
        if (link) view.dispatch(link);
        return false;
      },
    },
  });
}
