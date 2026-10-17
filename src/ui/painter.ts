// The format painter: copy the formatting of some text (Ctrl+Shift+C, or the
// brush) and put it on other text (Ctrl+Shift+V, or select it after the brush).
// Also pasting as plain text (Ctrl+Alt+V).

import type { Mark } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { wordAround } from "../editor/review";
import { schema } from "../editor/schema";
import { FORMATTING } from "../editor/track";

interface Copied {
  marks: Mark[];
  /** The paragraph style, when a whole paragraph was copied from. */
  style: string | null;
}

let copied: Copied | null = null;
let armed: { view: EditorView; done: () => void } | null = null;

const isFormatting = (m: Mark) => FORMATTING.has(m.type.name) || m.type === schema.marks.runExtra;

/** Remembers the formatting at the start of the selection. */
export function copyFormatting(view: EditorView): boolean {
  const { $from, from, to, empty } = view.state.selection;
  const marks = (empty ? (view.state.storedMarks ?? $from.marks()) : view.state.doc.resolve(Math.min(from + 1, to)).marks()).filter(isFormatting);
  const para = $from.parent;
  const whole = !empty && $from.parentOffset === 0 && view.state.selection.$to.parentOffset === para.content.size && para.content.size > 0;
  copied = { marks: [...marks], style: whole ? para.attrs.style : null };
  return true;
}

/** Puts the copied formatting on the selection (or the word at the cursor). */
export function pasteFormatting(view: EditorView): boolean {
  if (!copied) return false;
  const state = view.state;
  const { from, to } = wordAround(state);
  const tr = state.tr;
  if (to > from) {
    for (const name of [...FORMATTING, "runExtra"]) tr.removeMark(from, to, schema.marks[name as keyof typeof schema.marks]);
    for (const mark of copied.marks) tr.addMark(from, to, mark);
  } else tr.setStoredMarks(copied.marks);
  if (copied.style) {
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.type !== schema.nodes.paragraph) return true;
      tr.setNodeMarkup(pos, null, { ...node.attrs, style: copied!.style });
      return false;
    });
  }
  view.dispatch(tr.scrollIntoView());
  return true;
}

/** The brush: copies now, and paints the next text selected with the mouse. */
export function armPainter(view: EditorView, onDone: () => void) {
  if (armed) {
    disarmPainter();
    return;
  }
  copyFormatting(view);
  const dom = view.dom as HTMLElement;
  dom.classList.add("painting");
  const up = () => {
    setTimeout(() => {
      if (!armed || view.state.selection.empty) return;
      pasteFormatting(view);
      disarmPainter();
    }, 0);
  };
  const key = (event: KeyboardEvent) => {
    if (event.key === "Escape") disarmPainter();
  };
  dom.addEventListener("mouseup", up);
  document.addEventListener("keydown", key, true);
  armed = {
    view,
    done: () => {
      dom.classList.remove("painting");
      dom.removeEventListener("mouseup", up);
      document.removeEventListener("keydown", key, true);
      onDone();
    },
  };
}

export function disarmPainter() {
  const a = armed;
  armed = null;
  a?.done();
}

export const painterArmed = () => !!armed;

/** Pastes the clipboard's text without its formatting. */
export async function pastePlain(view: EditorView): Promise<boolean> {
  let text = "";
  try {
    text = await navigator.clipboard.readText();
  } catch {
    return false;
  }
  if (!text) return true;
  view.focus();
  view.pasteText(text);
  return true;
}
