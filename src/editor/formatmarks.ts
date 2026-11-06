// Show formatting (¶), as Word's button: a ¶ at the end of each paragraph, a dot for each
// space, an arrow for each tab, ↵ for each line break, and where each section ends and how
// the next one starts. The marks are drawn over the text and take no room, so the
// pages break exactly as they do with them hidden. Ctrl+Shift+8.

import { Plugin, PluginKey, type EditorState } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import type { Node } from "prosemirror-model";
import { schema } from "./schema";

const key = new PluginKey<boolean>("formatMarks");
const STORE = "duckwriter.formatMarks";

const BREAK_NAMES: Record<string, string> = { continuous: "continuous", oddPage: "odd page", evenPage: "even page" };

function remembered(): boolean {
  try {
    return localStorage.getItem(STORE) === "on";
  } catch {
    return false;
  }
}

export const formatMarksOn = (state: EditorState) => key.getState(state) ?? false;

/** Shows or hides the marks in this editor, and remembers the choice. */
export function toggleFormatMarks(view: EditorView) {
  const on = !formatMarksOn(view.state);
  try {
    localStorage.setItem(STORE, on ? "on" : "off");
  } catch {
    // Just for now, then.
  }
  view.dispatch(view.state.tr.setMeta(key, on));
}

function marks(doc: Node): DecorationSet {
  const decorations: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      const end = pos + node.nodeSize - 1;
      const sectPr = /<w:sectPr\b[\s\S]*?(?:<\/w:sectPr>|\/>)/.exec(node.attrs.extra ?? "")?.[0];
      const label = sectPr ? `Section break (${BREAK_NAMES[/<w:type w:val="(\w+)"/.exec(sectPr)?.[1] ?? ""] ?? "next page"})` : null;
      decorations.push(Decoration.widget(end, () => fmMark(label), { side: 1, key: label ? `s${label}` : "p", ignoreSelection: true }));
      return true;
    }
    if (node.isText) {
      const text = node.text ?? "";
      for (let i = 0; i < text.length; i++) {
        if (text[i] === " " || text[i] === " ") decorations.push(Decoration.inline(pos + i, pos + i + 1, { class: "fm-space" }));
        else if (text[i] === "\t") decorations.push(Decoration.inline(pos + i, pos + i + 1, { class: "fm-tab" }));
      }
      return false;
    }
    if (node.type === schema.nodes.hard_break) decorations.push(Decoration.widget(pos, () => fmMark(null, "↵"), { side: -1, key: "br" }));
    return true;
  });
  return DecorationSet.create(doc, decorations);
}

function fmMark(sectionBreak: string | null, sign = "¶"): HTMLElement {
  const span = document.createElement("span");
  span.className = sectionBreak ? "fm-mark fm-section" : "fm-mark";
  span.contentEditable = "false";
  span.textContent = sign;
  if (sectionBreak) span.dataset.label = sectionBreak;
  return span;
}

export const formatMarksPlugin = new Plugin<boolean>({
  key,
  state: {
    init: () => remembered(),
    apply: (tr, on) => tr.getMeta(key) ?? on,
  },
  props: {
    decorations(state) {
      if (!key.getState(state)) return null;
      // Worked out again only when the text changes.
      if (cache?.doc !== state.doc) cache = { doc: state.doc, set: marks(state.doc) };
      return cache.set;
    },
    attributes: (state): Record<string, string> => (key.getState(state) ? { class: "show-marks" } : {}),
  },
});

let cache: { doc: Node; set: DecorationSet } | null = null;
