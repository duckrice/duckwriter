// Read aloud (Ctrl+Alt+Space, as in Word): the selection, or the text from the
// cursor on, in one of Windows' voices; hearing a sentence catches what the
// eye skips. The paragraph being read is marked; the same key stops it.

import type { Node } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { schema } from "../editor/schema";

interface Piece {
  /** The top-level paragraph's position, to mark it while it is read. */
  pos: number;
  text: string;
}

/** What to read: the selected text, or every paragraph from the cursor's on; deleted text left out. */
export function piecesToRead(doc: Node, from: number, to: number): Piece[] {
  const pieces: Piece[] = [];
  const whole = from === to;
  const end = whole ? doc.content.size : to;
  doc.forEach((block, pos) => {
    if (pos + block.nodeSize <= from || pos >= end) return;
    let text = "";
    block.descendants((node, at) => {
      const abs = pos + 1 + at;
      if (node.isText && !node.marks.some((m) => m.type === schema.marks.deletion)) {
        const start = whole ? 0 : Math.max(0, from - abs);
        const stop = whole ? node.text!.length : Math.min(node.text!.length, end - abs);
        if (stop > start) text += node.text!.slice(start, stop);
      } else if (node.type === schema.nodes.math) text += ` ${node.attrs.latex} `;
      else if (node.type === schema.nodes.field) text += node.attrs.result;
      else if (node.isBlock && text && !/\s$/.test(text)) text += ". ";
      return true;
    });
    if (text.trim()) pieces.push({ pos, text: text.replace(/\s+/g, " ").trim() });
  });
  return pieces;
}

let reading: { stop: () => void } | null = null;

export const isReading = () => reading !== null;

/** Starts reading, or stops if it already is; `onState` hears when it starts and stops. */
export function toggleReadAloud(view: EditorView, onState?: (on: boolean) => void): boolean {
  if (reading) {
    reading.stop();
    return false;
  }
  const speech = window.speechSynthesis;
  if (!speech) return false;
  const { from, to } = view.state.selection;
  const pieces = piecesToRead(view.state.doc, from, to);
  if (!pieces.length) return false;
  // Windows' voice for the document's language, if there is one; otherwise the default.
  const lang = (view.dom.closest("[lang]") as HTMLElement | null)?.lang || navigator.language;
  const voice = speech.getVoices().find((v) => v.lang.toLowerCase().startsWith(lang.slice(0, 2).toLowerCase()) && v.localService);
  let marked: Element | null = null;
  const mark = (pos: number | null) => {
    marked?.classList.remove("reading-aloud");
    marked = pos === null ? null : (view.nodeDOM(pos) as Element | null);
    marked?.classList.add("reading-aloud");
    marked?.scrollIntoView({ block: "nearest", behavior: "smooth" });
  };
  let stopped = false;
  const finish = () => {
    if (!reading) return;
    reading = null;
    mark(null);
    onState?.(false);
  };
  reading = {
    stop: () => {
      stopped = true;
      speech.cancel();
      finish();
    },
  };
  pieces.forEach((piece, i) => {
    const u = new SpeechSynthesisUtterance(piece.text);
    if (voice) u.voice = voice;
    u.onstart = () => !stopped && mark(piece.pos);
    if (i === pieces.length - 1) {
      u.onend = finish;
      u.onerror = finish;
    }
    speech.speak(u);
  });
  onState?.(true);
  return true;
}
