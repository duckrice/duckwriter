// Keeps fields up to date as the text changes: figure and table numbers
// (SEQ fields) count in order, and references to them (REF fields) show the
// current number. Word does this when asked to update fields; DuckWriter does
// it all the time, so what is saved is always current.

import type { Node } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import { fieldKind, schema } from "./schema";

const key = new PluginKey("fields");

/** The parts of a field instruction: SEQ Figure \* ARABIC → ["SEQ", "Figure", "\*", "ARABIC"]. */
export function fieldWords(instr: string): string[] {
  return [...instr.matchAll(/"([^"]*)"|(\S+)/g)].map((m) => m[1] ?? m[2]);
}

/** A bookmark's name and id, from the XML of a hidden bookmarkStart/End node. */
export function bookmarkOf(node: Node): { kind: "start" | "end"; id: string; name?: string } | null {
  if (node.type !== schema.nodes.raw_inline) return null;
  const xml: string = node.attrs.xml;
  const start = /^<w:bookmarkStart\b[^>]*\bw:id="([^"]*)"[^>]*?(?:\bw:name="([^"]*)")?/.exec(xml);
  if (start) {
    const name = /\bw:name="([^"]*)"/.exec(xml)?.[1];
    return { kind: "start", id: start[1], name };
  }
  const end = /^<w:bookmarkEnd\b[^>]*\bw:id="([^"]*)"/.exec(xml);
  return end ? { kind: "end", id: end[1] } : null;
}

/** The results all SEQ and REF fields should show, by position. */
export function fieldResults(doc: Node): Map<number, string> {
  const results = new Map<number, string>();
  const counts = new Map<string, number>();
  // First the numbers.
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.field || fieldKind(node.attrs.instr) !== "SEQ") return;
    const words = fieldWords(node.attrs.instr);
    const label = words[1] ?? "";
    let n = counts.get(label) ?? 0;
    const reset = words.indexOf("\\r");
    if (reset >= 0 && words[reset + 1]) n = Number(words[reset + 1]) - 1;
    if (!words.includes("\\c")) n++;
    counts.set(label, n);
    results.set(pos, words.includes("\\h") ? "" : String(n));
  });
  // Then what each bookmark holds, with those numbers in it.
  const marked = new Map<string, string>();
  doc.descendants((para, paraPos) => {
    if (para.type !== schema.nodes.paragraph) return true;
    const open = new Map<string, { name: string; text: string }>();
    para.forEach((child, offset) => {
      const mark = bookmarkOf(child);
      if (mark?.kind === "start" && mark.name) open.set(mark.id, { name: mark.name, text: "" });
      else if (mark?.kind === "end") {
        const b = open.get(mark.id);
        if (b) marked.set(b.name, b.text.trim());
        open.delete(mark.id);
      } else {
        let text = child.text ?? "";
        if (child.type === schema.nodes.field) text = results.get(paraPos + 1 + offset) ?? child.attrs.result;
        for (const b of open.values()) b.text += text;
      }
    });
    // A bookmark that runs to the end of the paragraph (a heading).
    for (const b of open.values()) marked.set(b.name, b.text.trim());
    return false;
  });
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.field || fieldKind(node.attrs.instr) !== "REF") return;
    const name = fieldWords(node.attrs.instr)[1];
    if (name && marked.has(name)) results.set(pos, marked.get(name)!);
  });
  return results;
}

export function fieldsPlugin() {
  return new Plugin({
    key,
    appendTransaction(trs, _old, state) {
      if (!trs.some((tr) => tr.docChanged)) return null;
      const tr = state.tr;
      for (const [pos, result] of fieldResults(state.doc)) {
        const node = state.doc.nodeAt(pos);
        if (node && node.attrs.result !== result) tr.setNodeMarkup(pos, undefined, { ...node.attrs, result });
      }
      return tr.docChanged ? tr.setMeta("addToHistory", false) : null;
    },
  });
}
