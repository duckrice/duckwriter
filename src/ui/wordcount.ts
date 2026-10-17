// Counting words, characters and paragraphs, of the document or the selection.
// Deleted text (a tracked change) isn't counted.

import type { Node } from "prosemirror-model";
import { h } from "../dom";
import { schema } from "../editor/schema";

export interface Counts {
  words: number;
  characters: number;
  charactersNoSpaces: number;
  paragraphs: number;
}

const WORD = /[\p{L}\p{N}][\p{L}\p{N}'’\-.]*/gu;

export function count(doc: Node, from = 0, to = doc.content.size): Counts {
  const out: Counts = { words: 0, characters: 0, charactersNoSpaces: 0, paragraphs: 0 };
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isTextblock) return true;
    let text = "";
    node.forEach((child, offset) => {
      // Deleted text is left out (a word it was in the middle of stays one word); a picture ends a word.
      if (child.isText && child.marks.some((m) => m.type === schema.marks.deletion)) return;
      if (!child.isText) {
        text += " ";
        return;
      }
      const start = pos + 1 + offset;
      text += child.text!.slice(Math.max(0, from - start), Math.max(0, to - start));
    });
    if (text.trim()) out.paragraphs++;
    out.words += (text.match(WORD) ?? []).length;
    out.characters += text.replace(/ +$/, "").length;
    out.charactersNoSpaces += text.replace(/\s/g, "").length;
    return false;
  });
  return out;
}

/** The counts, for the document and (if any) the selection. */
export function showCounts(doc: Node, selection: { from: number; to: number; empty: boolean }, pages: number) {
  const all = count(doc);
  const sel = selection.empty ? null : count(doc, selection.from, selection.to);
  const n = (v: number) => v.toLocaleString();
  const row = (label: string, a: number | string, b?: number) => h("tr", {}, h("th", {}, label), h("td", {}, typeof a === "number" ? n(a) : a), sel ? h("td", {}, b === undefined ? "" : n(b)) : null);
  const dialog = h(
    "dialog",
    { class: "dialog" },
    h("h2", {}, "Word count"),
    h(
      "table",
      { class: "count-table" },
      h("thead", {}, h("tr", {}, h("th", {}), h("th", {}, "Document"), sel ? h("th", {}, "Selection") : null)),
      h(
        "tbody",
        {},
        row("Pages", pages),
        row("Words", all.words, sel?.words),
        row("Characters (no spaces)", all.charactersNoSpaces, sel?.charactersNoSpaces),
        row("Characters (with spaces)", all.characters, sel?.characters),
        row("Paragraphs", all.paragraphs, sel?.paragraphs),
      ),
    ),
    h("p", { class: "dialog-hint" }, "Footnotes, headers and footers aren't counted, and deleted text (tracked changes) isn't either."),
    h("div", { class: "dialog-footer" }, h("button", { type: "button", class: "btn primary", onclick: () => dialog.close() }, "Close")),
  );
  dialog.addEventListener("close", () => dialog.remove());
  document.body.append(dialog);
  dialog.showModal();
  (dialog.querySelector(".btn.primary") as HTMLElement).focus();
}
