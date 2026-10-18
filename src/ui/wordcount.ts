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

export interface Readability {
  sentences: number;
  wordsPerSentence: number;
  /** Words of more than six letters, as a share of all words (0–1). */
  longWords: number;
  /** Björnsson's readability index: words per sentence plus the percentage of long words. */
  lix: number;
}

/** How hard the text is to read, from its sentences and long words; headings and captions left out. */
export function readability(doc: Node): Readability {
  let sentences = 0;
  let words = 0;
  let long = 0;
  doc.descendants((node) => {
    if (!node.isTextblock) return true;
    if (/^(Heading|Title|Subtitle|Caption|TOC)/.test(node.attrs.style ?? "")) return false;
    let text = "";
    node.forEach((child) => {
      if (child.isText && !child.marks.some((m) => m.type === schema.marks.deletion)) text += child.text;
      else if (!child.isText) text += " ";
    });
    const found = text.match(WORD) ?? [];
    if (!found.length) return false;
    words += found.length;
    long += found.filter((w) => w.replace(/[^\p{L}]/gu, "").length > 6).length;
    // Full stops, question and exclamation marks and colons end sentences; so does the paragraph's end.
    const ends = (text.match(/[.!?:]+(?=\s|$)/g) ?? []).length;
    sentences += Math.max(1, ends + (/[.!?:]["'”’)]*\s*$/.test(text) ? 0 : 1));
    return false;
  });
  const wordsPerSentence = sentences ? words / sentences : 0;
  const longWords = words ? long / words : 0;
  return { sentences, wordsPerSentence, longWords, lix: Math.round(wordsPerSentence + longWords * 100) };
}

/** What a LIX value means, roughly (Björnsson's scale). */
export function lixLevel(lix: number): string {
  if (lix < 25) return "very easy";
  if (lix < 35) return "easy";
  if (lix < 45) return "medium, like a newspaper";
  if (lix < 55) return "hard, like academic text";
  return "very hard, like technical writing";
}

export interface Progress {
  /** The words the document should reach, if set. */
  goal: number | null;
  /** Words written today. */
  today: number;
}

/** The counts, for the document and (if any) the selection, and the goal to set. */
export function showCounts(doc: Node, selection: { from: number; to: number; empty: boolean }, pages: number, progress?: Progress, setGoal?: (goal: number | null) => void) {
  const all = count(doc);
  const read = readability(doc);
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
        row("Sentences", read.sentences),
        row("Words per sentence", read.wordsPerSentence.toLocaleString(undefined, { maximumFractionDigits: 1 })),
        row("Readability (LIX)", read.sentences ? `${read.lix}: ${lixLevel(read.lix)}` : "–"),
      ),
    ),
    h("p", { class: "dialog-hint" }, "Footnotes, headers and footers aren't counted, and deleted text (tracked changes) isn't either."),
    ...(progress && setGoal ? [goalForm(all.words, progress, setGoal)] : []),
    h("div", { class: "dialog-footer" }, h("button", { type: "button", class: "btn primary", onclick: () => dialog.close() }, "Close")),
  );
  dialog.addEventListener("close", () => {
    const box = dialog.querySelector<HTMLInputElement>(".goal-input");
    if (box && setGoal) setGoal(Number(box.value) > 0 ? Number(box.value) : null);
    dialog.remove();
  });
  document.body.append(dialog);
  dialog.showModal();
  (dialog.querySelector(".btn.primary") as HTMLElement).focus();
}

/** "Goal: [5000] words", with how far along it is and today's words. */
function goalForm(words: number, progress: Progress, setGoal: (goal: number | null) => void): HTMLElement {
  const box = h("input", { class: "tb-input dialog-input number goal-input", type: "number", min: "0", step: "100", placeholder: "none", "aria-label": "Word goal" }) as HTMLInputElement;
  if (progress.goal) box.value = String(progress.goal);
  const note = h("span", { class: "goal-note" });
  const describe = () => {
    const goal = Number(box.value);
    const parts = [goal > 0 ? `${Math.min(100, Math.round((words / goal) * 100))}% there` : "", progress.today > 0 ? `${progress.today.toLocaleString()} written today` : ""].filter(Boolean);
    note.textContent = parts.join(" · ");
  };
  box.addEventListener("input", describe);
  box.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      setGoal(Number(box.value) > 0 ? Number(box.value) : null);
      box.closest("dialog")?.close();
    }
  });
  describe();
  return h("label", { class: "goal-form" }, "Goal for this document: ", box, " words ", note);
}
